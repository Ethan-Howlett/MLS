// ============================================================================
// CONFIG - Everything specific to the workbook layout lives here.
// If a sheet, cell, table, or column is renamed or moved, update it HERE ONCE.
// You should not need to touch the code below this block.
// ============================================================================
const CONFIG = {
  sheets: {
    main: "MP_DistrictMatch", // Sheet that holds the day selector and the tables
    log: "Logs",              // Sheet where a run history is recorded (auto-created)
  },
  cells: {
    daySelector: "AE1", // Dropdown the user sets to "Thursday" or "Friday"
    status: "AG1",      // Cell where progress / errors are shown to the user
  },
  tables: {
    assignedDistricts: "TYM_Numbers", // Table listing districts assigned to each mission leader
    thursdayMissionaries: "TYM_Missionaries",  // Missionary table for Thursday
    fridayMissionaries: "TYM_Missionaries2",   // Missionary table for Friday
  },
  // Column HEADER names. These are looked up by name, so reordering columns is safe
  // as long as the header text stays the same.
  columns: {
    thursdayAssignedDistrict: "District",  // Column in assignedDistricts table (Thursday)
    fridayAssignedDistrict: "District2",   // Column in assignedDistricts table (Friday)
    slicer: "Slicer",                      // Column written back to the missionary table
  },
  // Column POSITIONS (0-based) inside the missionary tables, used to read each
  // missionary's details. If you reorder columns in TYM_Missionaries /
  // TYM_Missionaries2, update these numbers.
  missionaryColumnIndexes: {
    district: 0,
    type: 1,
    name: 2,
    companionshipId: 3,
    slicer: 9,
  },
  maxLogRows: 5000, // The Logs sheet resets itself once it passes this many rows
}

type ExcelInput = (string | number | boolean)[][]
interface Missionary {
  district: string;
  type: string;
  name: string;
  compId: string | number;
  destination: number;
  originalIndex: number; // Keeps track of the row for writing back to Excel
}
// A district's split directive: how many mission leaders it was assigned to.
type DistrictMap = Record<string, number>
interface SelectedDay { thur: boolean; fri: boolean }

function setDay(workbook: ExcelScript.Workbook, logger: Logger): SelectedDay {
  const sheet = workbook.getWorksheet(CONFIG.sheets.main)
  const day = String(sheet.getRange(CONFIG.cells.daySelector).getValue()).trim()
  const selectedDay: SelectedDay = { thur: false, fri: false }

  if (day === "Thursday") {
    selectedDay.thur = true
  } else if (day === "Friday") {
    selectedDay.fri = true
  } else {
    logger.error('setDay',
      `No valid day selected. Please choose "Thursday" or "Friday" in cell ${CONFIG.cells.daySelector}. (Found: "${day}")`)
  }

  logger.info('setDay', `Selected day: ${day}`)
  return selectedDay
}

function getRequiredColumn(table: ExcelScript.Table, name: string, logger: Logger): ExcelScript.TableColumn {
  const col = table.getColumn(name)
  if (!col) {
    logger.error('getRequiredColumn', `Could not find a column named "${name}" in table "${table.getName()}".`)
  }
  return col
}

function getValuesFromTables(workbook: ExcelScript.Workbook, selectedDay: SelectedDay, logger: Logger)
  : { missionaryTable: ExcelScript.Table, assignedDistrictRange: string[], missionaryRange: ExcelInput } {
  const sheet = workbook.getWorksheet(CONFIG.sheets.main)

  const assignedTable = sheet.getTable(CONFIG.tables.assignedDistricts)
  if (!assignedTable) {
    logger.error('getValuesFromTables', `Could not find the table "${CONFIG.tables.assignedDistricts}".`)
  }

  let assignedDistrictRange: string[]
  let missionaryTable: ExcelScript.Table

  if (selectedDay.thur) {
    assignedDistrictRange = flattenInput(
      getRequiredColumn(assignedTable, CONFIG.columns.thursdayAssignedDistrict, logger)
        .getRangeBetweenHeaderAndTotal().getValues())
    missionaryTable = sheet.getTable(CONFIG.tables.thursdayMissionaries)
  } else if (selectedDay.fri) {
    assignedDistrictRange = flattenInput(
      getRequiredColumn(assignedTable, CONFIG.columns.fridayAssignedDistrict, logger)
        .getRangeBetweenHeaderAndTotal().getValues())
    missionaryTable = sheet.getTable(CONFIG.tables.fridayMissionaries)
  } else {
    // Should be unreachable because setDay throws first, but guard anyway so the
    // returned values can never silently be undefined.
    logger.error('getValuesFromTables', 'No day was selected, so no tables could be read.')
  }

  if (!missionaryTable) {
    logger.error('getValuesFromTables', 'Could not find the missionary table for the selected day.')
  }

  const missionaryRange = missionaryTable.getRangeBetweenHeaderAndTotal().getValues()
  return { missionaryTable, assignedDistrictRange, missionaryRange }
}

function flattenInput(arr: ExcelInput): string[] {
  const rowCount = arr.length
  const colCount = arr[0]?.length || 0

  const flatArr = new Array<string>(rowCount * colCount)

  let index = 0
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      flatArr[index++] = String(arr[r][c])
    }
  }
  return flatArr
}

function buildDistrictMap(arr: string[]): DistrictMap {
  const map: DistrictMap = {}

  for (const raw of arr) {
    // Ignore blank assignment slots so an empty cell is never treated as a district.
    const str = raw.trim()
    if (!str) continue
    map[str] = (map[str] || 0) + 1
  }

  return map
}

function assignDestinations(districtMap: DistrictMap, missionaries: Missionary[], logger: Logger): void {
  // 1. Group all missionaries by District
  const districts: Record<string, Missionary[]> = {}
  for (const missionary of missionaries) {
    if (!districts[missionary.district]) districts[missionary.district] = []
    districts[missionary.district].push(missionary)
  }

  // 2. Process each district
  for (const districtName in districts) {
    const members = districts[districtName]
    const splitDirective = districtMap[districtName]

    // Scenario: District exists in the missionary list but was never assigned to a
    // mission leader. Leave it blank and warn instead of silently splitting it.
    if (splitDirective === undefined) {
      logger.warn('assignDestinations',
        `District "${districtName}" has ${members.length} missionaries but was not assigned to any mission leader. Leaving its slicer blank.`)
      for (const member of members) member.destination = 0
      continue
    }

    // Scenario A: District is fully assigned to a single mission leader
    if (splitDirective <= 1) {
      for (const member of members) {
        member.destination = 1
      }
      continue
    }

    // Scenario B: District needs to be split N ways (2, 3, 4, ...)
    // Group by comp ID first
    const companionships: Record<string, Missionary[]> = {}
    for (const member of members) {
      const compId = String(member.compId)
      if (!companionships[compId]) companionships[compId] = []
      companionships[compId].push(member)
    }
    // Separate companionships by Type to balance them individually
    const elders: Missionary[][] = []
    const sisters: Missionary[][] = []

    for (const compId in companionships) {
      const comp = companionships[compId]
      // Assume companions share the same type; check the first person
      if (comp[0].type.toLowerCase() === 'elder') {
        elders.push(comp)
      } else {
        sisters.push(comp)
      }
    }
    // Determine the number of buckets for this specific district.
    const totalCompanionships = Object.keys(companionships).length;

    // True N-way split: honor however many mission leaders the district was assigned
    // to, but never create more buckets than there are companionships to fill them.
    const numBuckets = Math.max(1, Math.min(splitDirective, totalCompanionships));

    if (splitDirective > totalCompanionships) {
      logger.warn('assignDestinations',
        `District "${districtName}" was assigned to ${splitDirective} mission leaders but only has ${totalCompanionships} companionship(s). Splitting ${numBuckets} way(s).`)
    }

    // Array to track the current population of each bucket
    // Index 0 = Destination 1, Index 1 = Destination 2, Index 2 = Destination 3, ...
    const bucketCounts: number[] = new Array(numBuckets).fill(0);

    // Greedy N-way distribution function
    const distribute = (groups: Missionary[][]) => {
      // Sort groups by size descending (packs larger groups first for better balance)
      groups.sort((a, b) => b.length - a.length)

      for (const group of groups) {
        // 1. Find the bucket with the lowest current headcount
        let minIndex = 0
        let minCount = bucketCounts[0]

        for (let i = 1; i < numBuckets; i++) {
          if (bucketCounts[i] < minCount) {
            minCount = bucketCounts[i]
            minIndex = i
          }
        }

        // 2. Assign the groups to that bucket (minIndex 0 becomes destination 1)
        const targetDestination = minIndex + 1
        group.forEach(m => m.destination = targetDestination)

        // 3. Update the tracked count for the chosen bucket
        bucketCounts[minIndex] += group.length
      }
    };

    // Distributing demographics individually ensures Elders and Sisters
    // are spread as evenly as possible across the buckets.
    distribute(elders)
    distribute(sisters)
  }
}

function processMissionaryData(missionaryRange: ExcelInput, map: DistrictMap, logger: Logger) {
  const idx = CONFIG.missionaryColumnIndexes
  const missionaries: Missionary[] = []

  // 1. Map 2D array to Objects, ignoring rows with blank districts
  for (let i = 0; i < missionaryRange.length; i++) {
    // Extract and trim to catch strings with only spaces
    const district = String(missionaryRange[i][idx.district] || '').trim()
    // Skip this row if the district is empty
    if (!district) continue

    missionaries.push({
      district: district,
      type: String(missionaryRange[i][idx.type]),
      name: String(missionaryRange[i][idx.name]),
      compId: missionaryRange[i][idx.companionshipId] as string | number,
      destination: 0,
      originalIndex: i
    })
  }

  // 2. Run the partitioning algorithm
  assignDestinations(map, missionaries, logger)

  // 3. Write destinations into an array for Excel output.
  // Start from the existing slicer values so untouched rows are preserved.
  const slicerOutput: ExcelInput = missionaryRange.map(row => [row[idx.slicer]])

  for (const m of missionaries) {
    // Unassigned districts (destination 0) get a blank slicer rather than a 0.
    slicerOutput[m.originalIndex][0] = m.destination > 0 ? m.destination : ""
  }

  logger.info('processMissionaryData',
    `Processed ${missionaries.length} missionaries across ${Object.keys(map).length} assigned district(s).`)

  return slicerOutput
}

function writeOutput(missionaryTable: ExcelScript.Table, data: ExcelInput, logger: Logger): void {
  getRequiredColumn(missionaryTable, CONFIG.columns.slicer, logger)
    .getRangeBetweenHeaderAndTotal()
    .setValues(data)
}

function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getWorksheet(CONFIG.sheets.main)
  const statusCell = sheet ? sheet.getRange(CONFIG.cells.status) : undefined
  const logger = new Logger(workbook, CONFIG.sheets.log, statusCell, CONFIG.maxLogRows)

  try {
    if (!sheet) {
      logger.error('main', `Required sheet "${CONFIG.sheets.main}" was not found. Please check the sheet name.`)
    }

    logger.setStatus("Running...")
    logger.info('main', 'District splitter started.')

    const selectedDay = setDay(workbook, logger)
    const { missionaryTable, assignedDistrictRange, missionaryRange } = getValuesFromTables(workbook, selectedDay, logger)

    const districtMap = buildDistrictMap(assignedDistrictRange)
    const destinationOutput = processMissionaryData(missionaryRange, districtMap, logger)

    writeOutput(missionaryTable, destinationOutput, logger)
    workbook.refreshAllPivotTables()

    logger.info('main', 'District splitter finished successfully.')
    logger.setStatus("Done.")
  } catch (e) {
    // Surface the failure in the spreadsheet so a non-coder can see what happened.
    logger.setStatus(`Error: ${e.message}`)
    console.log(`District splitter failed: ${e.message}`)
  } finally {
    logger.flush()
  }
}

// ============================================================================
// Logger - shared, identical across all three MLS scripts.
// Buffers messages and writes them to the Logs sheet in one operation, shows
// progress in a status cell, and resets the Logs sheet when it gets too large.
// ============================================================================
type LogLevel = 'INFO' | 'WARN' | 'ERROR'

class Logger {
  private workbook: ExcelScript.Workbook
  private logSheetName: string
  private statusCell: ExcelScript.Range | undefined
  private maxLogRows: number
  private buffer: string[][]
  private timers: Record<string, number>
  private startTime: number

  constructor(
    workbook: ExcelScript.Workbook,
    logSheetName: string = "Logs",
    statusCell?: ExcelScript.Range,
    maxLogRows: number = 5000
  ) {
    this.workbook = workbook
    this.logSheetName = logSheetName
    this.statusCell = statusCell
    this.maxLogRows = maxLogRows
    this.buffer = []
    this.timers = {}
    this.startTime = Date.now()
  }

  private timestamp(): string {
    return new Date().toISOString()
  }

  private add(level: LogLevel, fn: string, msg: string) {
    console.log(`[${level}][${fn}] ${msg}`)
    this.buffer.push([this.timestamp(), level, fn, msg])
  }

  info(fn: string, msg: string) {
    this.add('INFO', fn, msg)
  }

  warn(fn: string, msg: string) {
    this.add('WARN', fn, msg)
  }

  error(fn: string, msg: string): never {
    // Records an error, shows it in the status cell, writes the log, then stops the run.
    this.add('ERROR', fn, msg)
    this.setStatus(`Error: ${msg}`)
    this.flush()
    throw new Error(msg)
  }

  setStatus(msg: string) {
    // Updates the user-facing status cell immediately so progress is visible live.
    if (this.statusCell) {
      this.statusCell.setValue(msg)
    }
  }

  startTimer(label: string) {
    this.timers[label] = Date.now()
  }

  endTimer(label: string): number {
    const elapsed = Date.now() - (this.timers[label] ?? Date.now())
    this.info('Timer', `${label} took ${elapsed}ms`)
    return elapsed
  }

  summary() {
    const totalTime = Date.now() - this.startTime
    this.info('Summary', `Completed in ${totalTime}ms total`)
  }

  flush() {
    // Writes all buffered rows to the Logs sheet in a single operation.
    if (this.buffer.length === 0) return
    const sheet = this.initLogSheet()

    // Reset the sheet if it is getting huge so it never grows without bound.
    const used = sheet.getUsedRange()
    let nextRow = used ? used.getRowCount() : 1
    if (nextRow + this.buffer.length > this.maxLogRows) {
      sheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.contents)
      sheet.getRange("A1:D1").setValues([["Timestamp", "Level", "Function", "Message"]])
      nextRow = 1
    }

    const startRow = nextRow + 1
    sheet.getRange(`A${startRow}:D${startRow + this.buffer.length - 1}`).setValues(this.buffer)
    this.buffer = []
  }

  private initLogSheet(): ExcelScript.Worksheet {
    let sheet = this.workbook.getWorksheet(this.logSheetName)
    if (!sheet) {
      sheet = this.workbook.addWorksheet(this.logSheetName)
      sheet.getRange("A1:D1").setValues([["Timestamp", "Level", "Function", "Message"]])
    }
    return sheet
  }
}
