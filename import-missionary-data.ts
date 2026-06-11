// ============================================================================
// CONFIG - Everything specific to the workbook layout and the MTC Tools API
// lives here. If a sheet, cell, table, URL, or limit changes, update it HERE.
// ============================================================================
const CONFIG = {
  sheets: {
    control: "ControlPanel", // Sheet with the API key, dates, and status cell
    rawData: "Raw Data",     // Sheet whose table receives the imported records
    log: "Logs",             // Sheet where a run history is recorded (auto-created)
  },
  cells: {
    apiToken: "C8",       // Paste the MTC Tools "Bearer ..." key here
    onSiteStart: "D8",    // On-site window start date
    onSiteEnd: "D9",      // On-site window end date
    departureStart: "E8", // Departure window start date
    departureEnd: "E9",   // Departure window end date
    status: "G8",         // Cell where progress / errors are shown to the user
  },
  table: "Raw_Mission",   // Table on the Raw Data sheet
  apiUrl: "https://api.tools.mtc.byu.edu/search/missionaries",
  pageSize: 400,          // Records requested per page (payload.take)
  // Validation floor: a successful run should return at least this many records.
  // If a year's group is genuinely small, lower this number.
  minExpectedRecords: 100,
  maxPages: 100,          // Safety cap so a bad API response can never loop forever
  maxLogRows: 5000,       // The Logs sheet resets itself once it passes this many rows
  // The MTC(s) to search. The API uses this list to scope the search. Today it
  // holds only "Provo"; add more MTC objects here to search additional MTCs.
  // The leader name/email/phone fields are informational and WILL go stale over
  // the years (MTC presidents change). The "id"/"unitId"/"typeId" are the values
  // that actually matter to the API.
  mtcs: [
    {
      id: 2010852,
      name: "Provo",
    },
  ] as MTC[],
}

interface ResponseJSON {
  data: [{}];
  pagination: Pagination;
  message: string;
  statusCode: number;
}
interface Pagination {
  hasMore: boolean;
  take: number;
  skip: number;
  total: number;
}
interface Status {
  id: number;
  name: string;
}
interface Types {
  name: string;
}
interface MTC {
  id: number;
  name: string;
  typeId: number;
  unitId: number;
  leaderName: string;
  leaderEmail: string;
  leaderPhone: string | null;
  leaderMobile: string | null;
  homeAddress: string | null;
  officePhone: string;
  officeEmail: string;
  officeAddress: string;
}
interface Payload {
  statuses: Status[];
  types: Types[];
  mtcs: MTC[];
  requestedFields: string[];
  scheduledDepartureEnd: string;
  scheduledDepartureStart: string;
  scheduledOnSiteEnd: string;
  scheduledOnSiteStart: string;
  skip: number;
  take: number;
}
interface RestrictedData {
  branchDistrict: string;
  scheduleName: string;
  missionaryType: string;
  missionaryId: string;
  fullName: string;
  trainingLanguage: string;
  companionshipId: string;
  scheduledOnSiteStartDate: string;
  scheduledDepartureDate: string;
  missionEmail: string;
  missionName: string;
  classroom: string;
  mtc: string;
}

/**
 * Makes the fetch request to MTC Tools and surfaces clear, user-friendly errors.
 * @param url Tools API endpoint.
 * @param token the bearer token (the whole "Bearer ..." string).
 * @param method the HTTP method.
 * @param payload JSON body describing what to ask Tools for.
 * @param logger shared logger; on failure it writes to the status cell and stops.
 */
async function fetchWithAuth(url: string, token: string, method: string, payload: string | null, logger: Logger): Promise<ResponseJSON> {
  const options = {
    method: method,
    headers: {
      'Content-Type': 'application/json',
      'Authorization': token
    },
    body: payload
  }

  const response = await fetch(url, options)

  // The single most common failure: the 15-minute API key has expired.
  if (response.status === 401 || response.status === 403) {
    logger.error('fetchWithAuth',
      `Authorization failed (HTTP ${response.status}). Your MTC Tools API key has likely expired. Get a fresh key and paste it into cell ${CONFIG.cells.apiToken}, then run again.`)
  }
  if (!response.ok) {
    logger.error('fetchWithAuth',
      `The request to MTC Tools failed (HTTP ${response.status}). Please wait a moment and try again.`)
  }

  let responseJson: ResponseJSON
  try {
    responseJson = await response.json() as ResponseJSON
  } catch (e) {
    logger.error('fetchWithAuth',
      `Could not read the response from MTC Tools (it was not valid data). HTTP status was ${response.status}.`)
  }

  if (responseJson.message === 'Unauthorized') {
    logger.error('fetchWithAuth',
      `MTC Tools rejected the request as "Unauthorized". Your API key has likely expired. Get a fresh key and paste it into cell ${CONFIG.cells.apiToken}.`)
  }

  return responseJson
}

async function MTCSearch(token: string, payload: Payload, logger: Logger): Promise<RestrictedData[]> {
  let skipCount = 0
  let hasMore = true
  let allMissionaries: RestrictedData[] = []
  const url = CONFIG.apiUrl
  let pageCount = 0

  while (hasMore) {
    // Safety cap: never loop more than maxPages times no matter what the API says.
    if (pageCount >= CONFIG.maxPages) {
      logger.error('MTCSearch',
        `Stopped after ${CONFIG.maxPages} pages (safety limit). The API may be misbehaving.`)
    }

    payload.skip = skipCount

    const response = await fetchWithAuth(url, token, 'POST', JSON.stringify(payload), logger)

    if (!response || !response.pagination) {
      logger.error('MTCSearch',
        'The response from MTC Tools was missing pagination info. The API format may have changed.')
    }

    pageCount++
    // Add cleaned results of the data:
    allMissionaries = allMissionaries.concat(response.data.map((m: Record<string, unknown>) => filterToolsJSONInfo(m)))

    const take = response.pagination.take
    // Guard against an invalid page size that would cause an endless loop.
    if (!take || take <= 0) {
      logger.error('MTCSearch',
        `The API returned an invalid page size (take=${take}). Stopping to avoid an endless loop.`)
    }

    // Update skip for next request
    skipCount += take
    // Check if we need to search more
    hasMore = response.pagination.hasMore

    logger.info('MTCSearch', `Fetched page ${pageCount}: ${skipCount} of ${response.pagination.total} missionaries.`)
  }

  logger.info('MTCSearch', `Total missionaries fetched: ${allMissionaries.length}`)
  return allMissionaries
}


function filterToolsJSONInfo(item: Record<string, unknown>): RestrictedData {
  return {
    branchDistrict: (item.branchDistrict as string) ?? "",
    scheduleName: (item.scheduleName as string) ?? "",
    missionaryType: (item.missionaryType as string) ?? "",
    missionaryId: (item.missionaryId as string) ?? "",
    fullName: (item.fullName as string) ?? "",
    trainingLanguage: (item.trainingLanguage as string) ?? "",
    companionshipId: (item.companionshipId as string) ?? "",
    scheduledOnSiteStartDate: ((item.scheduledOnSiteStartDate as string) ?? "").split("T")[0],
    scheduledDepartureDate: ((item.scheduledDepartureDate as string) ?? "").split("T")[0],
    missionEmail: (item.missionEmail as string) ?? "",
    missionName: (item.missionName as string) ?? "",
    classroom: (item.classroom as string) ?? "",
    mtc: (item.mtc as string) ?? "",
  };
}

function writeToSheet(workbook: ExcelScript.Workbook, data: RestrictedData[], logger: Logger) {
  const sheet = workbook.getWorksheet(CONFIG.sheets.rawData);
  if (!sheet) {
    logger.error('writeToSheet', `Sheet "${CONFIG.sheets.rawData}" was not found. Please check the sheet name.`)
  }
  const table = sheet.getTable(CONFIG.table);
  if (!table) {
    logger.error('writeToSheet', `Table "${CONFIG.table}" was not found on the "${CONFIG.sheets.rawData}" sheet.`)
  }

  table.clearFilters();
  const rowCount = table.getRowCount();

  // Convert RestrictedData objects to 2D array
  const rows: (string | number | boolean)[][] = data.map(m => [
    m.branchDistrict,
    m.scheduleName,
    m.missionaryType,
    m.missionaryId,
    m.fullName,
    m.trainingLanguage,
    m.companionshipId,
    m.scheduledOnSiteStartDate,
    m.scheduledDepartureDate,
    m.missionEmail,
    m.missionName,
    m.classroom,
    m.mtc,
  ]);

  if (rowCount === 0) {
    // Table is empty, just add all rows
    table.addRows(-1, rows);

  } else if (data.length >= rowCount) {
    // New data is same size or larger - overwrite existing rows, add the rest
    const overwriteRows = rows.slice(0, rowCount);
    const newRows = rows.slice(rowCount);

    table.getRangeBetweenHeaderAndTotal().setValues(overwriteRows);

    if (newRows.length > 0) {
      table.addRows(-1, newRows);
    }

  } else {
    // New data is shorter - overwrite what we can, clear the remainder, shrink the table
    const columnCount = table.getColumns().length;
    const overwriteRange = table.getRangeBetweenHeaderAndTotal()
      .getCell(0, 0)
      .getResizedRange(data.length - 1, columnCount - 1)
    overwriteRange.setValues(rows)

    const leftoverCount = rowCount - data.length;
    logger.info('writeToSheet', `Clearing ${leftoverCount} leftover row(s).`)

    // Build distinct empty rows (do NOT reuse one array reference for every row).
    const emptyRows: (string | number | boolean)[][] =
      Array.from({ length: leftoverCount }, () => new Array(columnCount).fill(""));
    const leftoverRange = table.getRangeBetweenHeaderAndTotal()
      .getCell(data.length, 0)
      .getResizedRange(leftoverCount - 1, columnCount - 1);
    leftoverRange.setValues(emptyRows)

    // Get the table's actual position
    const headerRange = table.getHeaderRowRange()
    const startRow = headerRange.getRowIndex()
    const startCol = headerRange.getColumnIndex()

    // Build new range from header to last data row and resize the table to fit
    const newRange = sheet.getRangeByIndexes(
      startRow,
      startCol,
      data.length + 1,
      columnCount
    )
    logger.info('writeToSheet', `Resizing table to: ${newRange.getAddress()}`)
    table.resize(newRange.getAddress())
  }

  logger.info('writeToSheet', `Wrote ${data.length} rows to the table.`);
}

function isDataValid(data: RestrictedData[], logger: Logger): boolean {
  // 1. Check we actually got data back
  if (data.length === 0) {
    logger.warn('isDataValid', 'No data was returned from MTC Tools.')
    return false
  }

  // 2. Check we got a reasonable amount (not suspiciously low)
  if (data.length < CONFIG.minExpectedRecords) {
    logger.warn('isDataValid',
      `Only ${data.length} records returned, expected at least ${CONFIG.minExpectedRecords}. If this year's group really is this small, lower CONFIG.minExpectedRecords.`)
    return false
  }

  // 3. Spot check required fields on a few records
  const sample = data.slice(0, 5);
  for (const m of sample) {
    if (!m.missionaryId || !m.fullName || !m.missionEmail) {
      logger.warn('isDataValid', 'Sampled records are missing critical fields (missionaryId / fullName / missionEmail).')
      return false
    }
  }

  return true
}

function serialToISO(serial: number, logger: Logger, label: string): string {
  // Excel's epoch starts on Jan 1, 1900, but has a leap year bug for 1900,
  // so we subtract 25569 days to get to the Unix epoch (Jan 1, 1970).
  if (isNaN(serial)) {
    logger.error('serialToISO', `The date in "${label}" is not a valid date. Please check that cell.`)
  }
  const unixTimestamp = (serial - 25569) * 86400 * 1000
  const date = new Date(unixTimestamp)
  if (isNaN(date.getTime())) {
    logger.error('serialToISO', `Could not convert the date in "${label}" (value ${serial}).`)
  }
  return date.toISOString()
}

async function runSearchPipeline(workbook: ExcelScript.Workbook, logger: Logger) {
  const controlSheet = workbook.getWorksheet(CONFIG.sheets.control)

  const token = String(controlSheet.getRange(CONFIG.cells.apiToken).getValue()).trim()
  if (!token) {
    logger.error('runSearchPipeline',
      `No API key found in cell ${CONFIG.cells.apiToken}. Get a fresh key from MTC Tools and paste it there.`)
  }

  const onSiteStart = serialToISO(Number(controlSheet.getRange(CONFIG.cells.onSiteStart).getValue()), logger, 'On-Site Start')
  const onSiteEnd = serialToISO(Number(controlSheet.getRange(CONFIG.cells.onSiteEnd).getValue()), logger, 'On-Site End')
  const departureStart = serialToISO(Number(controlSheet.getRange(CONFIG.cells.departureStart).getValue()), logger, 'Departure Start')
  const departureEnd = serialToISO(Number(controlSheet.getRange(CONFIG.cells.departureEnd).getValue()), logger, 'Departure End')

  // The MTC scope ("mtcs") now lives in CONFIG.mtcs so it is easy to maintain.
  // "requestedFields" lists what we ask Tools to return; the fields we actually read
  // are defined in filterToolsJSONInfo().
  const payload: Payload = {
    "statuses": [], "types": [{ name: "Sister" }, { name: "Elder" }], "scheduledOnSiteStart": onSiteStart, "scheduledOnSiteEnd": onSiteEnd, "scheduledDepartureStart": departureStart, "scheduledDepartureEnd": departureEnd, "requestedFields": ["branchDistrict", "scheduleName", "type", "missionaryId", "fullName", "trainingLanguage", "companionshipId", "scheduledRemoteStartDate", "scheduledOnSiteStartDate", "scheduledDepartureDate", "missionEmail", "missionName", "classroom"], "mtcs": CONFIG.mtcs, "skip": 0, "take": CONFIG.pageSize};

  // Run MTC Tools Custom Search
  logger.setStatus('Running custom search on MTC Tools. Please wait ...')
  logger.info('runSearchPipeline', 'Starting custom search on MTC Tools.')
  const searchedData = await MTCSearch(token, payload, logger)

  // If data looks good, overwrite the sheet
  if (isDataValid(searchedData, logger)) {
    logger.setStatus('Data is valid!')
    writeToSheet(workbook, searchedData, logger);
    logger.setStatus(`Success: wrote ${searchedData.length} records.`)
    logger.info('runSearchPipeline', `Success: wrote ${searchedData.length} records.`)
  } else {
    logger.setStatus('Data validation failed. Existing table data was preserved.')
    logger.warn('runSearchPipeline', 'Data validation failed. Existing table data was preserved.')
  }
}

async function main(workbook: ExcelScript.Workbook) {
  const controlSheet = workbook.getWorksheet(CONFIG.sheets.control)
  const statusCell = controlSheet ? controlSheet.getRange(CONFIG.cells.status) : undefined
  const logger = new Logger(workbook, CONFIG.sheets.log, statusCell, CONFIG.maxLogRows)

  try {
    if (!controlSheet) {
      logger.error('main', `Required sheet "${CONFIG.sheets.control}" was not found. Please check the sheet name.`)
    }
    logger.info('main', 'Missionary import started.')
    await runSearchPipeline(workbook, logger)
    logger.info('main', 'Missionary import finished.')
  } catch (e) {
    // Surface the failure in the spreadsheet so a non-coder can see what happened.
    logger.setStatus(`Error: ${e.message}`)
    console.log(`Missionary import failed: ${e.message}`)
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
