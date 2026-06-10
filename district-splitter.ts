type ExcelInput = (string | number | boolean)[][]
interface Missionary {
  district: string;
  type: string;
  name: string;
  compId: string | number;
  destination: number;
  originalIndex: number; // Keeps track of the row for writing back to Excel
}
type DistrictMap = Record<string, 1|2|3>

function setDay(workbook: ExcelScript.Workbook): {} {
  const thurCheck = Boolean(workbook.getWorksheet("MP_DistrictMatch").getRange("AD1").getValue())
  const friCheck = Boolean(workbook.getWorksheet("MP_DistrictMatch").getRange("AE1").getValue())
  let selectedDay = {"thur": false, "fri": false}

  if (thurCheck && friCheck) {
    selectedDay["thur"] = true
  } else if (thurCheck) {
    selectedDay["thur"] = true
  } else if (friCheck) {
    selectedDay["fri"] = true
  } else {
    workbook.getWorksheet("MP_DistrictMatch").getRange("AF1").setValue("Please select a day!")
    throw Error("func [setDay]: No day was selected")
  }

  return selectedDay
}

function getValuesFromTables(workbook: ExcelScript.Workbook, selectedDay: {})
  : { assignedDistrictRange: string[], missionaryRange: ExcelInput} {
  const sheet = workbook.getWorksheet("MP_DistrictMatch")
  let assignedDistrictRange: string[]
  let missionaryRange: ExcelInput

  if (selectedDay['thur']) {
    assignedDistrictRange = flattenInput(sheet.getTable('Mission_Numbers4').getColumn('District').getRangeBetweenHeaderAndTotal().getValues())

    missionaryRange = sheet.getTable('TYM_Missionaries').getRangeBetweenHeaderAndTotal().getValues()

  } else if (selectedDay['fri']) {
    assignedDistrictRange = flattenInput(sheet.getTable('Mission_Numbers4').getColumn('District2').getRangeBetweenHeaderAndTotal().getValues())

    missionaryRange = sheet.getTable('TYM_Missionaries2').getRangeBetweenHeaderAndTotal().getValues()
  }
  return {assignedDistrictRange, missionaryRange}
}

function flattenInput(arr: ExcelInput): string[] {
  const rowCount = arr.length
  const colCount = arr[0]?.length || 0

  const flatArr = new Array<string>(rowCount * colCount)

  let index = 0
  for (let r=0; r<rowCount; r++) {
    for (let c=0; c<colCount; c++) {
      flatArr[index++] = String(arr[r][c])
    }
  }
  return flatArr
}

function buildDistrictMap(arr: string[]): DistrictMap {
  const map: DistrictMap = {}

  for (const str of arr) {
    map[str] = (map[str] || 0) + 1 as 1 | 2 | 3
  }

  return map
}

function assignDestinations(districtMap: DistrictMap, missionaries: Missionary[]): void {
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

    // Scenario A: District is fully assigned to 1
    if (splitDirective === 1) {
      for (const member of members) {
        member.destination = 1
      }
      continue
    }

    // Scenario B & C: District needs to be split 2-way or 3-way
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
    // Determine the number of buckets for this specific district
    const totalCompanionships = Object.keys(companionships).length;

    // Only use 3 buckets if the map dictates it AND there are enough groups to split
    const numBuckets = (splitDirective === 3 && totalCompanionships >= 3) ? 3 : 2;

    // Array to track the current population of each bucket
    // Index 0 = Destination 1, Index 1 = Destination 2, Index 2 = Destination 3
    const bucketCounts: number[] = new Array(numBuckets).fill(0);

    // Greedy N-way distribution function
    const distribute = (groups: Missionary[][]) => {
      // Sort groups by size descending (packs larger groups first for better balance)
      groups.sort((a, b) => b.length - a.length)

      for (const group of groups) {
        // 1. Find the bucket with the lowest current headcount
        let minIndex = 0
        let minCount = bucketCounts[0]

        for (let i=1; i<numBuckets; i++) {
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
    // are spread as evenly as possible across the 2 or 3 buckets.
    distribute(elders)
    distribute(sisters)
  }
}

function processMissionaryData(rawValues: ExcelInput, map: DistrictMap) {
  const missionaries: Missionary[] = []

  // 1. Map 2D array to Objects, ignoring rows with blank districts
  for (let i=0; i<rawValues.length; i++) {
    // Extract and trim to catch strings with only spaces
    const district = String(rawValues[i][0] || '').trim()
    // Skip this row if the district is empty
    if (!district) continue

    missionaries.push({
      district: district,
      type: String(rawValues[i][1]),
      name: String(rawValues[i][2]),
      compId: rawValues[i][3] as string | number,
      destination: 0,
      originalIndex: i
    })
  }

  // 2. Run the partitioning algorithm
  assignDestinations(map, missionaries)

  // 3. Write destinations into an array for Excel output
  const slicerOutput = rawValues.map(row => [row[9]])

  for (const m of missionaries) {
    slicerOutput[m.originalIndex][0] = m.destination
  }

  return slicerOutput
}

function writeOutput(workbook: ExcelScript.Workbook, selectedDay: {}, data: ExcelInput): void {
  const sheet = workbook.getWorksheet('MP_DistrictMatch')
  
  if (selectedDay['thur']) {
    sheet.getTable('TYM_Missionaries').getColumn('Slicer').getRangeBetweenHeaderAndTotal().setValues(data)
  } else {
    sheet.getTable('TYM_Missionaries2').getColumn('Slicer').getRangeBetweenHeaderAndTotal().setValues(data)
  }
}

function main(workbook: ExcelScript.Workbook) {
  const logCell = workbook.getWorksheet("MP_DistrictMatch").getRange("AF1")
  logCell.setValue("Running.")
  const selectedDay = setDay(workbook)
  console.log(selectedDay)
  const {assignedDistrictRange, missionaryRange} = getValuesFromTables(workbook, selectedDay)

  logCell.setValue("Running..")
  const districtMap = buildDistrictMap(assignedDistrictRange)
  // console.log(districtMap)

  logCell.setValue("Running...")
  const destinationOutput = processMissionaryData(missionaryRange, districtMap)

  logCell.setValue("Running....")
  writeOutput(workbook, selectedDay, destinationOutput)
  logCell.setValue("Done.")
}
