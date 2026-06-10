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
  leaderPhone: string;
  leaderMobile: string;
  homeAddress: string;
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
 * This function makes the fetch request to MTC Tools.
 * @url takes the url of Tools api endpoint.
 * @token is the bearer token
 * @method is the http fetch method
 * @payload JSON asking Tools for specific info. What we want.
 */
async function fetchWithAuth(url: string, token: string, method: string, payload: string = null) {

  const options = {
    method: method,
    headers: {
      'Content-Type': 'application/json',
      'Authorization': token
    },
    body: payload
  }
  //console.log(url)
  //console.log(method)
  //console.log(token)
  //console.log(options)
  try {
    const response = await fetch(url, options);
    const responseJson: ResponseJSON = await response.json()
    if (responseJson.message === 'Unauthorized') {
      throw Error('There was an error fetching data. Check your API key.')
    }
    return responseJson;

  } catch (e) {
    console.log(`Error message: ${e.message}`)
    // let newToken = ui.prompt('Expired or Incorrect Authorization Token. Enter Updated Token:', ui.ButtonSet.OK_CANCEL).getResponseText();
    // PropertiesService.getUserProperties().setProperty('token', newToken);

    // return fetchWithAuth(url, method, payload);
  }
}

async function MTCSearch(token: string, payload: Payload) {
  let skipCount = 0
  var hasMore = true;
  let allMissionaries: RestrictedData[] = [];
  const url = `https://api.tools.mtc.byu.edu/search/missionaries`
  let count = 0
  while (hasMore) {
    payload.skip = skipCount

    const response = await fetchWithAuth(url, token, 'POST', JSON.stringify(payload))
    if (response === undefined) {
      throw Error('Response is undefined!')
    }
    console.log(response.pagination)
    count++
    //Add cleaned results of the data:
    allMissionaries = allMissionaries.concat(response.data.map((m: Record<string, unknown>) => filterToolsJSONInfo(m)))

    //Update skip for next request
    skipCount += response.pagination.take
    //Check if we need to search more
    hasMore = response.pagination.hasMore

    console.log(`Fetch: ${count} for: ${skipCount} of ${response.pagination.total} missionaries. `)
  }
  console.log(`Length of allMissionaries: ${allMissionaries.length}`)
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

function writeToSheet(workbook: ExcelScript.Workbook, data: RestrictedData[]) {
  const sheet = workbook.getWorksheet("Raw Data");
  const table = sheet.getTable("Raw_Mission");

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
    // New data is same size or larger — overwrite existing rows, add the rest
    const overwriteRows = rows.slice(0, rowCount);
    const newRows = rows.slice(rowCount);

    table.getRangeBetweenHeaderAndTotal().setValues(overwriteRows);

    if (newRows.length > 0) {
      table.addRows(-1, newRows);
    }

  } else {
    // New data is shorter — overwrite what we can, delete the remainder
    const columnCount = table.getColumns().length;
    // const overwriteRange = table.getRange().getCell(1, 0).getResizedRange(data.length - 1, columnCount - 1);
    const overwriteRange = table.getRangeBetweenHeaderAndTotal()
      .getCell(0, 0)
      .getResizedRange(data.length - 1, columnCount - 1)
    overwriteRange.setValues(rows)

    console.log(`Deleting ${rowCount-data.length} rows.`)
    // Overwrite leftover rows with empty strings
    const leftoverCount = rowCount - data.length;
    const emptyRows: (string | number | boolean)[][] = Array(leftoverCount).fill(Array(columnCount).fill(""));
    const leftoverRange = table.getRangeBetweenHeaderAndTotal()
      .getCell(data.length, 0)
      .getResizedRange(leftoverCount - 1, columnCount - 1);
    leftoverRange.setValues(emptyRows)

    // Get the table's actual position
    const headerRange = table.getHeaderRowRange()
    const startRow = headerRange.getRowIndex()
    const startCol = headerRange.getColumnIndex()

    // Build new range from header to last data row
    const newRange = sheet.getRangeByIndexes(
      startRow,
      startCol,
      data.length + 1,
      columnCount
    )
    console.log(`Resizing table to: ${newRange.getAddress()}`)
    table.resize(newRange.getAddress())
  }

  console.log(`Wrote ${data.length} rows to table.`);
}

function isDataValid(data: RestrictedData[]): boolean {
  // 1. Check we actually got data back
  if (data.length === 0) {
    console.log("Validation failed: No data returned.")
    return false
  }

  // 2. Check we got a reasonable amount (not suspiciously low)
  const MIN_EXPECTED = 400
  if (data.length < MIN_EXPECTED) {
    console.log(`Validation failed: Only ${data.length} records returned, expected at least ${MIN_EXPECTED}.`)
    return false
  }

  // 3. Spot check required fields on a few records
  const sample = data.slice(0, 5);
  for (const m of sample) {
    if (!m.missionaryId || !m.fullName || !m.missionEmail) {
      console.log("Validation failed: Records missing critical fields.")
      return false
    }
  }

  return true
}

function serialToISO(serial: number): string {
  // Excel's epoch starts on Jan 1, 1900, but has a leap year bug for 1900
  // So we subtract 25569 days to get to Unix epoch (Jan 1, 1970)
  const unixTimestamp = (serial - 25569) * 86400 * 1000
  return new Date(unixTimestamp).toISOString()
}

async function runSearchPipeline(workbook: ExcelScript.Workbook) {
  const controlSheet = workbook.getWorksheet("ControlPanel")
  const token = String(controlSheet.getRange("C8").getValue())
  const onSiteStart = serialToISO(Number(controlSheet.getRange('D8').getValue()))
  const onSiteEnd = serialToISO(Number(controlSheet.getRange('D9').getValue()))
  const departureStart = serialToISO(Number(controlSheet.getRange('E8').getValue()))
  const departureEnd = serialToISO(Number(controlSheet.getRange('E9').getValue()))
  const logCell = controlSheet.getRange('G8')

  const payload: Payload = {
    "statuses": [], "types": [{ name: "Sister" }, { name: "Elder" }], "scheduledOnSiteStart": onSiteStart, "scheduledOnSiteEnd": onSiteEnd, "scheduledDepartureStart": departureStart, "scheduledDepartureEnd": departureEnd, "requestedFields": ["branchDistrict", "scheduleName", "type", "missionaryId", "fullName", "trainingLanguage", "companionshipId", "scheduledRemoteStartDate", "scheduledOnSiteStartDate", "scheduledDepartureDate", "missionEmail", "missionName", "classroom"], "mtcs": [{ "id": 2010852, "name": "Provo", "typeId": 26, "unitId": 2010852, "leaderName": "Stephen W. Owen", "leaderEmail": "stephen.owen@churchofJesusChrist.org", "officePhone": "+1 385-384-2900", "officeEmail": "tnisse@churchofjesuschrist.org", "officeAddress": "2005 N 900 East\r\nProvo UT  84602\r\nUnited States", "leaderPhone": null, "leaderMobile": null, "homeAddress": null }], "skip": 0, "take": 400};
  
  console.log(payload)
  // Run MTC Tools Custom Search
  logCell.setValue('Running custom search on MTC Tools. Please wait ...')
  const searchedData = await MTCSearch(token, payload)
  // If data looks good, overwrite the sheet
  if (searchedData === undefined) {
    throw Error('SearchData is undefined')
  }
  if (isDataValid(searchedData)) {
    logCell.setValue('Data is valid!')
    writeToSheet(workbook, searchedData);
    logCell.setValue(`Success: wrote ${searchedData.length} records.`)
    console.log(`Success: wrote ${searchedData.length} records.`);
  } else {
    logCell.setValue('Data validation failed. Existing table data was preserved.')
    console.log("Data validation failed. Existing table data was preserved.");
  }
}

async function main(workbook: ExcelScript.Workbook) {
  await runSearchPipeline(workbook)
}
