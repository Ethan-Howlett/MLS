let SEMSTART: Date | number | undefined;
let SEMEND: Date | number | undefined;
let EVENTDAY: string;
let EVENTTIME: string;

type TempleAppts = {
  bap: string | Date,
  init: string | Date,
  endow: string | Date,
  sealing: string | Date
}
interface TempleOrdiances {
  bap: string;
  init: string;
  endow: string;
  sealing: string;
}
interface TempleApp {
  district: string;
  lang: string;
  temples: TempleOrdiances;
}

const SCHEDULEDCLASS: { [key: string]: string } = {
  "M": "ME N",
  "T": "MA N",
  "W": "ME N",
  "R": "MA N",
  "F": "AE N",
  "S": "AE N"
}

class BuildingRooms {
  rooms: Room[] = [];

  constructor() { }

  // Helper to find a room by it's name
  getRoomByName(nameToFind: string): Room | undefined {
    return this.rooms.find(room => room.name === nameToFind);
  }

  addRoom(newRoom: Room | string): Room {
    let roomName: string;
    let potentialRoomInstance: Room | undefined;

    if (typeof newRoom === 'string') {
      roomName = newRoom;
    } else {
      roomName = newRoom.name;
      potentialRoomInstance = newRoom; // Keep for later
    }

    const existingRoom = this.getRoomByName(roomName);

    if (existingRoom) {
      console.log(`Room "${roomName}" already exists in the building.`);
      return existingRoom;
    }

    // If we reach here, no room with the same name exists.
    let roomToAdd: Room;
    if (typeof newRoom === 'string') {
      roomToAdd = new Room(newRoom); // Create a new Room from passed-in string
    } else {
      roomToAdd = potentialRoomInstance!;
    }

    this.rooms.push(roomToAdd);
    // console.log(`Room "${roomToAdd.name}" has been added.`);
    return roomToAdd;
  }

  listAllRooms() {
    if (this.rooms.length === 0) {
      console.log("There are no rooms.");
      return;
    }
    console.log("\nAll rooms in building:");
    this.rooms.forEach(room => {
      room.displayInfo();
    });
  }
}

class Room {
  name: string;
  building: string;
  floor: number;
  roomNum: number;
  capacity: number;
  assignments: District[] = [];

  constructor(name: string) {
    this.name = name;
  }

  displayInfo(): void {
    console.log(`Room: ${this.name}, Building: ${this.building}`);
  }

  addAssignment(dist: District): void {
    this.assignments.push(dist);
    // console.log("Assignment added.");
  }

  listAssignments() {
    if (this.assignments.length != 0) {
      this.assignments.forEach(district => {
        console.log(district);
      });
    }
  }
}

class Districts {
  districts: District[] = [];

  constructor() { }

  // Helper to find a district by its name and language
  getDistrictByNameAndLang(nameToFind: string, langToFind: string): District | undefined {
    return this.districts.find(d => d.name === nameToFind && d.lang === langToFind);
  }

  addDistrict(districtInput: District | string, lang?: string, sc?: string,
    room?: string | Room, arr?: string, dep?: string, temple?: TempleAppts): District {
    let districtName: string;
    let districtLang: string;
    let potentialDistrictInstance: District | undefined;

    if (typeof districtInput === 'string') {
      districtName = districtInput;
      if (!lang) {
        // This case should ideally not happen if lang is always logically paired with a string name input
        console.log(`Warning: Adding district by name "${districtName}" without a specified language.`);
        districtLang = "N/A"; // Default or handle error
      } else {
        districtLang = lang;
      }
    } else {
      districtName = districtInput.name;
      districtLang = districtInput.lang;
      potentialDistrictInstance = districtInput;
    }

    const existingDistrict = this.getDistrictByNameAndLang(districtName, districtLang);

    if (existingDistrict) {
      // console.log(`District "${districtName}" with language "${districtLang}" already exists globally.`);
      return existingDistrict;
    }

    let districtToAdd: District;
    if (typeof districtInput === 'string') {
      districtToAdd = new District(districtName, districtLang, sc, room, arr, dep, temple);
    } else {
      // This implies the District object was pre-created with a unique name/lang pair and not found by getDistrictByNameAndLang
      districtToAdd = potentialDistrictInstance!;
    }

    this.districts.push(districtToAdd);
    return districtToAdd;
  }

  checkConflicts() {
    this.districts.forEach(district => {
      district.checkConflicts();
    });
  }
}

class District {
  name: string;
  lang: string;
  arr: string;
  dep: string;
  sc: string;
  room: string | Room;
  temple: TempleAppts;
  classConflict: boolean = false;
  templeConflict: boolean = false;
  assignedRoom: Room | string | undefined = undefined;

  constructor(name: string, lang?: string, sc?: string,
    room?: string | Room, arr?: string, dep?: string, temple?: TempleAppts) {
    this.name = name;
    this.lang = lang || "N/A";
    this.sc = sc || "N/A";
    this.room = room || "N/A";
    this.arr = arr || "N/A";
    this.dep = dep || "N/A";
    this.temple = temple || {
      bap: "N/A",
      init: "N/A",
      endow: "N/A",
      sealing: "N/A"
    };
  }

  setTempleAppts(templeAppts: TempleAppts, logger: Logger) {
    this.temple = templeAppts;
    Object.entries(this.temple).forEach(([key, value]) => {
      if (typeof value == "string") {
        this.temple[key] = convertTempleDateTimeToDate(value, logger);
      }
    });
  }

  checkConflicts() {
    this.classConflict = false; // Initialize conflict state
    this.templeConflict = false;

    // --- 1. Temple Appointment Conflict Check ---
    // TODO: Option to switch if temple includes them or excludeds them.

    // Check if SEMSTART is a valid Date object before proceeding
    if (this.temple.bap !== "N/A" && SEMSTART instanceof Date) {
      const semStartDay = SEMSTART.getDay(); // 0 (Sun) - 6 (Sat)

      // Map for converting EVENTDAY character to day number (0-6)
      const eventDayMap = {
        'Sunday': 0,
        'Monday': 1,
        'Tuesday': 2,
        'Wednesday': 3,
        'Thursday': 4,
        'R': 4, // Alias for Thursday
        'Friday': 5,
        'Saturday': 6
      };

      const eventDayNumber: number = eventDayMap[EVENTDAY];

      if (eventDayNumber !== undefined) {
        const daysUntilEvent = (eventDayNumber - semStartDay + 7) % 7;

        const eventDateForConflict = new Date(SEMSTART);
        eventDateForConflict.setDate(eventDateForConflict.getDate() + daysUntilEvent);

        let eventSlotStartTime = new Date(eventDateForConflict);
        let eventSlotEndTime = new Date(eventDateForConflict);
        let validEventSlot = true;

        // Define event slots (start and end times)
        // Using your specified times as end times and assuming common start times for M/A/E blocks
        switch (EVENTTIME) {
          case "M": // Morning: e.g., 9:00 AM - 12:00 PM
            eventSlotStartTime.setHours(9, 0, 0, 0);
            eventSlotEndTime.setHours(12, 0, 0, 0);
            break;
          case "A": // Afternoon: e.g., 1:00 PM - 5:00 PM
            eventSlotStartTime.setHours(13, 0, 0, 0);
            eventSlotEndTime.setHours(17, 0, 0, 0);
            break;
          case "E": // Evening: e.g., 6:00 PM - 9:00 PM
            eventSlotStartTime.setHours(18, 0, 0, 0);
            eventSlotEndTime.setHours(21, 0, 0, 0);
            break;
          default:
            validEventSlot = false;
            console.log(`Unknown EVENTTIME: ${EVENTTIME} for temple conflict check.`);
            break;
        }

        if (validEventSlot) {
          const TEMPLE_APPT_DURATION_HOURS = 2.5; // Assumption: Temple appointments are 2.5 hours long

          Object.entries(this.temple).forEach(([key, value]) => {
            if (this.templeConflict) return; // Already found a conflict, no need to check further temple appts

            if (value instanceof Date) {
              const templeApptStartTime = value;
              const templeApptEndTime = new Date(templeApptStartTime.getTime() + TEMPLE_APPT_DURATION_HOURS * 60 * 60 * 1000);

              // Check for overlap: TempleApptStart < EventSlotEnd AND TempleApptEnd > EventSlotStart
              if (templeApptStartTime < eventSlotEndTime && templeApptEndTime > eventSlotStartTime) {
                this.templeConflict = true;
                console.log(`District ${this.name} has a conflict with the temple.`)
              }
            }
          });
        }
      } else {
        console.log(`Could not map EVENTDAY "${EVENTDAY}" to a day number for temple conflict check.`);
      }
    }

    // --- 2. Scheduled Class Conflict Check ---
    if (this.sc && this.sc !== "N/A" && EVENTDAY && EVENTTIME && typeof EVENTDAY === 'string') {
      const districtClassDayChar = this.sc.charAt(0);
      const eventDayInitial = EVENTDAY.charAt(0); // Take first char of EVENTDAY, assuming it might be full word initially

      // Check if the district's class is on the same day as the event
      if (districtClassDayChar != eventDayInitial || eventDayInitial == "W") {
        // District has a class on the event day.
        // Check if EVENTTIME falls into the busy slots for that district's class day type.
        // SCHEDULEDCLASS maps the day character of a class to its busy time slots.
        // e.g., SCHEDULEDCLASS['M'] (for a Monday class) gives "ME N" (busy Morning and Evening)
        const busyTimesForDayType = SCHEDULEDCLASS[districtClassDayChar];
        if (busyTimesForDayType && busyTimesForDayType.includes(EVENTTIME)) {
          this.classConflict = true;
        }
      }
    }
    // No final "else { this.conflict = false; }" - preserve conflict if set by temple or schedule
  }
}

/**
 * Converts a date string in "DD-Mon-YY" format (e.g., "18-Jun-25") to a JavaScript Date object.
 * Assumes two-digit years "YY" refer to the 21st century (20YY).
 * @param dateValue The date string to convert.
 * @returns A Date object if parsing is successful, or null if the format is invalid.
 */
function convertDdMonYyToDate(dateValue: string | number, logger: Logger): Date | undefined {
  if (typeof dateValue !== 'number' && typeof dateValue !== 'string') {
    logger.error('convertDdMonYytoDate()', `Invalid input: dateValue must be a non-empty string or excel serial. Got ${dateValue}: ${typeof dateValue}`);
  }
  
  if (typeof dateValue === 'number') {
    return excelSerialToDate(dateValue);

  } else if (typeof dateValue === 'string') {
    const parts = dateValue.split('-');
    if (parts.length !== 3) {
      logger.warn('convertDdMonYytoDate()', `Invalid date format for "${dateValue}". Expected DD-Mon-YY (e.g., "18-Jun-25").`);
      return undefined;
    }

    const dayPart = parts[0];
    const monthPart = parts[1];
    const yearPart = parts[2];

    const day = parseInt(dayPart, 10);
    const yearShort = parseInt(yearPart, 10);

    // Basic validation for day and year parts
    if (isNaN(day) || day < 1 || day > 31) { // Simple day check, more precise check later
      logger.warn('convertDdMonYytoDate()', `Invalid day "${dayPart}" in date string: "${dateValue}".`);
      return undefined;
    }
    if (isNaN(yearShort) || yearShort < 0 || yearShort > 99) {
      logger.warn('convertDdMonYytoDate()', `Invalid year "${yearPart}" in date string: "${dateValue}". Expected 00-99.`);
      return undefined;
    }

    const monthMap: { [key: string]: number } = {
      Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5,
      Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11
    }

    // Normalize month abbv for map lookup
    const normMonth = monthPart.charAt(0).toUpperCase() + monthPart.slice(1).toLowerCase();
    const monthIndex = monthMap[normMonth];

    if (monthIndex === undefined) {
      logger.warn('convertDdMonYytoDate()', `Invalid month abbreviation "${monthPart}" in date string: "${dateValue}".`);
      return undefined;
    }

    // Convert two-digit year "YY" to a four-digit year "YYYY".
    // This assumes "YY" means 20YY (e.g., "25" becomes 2025).
    // This is generally safe for current and near-future dates.
    const fullYear = 2000 + yearShort;

    // Create the Date object (month is 0-indexed in JavaScript Date constructor)
    const dateObject = new Date(fullYear, monthIndex, day);

    // Final validation: Check if the created date matches the input parts.
    // This helps catch invalid dates like "31-Feb-25", which `new Date()` might adjust.
    if (
      dateObject.getFullYear() !== fullYear ||
      dateObject.getMonth() !== monthIndex ||
      dateObject.getDate() !== day
    ) {
      logger.warn('convertDdMonYytoDate()', `Invalid date values in "${dateValue}" (e.g., day out of range for month). Date interpreted as ${dateObject.toLocaleDateString()}`);
      return undefined;
    }

    return dateObject;
  }
}

function excelSerialToDate(serial: number): Date | undefined {
  if (typeof serial !== 'number' || isNaN(serial)) {
    console.log("Invalid input: excelSerialNumber must be a valid number.");
    return undefined;
  }

  // Excel epoch: 1899-12-30 (where serial 0 would be, and serial 1 is 1900-01-01).
  // The number of days from 1899-12-30 to 1970-01-01 UTC is 25569.
  // This constant correctly handles Excel's 1900 leap year bug for dates March 1, 1900, and onwards.
  const excelBaseDateOffset = 25569;
  const millisecondsPerDay = 86400000; // 24 * 60 * 60 * 1000

  // Calculate the UTC timestamp corresponding to the Excel serial date.
  // This timestamp will be at midnight UTC if 'serial' is an integer,
  // or will include time if 'serial' has a fractional part.
  const utcTimestamp = (serial - excelBaseDateOffset) * millisecondsPerDay;

  // Create a temporary Date object based on the UTC timestamp.
  const tempDate = new Date(utcTimestamp);

  if (isNaN(tempDate.getTime())) {
    console.log(`Failed to convert Excel serial ${serial} to a valid intermediate Date object.`);
    return undefined;
  }

  // Extract year, month (0-indexed), and day components as UTC values from the intermediate date.
  // This gives the actual calendar date (Y, M, D) represented by the serial number.
  const year = tempDate.getUTCFullYear();
  const month = tempDate.getUTCMonth(); // 0 for January, 11 for December
  const day = tempDate.getUTCDate();

  // Create the final Date object using these year, month, day components.
  // new Date(year, month, day) constructs a Date object for midnight in the LOCAL timezone.
  // This approach is chosen for consistency with other date creation methods in this file
  // (e.g., convertDdMonYyToDate) that also produce dates at local midnight.
  const resultDate = new Date(year, month, day);

  // Validate the final constructed Date object.
  if (isNaN(resultDate.getTime())) {
    console.log(`Failed to construct final local Date object for Excel serial ${serial} (Y/M/D: ${year}/${month}/${day}).`);
    return undefined;
  }

  // Optional: A check to see if local Date construction significantly altered the calendar date.
  // This is rare but could happen with extreme timezone offsets or DST transitions exactly at midnight.
  if (resultDate.getFullYear() !== year || resultDate.getMonth() !== month || resultDate.getDate() !== day) {
    console.log(`Local date components for serial ${serial} (${resultDate.getFullYear()}-${resultDate.getMonth()}-${resultDate.getDate()}) ` +
      `differ from UTC-derived components (${year}-${month}-${day}). This may indicate unusual timezone behavior.`);
  }

  return resultDate;
}

/**
 * Converts a verbose date-time string (e.g., "Tuesday, June 17 at 9:00 am") to a JavaScript Date object.
 * Assumes the current year if the year is not specified in the string.
 * @param dateTimeString The date-time string to convert.
 * @returns A Date object if parsing is successful, or undefined if the format is invalid.
 */
function convertTempleDateTimeToDate(dateTimeString: string, logger: Logger): Date | undefined {
  if (!dateTimeString || typeof dateTimeString !== 'string') {
    logger.warn('convertTempleDateTimeToDate()', 'Invalid input: dateTimeString must be a non-empty string.');
    return undefined;
  }

  const regex = /^\w+,\s*(\w+)\s*(\d{1,2})\s*at\s*(\d{1,2}):(\d{2})\s*(am|pm)$/i;
  const match = dateTimeString.match(regex);

  if (!match) {
    logger.warn('convertTempleDateTimeToDate()', `Invalid date-time format for "${dateTimeString}". Expected format like "Tuesday, June 17 at 9:00 am".`);
    return undefined;
  }

  const [, monthName, dayStr, hourStr, minuteStr, period] = match;

  const monthMap: { [key: string]: number } = {
    January: 0, February: 1, March: 2, April: 3, May: 4, June: 5,
    July: 6, August: 7, September: 8, October: 9, November: 10, December: 11
  };

  const normalizedMonthName = monthName.charAt(0).toUpperCase() + monthName.slice(1).toLowerCase();
  const monthIndex = monthMap[normalizedMonthName];

  if (monthIndex === undefined) {
    logger.warn('convertTempleDateTimeToDate()', `Invalid month name "${monthName}" in string: "${dateTimeString}".`);
    return undefined;
  }

  const day = parseInt(dayStr, 10);
  let hour = parseInt(hourStr, 10);
  const minute = parseInt(minuteStr, 10);

  if (isNaN(day) || day < 1 || day > 31 || isNaN(hour) || hour < 1 || hour > 12 || isNaN(minute) || minute < 0 || minute > 59) {
    logger.warn('convertTempleDateTimeToDate()', `Invalid day, hour, or minute in string: "${dateTimeString}".`);
    return undefined;
  }

  const lowerPeriod = period.toLowerCase();
  if (lowerPeriod === 'pm' && hour < 12) {
    hour += 12;
  } else if (lowerPeriod === 'am' && hour === 12) { // Midnight case: 12 am is 00 hours
    hour = 0;
  }

  const currentYear = new Date().getFullYear();
  const dateObject = new Date(currentYear, monthIndex, day, hour, minute);

  // Final validation to catch issues like Feb 30
  if (dateObject.getFullYear() !== currentYear ||
    dateObject.getMonth() !== monthIndex ||
    dateObject.getDate() !== day ||
    dateObject.getHours() !== hour ||
    dateObject.getMinutes() !== minute) {
    logger.warn('convertTempleDateTimeToDate()', `Date values resulted in an invalid date (e.g., Feb 30) for string: "${dateTimeString}". Interpreted as ${dateObject.toString()}`);
    return undefined;
  }

  return dateObject;
}

function main(workbook: ExcelScript.Workbook) {
  console.log("main started")
  const logSheet = initLogSheet(workbook)
  const logger = new Logger(logSheet)

  logger.startTimer('pipeline')
  const controlPanel = workbook.getWorksheet('ControlPanel');
  controlPanel.getRange("G21").setValue("Running!");

  EVENTDAY = String(controlPanel.getRange('E21').getValue());
  EVENTDAY = EVENTDAY == "Thursday" ? "R" : EVENTDAY;

  SEMSTART = Number(controlPanel.getRange('C21').getValue());
  SEMSTART = excelSerialToDate(SEMSTART);

  SEMEND = Number(controlPanel.getRange('D21').getValue());
  SEMEND = excelSerialToDate(SEMEND);

  EVENTTIME = String(controlPanel.getRange("F21").getValue()).charAt(0);

  logger.info("main", `Finding open classrooms with these details: Event Day - ${EVENTDAY}, Event Time - ${EVENTTIME}`)

  const T4 = new BuildingRooms();
  const districts = new Districts();

  const range = workbook.getWorksheet('Classroom Assignments').getUsedRange(true);

  const values = range.getValues();
  const processedValues: (string | number)[][] = values.map(row => {
    return row.map(cellValue => {
      if (typeof cellValue === 'boolean') {
        return String(cellValue);
      }
      return cellValue;
    });
  });

  splitRooms(processedValues, districts, T4, logger);
  // console.log(districts)

  const templeAppts = getTempleApptsFromSheet(workbook);
  setTempleAppts(templeAppts, districts, logger);

  const { inClass, atTemple, emptyRooms } = filterAvailableDistricts(T4, districts);

  //   const dist = districts.getDistrictByNameAndLang("01-G", "Tagalog 9");
  //   console.log(dist?.temple);

  const outputSheet = workbook.getWorksheet("Clean Classrooms") || workbook.addWorksheet("Clean Classrooms");
  writeToSheet(inClass, atTemple, emptyRooms, outputSheet, controlPanel);
  logger.endTimer('pipeline')
  logger.summary()
}

function splitRooms(rawData: (string | number)[][], districts: Districts, building: BuildingRooms, logger: Logger) {
  let currRoom: Room | undefined = undefined; // Initialized

  rawData.forEach(item => {
    const firstCellContent = item[0];

    if (typeof firstCellContent == 'string' && firstCellContent.length > 0) {
      if (firstCellContent.startsWith("T")) {
        // console.log('Classroom found: ' + firstCellContent);
        currRoom = building.addRoom(firstCellContent);
      }
    }

    const dist = item[12] ? String(item[12]) : undefined;

    if (currRoom && dist) {
      const lang = item[6] ? String(item[6]) : "N/A";

      // Check if the district (name + lang) is already assigned to the current room
      const alreadyAssignedToCurrentRoom = currRoom.assignments.some(
        assignedDist => assignedDist.name === dist && assignedDist.lang === lang
      );
      // console.log(`Dist: ${dist}, Lang: ${lang}, Assigned: ${alreadyAssignedToCurrentRoom}`)

      if (!alreadyAssignedToCurrentRoom) {
        const arrString = item[10] !== undefined && item[10] !== null ? item[10] : "N/A";
        const depString = item[11] !== undefined && item[11] !== null ? item[11] : "N/A";
        const sc = item[2] ? String(item[2]) : "N/A";
        
        // Skip date processing if the room is empty
        if (sc.length < 3) return;

        const arrivalDate = convertDdMonYyToDate(arrString, logger);
        const departureDate = convertDdMonYyToDate(depString, logger);

        if (arrivalDate && departureDate && SEMSTART instanceof Date && SEMEND instanceof Date) {
          if (arrivalDate <= SEMSTART && departureDate > SEMEND) {
            // Pass currRoom.name as the 'room' property for the District instance
            const newDist: District = districts.addDistrict(dist, lang, sc, currRoom.name, arrivalDate.toLocaleDateString(), departureDate.toLocaleDateString());
            currRoom.addAssignment(newDist);
          }
        } else {
          console.log(`Skipping district ${dist} (lang: ${lang}) due to invalid date, uninitialized SEMSTART/SEMEND, or empty classroom.`);
        }
      }
    } else if (!currRoom && dist) {
      if (dist === 'BN-D') return
      console.log(`Attempted to add district ${dist}, but no current room was set.`);
    }
  });
}

function getTempleApptsFromSheet(workbook: ExcelScript.Workbook) {
  const templeSheet = workbook.getWorksheet("Temple Appointments");
  const range = templeSheet.getUsedRange(true);
  const values = range.getValues();
  const districtsSeen = new Set<string>();
  const templeAppts: TempleApp[] = [];

  // Skip first row (header) by starting from index 1
  values.slice(1).forEach(row => {
    const district = String(row[0]);
    const lang = String(row[5]);
    if (districtsSeen.has(district + lang)) {
      return;
    }
    districtsSeen.add(district + lang);

    templeAppts.push({
      district: district, // District name
      lang: lang,
      temples: {
        bap: String(row[6]),
        init: String(row[8]),
        endow: String(row[10]),
        sealing: String(row[12])
      }
    });
  });

  return templeAppts;
}

function setTempleAppts(templeAppts: TempleApp[], districts: Districts, logger: Logger) {
  templeAppts.forEach(appt => {
    const district = districts.getDistrictByNameAndLang(appt.district, appt.lang);
    if (district) {
      district.setTempleAppts(appt.temples, logger);
    }
  });
}

function filterAvailableDistricts(T4: BuildingRooms, districts: Districts) {
  const inClass: (string)[][] = [["District", "Lang", "Arrival", "Departure", "Schedule", "Room"]];
  const atTemple: (string)[][] = [["District", "Lang", "Arrival", "Departure", "Schedule", "Room"]];
  const emptyRooms: (string)[][] = [["Empty Rooms:", ""]];

  districts.checkConflicts();

  T4.rooms.forEach(room => {
    // Handle empty rooms
    if (room.assignments.length === 0) {
      emptyRooms.push([room.name, ""]);
      return;
    }

    // Process room assignments
    room.assignments.forEach(district => {
      const rowForSheet: (string)[] = [
        district.name, district.lang,
        district.arr, district.dep, district.sc, room.name
      ];
      // If district has class and is not at the temple
      if (district.classConflict && !district.templeConflict) {
        inClass.push(rowForSheet);
      }
      
      if (district.templeConflict) {
        atTemple.push(rowForSheet);
      }

      if ((!district.classConflict && !district.templeConflict) || district.templeConflict) {
        emptyRooms.push([room.name, ""]);
      }
    });
  });

  return { inClass, atTemple, emptyRooms };
}

function writeToSheet(inClass: (string)[][], atTemple: (string)[][], emptyRooms: (string)[][], outputSheet: ExcelScript.Worksheet, controlPanel: ExcelScript.Worksheet) {
  if (inClass.length > 1) {
    outputSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.contents);
    outputSheet.getRange("A1").setValue("Districts in class")
    outputSheet.getRangeByIndexes(1, 0, inClass.length, inClass[1].length).setValues(inClass);
  } else {
    console.log("Nothing to write to the sheet.");
  }

  if (emptyRooms.length > 1) {
    outputSheet.getRangeByIndexes(1, 7, emptyRooms.length, 2).setValues(emptyRooms);
    outputSheet.getRange("I2").setValue(emptyRooms.length - 1);

    controlPanel.getRange("C25").setValue(`Found ${emptyRooms.length - 1} empty rooms with the set parameters.`);
  }

  if (atTemple.length > 1) {
    outputSheet.getRange("L1").setValue("Districts at the temple")
    outputSheet.getRangeByIndexes(1, 11, atTemple.length, atTemple[0].length).setValues(atTemple);
  } else {
    console.log("No temple districts to write to the sheet.");
  }

  controlPanel.getRange("G21").setValue("Done!");
}

type LogLevel = 'info' | 'warn' | 'error'

class Logger {
  private logSheet: ExcelScript.Worksheet
  private timers: Record<string, number>
  private fetchCount: number
  private recordCount: number
  private startTime: number

  constructor(logSheet: ExcelScript.Worksheet) {
    this.logSheet = logSheet
    this.startTime = Date.now()
    this.timers = {}
    this.fetchCount = 0
    this.recordCount = 0
  }

  private timestamp(): string {
    return new Date().toISOString()
  }

  private writeToSheet(level: LogLevel | string, fn: string, msg: string) {
    const usedRange = this.logSheet.getUsedRange()
    const nextRow = usedRange ? usedRange.getRowCount() : 1
    this.logSheet.getRange(`A${nextRow + 1}:D${nextRow + 1}`).setValues([[this.timestamp(), level, fn, msg]])
  }

  log(level: LogLevel, fn: string, msg: string) {
    const entry = `[${level.toUpperCase()}][${fn}] ${msg}`
    console.log(entry)
    this.writeToSheet(level, fn, msg)
    if (level === 'error') {
      throw new Error(entry)
    }
  }

  info(fn: string, msg: string) {
    const entry = `[INFO][${fn}] ${msg}`
    console.log(entry)
    this.writeToSheet('INFO', fn, msg)
  }

  warn(fn: string, msg: string) {
    const entry = `[WARN][${fn}] ${msg}`;
    console.log(entry);
    this.writeToSheet("WARN", fn, msg);
  }

  error(fn: string, msg: string) {
    const entry = `[ERROR][${fn}] ${msg}`;
    console.log(entry);
    this.writeToSheet("ERROR", fn, msg);
    throw new Error(entry);
  }

  startTimer(label: string) {
    this.timers[label] = Date.now()
  }

  endTimer(label: string): number {
    const elapsed = Date.now() - this.timers[label]
    this.log('info', 'Timer', `${label} took ${elapsed}ms`)
    return elapsed
  }

  summary() {
    const totalTime = Date.now() - this.startTime
    const msg = `Completed - ${this.fetchCount} fetches, ${this.recordCount} records, ${totalTime}ms total`
    this.log('info', 'Summary', msg)
  }
}

function initLogSheet(workbook: ExcelScript.Workbook, name: string = "Logs"): ExcelScript.Worksheet {
  let logSheet = workbook.getWorksheet(name)

  if (!logSheet) {
    logSheet = workbook.addWorksheet(name)
    logSheet.getRange("A1:D1").setValues([["Timestamp", "Level", "Function", "Message"]])
    console.log('No logs sheet found. Creating and initializing.')
  } else {
    console.log(`Logs sheet is found`)
  }

  return logSheet
}
