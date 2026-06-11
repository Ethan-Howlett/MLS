// ============================================================================
// CONFIG - Everything specific to the workbook layout lives here.
// If a sheet, cell, or report column moves, update it HERE ONCE.
// ============================================================================
const CONFIG = {
  sheets: {
    control: "ControlPanel",                  // Sheet with the event parameters + status
    classroomAssignments: "Classroom Assignments", // Pasted report 119 data
    templeAppointments: "Temple Appointments", // Optional temple appointment list
    output: "Clean Classrooms",                // Where results are written
    log: "Logs",                               // Run history (auto-created)
  },
  cells: {
    semStart: "C21",          // Semester start date
    semEnd: "D21",            // Semester end date
    eventDay: "E21",          // Day of the event (e.g. "Wednesday", "Thursday")
    eventTime: "F21",         // Time of the event (first letter used: M/A/E)
    status: "G21",            // Progress / errors shown to the user
    emptyRoomsSummary: "C25", // Human-readable count of empty rooms
  },
  // Column POSITIONS (0-based) in the pasted "Classroom Assignments" report.
  // If MTC Tools changes the report layout, update these numbers.
  assignmentColumns: {
    roomName: 0,  // Cells starting with the room prefix mark a new classroom
    schedule: 2,  // Scheduled class code
    lang: 6,
    arrival: 10,
    departure: 11,
    district: 12,
  },
  // Column POSITIONS (0-based) in the "Temple Appointments" sheet.
  templeColumns: {
    district: 0,
    lang: 5,
    bap: 6,
    init: 8,
    endow: 10,
    sealing: 12,
  },
  building: {
    roomPrefix: "T",      // A first-cell value starting with this is a classroom name
    skipDistrict: "BN-D", // District that legitimately has no room; do not warn about it
  },
  maxLogRows: 5000,       // The Logs sheet resets itself once it passes this many rows
}

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

    // Proceed if the semester start is valid AND this district has at least one
    // real temple appointment (any of the four ordinances parsed to a Date).
    // Previously this only checked "bap", which skipped districts whose only
    // appointment was an endowment/sealing.
    const hasAnyTempleAppt = Object.keys(this.temple).some(
      key => this.temple[key] instanceof Date
    );

    if (hasAnyTempleAppt && SEMSTART instanceof Date) {
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
    // Goal: decide whether this district is sitting in its own scheduled class at
    // the exact day/time of the event (which would make its room unavailable).
    //
    // How it works:
    //  - this.sc starts with the day letter of the district's class (M/T/W/R/F/S).
    //  - SCHEDULEDCLASS maps that day letter to the time blocks the class occupies,
    //    encoded as letters: M(orning) A(fternoon) E(vening). For example a Monday
    //    class -> "ME N" means it meets Morning and Evening.
    //  - EVENTTIME is the event's time block letter (M/A/E).
    //
    // NOTE: The boolean below is intentionally preserved exactly as it has worked
    // in production. It treats Wednesday ("W") events specially and otherwise keys
    // off whether the class day differs from the event day. Do not "simplify" it
    // without re-validating against real report data.
    if (this.sc && this.sc !== "N/A" && EVENTDAY && EVENTTIME && typeof EVENTDAY === 'string') {
      const districtClassDayChar = this.sc.charAt(0);
      const eventDayInitial = EVENTDAY.charAt(0); // First char of EVENTDAY

      if (districtClassDayChar != eventDayInitial || eventDayInitial == "W") {
        // Look up the time blocks the district's class occupies, then see if the
        // event's time block falls inside them.
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
 * @returns A Date object if parsing is successful, or undefined if the format is invalid.
 */
function convertDdMonYyToDate(dateValue: string | number, logger: Logger): Date | undefined {
  // Skip-and-warn on unexpected input instead of aborting the whole run.
  if (typeof dateValue !== 'number' && typeof dateValue !== 'string') {
    logger.warn('convertDdMonYyToDate', `Invalid input: expected a date string or Excel serial number. Got ${dateValue}: ${typeof dateValue}. Skipping.`);
    return undefined;
  }

  if (typeof dateValue === 'number') {
    return excelSerialToDate(dateValue);

  } else if (typeof dateValue === 'string') {
    const parts = dateValue.split('-');
    if (parts.length !== 3) {
      logger.warn('convertDdMonYyToDate', `Invalid date format for "${dateValue}". Expected DD-Mon-YY (e.g., "18-Jun-25").`);
      return undefined;
    }

    const dayPart = parts[0];
    const monthPart = parts[1];
    const yearPart = parts[2];

    const day = parseInt(dayPart, 10);
    const yearShort = parseInt(yearPart, 10);

    // Basic validation for day and year parts
    if (isNaN(day) || day < 1 || day > 31) { // Simple day check, more precise check later
      logger.warn('convertDdMonYyToDate', `Invalid day "${dayPart}" in date string: "${dateValue}".`);
      return undefined;
    }
    if (isNaN(yearShort) || yearShort < 0 || yearShort > 99) {
      logger.warn('convertDdMonYyToDate', `Invalid year "${yearPart}" in date string: "${dateValue}". Expected 00-99.`);
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
      logger.warn('convertDdMonYyToDate', `Invalid month abbreviation "${monthPart}" in date string: "${dateValue}".`);
      return undefined;
    }

    // Convert two-digit year "YY" to a four-digit year "YYYY".
    // This assumes "YY" means 20YY (e.g., "25" becomes 2025).
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
      logger.warn('convertDdMonYyToDate', `Invalid date values in "${dateValue}" (e.g., day out of range for month). Date interpreted as ${dateObject.toLocaleDateString()}`);
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

  const utcTimestamp = (serial - excelBaseDateOffset) * millisecondsPerDay;

  const tempDate = new Date(utcTimestamp);

  if (isNaN(tempDate.getTime())) {
    console.log(`Failed to convert Excel serial ${serial} to a valid intermediate Date object.`);
    return undefined;
  }

  // Extract the calendar date (Y, M, D) as UTC values, then rebuild at local midnight
  // for consistency with convertDdMonYyToDate.
  const year = tempDate.getUTCFullYear();
  const month = tempDate.getUTCMonth();
  const day = tempDate.getUTCDate();

  const resultDate = new Date(year, month, day);

  if (isNaN(resultDate.getTime())) {
    console.log(`Failed to construct final local Date object for Excel serial ${serial} (Y/M/D: ${year}/${month}/${day}).`);
    return undefined;
  }

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
    logger.warn('convertTempleDateTimeToDate', 'Invalid input: dateTimeString must be a non-empty string.');
    return undefined;
  }

  const regex = /^\w+,\s*(\w+)\s*(\d{1,2})\s*at\s*(\d{1,2}):(\d{2})\s*(am|pm)$/i;
  const match = dateTimeString.match(regex);

  if (!match) {
    // Many cells are legitimately "N/A" or blank; keep this quiet but logged.
    logger.warn('convertTempleDateTimeToDate', `Could not parse date-time "${dateTimeString}". Expected like "Tuesday, June 17 at 9:00 am".`);
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
    logger.warn('convertTempleDateTimeToDate', `Invalid month name "${monthName}" in string: "${dateTimeString}".`);
    return undefined;
  }

  const day = parseInt(dayStr, 10);
  let hour = parseInt(hourStr, 10);
  const minute = parseInt(minuteStr, 10);

  if (isNaN(day) || day < 1 || day > 31 || isNaN(hour) || hour < 1 || hour > 12 || isNaN(minute) || minute < 0 || minute > 59) {
    logger.warn('convertTempleDateTimeToDate', `Invalid day, hour, or minute in string: "${dateTimeString}".`);
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
    logger.warn('convertTempleDateTimeToDate', `Date values resulted in an invalid date (e.g., Feb 30) for string: "${dateTimeString}". Interpreted as ${dateObject.toString()}`);
    return undefined;
  }

  return dateObject;
}

function splitRooms(rawData: (string | number)[][], districts: Districts, building: BuildingRooms, logger: Logger) {
  const col = CONFIG.assignmentColumns;
  let currRoom: Room | undefined = undefined; // Initialized

  rawData.forEach(item => {
    const firstCellContent = item[col.roomName];

    if (typeof firstCellContent == 'string' && firstCellContent.length > 0) {
      if (firstCellContent.startsWith(CONFIG.building.roomPrefix)) {
        currRoom = building.addRoom(firstCellContent);
      }
    }

    const dist = item[col.district] ? String(item[col.district]) : undefined;

    if (currRoom && dist) {
      const lang = item[col.lang] ? String(item[col.lang]) : "N/A";

      // Check if the district (name + lang) is already assigned to the current room
      const alreadyAssignedToCurrentRoom = currRoom.assignments.some(
        assignedDist => assignedDist.name === dist && assignedDist.lang === lang
      );

      if (!alreadyAssignedToCurrentRoom) {
        const arrString = item[col.arrival] !== undefined && item[col.arrival] !== null ? item[col.arrival] : "N/A";
        const depString = item[col.departure] !== undefined && item[col.departure] !== null ? item[col.departure] : "N/A";
        const sc = item[col.schedule] ? String(item[col.schedule]) : "N/A";

        // A schedule code shorter than 3 characters indicates an empty/placeholder
        // classroom row, so there is nothing to schedule here.
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
          logger.warn('splitRooms', `Skipping district ${dist} (lang: ${lang}) due to invalid date or uninitialized SEMSTART/SEMEND.`);
        }
      }
    } else if (!currRoom && dist) {
      if (dist === CONFIG.building.skipDistrict) return;
      logger.warn('splitRooms', `Attempted to add district ${dist}, but no current room was set.`);
    }
  });
}

function getTempleApptsFromSheet(workbook: ExcelScript.Workbook, logger: Logger): TempleApp[] {
  const templeSheet = workbook.getWorksheet(CONFIG.sheets.templeAppointments);
  if (!templeSheet) {
    logger.warn('getTempleApptsFromSheet', `Sheet "${CONFIG.sheets.templeAppointments}" not found. Continuing without temple appointments.`);
    return [];
  }

  const range = templeSheet.getUsedRange(true);
  if (!range) {
    logger.warn('getTempleApptsFromSheet', 'Temple Appointments sheet is empty. Continuing without temple appointments.');
    return [];
  }

  const col = CONFIG.templeColumns;
  const values = range.getValues();
  const districtsSeen = new Set<string>();
  const templeAppts: TempleApp[] = [];

  // Skip first row (header) by starting from index 1
  values.slice(1).forEach(row => {
    const district = String(row[col.district]);
    const lang = String(row[col.lang]);
    if (districtsSeen.has(district + lang)) {
      return;
    }
    districtsSeen.add(district + lang);

    templeAppts.push({
      district: district, // District name
      lang: lang,
      temples: {
        bap: String(row[col.bap]),
        init: String(row[col.init]),
        endow: String(row[col.endow]),
        sealing: String(row[col.sealing])
      }
    });
  });

  logger.info('getTempleApptsFromSheet', `Loaded ${templeAppts.length} temple appointment record(s).`);
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

  const rowFor = (district: District, room: Room): string[] => [
    district.name, district.lang,
    district.arr, district.dep, district.sc, room.name
  ];

  T4.rooms.forEach(room => {
    // A district physically occupies its room during the event only if it has a
    // scheduled class at that time AND is not away at the temple.
    const occupiers = room.assignments.filter(d => d.classConflict && !d.templeConflict);
    const templeDistricts = room.assignments.filter(d => d.templeConflict);

    occupiers.forEach(d => inClass.push(rowFor(d, room)));
    templeDistricts.forEach(d => atTemple.push(rowFor(d, room)));

    // The room is available for the activity only when nobody is in class in it.
    // This covers truly empty rooms and rooms whose districts are all at the temple.
    // Evaluated once per room, so a room is never listed more than once.
    if (occupiers.length === 0) {
      emptyRooms.push([room.name, ""]);
    }
  });

  return { inClass, atTemple, emptyRooms };
}

function writeToSheet(
  inClass: (string)[][],
  atTemple: (string)[][],
  emptyRooms: (string)[][],
  outputSheet: ExcelScript.Worksheet,
  controlPanel: ExcelScript.Worksheet,
  logger: Logger
) {
  // Always clear stale results from previous runs before writing anything new.
  outputSheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.contents);

  if (inClass.length > 1) {
    outputSheet.getRange("A1").setValue("Districts in class");
    outputSheet.getRangeByIndexes(1, 0, inClass.length, inClass[0].length).setValues(inClass);
  } else {
    logger.info('writeToSheet', 'No "in class" districts to write.');
  }

  if (emptyRooms.length > 1) {
    outputSheet.getRangeByIndexes(1, 7, emptyRooms.length, 2).setValues(emptyRooms);
    outputSheet.getRange("I2").setValue(emptyRooms.length - 1);
    controlPanel.getRange(CONFIG.cells.emptyRoomsSummary).setValue(`Found ${emptyRooms.length - 1} empty rooms with the set parameters.`);
  } else {
    controlPanel.getRange(CONFIG.cells.emptyRoomsSummary).setValue("Found 0 empty rooms with the set parameters.");
  }

  if (atTemple.length > 1) {
    outputSheet.getRange("L1").setValue("Districts at the temple");
    outputSheet.getRangeByIndexes(1, 11, atTemple.length, atTemple[0].length).setValues(atTemple);
  } else {
    logger.info('writeToSheet', 'No temple districts to write.');
  }

  logger.info('writeToSheet',
    `Wrote ${inClass.length - 1} in-class, ${atTemple.length - 1} temple, and ${emptyRooms.length - 1} empty-room rows.`);
}

function main(workbook: ExcelScript.Workbook) {
  const controlPanel = workbook.getWorksheet(CONFIG.sheets.control);
  const statusCell = controlPanel ? controlPanel.getRange(CONFIG.cells.status) : undefined;
  const logger = new Logger(workbook, CONFIG.sheets.log, statusCell, CONFIG.maxLogRows);

  try {
    if (!controlPanel) {
      logger.error('main', `Required sheet "${CONFIG.sheets.control}" was not found. Please check the sheet name.`);
    }

    logger.startTimer('pipeline');
    logger.setStatus("Running!");
    logger.info('main', 'Clean classroom assignments started.');

    EVENTDAY = String(controlPanel.getRange(CONFIG.cells.eventDay).getValue());
    EVENTDAY = EVENTDAY == "Thursday" ? "R" : EVENTDAY;

    SEMSTART = excelSerialToDate(Number(controlPanel.getRange(CONFIG.cells.semStart).getValue()));
    SEMEND = excelSerialToDate(Number(controlPanel.getRange(CONFIG.cells.semEnd).getValue()));

    EVENTTIME = String(controlPanel.getRange(CONFIG.cells.eventTime).getValue()).charAt(0);

    if (!(SEMSTART instanceof Date) || !(SEMEND instanceof Date)) {
      logger.error('main',
        `The semester start/end dates in ${CONFIG.cells.semStart}/${CONFIG.cells.semEnd} are not valid dates. Please check those cells.`);
    }

    logger.info('main', `Finding open classrooms with these details: Event Day - ${EVENTDAY}, Event Time - ${EVENTTIME}`);

    const T4 = new BuildingRooms();
    const districts = new Districts();

    const assignmentsSheet = workbook.getWorksheet(CONFIG.sheets.classroomAssignments);
    if (!assignmentsSheet) {
      logger.error('main', `Sheet "${CONFIG.sheets.classroomAssignments}" was not found. Paste the report data there first.`);
    }
    const range = assignmentsSheet.getUsedRange(true);
    if (!range) {
      logger.error('main', `Sheet "${CONFIG.sheets.classroomAssignments}" appears to be empty. Paste the report data there first.`);
    }

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

    const templeAppts = getTempleApptsFromSheet(workbook, logger);
    setTempleAppts(templeAppts, districts, logger);

    const { inClass, atTemple, emptyRooms } = filterAvailableDistricts(T4, districts);

    const outputSheet = workbook.getWorksheet(CONFIG.sheets.output) || workbook.addWorksheet(CONFIG.sheets.output);
    writeToSheet(inClass, atTemple, emptyRooms, outputSheet, controlPanel, logger);

    logger.endTimer('pipeline');
    logger.summary();
    logger.setStatus("Done!");
  } catch (e) {
    // Surface the failure in the spreadsheet so a non-coder can see what happened.
    logger.setStatus(`Error: ${e.message}`);
    console.log(`Clean classrooms failed: ${e.message}`);
  } finally {
    logger.flush();
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
