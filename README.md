# Meet Your Missionaries and Teaching Missionary Activities

This is the code I wrote for The Church of Jesus Christ of Latter-day Saints Mission Leader Seminar. The code is to be used in Excel. There are three files:

1. import-missionary-data.ts
2. clean-classroom-assignments.ts
3. district-splitter.ts

> **New to maintaining these scripts?** Read [Maintaining the scripts](#maintaining-the-scripts-for-future-maintainers) at the bottom first. It explains where to change sheet/cell names, where errors show up, and the new `Logs` sheet. You should not need to read the rest of the code.

## Importing missionary data (import-missionary-data.ts)

This script makes an API fetch request to [tools.mtc.byu.edu](https://tools.mtc.byu.edu) to run a custom search for missionaries who will be in the Provo MTC during the seminar. The response is cleaned and written to the _Raw Data_ sheet.

### How to set up

First, copy the code in the file and paste it into a new script in Excel → **Automate** → **New Script** → **Create in Code Editor**. Give it a useful name, like the name of the file. You can then add it to the workbook and replace the old button. Note: the old button will not work anymore because the script it was linked to was likely deleted.

On your copy of the workbook, double-check that the dates are correct. They should be a large range like this:

| | On-Site | Departure |
| - | ------ | ------- |
| Start | 23-Mar-26 | 18-Jun-26 |
| End | 16-Jun-26 | 28-Sep-26 |

These cells are formulas and should automatically update the dates for the year given in cell A2.

### How to use

In the API Key input you will need to get your MTC Tools API key by inspecting the site. Since the API key expires after 15 minutes, you will have to keep getting a new one. To get the API key:

1. Open [MTC Tools](https://tools.mtc.byu.edu) and log in.
2. Once on the dashboard, right-click the window and select **Inspect**.
   - This opens a panel with information about the webpage. Click the **Network** tab.
3. With the **Network** tab open, click **Custom Search** on MTC Tools. As the page loads, network traffic appears in the Inspector panel. Look for a request named **custom-search** and click it.

   ![Screenshot of the custom-search packet to select](custom-search.png)

   Scroll down in the request details and find the **Authorization** field. The value is the API key. Triple-click the long string that starts with `Bearer` and copy it (press <kbd>⌘</kbd> + <kbd>C</kbd> or <kbd>Ctrl</kbd> + <kbd>C</kbd>).

   ![Screenshot of the inspect panel with the authorization key for custom-search](inspect-panel.png)

4. Paste the key into the *ControlPanel* sheet in cell C8. (Press <kbd>⌘</kbd> + <kbd>Shift</kbd> + <kbd>V</kbd> or <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>V</kbd> to paste values only.)
5. You are ready to go! Click the button you added or run the script from the *Office Script* panel in Excel.

## Finding open classrooms (clean-classroom-assignments.ts)

This script processes the classroom assignments report from MTC Tools and returns a list of open classrooms based on the activity day and time, plus district temple assignments.

### How to set up

1. Open [MTC Tools](https://tools.mtc.byu.edu) and log in.
2. Go to **Reports** and find report 119, Classroom Assignments.
3. For the date, use the week of seminar and **T4** as the building. Click **OK**.
4. Export the report to Excel using the export icon (paper with an arrow) at the top right, under the report name.
5. Open the Excel file and copy everything in columns A through M. There will be many rows.
6. Paste into the MYM & TYM *classroom assignments* sheet.
7. Do this once a week to keep room information up to date.
8. Temple appointments are optional for the script to run, but having them unlocks a few more rooms. **More importantly**, they show which districts **will not be on campus** to participate. Get the appointment list from Tyler Peck and copy it to the *Temple Appointments* sheet. New appointments are added each week—stay current until the seminar week is scheduled.

### How to use

1. After *classroom assignments* has updated information, run the script to filter through classrooms and return a list of open rooms for an activity.
2. Select the parameters you need for **Day of Event** and **Time of Event**.
   - Meet Your Missionaries is Wednesday evening.
   - Teaching Missionaries is Thursday and Friday afternoon. You can only filter one day at a time, so use the parameter for the day you are planning for.
3. When the script finishes, the output appears in *Clean Classrooms*.
4. On both *MYMRooms* and *TYMRooms*, used rooms show up in AF:AG.
   - Select all rows of districts and rooms and copy to AC:AD using paste values only. (Press <kbd>⌘</kbd> + <kbd>Shift</kbd> + <kbd>V</kbd> or <kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>V</kbd> to paste values only.) We do not want the formula. Rooms that cannot be used for this activity will turn red.
   - **For TYM only**, you may assign part of a district to its own room. The room may highlight red if they are in class during that time—that is expected. Simply remove the district and room from AF:AG and the conditional formatting will return to normal.

## Splitting districts (district-splitter.ts)

This script uses the number of times a district is assigned and splits district members as evenly as possible while keeping companions together. The output is the slicer number on TYM_Missionary tables. This number assigns each missionary to a mission leader. A district that was not split will all have 1's; a district split two ways will have 1's and 2's, and so on.

On the sheet, the **Use Split** column only affects conditional formatting and helps you see whether a district has already been split.

### How to set up

As you assign districts to mission leaders, missionaries appear in the tables to the right. Orange is Thursday and green is Friday. Counts will not update until you assign missionaries to that ML, even after you add a district more than once for splitting. Make sure that the tables 'T:AD' and 'AE:AO' need to be the same number of rows as 'AR:BD' and 'BG:BS' respectivily.

### How to use

1. Select the day to split using the dropdown, then click your button or run the script from the **Automate** tab.
2. The script assigns companionships to mission leaders (MLs), which updates the counts.

## Maintaining the scripts (for future maintainers)

These scripts are written so you can keep them working **without reading the code**. Everything that depends on the workbook layout lives in one place, and any problem is reported right in the spreadsheet.

### The CONFIG block

At the very top of each `.ts` file there is a block that looks like this:

```ts
const CONFIG = {
  sheets: { ... },   // names of the worksheets the script uses
  cells:  { ... },   // specific cells it reads from and writes to
  ...
}
```

If you ever **rename a sheet, move a cell, or rename a table**, update the matching value in `CONFIG` and the script will keep working. You should not need to change anything below the `CONFIG` block. Each line in `CONFIG` has a comment explaining what it controls.

A few values worth knowing about:

- **import-missionary-data.ts** → `CONFIG.minExpectedRecords`: a run is rejected if it returns fewer records than this. If a future year's group is genuinely small and you see "validation failed", lower this number.
- **clean-classroom-assignments.ts** → `CONFIG.assignmentColumns` / `CONFIG.templeColumns`: these are the column positions in the pasted reports. If MTC Tools changes a report's column order, update these numbers.
- **district-splitter.ts** → `CONFIG.missionaryColumnIndexes`: column positions in the missionary tables.

### Where errors and progress show up

Every script writes its progress and any error message into a **status cell** on the sheet (the same cell that used to say "Running..."):

- import-missionary-data.ts → `ControlPanel!G8`
- clean-classroom-assignments.ts → `ControlPanel!G21`
- district-splitter.ts → `MP_DistrictMatch!AF1`

If something goes wrong, that cell will show a plain-language message such as `Error: Your MTC Tools API key has likely expired...` instead of staying stuck on "Running". Read that cell first when a script does not behave.

### The Logs sheet

The first time any script runs it creates a sheet named **Logs** with four columns: Timestamp, Level (INFO/WARN/ERROR), Function, and Message. This is a running history of what each script did, useful for figuring out why a run produced unexpected results. It cleans itself up automatically once it gets large, so you can ignore it day to day and just look when troubleshooting.
