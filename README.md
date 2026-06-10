# Meet Your Missionaries and Teaching Missionary Activities
This is the code I wrote for The Church of Jesus Christ of Latter-day Saints Mission Leader Seminar. The code is to be used in Excel. There are three files:
1. import-missionary-data.ts
2. clean-classroom-assignments.ts
3. district-splitter.ts

## Importing missionary data (import-missionary-data.ts)
This script will make an API fetch request to tools.mtc.byu.edu to make a custom search for missionaries that will be in the Provo MTC during the seminar. The response is cleaned and written to the _Raw Data_ sheet.

### How to set up
First, copy the code in the file and paste it into a new script in Excel -> Automate -> New Script -> Create in Code Editor. Give it a useful name, like the name of the file. You can then add it to the workbook and replace the old button. Note: the old button will not work anymore because the script that it was linked to was likely deleted.

On your copy of the workbook double check the dates are correct. They should be a large range like this:
| | On-Site | Departure |
| - | ------ | ------- |
| Start | 23-Mar-26 | 18-Jun-26 |
| End | 16-Jun-26 | 28-Sep-26 |

These cells are formulas and should automatically update the dates for you each year.

In the API Key input you will have to get your MTC Tools API key by inspecting Tools.

### How to use

## Finding open classrooms (clean-classroom-assignments.ts)
This script will process the classroom assignments report from MTC Tools and return a list of open classrooms based on what day and time the activity is and district temple assignments.

### How to set up

### How to use

## Splitting districts (district-splitter.ts)
This script will use the number of times a district is being used and split the district members up as evenly as possible all while keeping companions together. The output of this script is the slicer number on TYM_Missionary tables. This number assigns that missionary to the mission leader. A district that was not split will all have 1's, a district that was split two ways will have 1's and 2's, ...

### How to set up

### How to use
