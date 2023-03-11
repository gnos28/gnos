# GNOS - Googleapi Now Over Simplificated

## google api wrapper to make it easier to use !

### features :

- cache handling in read operations
- api call delays & queue to stay in the range of google limitations
- error management : will await until google API is available again in case of error, will throw an error if not available during 120s

### installation :

`npm i gnos`

`import { sheetAPI } from "gnos";`

`sheetAPI.setAuthJsonPath(<YOUR PATH TO AUTH.JSON>)`

### usage :

`sheetAPI.getTabIds(sheetId)`
return list of sheet tabs with their respective IDs

`sheetAPI.getTabData(sheetId, tabList, tabName, headerRowIndex?)`
return tab data in the form of an array of object built according to header values
