# GNOS - Googleapi Now Over Simplificated

## Google api wrapper to make it easier to use !

### Features :

- **cache handling** in read operations
- **api call delays & queue** to stay in the range of google limitations
- **error management** : will await until google API is available again in case of error, will throw an error if not available during 120s or if a single request takes more than 120s to be executed

### Authentication and authorization using Service account credentials

Service accounts allow you to perform server to server, app-level authentication using a robot account.  You will create a service account, download a keyfile, and use that to authenticate to Google APIs.  To create a service account:
- Go to the [Create Service Account Key page](https://console.cloud.google.com/apis/credentials/serviceaccountkey)
- Select `New Service Account` in the drop down
- Click the `Create` button

Save the service account credential file somewhere safe, and *do not check this file into source control* ! 

### Enable spreadsheet API :

On the left menu go to `APIs and services` / `Enabled APIs and services`
Click on the top button `+ ENABLE APIS AND SERVICES`
Type `spreadsheet` in the search bar
Click on the result, then on the button `ENABLE`

`https://www.googleapis.com/auth/spreadsheets`

### Installation :

install package
`npm i gnos`

import library
`import { sheetAPI } from "gnos";`

set your path to auth.json, include filename in the path (ex. `"./auth.json"`)
`sheetAPI.setAuthJsonPath(<YOUR PATH TO AUTH.JSON>)`

### Usage :

#### READ operations :

- **getTabIds**
`sheetAPI.getTabIds({sheetId:string, disableCache?: boolean}) => Promise<TabListItem[]>`
return list of sheet tabs with their respective IDs

- **getTabData**
`sheetAPI.getTabData({sheetId:string, tabName:string, headerRowIndex?:number, tabList?:TabListItem[], disableCache?: boolean}) => Promise<TabDataItem[]>`
return tab data in the form of an array of objects built according to header values

- **getTabSize**
`sheetAPI.getTabData({sheetId: string, tabName: string) => Promise<TabSize>`
return columns & rows count for a tab

- **clearCache**
`sheetAPI.clearCache() => void`
clear all of the above operations cache

- **getProtectedRangeIds**
`sheetAPI.getProtectedRangeIds({sheetId: string, tabName: string) => Promise<number[]>`
get all protectedRange ids of a defined tab

#### WRITE operations :

- **updateRange**
`sheetAPI.updateRange({sheetId: string, tabName: string, startCoords: [number, number], data: any[][]}) => Promise<void>`
update a specific range

- **appendToSheet**
`sheetAPI.appendToSheet({sheetId: string, tabName: string, data: Data[]}) => Promise<void>`
append a line of data to second line of a tab (matching headers)

- **addBatchProtectedRange**
`sheetAPI.addBatchProtectedRange({sheetId: string, editors: string[], namedRangeId?: string, tabId: number, startColumnIndex: number, startRowIndex: number, endColumnIndex: number, endRowIndex: number}) => void`
add a request to batch buffer

- **runBatchProtectedRange**
`sheetAPI.runBatchProtectedRange({sheetId: string, onTimeoutCallback:() => Promise<void>}) => Promise<void>`
run all requests in batch buffer, "onTimeoutCallback" will be run if request takes more than 120s to be executed

- **deleteProtectedRange**
`sheetAPI.deleteProtectedRange({sheetId: string, protectedRangeIds:number[]}) => Promise<void>`
delete a list of protected range ids

- **clearTabData**
`sheetAPI.runBatchProtectedRange({ sheetId: string, tabName: string, headerRowIndex?: number, tabList?: TabListItem[]}) => Promise<void>`
clear all values of a defined tab

#### MISC

- **enableConsoleLog**
`sheetAPI.enableConsoleLog() => void`
turns on logs with detailed infos over cache usage & delay between requests

- **clearBuffer**
`sheetAPI.clearBuffer() => void`
clear all requests batch buffer

- **getBatchProtectedRange**
`sheetAPI.getBatchProtectedRange() => {[key: string]: sheets_v4.Schema$Request[];}`
return all requests in batch buffer