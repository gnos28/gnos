# GNOS - Googleapi Now Over Simplificated

## Google api wrapper to make it easier to use !

### Features :

- cache handling in read operations
- api call delays & queue to stay in the range of google limitations
- error management : will await until google API is available again in case of error, will throw an error if not available during 120s

### Authentication and authorization using Service account credentials

Service accounts allow you to perform server to server, app-level authentication using a robot account.  You will create a service account, download a keyfile, and use that to authenticate to Google APIs.  To create a service account:
- Go to the [Create Service Account Key page](https://console.cloud.google.com/apis/credentials/serviceaccountkey)
- Select `New Service Account` in the drop down
- Click the `Create` button

Save the service account credential file somewhere safe, and *do not check this file into source control*!  To reference the service account credential file, you have a few options.

### Enable spreadsheet API :

On the left menu go to `APIs and services` / `Enabled APIs and services`
Click on the top button `+ ENABLE APIS AND SERVICES`
Type `spreadsheet` in the search bar
Click on the result, then on the button `ENABLE`

`https://www.googleapis.com/auth/spreadsheets`

### Installation :

`npm i gnos`

`import { sheetAPI } from "gnos";`

`sheetAPI.setAuthJsonPath(<YOUR PATH TO AUTH.JSON>)`

### Usage :

- **getTabIds**
`sheetAPI.getTabIds(sheetId:string) => Promise<TabListItem[]>`
return list of sheet tabs with their respective IDs

- **getTabData**
`sheetAPI.getTabData(sheetId:string, tabList:TabListItem[], tabName:string, headerRowIndex?:number) => Promise<TabDataItem[]>`
return tab data in the form of an array of object built according to header values
