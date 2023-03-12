import { appSheet } from "../google";

type GetTabSizeProps = {
  sheetId: string;
  tabName: string;
  AUTH_JSON_PATH: string;
};

export const getTabSize = async ({
  sheetId,
  tabName,
  AUTH_JSON_PATH,
}: GetTabSizeProps) => {
  const sheetApp = appSheet(AUTH_JSON_PATH);
  const fields = "*";
  const ranges = [`'${tabName}'!A:A`];
  const errorReturn = {
    nbRows: undefined,
    nbColumns: undefined,
  };

  const metaData = await sheetApp.spreadsheets.get({
    spreadsheetId: sheetId,
    fields,
    ranges,
  });

  const sheets = metaData.data.sheets;
  if (!sheets) return errorReturn;
  const gridProperties = sheets[0].properties?.gridProperties;

  if (!gridProperties) return errorReturn;

  return {
    nbRows: gridProperties.rowCount || undefined,
    nbColumns: gridProperties.columnCount || undefined,
  };
};
