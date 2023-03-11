export type BaseRow = {
  [key: string]: string;
};

export type ExtRow = BaseRow & {
  rowIndex: number;
  a1Range: string;
};

export type TabListItem = {
  sheetId: string;
  sheetName: string;
};

export type TabDataItem = {
  [key: string]: string;
} & {
  rowIndex: number;
  a1Range: string;
};
