export type BaseRow = {
  [key: string]: string;
};

export type ExtRow = BaseRow & {
  rowIndex: number;
  a1Range: string;
};

export type TabListItem = {
  tabId: string;
  tabName: string;
};

export type TabDataItem = {
  [key: string]: string;
} & {
  rowIndex: number;
  a1Range: string;
};

export type DataRowWithId = {
  id: number | string;
  [key: string]: string | number | undefined;
};
