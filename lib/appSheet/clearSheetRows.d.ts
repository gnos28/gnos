import { TabListItem } from "../interfaces";
export declare const clearSheetRows: (sheetId: string, tabId: string, headerRowIndex?: number) => Promise<void>;
export declare const clearTabData: (sheetId: string, tabList: TabListItem[], tabName: string, headerRowIndex?: number) => Promise<void>;
