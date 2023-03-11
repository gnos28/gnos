import { AddProtectedRangeProps } from "./appSheet/batchUpdate";
import { TabListItem } from "./interfaces";
type GetTabMetaDataProps = {
    spreadsheetId: string;
    fields: string;
    ranges: string[];
};
type UpdateSheetRangeProps = {
    sheetId: string;
    tabName: string;
    startCoords: [number, number];
    data: any[][];
};
type ClearTabDataProps = {
    sheetId: string;
    tabList: TabListItem[];
    tabName: string;
    headerRowIndex?: number;
};
export declare const sheetAPI: {
    setAuthJsonPath: (path: string) => void;
    getTabIds: (sheetId: string | undefined) => Promise<{
        sheetId: string;
        sheetName: string;
    }[]>;
    getTabData: (sheetId: string, tabList: TabListItem[], tabName: string, headerRowIndex?: number) => Promise<({
        [key: string]: string;
    } & {
        rowIndex: number;
        a1Range: string;
    })[]>;
    getTabMetaData: ({ spreadsheetId, fields, ranges, }: GetTabMetaDataProps) => Promise<import("googleapis-common").GaxiosResponse<import("googleapis").sheets_v4.Schema$Spreadsheet>>;
    clearCache: () => void;
    updateRange: ({ sheetId, tabName, startCoords, data, }: UpdateSheetRangeProps) => Promise<void>;
    addBatchProtectedRange: ({ spreadsheetId, editors, namedRangeId, sheetId, startColumnIndex, startRowIndex, endColumnIndex, endRowIndex, }: AddProtectedRangeProps) => void;
    runBatchProtectedRange: (spreadsheetId: string) => Promise<void>;
    clearTabData: ({ sheetId, tabList, tabName, headerRowIndex, }: ClearTabDataProps) => Promise<void>;
};
export {};
