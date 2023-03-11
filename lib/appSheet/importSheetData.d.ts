export declare const importSheetData: (sheetId: string | undefined, tabId: string | undefined, headerRowIndex?: number) => Promise<({
    [key: string]: string;
} & {
    rowIndex: number;
    a1Range: string;
})[]>;
