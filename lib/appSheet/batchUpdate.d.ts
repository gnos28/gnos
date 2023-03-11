export type AddProtectedRangeProps = {
    spreadsheetId: string;
    editors: string[];
    namedRangeId?: string;
    sheetId: number;
    startColumnIndex: number;
    startRowIndex: number;
    endColumnIndex: number;
    endRowIndex: number;
};
export declare const batchUpdate: {
    addProtectedRange: ({ spreadsheetId, editors, namedRangeId, sheetId, startColumnIndex, startRowIndex, endColumnIndex, endRowIndex, }: AddProtectedRangeProps) => void;
    runProtectedRange: (spreadsheetId: string) => Promise<void>;
};
