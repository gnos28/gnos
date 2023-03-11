type UpdateSheetRange = {
    sheetId: string;
    tabName: string;
    startCoords: [number, number];
    data: any[][];
};
export declare const updateSheetRange: ({ sheetId, tabName, startCoords, data, }: UpdateSheetRange) => Promise<void>;
export {};
