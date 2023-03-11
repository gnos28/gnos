type Data = {
    id: number | string;
    [key: string]: string | number | undefined;
};
export declare const exportToSheet: (datas: Data[], sheetId: string) => Promise<number | undefined>;
export {};
