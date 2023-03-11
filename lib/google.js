"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.appGmail = exports.appSheet = exports.appDrive = exports.appCalendar = void 0;
const googleapis_1 = require("googleapis");
const getAuth = (scopes) => new googleapis_1.google.auth.GoogleAuth({
    keyFile: "./auth.json",
    scopes,
});
const appCalendar = () => {
    const scopes = [
        "https://www.googleapis.com/auth/calendar",
        "https://www.googleapis.com/auth/calendar.events",
    ];
    const auth = getAuth(scopes);
    const agenda = googleapis_1.google.calendar({
        version: "v3",
        auth,
    });
    return agenda;
};
exports.appCalendar = appCalendar;
const appDrive = () => {
    const scopes = ["https://www.googleapis.com/auth/drive"];
    const auth = getAuth(scopes);
    const drive = googleapis_1.google.drive({
        version: "v3",
        auth,
    });
    return drive;
};
exports.appDrive = appDrive;
const appSheet = () => {
    const scopes = ["https://www.googleapis.com/auth/spreadsheets"];
    const auth = getAuth(scopes);
    const sheets = googleapis_1.google.sheets({
        version: "v4",
        auth,
    });
    return sheets;
};
exports.appSheet = appSheet;
const appGmail = () => {
    const scopes = ["https://www.googleapis.com/auth/gmail.send"];
    const auth = getAuth(scopes);
    const gmail = googleapis_1.google.gmail({
        version: "v1",
        auth,
    });
    return gmail;
};
exports.appGmail = appGmail;
