import { readFileSync } from "fs";

export const getAuthJson = (AUTH_JSON_PATH: string) => {
  const rawAuthJson = readFileSync(AUTH_JSON_PATH, { encoding: "utf8" });

  const authJson = JSON.parse(rawAuthJson);

  const GOOGLE_SERVICE_ACCOUNT_EMAIL = authJson.client_email;
  const GOOGLE_PRIVATE_KEY = authJson.private_key;

  return {
    GOOGLE_SERVICE_ACCOUNT_EMAIL,
    GOOGLE_PRIVATE_KEY,
  };
};
