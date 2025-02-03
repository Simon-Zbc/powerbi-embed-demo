import { Configuration } from "@azure/msal-browser";

export const msalConfig: Configuration = {
    auth: {
        clientId: process.env.REACT_APP_CLIENT_ID!,
        authority: process.env.REACT_APP_AUTHORITY,
        redirectUri: process.env.REACT_APP_REDIRECT_URI,
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
    }
};