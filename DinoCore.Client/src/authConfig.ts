import type { Configuration } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: "f9463dfa-f304-4f84-a8e9-91e72cd58730",
    authority: "https://login.microsoftonline.com/d0ae9dcf-81d1-4912-b068-fe80bba14ca7",
    redirectUri: "http://localhost:5173/",
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const loginRequest = {
    scopes: ["openid", "profile", "email"],
};

export const apiRequest = {
  scopes: ["api://08dbc65f-9d9c-471d-8bf6-5a0395b41da0/Dino.Access"],
};