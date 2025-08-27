import { Configuration, PopupRequest } from "@azure/msal-browser";

// MSAL configuration
export const msalConfig: Configuration = {
  auth: {
    clientId: "your-client-id-here", // Replace with your Azure App Registration Client ID
    authority: "https://login.microsoftonline.com/common", // Can be specific tenant or common
    redirectUri: window.location.origin + "/", // Must match redirect URI in Azure App Registration
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  },
};

// Add scopes here for ID token to be used at Microsoft identity platform endpoints.
export const loginRequest: PopupRequest = {
  scopes: ["User.Read", "Mail.Read", "Files.Read"],
};

// Add the endpoints here for Microsoft Graph API services you'd like to use.
export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
  graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages",
  graphFilesEndpoint: "https://graph.microsoft.com/v1.0/me/drive/root/children",
};
