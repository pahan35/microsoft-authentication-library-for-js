/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
import { PublicClientApplication, LogLevel, AuthorizationUrlRequest, Configuration } from "@azure/msal-node";
import open from "open";

// Before running the sample, you will need to replace the values in the config.
const config: Configuration = {
    auth: {
        clientId: "c3a8e9df-f1d4-427d-be73-acab139c40fd",
        authority: "https://login.microsoftonline.com/common"
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel: LogLevel, message: string, containsPii: boolean) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Verbose,
        }
    }
};

// Create msal application object
const pca = new PublicClientApplication(config);

/**
 * Proof Key for Code Exchange (PKCE) Setup
 *
 * MSAL enables PKCE in the Authorization Code Grant Flow by including the codeChallenge and codeChallengeMethod parameters
 * in the request passed into getAuthCodeUrl() API, as well as the codeVerifier parameter in the
 * second leg (acquireTokenByCode() API).
 *
 * MSAL Node provides PKCE Generation tools through the CryptoProvider class, which exposes
 * the generatePkceCodes() asynchronous API. As illustrated in the example below, the verifier
 * and challenge values should be generated previous to the authorization flow initiation.
 *
 * For details on PKCE code generation logic, consult the
 * PKCE specification https://tools.ietf.org/html/rfc7636#section-4
 */

// Add PKCE code challenge and challenge method to authCodeUrl request objectgit st
const authCodeUrlParameters: AuthorizationUrlRequest = {
    scopes: ["user.read"],
    redirectUri: "http://localhost:3000"
};

// Get url to sign user in and consent to scopes needed for application
const startNavigation = async (url: string): Promise<void> => {
    open(url);
};

const endNavigation = async (): Promise<void> => {
    // do nothing
}
pca.acquireTokenInteractive(authCodeUrlParameters, startNavigation, endNavigation)
    .then((response => {
        console.log("SUCCESS!");
        console.log(response);
    }))
    .catch((error) => {
        console.log("FAILED");
        console.log(error);
    })
    .finally(() => {
        process.exit(0);
    });
