import { LogLevel, Configuration, RedirectRequest, PublicClientApplication } from "@azure/msal-browser"
import { Operation, makeOperation } from "urql"

// enable msal logging
export const MSAL_LOGGING = `${process.env.REACT_APP_MSAL_LOGGING ?? "false"}` === "true"

export const msalConfig: Configuration = {
  auth: {
    clientId: `${process.env.REACT_APP_WEB_CLIENT_ID ?? "853fcc5e-a313-440e-861f-c65ceb84974d"}`,
    authority: `https://login.microsoftonline.com/${
      process.env.REACT_APP_TENANT_ID ?? "213ac3b9-2d86-4223-aa17-c7727e2e175f"
    }`,
    redirectUri: `${process.env.REACT_APP_REDIRECT_URI ?? "http://localhost:3000"}`,
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
  },
}

if (MSAL_LOGGING) {
  msalConfig.system = {
    loggerOptions: {
      logLevel: LogLevel.Trace,
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message)
            return
          case LogLevel.Info:
            console.info(message)
            return
          case LogLevel.Verbose:
            console.debug(message)
            return
          case LogLevel.Warning:
            console.warn(message)
            return
        }
      },
      piiLoggingEnabled: false,
    },
  }
}

// Scopes for id token to be used with the backend LanzaLake API
export const scopes = [
  `api://${process.env.REACT_APP_API_CLIENT_ID ?? "e8b75572-9a46-4efd-9744-c4089dc4a352"}/api`,
]

// wrap scopes in a request object
export const loginRequest: RedirectRequest = {
  scopes,
}

// get the msal access token from cache or acquire a new one
export const acquireAccessToken = async (msalInstance: PublicClientApplication): Promise<string> => {
  const activeAccount = msalInstance.getActiveAccount() // This will only return a non-null value if you have logic somewhere else that calls the setActiveAccount API
  const accounts = msalInstance.getAllAccounts()

  if (!activeAccount && accounts.length === 0) {
    console.error("No accounts found. Please login.")
    /*
     * User is not signed in. Throw error or wait for user to login.
     * Do not attempt to log a user in outside of the context of MsalProvider
     */
  }
  const request = {
    scopes,
    account: activeAccount || accounts[0],
  }

  const authResult = await msalInstance.acquireTokenSilent(request)
  return authResult.accessToken
}

// add the access token to the exchange operation context
export const addTokenToContext = (msalInstance: PublicClientApplication) => async (operation: Operation) => {
  // if operation is not a query or mutation, return the operation
  if (operation.kind !== "query" && operation.kind !== "mutation") {
    return operation.context
  }

  const token = await acquireAccessToken(msalInstance)

  const fetchOptions =
    typeof operation.context.fetchOptions === "function"
      ? operation.context.fetchOptions()
      : operation.context.fetchOptions || {}

  return makeOperation(operation.kind, operation, {
    ...operation.context,
    fetchOptions: {
      ...fetchOptions,
      headers: {
        ...fetchOptions.headers,
        Authorization: `Bearer ${token}`,
      },
    },
  }).context
}
