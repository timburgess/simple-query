import { useEffect } from "react"
import ReactDOM from "react-dom/client"
import { InteractionType, PublicClientApplication } from "@azure/msal-browser"
import { MsalAuthenticationTemplate, MsalProvider } from "@azure/msal-react"
import { useQuery, Operation, makeOperation, createClient, Provider, fetchExchange } from "urql"

import "./index.css"
import { msalConfig } from "./ssoConfig"
import { contextExchange } from "./contextExchange"


const GRAPHQL_URL = "https://localhost:5001/graphql"

const SITES = `
query utags {
  sites {
    id
    name
    description
  }
}
`
const msalInstance = new PublicClientApplication(msalConfig)

// get the msal access token from cache or acquire a new one
const acquireAccessToken = async (): Promise<string> => {
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
    scopes: ["User.Read"],
    account: activeAccount || accounts[0],
  }

  const authResult = await msalInstance.acquireTokenSilent(request)
  return authResult.accessToken
}

// add the access token to the exchange operation context
const addTokenToContext = async (operation: Operation) => {
  // if operation is not a query or mutation, return the operation
  if (operation.kind !== "query" && operation.kind !== "mutation") {
    return operation.context
  }

  const token = await acquireAccessToken()

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

const root = ReactDOM.createRoot(document.getElementById("root") as HTMLElement)

const exchanges = [
  contextExchange({
    getContext: addTokenToContext,
  }),
  fetchExchange,
]

// setup urql client for use with msal
const client = createClient({
  url: GRAPHQL_URL,
  exchanges,
})

root.render(
  <MsalProvider instance={msalInstance}>
    <MsalAuthenticationTemplate interactionType={InteractionType.Redirect}>
      <Provider value={client}>
        <SimpleQuery />
      </Provider>
    </MsalAuthenticationTemplate>
  </MsalProvider>,
)

export function SimpleQuery() {
  const [result] = useQuery({
    query: SITES,
  })
  const { data, fetching, error } = result

  useEffect(() => {
    if (data?.sites) {
      console.log(data.sites)
    }
  }, [data])

  if (fetching) {
    return <div>Fetching universal tags</div>
  }
  if (error) {
    return <div>Unable to get a valid response</div>
  }

  return <div>Simple Query</div>
}
