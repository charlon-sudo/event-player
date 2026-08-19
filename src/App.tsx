import "./App.css";
import { useState } from "react";
import { useMsal } from "@azure/msal-react";

function App() {
  const { instance, accounts } = useMsal();

  const [graphResult, setGraphResult] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);

  const signIn = async () => {
    await instance.loginRedirect({
      scopes: ["User.Read", "Files.Read.All", "Sites.Read.All"],
    });
  };

  const signOut = async () => {
    await instance.logoutRedirect();
  };

  const testGraph = async () => {
    if (accounts.length === 0) {
      return;
    }

    setLoading(true);
    setGraphResult("");

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        scopes: ["User.Read", "Files.Read.All", "Sites.Read.All"],
        account: accounts[0],
      });

      const response = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
          Authorization: `Bearer ${tokenResponse.accessToken}`,
        },
      });

      if (!response.ok) {
        throw new Error(`Graph returned ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();

      setGraphResult(JSON.stringify(data, null, 2));
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Unable to call Microsoft Graph.";

      setGraphResult(message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="app">
      <div className="container">
        <h1>Event Player</h1>

        {accounts.length > 0 ? (
          <>
            <h2>Welcome</h2>
            <p>{accounts[0].username}</p>

            <div className="buttonRow">
              <button className="signInButton" onClick={testGraph}>
                Test Microsoft Graph
              </button>

              <button className="secondaryButton" onClick={signOut}>
                Sign out
              </button>
            </div>

            {loading && <p>Loading...</p>}

            {graphResult && (
              <pre className="resultBox">
                {graphResult}
              </pre>
            )}
          </>
        ) : (
          <button className="signInButton" onClick={signIn}>
            Sign in with Microsoft
          </button>
        )}
      </div>
    </div>
  );
}

export default App;