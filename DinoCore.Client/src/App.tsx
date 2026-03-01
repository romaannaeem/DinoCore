import "./App.css";
import { useEffect } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  useEffect(() => {
    console.log("MSAL inProgress:", inProgress);
    console.log("MSAL accounts:", accounts);
    console.log("MSAL activeAccount:", instance.getActiveAccount());
    console.log("MSAL isAuthenticated:", isAuthenticated);
  }, [accounts, inProgress, isAuthenticated, instance]);

  const signIn = async () => {
    try {
      // Redirect flow: main window goes to Azure, then comes back with tokens.
      // More reliable than popup (avoids timed_out, popup blockers, blank popup issues).
      await instance.loginRedirect(loginRequest);
    } catch (err) {
      console.error("loginRedirect error:", err);
    }
  };

  const signOut = async () => {
    try {
      await instance.logoutRedirect();
    } catch (err) {
      console.error("logoutRedirect error:", err);
    }
  };

  return (
    <div style={{ padding: 24 }}>
      <h1>DinoCore Client</h1>

      {!isAuthenticated ? (
        <button onClick={signIn} disabled={inProgress !== "none"}>
          Sign in
        </button>
      ) : (
        <>
          <div style={{ marginBottom: 12 }}>
            Signed in as: {accounts[0]?.username ?? "(unknown)"}
          </div>
          <button onClick={signOut}>Sign out</button>
        </>
      )}
    </div>
  );
}