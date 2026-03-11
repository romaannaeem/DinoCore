import "./App.css";
import { useEffect, useState } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest, apiRequest } from "./authConfig";

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [helloMessage, setHelloMessage] = useState("");
  const [error, setError] = useState("");

  const API_BASE_URL = "http://localhost:5067";

  useEffect(() => {
    console.log("MSAL inProgress:", inProgress);
    console.log("MSAL accounts:", accounts);
    console.log("MSAL activeAccount:", instance.getActiveAccount());
    console.log("MSAL isAuthenticated:", isAuthenticated);
  }, [accounts, inProgress, isAuthenticated, instance]);

  const signIn = async () => {
    try {
      setError("");
      setHelloMessage("");
      await instance.loginRedirect(loginRequest);
    } catch (err) {
      console.error("loginRedirect error:", err);
      setError("Login failed.");
    }
  };

  const signOut = async () => {
    try {
      setError("");
      setHelloMessage("");
      await instance.logoutRedirect();
    } catch (err) {
      console.error("logoutRedirect error:", err);
      setError("Logout failed.");
    }
  };

  const callHello = async () => {
    try {
      setError("");
      setHelloMessage("");

      const account = instance.getActiveAccount() ?? accounts[0];
      if (!account) {
        setError("No signed-in account found. Please sign in first.");
        return;
      }

      const tokenResponse = await instance.acquireTokenSilent({
        ...apiRequest,
        account,
      });

      const response = await fetch(`${API_BASE_URL}/hello`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${tokenResponse.accessToken}`,
        },
      });

      if (!response.ok) {
        throw new Error(`API request failed: ${response.status}`);
      }

      const text = await response.text();
      setHelloMessage(text);
    } catch (err) {
      console.error("callHello error:", err);
      setError("Failed to call secured API.");
    }
  };

  return (
    <div style={{ padding: 24 }}>
      <h1>DinoCore Client</h1>

      {isAuthenticated && (
        <div style={{ marginBottom: 12 }}>
          Signed in as: {accounts[0]?.username ?? "(unknown)"}
        </div>
      )}

      <div style={{ display: "flex", flexDirection: "column", gap: 12, marginBottom: 16 }}>
        <button onClick={callHello} disabled={inProgress !== "none"}>
          Call /hello
        </button>

        {!isAuthenticated ? (
          <button onClick={signIn} disabled={inProgress !== "none"}>
            Sign in
          </button>
        ) : (
          <button onClick={signOut} disabled={inProgress !== "none"}>
            Sign out
          </button>
        )}

        
      </div>

      {helloMessage && (
        <div>
          <strong>API response:</strong> {helloMessage}
        </div>
      )}

      {error && (
        <div style={{ color: "red", marginTop: 12 }}>
          {error}
        </div>
      )}
    </div>
  );
}