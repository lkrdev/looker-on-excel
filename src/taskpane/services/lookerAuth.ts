export interface AuthTokens {
  accessToken: string;
  refreshToken?: string;
  expiresIn?: number;
  expiresAt?: number;
  baseUrl: string;
  clientId: string;
}

const STORAGE_KEY = "looker_excel_auth";

export function generateRandomString(byteCount: number = 32): string {
  const array = new Uint8Array(byteCount);
  window.crypto.getRandomValues(array);
  return Array.from(array)
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

export async function generateCodeChallenge(verifier: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const digest = await window.crypto.subtle.digest("SHA-256", data);
  return btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

export function getStoredAuth(): AuthTokens | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY) || sessionStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const data = JSON.parse(raw);
    // If expired and has no refresh token, clear and return null
    if (data.expiresAt && Date.now() > data.expiresAt && !data.refreshToken) {
      clearAuth();
      return null;
    }
    return data;
  } catch {
    return null;
  }
}

export function saveAuth(tokens: AuthTokens, persist: boolean = true) {
  const payload = JSON.stringify({
    ...tokens,
    expiresAt: tokens.expiresIn ? Date.now() + tokens.expiresIn * 1000 : tokens.expiresAt,
  });
  if (persist) {
    localStorage.setItem(STORAGE_KEY, payload);
  } else {
    sessionStorage.setItem(STORAGE_KEY, payload);
  }
}

export function clearAuth() {
  localStorage.removeItem(STORAGE_KEY);
  sessionStorage.removeItem(STORAGE_KEY);
}

/**
 * Silently refreshes Looker OAuth access token using refresh_token grant.
 */
export async function refreshAccessToken(tokens: AuthTokens): Promise<AuthTokens> {
  if (!tokens.refreshToken) {
    throw new Error("No refresh token available. Please sign in again.");
  }

  const tokenUrl = `${tokens.baseUrl.replace(/\/$/, "")}/api/token`;
  const res = await fetch(tokenUrl, {
    method: "POST",
    mode: "cors",
    headers: {
      "Content-Type": "application/json;charset=UTF-8",
      "x-looker-appid": "Looker Microsoft Excel Add-in",
    },
    body: JSON.stringify({
      grant_type: "refresh_token",
      client_id: tokens.clientId,
      refresh_token: tokens.refreshToken,
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    clearAuth();
    throw new Error(`Session refresh failed (${res.status}): ${text}`);
  }

  const data = await res.json();
  const updatedTokens: AuthTokens = {
    ...tokens,
    accessToken: data.access_token,
    refreshToken: data.refresh_token || tokens.refreshToken,
    expiresIn: data.expires_in,
    expiresAt: data.expires_in ? Date.now() + data.expires_in * 1000 : undefined,
  };

  saveAuth(updatedTokens);
  return updatedTokens;
}

/**
 * Checks token expiration with 5-minute buffer and silently refreshes if needed.
 */
export async function ensureValidToken(tokens: AuthTokens): Promise<AuthTokens> {
  const BUFFER_MS = 5 * 60 * 1000;
  if (tokens.expiresAt && Date.now() > tokens.expiresAt - BUFFER_MS && tokens.refreshToken) {
    try {
      return await refreshAccessToken(tokens);
    } catch (err) {
      console.warn("Silent token refresh failed, proceeding with current token:", err);
    }
  }
  return tokens;
}

export function openOAuthDialog(authUrl: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const dialogUrl = `${window.location.origin}/dialog-auth.html?authUrl=${encodeURIComponent(authUrl)}`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 65, width: 40, displayInIframe: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          return reject(new Error(result.error.message || "Failed to open Office Dialog"));
        }
        const dialog = result.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
          dialog.close();
          try {
            const data = typeof arg.message === "string" ? JSON.parse(arg.message) : arg;
            if (data.error) {
              reject(new Error(data.errorDescription || data.error));
            } else if (data.code) {
              resolve(data.code);
            } else {
              reject(new Error("No code received from authentication dialog"));
            }
          } catch {
            reject(new Error("Failed to parse response from authentication dialog"));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          reject(new Error("Authentication dialog closed before completion."));
        });
      }
    );
  });
}

export async function exchangeCodeForToken(
  baseUrl: string,
  clientId: string,
  code: string,
  verifier: string,
  redirectUri: string
): Promise<AuthTokens> {
  // Official Looker CORS API Token Endpoint: POST /api/token
  const tokenUrl = `${baseUrl.replace(/\/$/, "")}/api/token`;
  const res = await fetch(tokenUrl, {
    method: "POST",
    mode: "cors",
    headers: {
      "Content-Type": "application/json;charset=UTF-8",
      "x-looker-appid": "Looker Microsoft Excel Add-in",
    },
    body: JSON.stringify({
      grant_type: "authorization_code",
      client_id: clientId,
      redirect_uri: redirectUri,
      code,
      code_verifier: verifier,
    }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Token exchange failed (${res.status}): ${text}`);
  }

  const data = await res.json();
  const tokens: AuthTokens = {
    accessToken: data.access_token,
    refreshToken: data.refresh_token,
    expiresIn: data.expires_in,
    baseUrl,
    clientId,
  };

  saveAuth(tokens);
  return tokens;
}

export async function loginWithApiCredentials(
  baseUrl: string,
  clientId: string,
  clientSecret: string
): Promise<AuthTokens> {
  const tokenUrl = `${baseUrl.replace(/\/$/, "")}/api/4.0/login`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Direct API login failed (${res.status}): ${text}`);
  }

  const data = await res.json();
  const tokens: AuthTokens = {
    accessToken: data.access_token,
    expiresIn: data.expires_in,
    baseUrl,
    clientId,
  };

  saveAuth(tokens);
  return tokens;
}
