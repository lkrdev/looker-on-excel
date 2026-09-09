import React, { useState } from "react";
import {
  makeStyles,
  tokens,
  Input,
  Label,
  Button,
  Spinner,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  MessageBar,
  MessageBarBody,
  Body1,
  Caption1,
} from "@fluentui/react-components";
import { PlugConnectedRegular, KeyRegular, GlobeRegular, ShieldCheckmarkRegular } from "@fluentui/react-icons";
import { LookerIcon } from "./LookerLogo";
import { lookerColors } from "../theme/lookerTheme";
import {
  generateRandomString,
  generateCodeChallenge,
  openOAuthDialog,
  exchangeCodeForToken,
  loginWithApiCredentials,
  AuthTokens,
} from "../services/lookerAuth";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "24px 16px",
  },
  hero: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    textAlign: "center",
    gap: "10px",
    padding: "12px 8px 8px 8px",
  },
  iconBadge: {
    width: "60px",
    height: "60px",
    borderRadius: "16px",
    backgroundColor: "#F8FAFD",
    border: "1px solid #E8F0FE",
    boxShadow: "0 4px 14px rgba(66, 133, 244, 0.14)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  heroTitle: {
    fontFamily:
      "'Google Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
    fontSize: "18px",
    fontWeight: 600,
    color: "#202124",
    margin: 0,
  },
  heroSubtitle: {
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
    lineHeight: "16px",
    maxWidth: "280px",
  },
  trustPill: {
    display: "inline-flex",
    alignItems: "center",
    gap: "4px",
    fontSize: "11px",
    color: lookerColors.navy,
    backgroundColor: lookerColors.blueLight,
    padding: "2px 8px",
    borderRadius: "12px",
    fontWeight: 500,
  },
  field: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  devSection: {
    marginTop: "4px",
  },
  buttonGroup: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "8px",
  },
});

interface AuthViewProps {
  onSuccess: (tokens: AuthTokens) => void;
}

export const AuthView: React.FC<AuthViewProps> = ({ onSuccess }) => {
  const styles = useStyles();

  const [baseUrl, setBaseUrl] = useState("");
  const [clientId, setClientId] = useState("");
  const [clientSecret, setClientSecret] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Standard OAuth 2.0 PKCE Flow
  const handleConnectOAuth = async () => {
    setError(null);
    setLoading(true);

    try {
      const cleanUrl = baseUrl.trim().replace(/\/$/, "");
      const verifier = generateRandomString(32);
      const challenge = await generateCodeChallenge(verifier);
      const state = generateRandomString(16);
      const redirectUri = `${window.location.origin}/dialog-callback.html`;

      const authUrl =
        `${cleanUrl}/auth?response_type=code` +
        `&client_id=${encodeURIComponent(clientId.trim())}` +
        `&redirect_uri=${encodeURIComponent(redirectUri)}` +
        `&scope=cors_api` +
        `&state=${encodeURIComponent(state)}` +
        `&code_challenge=${encodeURIComponent(challenge)}` +
        `&code_challenge_method=S256`;

      // Open Office Dialog
      const code = await openOAuthDialog(authUrl);

      // Exchange code for token
      const tokens = await exchangeCodeForToken(
        cleanUrl,
        clientId.trim(),
        code,
        verifier,
        redirectUri
      );
      onSuccess(tokens);
    } catch (err: any) {
      setError(err.message || "Failed to authenticate with Looker.");
    } finally {
      setLoading(false);
    }
  };

  // Direct API Credentials Login (Dev / Fallback)
  const handleConnectDirectApi = async () => {
    if (!clientSecret) {
      setError("Please provide a Client Secret for direct API authentication.");
      return;
    }
    setError(null);
    setLoading(true);

    try {
      const cleanUrl = baseUrl.trim().replace(/\/$/, "");
      const tokens = await loginWithApiCredentials(
        cleanUrl,
        clientId.trim(),
        clientSecret.trim()
      );
      onSuccess(tokens);
    } catch (err: any) {
      setError(err.message || "Direct API login failed.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.hero}>
        <div className={styles.iconBadge}>
          <LookerIcon size={36} variant="color" />
        </div>
        <div>
          <h2 className={styles.heroTitle}>Connect to Looker</h2>
          <div className={styles.heroSubtitle}>
            Stream governed semantic data directly into high-speed Excel tables and PivotTables.
          </div>
        </div>
        <div className={styles.trustPill}>
          <ShieldCheckmarkRegular style={{ fontSize: "12px" }} />
          OAuth 2.0 PKCE Verified
        </div>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      <div className={styles.field}>
        <Label required htmlFor="base-url">
          Looker Instance URL
        </Label>
        <Input
          id="base-url"
          contentBefore={<GlobeRegular />}
          value={baseUrl}
          onChange={(_, data) => setBaseUrl(data.value)}
          placeholder="https://company.looker.com"
          disabled={loading}
        />
      </div>

      <div className={styles.field}>
        <Label required htmlFor="client-id">
          OAuth Client ID
        </Label>
        <Input
          id="client-id"
          contentBefore={<KeyRegular />}
          value={clientId}
          onChange={(_, data) => setClientId(data.value)}
          placeholder="Looker API Client ID"
          disabled={loading}
        />
      </div>

      <div className={styles.buttonGroup}>
        <Button
          appearance="primary"
          icon={<PlugConnectedRegular />}
          onClick={handleConnectOAuth}
          disabled={loading || !baseUrl || !clientId}
          size="large"
          style={{
            backgroundColor: lookerColors.blueDark,
            boxShadow: "0 1px 3px rgba(0,0,0,0.12)",
          }}
        >
          {loading ? <Spinner size="tiny" /> : "Sign in with Looker (OAuth)"}
        </Button>
      </div>

      {/* Direct API Login Accordion for Dev Environments */}
      <div className={styles.devSection}>
        <Accordion collapsible>
          <AccordionItem value="api-secret">
            <AccordionHeader size="small">Direct API Secret (Dev & Automation)</AccordionHeader>
            <AccordionPanel>
              <div style={{ display: "flex", flexDirection: "column", gap: "10px", marginTop: "8px" }}>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  Bypasses OAuth browser popup using API Client Secret. Ideal for Cloudtop or headless test environments.
                </Caption1>
                <div className={styles.field}>
                  <Label htmlFor="client-secret">Client Secret</Label>
                  <Input
                    id="client-secret"
                    type="password"
                    value={clientSecret}
                    onChange={(_, data) => setClientSecret(data.value)}
                    placeholder="Looker API Client Secret"
                    disabled={loading}
                  />
                </div>
                <Button
                  appearance="secondary"
                  onClick={handleConnectDirectApi}
                  disabled={loading || !baseUrl || !clientId || !clientSecret}
                >
                  {loading ? <Spinner size="tiny" /> : "Connect via Direct Secret"}
                </Button>
              </div>
            </AccordionPanel>
          </AccordionItem>
        </Accordion>
      </div>
    </div>
  );
};
