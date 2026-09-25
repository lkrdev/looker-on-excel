Office.onReady(() => {
  const urlParams = new URLSearchParams(window.location.search);
  const authUrl = urlParams.get("authUrl");
  const status = document.getElementById("status-text");

  if (authUrl) {
    try {
      const parsed = new URL(authUrl);
      const isLocalhost =
        parsed.hostname === "localhost" || parsed.hostname === "127.0.0.1";
      if (parsed.protocol === "https:" || (parsed.protocol === "http:" && isLocalhost)) {
        window.location.href = parsed.toString();
        return;
      }
      if (status) {
        status.innerText = "Error: Authorization URL must use HTTPS.";
        status.style.color = "#a80000";
      }
    } catch {
      if (status) {
        status.innerText = "Error: Invalid authorization URL format.";
        status.style.color = "#a80000";
      }
    }
  } else if (status) {
    status.innerText = "Error: Missing authorization URL parameter.";
    status.style.color = "#a80000";
  }
});
