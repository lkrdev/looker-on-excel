Office.onReady(() => {
  const urlParams = new URLSearchParams(window.location.search);
  const authUrl = urlParams.get("authUrl");

  if (authUrl) {
    window.location.href = authUrl;
  } else {
    const status = document.getElementById("status-text");
    if (status) {
      status.innerText = "Error: Missing authorization URL parameter.";
      status.style.color = "#a80000";
    }
  }
});
