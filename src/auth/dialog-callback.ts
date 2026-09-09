Office.onReady(() => {
  const urlParams = new URLSearchParams(window.location.search);
  const code = urlParams.get("code");
  const state = urlParams.get("state");
  const error = urlParams.get("error");
  const errorDescription = urlParams.get("error_description");

  if (code) {
    Office.context.ui.messageParent(JSON.stringify({ code, state }));
  } else if (error) {
    Office.context.ui.messageParent(
      JSON.stringify({ error, errorDescription: errorDescription || error })
    );
  } else {
    Office.context.ui.messageParent(
      JSON.stringify({ error: "missing_code", errorDescription: "No authorization code returned from Looker." })
    );
  }
});
