import React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider } from "@fluentui/react-components";
import { lookerLightTheme } from "./theme/lookerTheme";
import { App } from "./components/App";

Office.onReady((info) => {
  const container = document.getElementById("root");
  if (!container) return;

  const root = createRoot(container);
  root.render(
    <FluentProvider theme={lookerLightTheme} style={{ height: "100%", width: "100%" }}>
      <App />
    </FluentProvider>
  );
});
