import * as React from "react";
import * as ReactDOM from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "./components/App";

/* global document, Office */

const rootElement: HTMLElement = document.getElementById("container");
const root = ReactDOM.createRoot(rootElement);

// Hàm khởi tạo chuẩn cho React Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const rootElement = document.getElementById("container");
    
    if (rootElement) {
      const root = ReactDOM.createRoot(rootElement);
      root.render(
        <FluentProvider theme={webLightTheme}>
          <App />
        </FluentProvider>
      );
    }
  }
});
