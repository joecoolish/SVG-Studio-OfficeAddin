import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";
import App from "./components/App";

/* global document, Office, window */

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

const Root: React.FC = () => {
  const [darkMode, setDarkMode] = React.useState(getPreferredDarkMode);

  return (
    <FluentProvider
      theme={darkMode ? webDarkTheme : webLightTheme}
      style={{ minHeight: "100vh", colorScheme: darkMode ? "dark" : "light" }}
    >
      <App darkMode={darkMode} onToggleTheme={() => setDarkMode((current) => !current)} />
    </FluentProvider>
  );
};

Office.onReady(() => root?.render(<Root />));

function getPreferredDarkMode(): boolean {
  const officeTheme = Office.context?.officeTheme;
  if (typeof officeTheme?.isDarkTheme === "boolean") {
    return officeTheme.isDarkTheme;
  }
  if (officeTheme?.bodyBackgroundColor) {
    return isDarkColor(officeTheme.bodyBackgroundColor);
  }
  return window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
}

function isDarkColor(color: string): boolean {
  const match = color.match(/^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i);
  if (!match) {
    return false;
  }
  const [red, green, blue] = match.slice(1).map((component) => Number.parseInt(component, 16));
  return (red * 299 + green * 587 + blue * 114) / 1000 < 128;
}
