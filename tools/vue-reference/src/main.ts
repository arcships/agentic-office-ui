import "@fontsource-variable/geist";
import "@fontsource-variable/geist-mono";
import "@arcships/vue-ui/theme.css";
import "@arcships/vue-ui/style.css";
import "../../../packages/vue-ui/src/styles/chrome.css";
import "./parity.css";

import { createApp } from "vue";

import App from "./App.vue";

const container = document.getElementById("app");
if (!container) throw new Error("#app is missing");

// #156 scopes every token under [data-office-theme]; set it on <html> so the
// page-level font/background inherit exactly like a host app that mounts
// OfficeThemeProvider at the root.
const theme = new URLSearchParams(window.location.search).get("theme") === "dark" ? "dark" : "light";
document.documentElement.dataset.officeTheme = theme;

createApp(App).mount(container);
