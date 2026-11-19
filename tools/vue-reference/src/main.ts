import "@fontsource-variable/geist";
import "@fontsource-variable/geist-mono";
import "./parity.css";

import { createApp } from "vue";

import App from "./App.vue";

const container = document.getElementById("app");
if (!container) throw new Error("#app is missing");

createApp(App).mount(container);
