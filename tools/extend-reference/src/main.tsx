import "@fontsource-variable/geist";
import "@fontsource-variable/geist-mono";
import "./styles/parity.css";

import { createRoot } from "react-dom/client";

import { App } from "./App";

const container = document.getElementById("root");
if (!container) throw new Error("#root is missing");

createRoot(container).render(<App />);
