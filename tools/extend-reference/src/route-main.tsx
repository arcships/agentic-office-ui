import "@fontsource-variable/geist";
import "@fontsource-variable/geist-mono";
import "./styles/parity.css";
import "./styles/route.css";

import { createRoot } from "react-dom/client";

import { RouteApp } from "./RouteApp";

const container = document.getElementById("route-root");
if (!container) throw new Error("#route-root is missing");

createRoot(container).render(<RouteApp />);
