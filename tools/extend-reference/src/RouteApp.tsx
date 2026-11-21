import * as React from "react";

import { PdfViewerRoute } from "./routes/PdfViewerRoute";

/** Query contract shared with `tools/parity-fixtures/routes.json`. */
export type RouteId = "pdf-viewer";

const routes: Record<RouteId, React.ComponentType> = {
  "pdf-viewer": PdfViewerRoute,
};

export function RouteApp(): React.ReactElement {
  const params = new URLSearchParams(window.location.search);
  const routeId = (params.get("route") ?? "pdf-viewer") as RouteId;
  const theme = params.get("theme") === "dark" ? "dark" : "light";

  // Upstream dark mode is the `.dark` class variant (next-themes
  // attribute="class"), applied on <html> exactly like the docs site.
  React.useLayoutEffect(() => {
    document.documentElement.classList.toggle("dark", theme === "dark");
    document.documentElement.style.colorScheme = theme;
    document.documentElement.dataset.parityRoute = routeId;
    document.documentElement.dataset.parityTheme = theme;
  }, [routeId, theme]);

  const Route = routes[routeId];
  if (!Route) {
    return (
      <p style={{ padding: 24 }}>
        Unknown route “{routeId}”. Known routes: {Object.keys(routes).join(", ")}.
      </p>
    );
  }

  return <Route />;
}
