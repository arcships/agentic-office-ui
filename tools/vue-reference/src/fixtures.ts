import fixtures from "../../parity-fixtures/fixtures.json";

export type ParityFixtures = typeof fixtures;
export type ParityTheme = "light" | "dark";

export { fixtures };

export function fixtureDocument(id: string): ParityFixtures["documents"][number] {
  const document = fixtures.documents.find((entry) => entry.id === id);
  if (!document) throw new Error(`unknown parity fixture document: ${id}`);
  return document;
}

export function readQuery(): { surface: string | null; theme: ParityTheme } {
  const params = new URLSearchParams(window.location.search);
  return {
    surface: params.get("surface"),
    theme: params.get("theme") === "dark" ? "dark" : "light",
  };
}
