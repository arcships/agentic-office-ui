import fixtures from "../../parity-fixtures/fixtures.json";

export type ParityFixtures = typeof fixtures;
export type ParityDocument = ParityFixtures["documents"][number];

export { fixtures };

export function fixtureDocument(id: string): ParityDocument {
  const document = fixtures.documents.find((entry) => entry.id === id);
  if (!document) throw new Error(`unknown parity fixture document: ${id}`);
  return document;
}
