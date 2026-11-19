// blocks barrel — document block/word extraction and page geometry.
export * from "./types";
export * from "./ids";
export * from "./segmentation";
export * from "./extract";
export * from "./geometry";
export { inspectDocumentSource } from "./source-graph";
export type {
  DocumentSourceAsset,
  DocumentSourceGraph,
  DocumentSourcePart,
  PreservedFeature,
  SourceAttribute,
  SourceElement,
} from "./source-graph";
