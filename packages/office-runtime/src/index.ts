export {
  OfficeLoadError,
  isOfficeAbortError,
  sanitizeOfficeUrl,
  toOfficeLoadError,
} from "./errors";
export type {
  OfficeErrorFallback,
  OfficeFormat,
  OfficeLoadErrorCode,
  OfficeLoadErrorInit,
  OfficeSourceKind,
} from "./errors";
export {
  detectOoxmlSourceFormat,
  detectSourceFormat,
  isMacroEnabledOfficeFormat,
} from "./source-format";
export type {
  DetectSourceFormatInput,
  OfficeSourceFamily,
  OfficeSourceFormat,
  OoxmlSourceFormat,
  SourceFormatDetection,
} from "./source-format";
export {
  SURFACE_ZOOM_MAX,
  SURFACE_ZOOM_MIN,
  SURFACE_ZOOM_PRECISION,
  clampSurfaceZoom,
  nextSurfaceZoom,
  normalizeWheelDelta,
} from "./gesture-zoom";
export {
  SURFACE_DOUBLE_TAP_ZOOM,
  TOUCH_CONTEXT_MENU_SUPPRESSION_MS,
  TOUCH_CONTEXT_MENU_SUPPRESSION_PX,
  TOUCH_DOUBLE_TAP_DISTANCE_PX,
  TOUCH_DOUBLE_TAP_MS,
  TOUCH_LONG_PRESS_MS,
  TOUCH_LONG_PRESS_MOVE_TOLERANCE_PX,
  TOUCH_MIN_TARGET_PX,
  createTouchGestureTracker,
  nextDoubleTapZoom,
  pinchSurfaceZoom,
  pointerDistance,
} from "./touch-gestures";
export type {
  TouchGesturePoint,
  TouchGesturePointer,
  TouchGesturePointerEvent,
  TouchGestureTracker,
  TouchGestureTrackerOptions,
} from "./touch-gestures";
export {
  assertOfficeInputBytes,
  resolveOfficeInputByteLimit,
  resolveOfficeLimits,
  snapshotOfficeLimits,
} from "./limits";
export type { OfficeLimits } from "./limits";
export {
  createOfficeImageBudget,
  inspectOfficeImage,
} from "./image-budget";
export type {
  OfficeImageBudget,
  OfficeImageBudgetSnapshot,
  OfficeImageFormat,
  OfficeImageMetadata,
} from "./image-budget";
export {
  inspectOfficeArchive,
  validateOfficeArchive,
  validateOfficeArchiveParts,
  validateOfficeXmlEntry,
  validateOfficeXmlText,
  utf8ByteLength,
} from "./archive-budget";
export type {
  OfficeArchiveEntry,
  OfficeArchivePart,
  OfficeArchivePartsOptions,
  OfficeArchivePartsValidation,
  OfficeArchiveSummary,
  OfficeArchiveValidationOptions,
  OfficeArchiveValidationResult,
  OfficeXmlEntrySummary,
} from "./archive-budget";
export { resolveOfficeUrl, snapshotOfficeUrlPolicy } from "./url-policy";
export {
  OFFICE_LINK_PROTOCOLS,
  isSafeOfficeLink,
  sanitizeOfficeLink,
} from "./link-policy";
export type { OfficeLinkProtocol } from "./link-policy";
export type {
  OfficeFetch,
  OfficeFetchBody,
  OfficeFetchBodyReader,
  OfficeFetchHeaders,
  OfficeFetchInit,
  OfficeFetchResponse,
  OfficeUrlPolicy,
} from "./url-policy";
export { loadOfficeSource } from "./source";
export type {
  LoadOfficeSourceOptions,
  OfficeFileLike,
  OfficeSource,
  ResolvedOfficeSource,
} from "./source";
export { createOfficeLoadContext } from "./load-context";
export type {
  CreateOfficeLoadContextOptions,
  OfficeLoadContext,
  OfficeResourceConfig,
} from "./load-context";
export { createLatestTaskCoordinator, createOfficeTaskSequence } from "./load-task";
export type {
  LatestTaskCoordinator,
  OfficeLoadTask,
  OfficeTaskSequence,
  StartOfficeTaskOptions,
} from "./load-task";
export { createOfficeDiagnostic, emitOfficeDiagnostic } from "./diagnostics";
export type {
  OfficeDiagnostic,
  OfficeDiagnosticType,
} from "./diagnostics";
export { createSurfaceSearchSession } from "./surface-search";
export type {
  SurfaceSearchAdapter,
  SurfaceSearchApi,
  SurfaceSearchError,
  SurfaceSearchSession,
  SurfaceSearchState,
  SurfaceSearchStatus,
  SurfaceSearchTaskContext,
} from "./surface-search";
