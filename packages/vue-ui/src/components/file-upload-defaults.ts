// Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn.
// Shipped defaults for FileUpload, mirroring extend-hq/ui
// apps/v4/components/extend/file-upload.tsx at that commit.
//
// The values live outside the SFC because `defineProps()` defaults are hoisted
// outside `setup()` and cannot reference locally declared variables.
import type { FileUploadAcceptedType } from "../types"

/** Extend's default accept list: PDF, Word, Excel/CSV and PNG/JPEG images. */
export const FILE_UPLOAD_DEFAULT_ACCEPT = [
  ".pdf",
  ".doc",
  ".docx",
  ".xlsx",
  ".csv",
  ".png",
  ".jpg",
  ".jpeg",
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "text/csv",
  "image/png",
  "image/jpeg",
].join(",")

/** Extend's default icon cluster: Image / PDF / Sheet. Labels stay optional
 * and are not rendered, so the shipped defaults carry no translatable copy. */
export const FILE_UPLOAD_DEFAULT_ACCEPTED_TYPES: FileUploadAcceptedType[] = [
  { icon: "image" },
  { icon: "pdf" },
  { icon: "sheet" },
]
