import type { FileSystemItem } from "@arcships/vue-ui"

/**
 * Demo fixture for `/file-system` (#179).
 *
 * Shaped like a document workspace a host would load from a drive listing:
 * two nested folders, an empty folder, image/PDF/Word/Excel files, sizes and
 * timestamps. `thumbnailUrl` points at the demo sample page so the block can
 * show a real `FileThumbnail` preview without a thumbnail server.
 */
export const FILE_SYSTEM_FIXTURE: FileSystemItem[] = [
  {
    id: "folder-contracts",
    name: "Contracts",
    kind: "folder",
    children: [
      {
        id: "folder-contracts-2026",
        name: "2026",
        kind: "folder",
        children: [
          {
            id: "file-northwind-msa",
            name: "northwind-msa.pdf",
            kind: "file",
            size: 482_304,
            mime: "application/pdf",
            updatedAt: "2026-10-14T09:30:00.000Z",
            url: "/samples/sample.pdf",
          },
          {
            id: "file-northwind-amendment",
            name: "northwind-amendment-2.docx",
            kind: "file",
            size: 96_256,
            mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            updatedAt: "2026-11-02T16:05:00.000Z",
          },
        ],
      },
      {
        id: "file-master-services",
        name: "master-services-agreement.pdf",
        kind: "file",
        size: 1_310_720,
        mime: "application/pdf",
        updatedAt: "2026-09-28T11:12:00.000Z",
      },
      {
        id: "file-nda",
        name: "mutual-nda-v4.pdf",
        kind: "file",
        size: 220_160,
        mime: "application/pdf",
        updatedAt: "2026-08-19T08:45:00.000Z",
      },
    ],
  },
  {
    id: "folder-invoices",
    name: "Invoices",
    kind: "folder",
    children: [
      {
        id: "file-invoice-0881",
        name: "invoice-0881.xlsx",
        kind: "file",
        size: 38_912,
        mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        updatedAt: "2026-11-18T07:20:00.000Z",
      },
      {
        id: "file-invoice-0881-scan",
        name: "invoice-0881-scan.png",
        kind: "file",
        size: 286_720,
        mime: "image/png",
        thumbnailUrl: "/samples/contract-page.png",
        updatedAt: "2026-11-18T07:22:00.000Z",
      },
    ],
  },
  {
    id: "folder-archive",
    name: "Archive",
    kind: "folder",
    children: [],
  },
  {
    id: "file-board-deck",
    name: "board-deck-q4.pptx",
    kind: "file",
    size: 4_718_592,
    mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    updatedAt: "2026-11-20T10:00:00.000Z",
  },
  {
    id: "file-reconciliation",
    name: "reconciliation-notes.txt",
    kind: "file",
    size: 5_120,
    mime: "text/plain",
    updatedAt: "2026-11-15T13:40:00.000Z",
  },
]
