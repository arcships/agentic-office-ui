// editor-clipboard.ts — copy/paste operations
//
// Extracted from useDocxEditor.ts.
// Placeholder module for clipboard integration.

export function createEditorClipboard() {
  // Copy — serializes selected content to clipboard
  const copy = async (): Promise<void> => {
    // Stub — full implementation uses navigator.clipboard.write
  }

  // Paste — reads from clipboard and inserts at selection
  const paste = async (): Promise<void> => {
    // Stub — full implementation uses navigator.clipboard.read
  }

  return { copy, paste }
}
