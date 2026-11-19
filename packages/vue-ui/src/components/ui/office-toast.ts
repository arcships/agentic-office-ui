/** Adapted from extend-hq/ui 70fe70c6f1e1eee53954c6923f3d9f583557a57e (MIT). Copyright (c) 2026 CrowdView Inc, dba Extend; portions Copyright (c) 2023 shadcn. */
/**
 * Toast API for the primitive layer.
 *
 * `OfficeToastProvider` owns the Reka toast store; components below it call
 * `useOfficeToast()` to show status toasts (success / error / warning / info /
 * loading) or use the full Reka manager for custom content.
 */

import {
  useToastManager,
  type ToastAddOptions,
  type ToastPromiseOptions,
  type ToastUpdateOptions,
} from "reka-ui"

export type OfficeToastStatus = "success" | "error" | "warning" | "info" | "loading"

export type OfficeToastOptions = ToastAddOptions
export type OfficeToastUpdate = ToastUpdateOptions
export type OfficeToastPromiseOptions<Value> = ToastPromiseOptions<Value>

export interface OfficeToastApi {
  /** Low-level Reka manager methods. */
  add(options: OfficeToastOptions): string
  close(id?: string): void
  update(id: string, updates: OfficeToastUpdate): void
  promise<Value>(
    promise: Promise<Value>,
    options: OfficeToastPromiseOptions<Value>,
  ): Promise<Value>
  /** Convenience helpers that set `status` for the token-coloured icon. */
  success(title: string, description?: string, options?: OfficeToastOptions): string
  error(title: string, description?: string, options?: OfficeToastOptions): string
  warning(title: string, description?: string, options?: OfficeToastOptions): string
  info(title: string, description?: string, options?: OfficeToastOptions): string
  loading(title: string, description?: string, options?: OfficeToastOptions): string
  dismiss(id?: string): void
}

/**
 * Access the nearest `OfficeToastProvider`.
 *
 * Throws a clear error when used without the provider, matching the repo rule
 * to fail fast on invalid composition.
 */
export function useOfficeToast(): OfficeToastApi {
  let manager: ReturnType<typeof useToastManager>
  try {
    manager = useToastManager()
  } catch {
    throw new Error(
      "[OfficeToast] useOfficeToast() must be called under an <OfficeToastProvider>.",
    )
  }

  function addStatus(
    status: OfficeToastStatus,
    title: string,
    description?: string,
    options: OfficeToastOptions = {},
  ): string {
    return manager.add({
      ...options,
      title,
      description,
      status,
      type: options.type ?? "foreground",
      duration: status === "loading" ? 0 : options.duration,
    })
  }

  return {
    add: manager.add,
    close: manager.close,
    update: manager.update,
    promise: manager.promise,
    success: (title, description, options) =>
      addStatus("success", title, description, options),
    error: (title, description, options) =>
      addStatus("error", title, description, options),
    warning: (title, description, options) =>
      addStatus("warning", title, description, options),
    info: (title, description, options) =>
      addStatus("info", title, description, options),
    loading: (title, description, options) =>
      addStatus("loading", title, description, options),
    dismiss: manager.close,
  }
}
