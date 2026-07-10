import { OfficeLoadError } from "./errors";
import {
  createOfficeLoadContext,
  type OfficeLoadContext,
  type OfficeResourceConfig,
} from "./load-context";
import type { OfficeLimits } from "./limits";
import type { OfficeSource } from "./source";
import type { OfficeUrlPolicy } from "./url-policy";

export interface OfficeTaskSequence {
  readonly runtimeId: string;
  next(): string;
  dispose(): void;
}

export interface StartOfficeTaskOptions {
  signal?: AbortSignal;
  resources?: OfficeResourceConfig;
  limits?: OfficeLimits;
  urlPolicy?: OfficeUrlPolicy;
}

export interface OfficeLoadTask {
  readonly context: Readonly<OfficeLoadContext>;
  readonly signal: AbortSignal;
  isCurrent(): boolean;
  assertCurrent(): void;
  finish(): void;
}

export interface LatestTaskCoordinator {
  start(source: OfficeSource, options?: StartOfficeTaskOptions): OfficeLoadTask;
  cancel(): void;
  dispose(): void;
}

export function createOfficeTaskSequence(runtimeId: string): OfficeTaskSequence {
  if (!runtimeId.trim()) {
    throw new OfficeLoadError({ code: "INVALID_ARGUMENT", message: "runtimeId 不能为空。" });
  }
  let sequence = 0;
  let disposed = false;
  return Object.freeze({
    runtimeId,
    next(): string {
      if (disposed) {
        throw new OfficeLoadError({ code: "RUNTIME_DISPOSED", message: "运行实例已经销毁。", runtimeId });
      }
      sequence += 1;
      return `${runtimeId}:${sequence}`;
    },
    dispose(): void {
      disposed = true;
    },
  });
}

export function createLatestTaskCoordinator(sequence: OfficeTaskSequence): LatestTaskCoordinator {
  let current: { controller: AbortController; cleanup: () => void } | undefined;
  let disposed = false;

  const cancel = (): void => {
    const active = current;
    current = undefined;
    active?.controller.abort();
    active?.cleanup();
  };

  return {
    start(source: OfficeSource, options: StartOfficeTaskOptions = {}): OfficeLoadTask {
      if (disposed) {
        throw new OfficeLoadError({
          code: "RUNTIME_DISPOSED",
          message: "加载器已经销毁。",
          runtimeId: sequence.runtimeId,
        });
      }
      cancel();
      const taskId = sequence.next();
      const controller = new AbortController();
      const abortFromExternal = (): void => controller.abort(options.signal?.reason);
      let externalListenerAttached = false;
      if (options.signal?.aborted) controller.abort(options.signal.reason);
      else if (options.signal) {
        options.signal.addEventListener("abort", abortFromExternal, { once: true });
        externalListenerAttached = true;
      }
      let cleaned = false;
      const active = {
        controller,
        cleanup: () => {
          if (cleaned) return;
          cleaned = true;
          if (externalListenerAttached) {
            externalListenerAttached = false;
            options.signal?.removeEventListener("abort", abortFromExternal);
          }
        },
      };
      current = active;
      const context = createOfficeLoadContext({
        runtimeId: sequence.runtimeId,
        taskId,
        source,
        resources: options.resources,
        limits: options.limits,
        urlPolicy: options.urlPolicy,
        signal: controller.signal,
      });
      const isCurrent = (): boolean => current === active && !disposed && !controller.signal.aborted;
      return Object.freeze({
        context,
        signal: controller.signal,
        isCurrent,
        assertCurrent(): void {
          if (isCurrent()) return;
          throw new OfficeLoadError({
            code: "STALE_RESULT",
            message: "加载结果已经过期。",
            runtimeId: sequence.runtimeId,
            taskId,
            sourceKind: source.kind,
          });
        },
        finish(): void {
          active.cleanup();
          if (current === active) current = undefined;
        },
      });
    },
    cancel,
    dispose(): void {
      if (disposed) return;
      disposed = true;
      cancel();
    },
  };
}
