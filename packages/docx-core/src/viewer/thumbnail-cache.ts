/**
 * Insertion-ordered LRU keyed by string. Values are typically surface
 * canvases (~4 bytes per pixel), so the entry cap bounds memory directly.
 */
export class DocxThumbnailSurfaceCache<T> {
  private readonly entries = new Map<string, T>();

  constructor(private readonly maxEntries: number) {}

  get size(): number {
    return this.entries.size;
  }

  get(key: string): T | undefined {
    const value = this.entries.get(key);
    if (value === undefined) {
      return undefined;
    }
    this.entries.delete(key);
    this.entries.set(key, value);
    return value;
  }

  set(key: string, value: T): void {
    this.entries.delete(key);
    this.entries.set(key, value);
    while (this.entries.size > this.maxEntries) {
      const oldestKey = this.entries.keys().next().value;
      if (oldestKey === undefined) {
        break;
      }
      this.entries.delete(oldestKey);
    }
  }

  clear(): void {
    this.entries.clear();
  }
}

interface SerialIdleTaskQueueEntry<K> {
  key: K;
  run: () => Promise<void>;
  resolvers: Array<() => void>;
  priority: number;
  sequence: number;
}

export interface SerialIdleTaskQueueOptions {
  /**
   * Schedules the next queue pump. Defaults to `requestIdleCallback` with a
   * timeout, falling back to a short `setTimeout`.
   */
  scheduleTask?: (callback: () => void) => void;
  /** Schedules a pump after a specific delay (throttle wake-ups). */
  scheduleDelayed?: (callback: () => void, delayMs: number) => void;
  /** Minimum interval between runs that share the same key. */
  minTaskIntervalMs?: number;
  now?: () => number;
}

export interface SerialIdleTaskQueueEnqueueOptions {
  /**
   * Lower values run first. Entries with the same priority keep FIFO order.
   *
   * @defaultValue `0`
   */
  priority?: number;
}

const IDLE_TASK_TIMEOUT_MS = 300;

function defaultScheduleTask(callback: () => void): void {
  const idleWindow =
    typeof window === "undefined"
      ? undefined
      : (window as Window & {
          requestIdleCallback?: (
            idleCallback: () => void,
            options?: { timeout?: number }
          ) => number;
          cancelIdleCallback?: (handle: number) => void;
        });
  if (!idleWindow || typeof idleWindow.requestIdleCallback !== "function") {
    setTimeout(callback, 16);
    return;
  }

  // Chrome suspends idle callbacks entirely while the document is hidden —
  // including ones with a timeout — which would starve the queue in
  // background tabs. Race the idle callback against a plain timer so the
  // queue always makes progress; whichever fires first wins.
  let invoked = false;
  const runOnce = (): void => {
    if (invoked) {
      return;
    }
    invoked = true;
    callback();
  };
  const idleHandle = idleWindow.requestIdleCallback(runOnce, {
    timeout: IDLE_TASK_TIMEOUT_MS,
  });
  setTimeout(() => {
    if (invoked) {
      return;
    }
    if (typeof idleWindow.cancelIdleCallback === "function") {
      idleWindow.cancelIdleCallback(idleHandle);
    }
    runOnce();
  }, IDLE_TASK_TIMEOUT_MS + 50);
}

function defaultScheduleDelayed(callback: () => void, delayMs: number): void {
  setTimeout(callback, delayMs);
}

/**
 * Runs async tasks strictly one at a time during idle periods. A newer task
 * with the same key replaces the queued one (its waiters resolve with the
 * newer run), and runs sharing a key are throttled to `minTaskIntervalMs`.
 */
export class SerialIdleTaskQueue<K> {
  private readonly pending: SerialIdleTaskQueueEntry<K>[] = [];
  private readonly lastRunAtByKey = new Map<K, number>();
  private readonly scheduleTask: (callback: () => void) => void;
  private readonly scheduleDelayed: (
    callback: () => void,
    delayMs: number
  ) => void;
  private readonly minTaskIntervalMs: number;
  private readonly now: () => number;
  private pumpScheduled = false;
  private running = false;
  private nextSequence = 0;

  constructor(options?: SerialIdleTaskQueueOptions) {
    this.scheduleTask = options?.scheduleTask ?? defaultScheduleTask;
    this.scheduleDelayed = options?.scheduleDelayed ?? defaultScheduleDelayed;
    this.minTaskIntervalMs = Math.max(0, options?.minTaskIntervalMs ?? 0);
    this.now = options?.now ?? (() => Date.now());
  }

  get pendingCount(): number {
    return this.pending.length;
  }

  enqueue(
    key: K,
    run: () => Promise<void>,
    options?: SerialIdleTaskQueueEnqueueOptions
  ): Promise<void> {
    const priority = Number.isFinite(options?.priority)
      ? Number(options?.priority)
      : 0;
    return new Promise<void>((resolve) => {
      const existing = this.pending.find((entry) => entry.key === key);
      if (existing) {
        existing.run = run;
        existing.resolvers.push(resolve);
        existing.priority = Math.min(existing.priority, priority);
      } else {
        this.pending.push({
          key,
          run,
          resolvers: [resolve],
          priority,
          sequence: this.nextSequence,
        });
        this.nextSequence += 1;
      }
      this.schedulePump();
    });
  }

  /** Drops queued work for a single key, resolving its waiters. */
  cancel(key: K): void {
    const remaining: SerialIdleTaskQueueEntry<K>[] = [];
    this.pending.forEach((entry) => {
      if (entry.key === key) {
        entry.resolvers.forEach((resolveEntry) => {
          resolveEntry();
        });
        return;
      }

      remaining.push(entry);
    });
    this.pending.splice(0, this.pending.length, ...remaining);
  }

  /** Drops all queued tasks, resolving their waiters without running them. */
  clear(): void {
    const dropped = this.pending.splice(0, this.pending.length);
    this.lastRunAtByKey.clear();
    dropped.forEach((entry) => {
      entry.resolvers.forEach((resolveEntry) => {
        resolveEntry();
      });
    });
  }

  private schedulePump(): void {
    if (this.pumpScheduled || this.running || this.pending.length === 0) {
      return;
    }
    this.pumpScheduled = true;
    this.scheduleTask(() => {
      this.pumpScheduled = false;
      void this.runNext();
    });
  }

  private takeNextEligibleEntry():
    | { entry: SerialIdleTaskQueueEntry<K> }
    | { retryDelayMs: number }
    | undefined {
    if (this.pending.length === 0) {
      return undefined;
    }

    const now = this.now();
    let earliestWaitMs: number | undefined;
    let bestIndex = -1;
    let bestEntry: SerialIdleTaskQueueEntry<K> | undefined;
    for (let index = 0; index < this.pending.length; index += 1) {
      const candidate = this.pending[index];
      if (!candidate) {
        continue;
      }
      const lastRunAt = this.lastRunAtByKey.get(candidate.key);
      const waitMs =
        lastRunAt === undefined
          ? 0
          : lastRunAt + this.minTaskIntervalMs - now;
      if (waitMs <= 0) {
        if (
          !bestEntry ||
          candidate.priority < bestEntry.priority ||
          (candidate.priority === bestEntry.priority &&
            candidate.sequence < bestEntry.sequence)
        ) {
          bestEntry = candidate;
          bestIndex = index;
        }
        continue;
      }
      earliestWaitMs =
        earliestWaitMs === undefined
          ? waitMs
          : Math.min(earliestWaitMs, waitMs);
    }

    if (bestEntry && bestIndex >= 0) {
      this.pending.splice(bestIndex, 1);
      return { entry: bestEntry };
    }

    return earliestWaitMs === undefined
      ? undefined
      : { retryDelayMs: earliestWaitMs };
  }

  private async runNext(): Promise<void> {
    if (this.running) {
      return;
    }

    const next = this.takeNextEligibleEntry();
    if (!next) {
      return;
    }
    if (!("entry" in next)) {
      this.scheduleDelayed(() => {
        this.schedulePump();
      }, next.retryDelayMs);
      return;
    }

    this.running = true;
    const { entry } = next;
    try {
      await entry.run();
    } catch {
      // Task bodies report their own failures; the queue only sequences them.
    } finally {
      this.lastRunAtByKey.set(entry.key, this.now());
      this.running = false;
      entry.resolvers.forEach((resolveEntry) => {
        resolveEntry();
      });
      this.schedulePump();
    }
  }
}