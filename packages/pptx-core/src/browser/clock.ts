export interface PptxClockDriver {
  now(): number
  requestFrame(callback: (now: number) => void): number
  cancelFrame(handle: number): void
}

export interface PptxClock {
  readonly positionMs: number
  readonly running: boolean
  play(): void
  pause(): void
  seek(positionMs: number): void
  subscribe(listener: (positionMs: number) => void): () => void
  dispose(): void
}

function browserDriver(): PptxClockDriver {
  return {
    now: () => performance.now(),
    requestFrame: (callback) => requestAnimationFrame(callback),
    cancelFrame: (handle) => cancelAnimationFrame(handle),
  }
}

export function createPptxClock(driver: PptxClockDriver = browserDriver()): PptxClock {
  let positionMs = 0
  let running = false
  let disposed = false
  let previousNow = 0
  let frameHandle: number | null = null
  const listeners = new Set<(positionMs: number) => void>()

  const assertUsable = () => {
    if (disposed) throw new Error("PPTX 时钟已经销毁。")
  }
  const notify = () => {
    for (const listener of [...listeners]) listener(positionMs)
  }
  const advance = (now: number) => {
    positionMs += Math.max(0, now - previousNow)
    previousNow = now
    notify()
  }
  const schedule = () => {
    frameHandle = driver.requestFrame((now) => {
      frameHandle = null
      if (!running || disposed) return
      advance(now)
      schedule()
    })
  }

  return {
    get positionMs() {
      return positionMs
    },
    get running() {
      return running
    },
    play() {
      assertUsable()
      if (running) return
      running = true
      previousNow = driver.now()
      schedule()
    },
    pause() {
      assertUsable()
      if (!running) return
      advance(driver.now())
      running = false
      if (frameHandle !== null) driver.cancelFrame(frameHandle)
      frameHandle = null
    },
    seek(value) {
      assertUsable()
      positionMs = Number.isFinite(value) ? Math.max(0, value) : 0
      if (running) previousNow = driver.now()
      notify()
    },
    subscribe(listener) {
      assertUsable()
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    dispose() {
      if (disposed) return
      disposed = true
      running = false
      if (frameHandle !== null) driver.cancelFrame(frameHandle)
      frameHandle = null
      listeners.clear()
    },
  }
}

