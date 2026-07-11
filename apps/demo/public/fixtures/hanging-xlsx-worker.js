self.addEventListener("message", () => {
  // This formal consumer fixture deliberately leaves the request pending.
  // The host configures it through the public createWorker option so timeout
  // and cancellation can be verified without a hidden test-only code path.
});
