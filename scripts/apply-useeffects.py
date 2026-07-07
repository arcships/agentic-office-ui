p = "packages/vue-xlsx/src/composables.ts"
s = open(p).read()
s = s.replace(
"""  /* @vue-useEffect-manual */ React.useEffect(() => () => {
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    revokeWorkbookImageAssets(imageAssetsRef.value);
    disposeWorkerClient();
  }, [disposeWorkerClient]);""",
"""  onUnmounted(() => {
    chartDisplayFallbackCleanupRef.value?.();
    chartDisplayFallbackCleanupRef.value = null;
    revokeWorkbookImageAssets(imageAssetsRef.value);
    disposeWorkerClient();
  });""")
s = s.replace(
"""  /* @vue-useEffect-manual */ React.useEffect(() => {
    if (!file && !src) {""",
"""  watch(() => [file, src], (_value, _oldValue, onCleanup) => {
    if (!file && !src) {""")
s = s.replace(
"""    return () => {
      isCurrent = false;
      abortController.abort();
      disposeWorkerClient();
    };
  }, [
    clearChartAssets,
    clearImageAssets,
    deferLoadingAboveBytes,
    disposeWorkerClient,
    file,
    getWorkerClient,
    hasIncompleteWorkerChartSnapshot,
    loadWorkbookOnMainThread,
    maxFileSizeBytes,
    setImageAssets,
    startChartDisplayHydration,
    shouldFallbackFromWorkerError,
    shouldDeferLoading,
    shouldForceReadOnlyForBuffer,
    shouldUseWorkerForReadOnlyLoad,
    src,
    showHiddenSheets
  ]);""",
"""    onCleanup(() => {
      isCurrent = false;
      abortController.abort();
      disposeWorkerClient();
    });
  }, { immediate: true });""")
s = s.replace(
"""  /* @vue-useEffect-manual */ React.useEffect(() => {
    activeCell.value = null;
    selection.value = null;
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
    selectionAnchorRef.value = null;
    sortState.value = null;
  }, [activeTabIndex.value]);""",
"""  watch(() => activeTabIndex.value, () => {
    activeCell.value = null;
    selection.value = null;
    selectedChartId.value = null;
    selectedChartElement.value = null;
    selectedImageId.value = null;
    selectionAnchorRef.value = null;
    sortState.value = null;
  });""")
s = s.replace(
"""  /* @vue-useEffect-manual */ React.useEffect(() => {
    activeTabIndex.value = (() => {
      if (tabs.value.length === 0) {
        return 0;
      }
      return Math.min(activeTabIndex.value, tabs.value.length - 1);
    })();
  }, [tabs.value.length]);""",
"""  watch(() => tabs.value.length, () => {
    activeTabIndex.value = (() => {
      if (tabs.value.length === 0) {
        return 0;
      }
      return Math.min(activeTabIndex.value, tabs.value.length - 1);
    })();
  });""")
s = s.replace(
"""  /* @vue-useEffect-manual */ React.useEffect(() => {
    if (!selectedChartId.value) {
      if (selectedChartElement.value) {
        selectedChartElement.value = null;
      }
      return;
    }

    if (!selectedChart.value) {
      selectedChartId.value = null;
      selectedChartElement.value = null;
      return;
    }

    if (!selectedChartElement.value) {
      selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
      return;
    }

    if (selectedChartElement.value.chartId !== selectedChartId.value) {
      selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
      return;
    }

    if (selectedChartElement.value.kind !== "chart") {
      const selectedSeries = selectedChart.value.series[selectedChartElement.value.seriesIndex];
      if (!selectedSeries || selectedSeries.id !== selectedChartElement.value.seriesId) {
        selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
      }
    }
  }, [selectedChart.value, selectedChartElement.value, selectedChartId.value]);""",
"""  watch(
    () => [selectedChart.value, selectedChartElement.value, selectedChartId.value],
    () => {
      if (!selectedChartId.value) {
        if (selectedChartElement.value) {
          selectedChartElement.value = null;
        }
        return;
      }

      if (!selectedChart.value) {
        selectedChartId.value = null;
        selectedChartElement.value = null;
        return;
      }

      if (!selectedChartElement.value) {
        selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
        return;
      }

      if (selectedChartElement.value.chartId !== selectedChartId.value) {
        selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
        return;
      }

      if (selectedChartElement.value.kind !== "chart") {
        const selectedSeries = selectedChart.value.series[selectedChartElement.value.seriesIndex];
        if (!selectedSeries || selectedSeries.id !== selectedChartElement.value.seriesId) {
          selectedChartElement.value = { chartId: selectedChartId.value, kind: "chart" };
        }
      }
    }
  );""")
s = s.replace(
"""  /* @vue-useEffect-manual */ React.useEffect(() => {
    if (!isWorkerBacked.value || !deferredMetadataSheet.value || !deferredMetadataCell.value) {
      return;
    }

    const cacheKey = `${deferredMetadataSheet.value.workbookSheetIndex}:${deferredMetadataCell.value.row}:${deferredMetadataCell.value.col}`;
    if (workerCellSnapshotCacheRef.value.has(cacheKey)) {
      return;
    }

    let isCurrent = true;
    void getCellSnapshotAsync(deferredMetadataSheet.value.workbookSheetIndex, deferredMetadataCell.value.row, deferredMetadataCell.value.col)
      .then((snapshot) => {
        if (!isCurrent) {
          return;
        }

        workerCellSnapshotCacheRef.value.set(cacheKey, snapshot);
        workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
      })
      .catch(() => {
        if (!isCurrent) {
          return;
        }

        workerCellSnapshotCacheRef.value.set(cacheKey, {
          displayValue: "",
          formula: ""
        });
        workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
      });

    return () => {
      isCurrent = false;
    };
  }, [deferredMetadataCell.value, deferredMetadataSheet.value, getCellSnapshotAsync, isWorkerBacked.value]);""",
"""  watch(
    () => [deferredMetadataCell.value, deferredMetadataSheet.value, isWorkerBacked.value],
    (_value, _oldValue, onCleanup) => {
      if (!isWorkerBacked.value || !deferredMetadataSheet.value || !deferredMetadataCell.value) {
        return;
      }

      const cacheKey = `${deferredMetadataSheet.value.workbookSheetIndex}:${deferredMetadataCell.value.row}:${deferredMetadataCell.value.col}`;
      if (workerCellSnapshotCacheRef.value.has(cacheKey)) {
        return;
      }

      let isCurrent = true;
      void getCellSnapshotAsync(deferredMetadataSheet.value.workbookSheetIndex, deferredMetadataCell.value.row, deferredMetadataCell.value.col)
        .then((snapshot) => {
          if (!isCurrent) {
            return;
          }

          workerCellSnapshotCacheRef.value.set(cacheKey, snapshot);
          workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
        })
        .catch(() => {
          if (!isCurrent) {
            return;
          }

          workerCellSnapshotCacheRef.value.set(cacheKey, {
            displayValue: "",
            formula: ""
          });
          workerCellSnapshotRevision.value = workerCellSnapshotRevision.value + 1;
        });

      onCleanup(() => {
        isCurrent = false;
      });
    }
  );""")
open(p, "w").write(s)
print("useEffect conversions applied")
