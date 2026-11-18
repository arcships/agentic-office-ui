/**
 * Shared user-facing message catalog for the office viewer packages.
 *
 * The catalog keeps English as the default locale and ships a Simplified
 * Chinese (`zh-CN`) translation. Hosts can override any message through the
 * `messages` prop exposed by each viewer.
 *
 * The module is framework-neutral: the component packages provide/inject the
 * resolved catalog themselves using {@link OFFICE_I18N_INJECTION_KEY}.
 */

export type OfficeBuiltInLocale = "en" | "zh-CN"

export type OfficeLocale = OfficeBuiltInLocale | (string & {})

type DeepPartial<T> = {
  [K in keyof T]?: T[K] extends string ? string : DeepPartial<T[K]>
}

export type OfficeMessagesOverride = DeepPartial<OfficeMessages>

export interface OfficeSearchMessages {
  placeholder: string
  previous: string
  next: string
  close: string
  searching: string
  error: string
  noResults: string
  /** Template with `{current}` and `{total}` placeholders. */
  resultOf: string
}

export interface OfficeZoomMessages {
  in: string
  out: string
  level: string
  reset: string
}

export interface OfficePageMessages {
  label: string
  current: string
  navigation: string
  previous: string
  next: string
  first: string
  last: string
  of: string
  /** Template with `{current}` and `{total}` placeholders. */
  indicator: string
}

export interface OfficeCommonMessages {
  loading: string
  readOnly: string
  selectedText: string
  search: OfficeSearchMessages
  zoom: OfficeZoomMessages
  page: OfficePageMessages
  reference: {
    /** Template with a `{format}` placeholder. */
    confirmFailed: string
    /** Template with a `{format}` placeholder. */
    createFailed: string
    /** Template with a `{format}` placeholder. */
    captureUnsupported: string
    /** Template with a `{format}` placeholder. */
    describeAborted: string
    /** Template with a `{format}` placeholder. */
    hitTestFailed: string
  }
  descriptors: {
    document: string
    /** Template with a `{number}` placeholder. */
    page: string
    image: string
    /** Template with a `{number}` placeholder. */
    table: string
    /** Template with a `{number}` placeholder. */
    row: string
    /** Template with a `{number}` placeholder. */
    cell: string
    /** Template with a `{number}` placeholder. */
    paragraph: string
    /** Template with a `{column}` placeholder. */
    column: string
    workbook: string
    presentation: string
    /** Template with a `{number}` placeholder. */
    slide: string
    pdf: string
    textRange: string
    selectedRegion: string
    /** Template with a `{number}` placeholder. */
    regionOnPage: string
    /** Template with a `{name}` placeholder. */
    regionOnSheet: string
    /** Template with a `{number}` placeholder. */
    regionOnSlide: string
    comment: string
    /** Template with a `{id}` placeholder. */
    commentById: string
    /** Template with a `{author}` placeholder. */
    commentBy: string
    /** Template with a `{text}` placeholder. */
    commentText: string
    chart: string
    /** Template with a `{name}` placeholder. */
    chartName: string
    /** Template with a `{number}` placeholder. */
    group: string
    /** Template with a `{id}` placeholder. */
    shape: string
    /** Template with a `{id}` placeholder. */
    object: string
    /** Template with a `{id}` placeholder. */
    textBox: string
    /** Template with a `{id}` placeholder. */
    imageId: string
    /** Template with a `{text}` placeholder. */
    imageText: string
    /** Template with a `{text}` placeholder. */
    textLabel: string
  }
}

export interface OfficeDocxMessages {
  viewer: {
    loadingTitle: string
    loadingDetail: string
    errorTitle: string
    chooseAnother: string
    emptyTitle: string
    emptyDetail: string
    chooseFile: string
    parseError: string
    unsupportedFile: string
    /** Template with `{actual}` and `{allowed}` placeholders. */
    fileTooLarge: string
    live: {
      loading: string
      error: string
      ready: string
      /** Template with `{page}` and `{total}` placeholders. */
      readyPages: string
    }
    /** Template with `{count}` and `{allowed}` placeholders. */
    pageLimitExceeded: string
  }
  toolbar: {
    sidebarToggle: string
    pageThumbnails: string
    searchPlaceholder: string
    trackedChanges: string
    comments: string
    /** Template with a `{count}` placeholder. */
    trackedChangesShow: string
    /** Template with a `{count}` placeholder. */
    trackedChangesHide: string
    trackedChangesNone: string
    /** Template with a `{count}` placeholder. */
    commentsShow: string
    /** Template with a `{count}` placeholder. */
    commentsHide: string
    commentsNone: string
    lightThemeAction: string
    darkThemeAction: string
    lightDocument: string
    darkDocument: string
    upload: string
    download: string
  }
  editor: {
    new: string
    newTitle: string
    newDocument: string
    importedDocument: string
    import: string
    importTitle: string
    export: string
    exportTitle: string
    edits: string
    toggleTheme: string
    undo: string
    redo: string
    fontFamily: string
    font: string
    fontSize: string
    size: string
    bold: string
    italic: string
    underline: string
    strikethrough: string
    superscript: string
    subscript: string
    textColor: string
    highlightColor: string
    placeholder: string
    link: string
    alignLeft: string
    alignCenter: string
    alignRight: string
    justify: string
    bulletList: string
    numberedList: string
    increaseIndent: string
    increaseIndentTitle: string
    decreaseIndent: string
    decreaseIndentTitle: string
    lineSpacing: string
    paragraphStyle: string
    normal: string
    borders: string
    insertImage: string
    insertTable: string
    zoom: string
    /** Template with a `{url}` placeholder. */
    removeLink: string
    unsupportedLink: string
    enterUrl: string
    liveError: string
    liveReady: string
    acceptAll: string
    acceptAllTitle: string
    rejectAll: string
    rejectAllTitle: string
    status: {
      ready: string
      copied: string
      pasted: string
      listParagraphNotFound: string
      listAppliedBullet: string
      listAppliedNumbered: string
      /** Template with a `{level}` placeholder. */
      maximumListLevel: string
      discardChanges: string
      /** Template with a `{name}` placeholder. */
      loadingFile: string
      /** Template with a `{message}` placeholder. */
      importFailed: string
      unsupportedFile: string
      exportCancelledChanged: string
      exportNotReady: string
      /** Template with a `{message}` placeholder. */
      exportFailed: string
      /** Template with a `{label}` placeholder. */
      exportRequiredField: string
      exportRequiredFieldEmpty: string
      exporting: string
      /** Template with a `{phase}` placeholder. */
      exportingPhase: string
      exported: string
      unknownImportError: string
      unknownExportError: string
      /** Template with a `{count}` placeholder. */
      acceptedChanges: string
      acceptedChange: string
      /** Template with a `{count}` placeholder. */
      rejectedChanges: string
      rejectedChange: string
      /** Template with a `{count}` placeholder. */
      formattingCleared: string
      formattingClearedOne: string
      /** Template with a `{feature}` placeholder. */
      unsupportedFeature: string
      imageTargetUnavailable: string
      imageSizePositive: string
      resizeTargetUnavailable: string
      lockedImageAnchorResize: string
      imageOverlapDisabled: string
      syntheticTextBoxUnavailable: string
      imageMoveTargetUnavailable: string
      lockedImageAnchorMove: string
      sectionImageTargetUnavailable: string
      dropCapTargetUnavailable: string
      dropCapFontSizeTargetUnavailable: string
      dropCapTextTargetUnavailable: string
      imageMoveSourceUnavailable: string
      selectImageFile: string
      imageInsertionCancelled: string
      /** Template with a `{name}` placeholder. */
      insertedImage: string
      /** Template with `{width}` and `{height}` placeholders. */
      resizedImage: string
      updatedTextBox: string
      /** Template with a `{mode}` placeholder. */
      imageWrapSet: string
      movedImage: string
      movedSectionImage: string
      movedDropCap: string
      /** Template with a `{size}` placeholder. */
      dropCapFontSize: string
      updatedDropCapText: string
      tableColumnWidthPositive: string
      tableColumnWidthTargetUnavailable: string
      cannotInsertRowTableNotFound: string
      cannotInsertTableRow: string
      cannotInsertColumnTableNotFound: string
      cannotInsertTableColumn: string
      cannotDeleteRowTableNotFound: string
      cannotDeleteOnlyTableRow: string
      cannotDeleteColumnTableNotFound: string
      cannotDeleteOnlyTableColumn: string
      cannotDeleteTableColumn: string
      cannotDeleteTableNotFound: string
      cannotClearCellsTableNotFound: string
      cannotMergeCellsTableNotFound: string
      cannotMergeSingleCell: string
      cannotSplitCellTableNotFound: string
      cannotSplitCellNotMerged: string
      cannotMoveTableNotFound: string
      cannotMoveEmbeddedHostTableNotFound: string
      cannotMoveEmbeddedHostCellNotFound: string
      cannotMoveEmbeddedTableNotFound: string
      insertedTable: string
      /** Template with a `{direction}` placeholder. */
      insertedTableRow: string
      /** Template with a `{direction}` placeholder. */
      insertedTableColumn: string
      deletedTableRow: string
      deletedTableColumn: string
      deletedTable: string
      clearedTableCells: string
      resizedTableColumns: string
      mergedTableCells: string
      splitTableCell: string
      movedTable: string
      movedEmbeddedTable: string
      /** Template with a `{preset}` placeholder. */
      unsupportedDiagonalTableCellBorders: string
      /** Template with a `{preset}` placeholder. */
      unsupportedDiagonalParagraphBorders: string
      revisionLossOne: string
      /** Template with a `{count}` placeholder. */
      revisionLossMany: string
    }
    borderPresets: {
      bottom: string
      top: string
      left: string
      right: string
      none: string
      all: string
      outside: string
      inside: string
      insideHorizontal: string
      insideVertical: string
      diagonalDown: string
      diagonalUp: string
      horizontalLine: string
    }
  }
  contextMenu: {
    label: string
    cut: string
    copy: string
    paste: string
    bold: string
    italic: string
    underline: string
    strikethrough: string
    highlight: string
    noHighlight: string
    textColor: string
    splitParagraph: string
    deleteParagraph: string
    table: string
    insertRowAbove: string
    insertRowBelow: string
    insertColumnLeft: string
    insertColumnRight: string
    deleteRow: string
    deleteColumn: string
    deleteTable: string
    mergeCells: string
    splitCell: string
  }
  colors: {
    black: string
    red: string
    orange: string
    yellow: string
    green: string
    cyan: string
    blue: string
    magenta: string
    purple: string
  }
  thumbnail: {
    show: string
    hide: string
    pages: string
    unavailable: string
    unavailableReason: string
  }
  image: {
    inline: string
    square: string
    behind: string
    front: string
    inlineTitle: string
    squareTitle: string
    behindTitle: string
    frontTitle: string
    alt: string
    /** Text shown for an image node without a source. */
    missing: string
    wrapInline: string
    wrapSquare: string
    wrapTight: string
    wrapThrough: string
    wrapTopAndBottom: string
    wrapBehind: string
    wrapFront: string
    moveWithText: string
    fixPositionOnPage: string
    editWrapBoundary: string
    moreLayoutOptions: string
  }
  formField: {
    placeholder: string
    selectDate: string
    errors: {
      invalidValue: string
      required: string
      noOptions: string
      chooseValue: string
      invalidDate: string
      /** Template with a `{max}` placeholder. */
      maxLength: string
    }
  }
  notes: {
    footnote: string
    endnote: string
  }
  documentSurface: {
    region: string
  }
  trackedChange: {
    resolved: string
    inserted: string
    deleted: string
    movedFrom: string
    movedTo: string
    format: string
    paragraphFormatting: string
    tableFormatting: string
    rowFormatting: string
    cellFormatting: string
    comment: string
    annotation: string
  }
  gutter: {
    accept: string
    acceptTitle: string
    reject: string
    rejectTitle: string
    hideComments: string
    showComments: string
    comments: string
    /** Template with a `{ordinal}` placeholder. */
    commentMarker: string
    /** Template with `{ordinal}` and `{author}` placeholders. */
    commentMarkerWithAuthor: string
  }
  dragOverlay: {
    drop: string
  }
  diagnostics: {
    imageDecodeFailed: string
    numberingImageDecodeFailed: string
  }
}

export interface OfficeXlsxMessages {
  viewer: {
    loading: string
    empty: string
    deferredUntitled: string
    deferredTitle: string
    /** Template with a `{size}` placeholder. */
    deferredDetail: string
    resumeParsing: string
    cancel: string
    live: {
      loading: string
      error: string
      sheetReady: string
      chartsheetReady: string
      empty: string
    }
  }
  toolbar: {
    zoomGroup: string
    search: string
    searchPlaceholder: string
    upload: string
    download: string
    export: string
  }
  ribbon: {
    undo: string
    redo: string
    cut: string
    copy: string
    paste: string
    pasteSpecial: string
    insertRowAbove: string
    insertRowBelow: string
    deleteRow: string
    insertColumnLeft: string
    insertColumnRight: string
    deleteColumn: string
    numberFormatDialog: string
    nameManager: string
    fontFamily: string
    fontSize: string
    bold: string
    italic: string
    underline: string
    strikethrough: string
    textColor: string
    fillColor: string
    alignLeft: string
    alignCenter: string
    alignRight: string
    alignTop: string
    alignMiddle: string
    alignBottom: string
    wrapText: string
    numberFormat: string
    currency: string
    percent: string
    borders: string
    allBorders: string
    outsideBorders: string
    leftBorder: string
    rightBorder: string
    topBorder: string
    bottomBorder: string
    noBorders: string
    borderColor: string
    merge: string
    mergeAndCenter: string
    unmerge: string
    readOnlyTitle: string
    numberFormats: {
      general: string
      number: string
      currency: string
      percent: string
      date: string
      text: string
    }
  }
  sheetTabs: {
    previous: string
    next: string
    add: string
    remove: string
    defaultName: string
  }
  contextMenu: {
    label: string
    copy: string
    cut: string
    paste: string
    pasteSpecial: string
    pasteAll: string
    pasteValues: string
    pasteFormats: string
    pasteTranspose: string
    editCell: string
    clear: string
    clearContents: string
    sortAscending: string
    sortDescending: string
    undo: string
    redo: string
    changeChartType: string
    insertRowAbove: string
    insertRowBelow: string
    deleteRow: string
    insertColumnLeft: string
    insertColumnRight: string
    deleteColumn: string
    mergeAndCenter: string
    mergeAcross: string
    mergeCells: string
    unmerge: string
    selectRow: string
    selectColumn: string
    selectAll: string
  }
  formulaBar: {
    activeCell: string
    placeholder: string
  }
  chart: {
    unsupported: string
    loading: string
    loadFailed: string
    unavailable: string
    noPayload: string
    ariaLabel: string
    /** Template with a `{type}` placeholder. */
    unsupportedType: string
    other: string
    /** Template with a `{number}` placeholder. */
    seriesName: string
    /** Template with a `{number}` placeholder. */
    surfaceSeries: string
    root: string
  }
  image: {
    alt: string
  }
  table: {
    /** Template with a `{number}` placeholder. */
    defaultName: string
    /** Template with a `{number}` placeholder. */
    defaultColumn: string
  }
  authoring: {
    close: string
    /** Template with a `{range}` placeholder. */
    rangeLabel: string
    chart: {
      title: string
      insertMode: string
      editMode: string
      family: string
      familyBar: string
      familyLine: string
      familyPie: string
      chartTitle: string
      chartTitlePlaceholder: string
      showDataLabels: string
      headerRow: string
      headerColumn: string
      insert: string
      update: string
      delete: string
      /** Template with a `{title}` placeholder. */
      selectedChart: string
      noneSelected: string
    }
    comment: {
      title: string
      text: string
      textPlaceholder: string
      author: string
      authorPlaceholder: string
      save: string
      delete: string
    }
    hyperlink: {
      title: string
      target: string
      targetPlaceholder: string
      tooltip: string
      tooltipPlaceholder: string
      save: string
      remove: string
    }
    table: {
      title: string
      name: string
      namePlaceholder: string
      hasHeaders: string
      create: string
      remove: string
      /** Template with a `{name}` and an `{range}` placeholder. */
      existing: string
    }
    conditional: {
      title: string
      kind: string
      kindColorScale: string
      kindDataBar: string
      kindIconSet: string
      thresholds: string
      typeMin: string
      typeMax: string
      typePercentile: string
      typeNumber: string
      typePercent: string
      value: string
      color: string
      iconSet: string
      showValue: string
      reverse: string
      add: string
      update: string
      delete: string
      edit: string
      empty: string
      /** Template with a `{kind}` and an `{ranges}` placeholder. */
      ruleSummary: string
    }
    validation: {
      title: string
      type: string
      typeList: string
      typeWhole: string
      typeDecimal: string
      typeDate: string
      typeTextLength: string
      source: string
      sourcePlaceholder: string
      value1: string
      value2: string
      allowBlank: string
      showDropdown: string
      inputMessage: string
      errorMessage: string
      apply: string
      remove: string
      /** Template with a `{type}` and an `{ranges}` placeholder. */
      existing: string
    }
    freeze: {
      freezeHere: string
      unfreeze: string
    }
    pasteSpecial: {
      title: string
      mode: string
      modeAll: string
      modeValues: string
      modeFormats: string
      transpose: string
      paste: string
      empty: string
    }
    numberFormat: {
      title: string
      category: string
      number: string
      currency: string
      percent: string
      date: string
      text: string
      custom: string
      decimals: string
      thousands: string
      pattern: string
      preview: string
      apply: string
    }
    names: {
      title: string
      name: string
      namePlaceholder: string
      refersTo: string
      refersToPlaceholder: string
      add: string
      update: string
      delete: string
      useSelection: string
      empty: string
    }
    context: {
      insertComment: string
      editComment: string
      insertHyperlink: string
      editHyperlink: string
      conditional: string
      createTable: string
      removeTable: string
      insertChart: string
      validation: string
      freezeHere: string
      unfreeze: string
    }
    status: {
      chartInserted: string
      chartUpdated: string
      chartDeleted: string
      commentSaved: string
      commentDeleted: string
      hyperlinkSaved: string
      hyperlinkRemoved: string
      tableCreated: string
      tableRemoved: string
      ruleAdded: string
      ruleUpdated: string
      ruleRemoved: string
      validationSet: string
      validationRemoved: string
      actionFailed: string
    }
  }
  drawing: {
    button: string
    dropdown: string
    editbox: string
    listbox: string
    spinner: string
    scrollbar: string
    label: string
    groupBox: string
    unknown: string
  }
  selection: {
    /** Template with an `{address}` placeholder. */
    current: string
    none: string
  }
  validation: {
    list: string
    whole: string
    decimal: string
    textLength: string
    date: string
    time: string
    generic: string
  }
  calculation: {
    sheetMissing: string
    snapshotUnavailable: string
    failed: string
  }
  clipboard: {
    copyFailed: string
    noDocument: string
    invalidUrl: string
    unverifiedBlob: string
  }
  grid: {
    openValidation: string
    closeValidation: string
    worksheet: string
    /** Template with a `{name}` placeholder. */
    worksheetNamed: string
    blank: string
  }
  surface: {
    region: string
  }
  diagnostics: {
    imageDecodeFailed: string
    chartDataIncomplete: string
    /** Template with a `{ms}` placeholder. */
    parseTimeout: string
    exportXmlEscapeFailed: string
    /** Template with a `{name}` placeholder. */
    macroLoss: string
    csvEncoding: string
    csvParseFailed: string
    xlsbParseFailed: string
    fileOrSrcRequired: string
    sheetNameInvalid: string
  }
  errors: {
    sourceNotAllowed: string
    fetchFailed: string
    invalidWorkbook: string
    limitExceeded: string
    imageLimitExceeded: string
    invalidImage: string
    imageDecodeFailed: string
    timeout: string
    workerUnavailable: string
    aborted: string
    unknown: string
  }
}

export interface OfficePptxMessages {
  viewer: {
    previousSlide: string
    nextSlide: string
    hiddenSlide: string
    searchPlaceholder: string
    search: string
    sidebar: string
    empty: string
    loading: string
    waiting: string
    cancelled: string
    reload: string
    retry: string
    skip: string
    slideRenderFailed: string
    previousStep: string
    nextStep: string
    replay: string
    play: string
    pause: string
    resumeMedia: string
    resumeMediaAction: string
    fullscreen: string
    capability: {
      /** Template with `{count}` and `{total}` placeholders. */
      strict: string
      /** Template with a `{count}` placeholder. */
      approximate: string
      /** Template with a `{count}` placeholder. */
      static: string
      /** Template with a `{count}` placeholder. */
      unparsed: string
    }
    status: {
      idle: string
      ready: string
      transitioning: string
      running: string
      waiting: string
      paused: string
      blocked: string
      ended: string
      disposed: string
    }
    error: {
      controllerNotReady: string
      operationFailed: string
      previewNotReady: string
      fullscreenUnsupported: string
      fullscreenEnterRejected: string
      fullscreenExitRejected: string
      documentOperationFailed: string
    }
    live: {
      empty: string
      loading: string
      error: string
      ready: string
      hiddenSuffix: string
      playbackSuffix: string
      unknownError: string
    }
  }
  documentErrors: {
    loadFailed: string
    previewNotReady: string
    disposed: string
    staleOpen: string
    timeout: string
    /** Template with an `{ms}` placeholder. */
    timeoutMs: string
    staleResult: string
    aborted: string
    slideRenderFailedWithMessage: string
    slideRenderFailed: string
  }
  thumbnail: {
    /** Template with a `{number}` placeholder. */
    label: string
    /** Template with a `{number}` placeholder. */
    hiddenLabel: string
    hidden: string
  }
  stage: {
    region: string
  }
}

export interface OfficePdfMessages {
  viewer: {
    noPdf: string
    opening: string
    toggleThumbnails: string
    pageInput: string
    rotateClockwise: string
    searchLabel: string
    searchPlaceholder: string
    search: string
    download: string
    thumbnails: string
    /** Template with a `{page}` placeholder. */
    thumbnailAlt: string
    /** Template with a `{page}` placeholder. */
    pageLabel: string
    /** Template with a `{page}` placeholder. */
    renderingPage: string
    /** Template with a `{page}` placeholder. */
    pageAlt: string
    searchSearching: string
    searchNoResults: string
    searchError: string
    /** Template with `{current}`, `{total}` and `{page}` placeholders. */
    searchResult: string
    renderFailed: string
    invalidPdf: string
    selectionLayer: string
  }
  surface: {
    noPdf: string
    opening: string
    /** Template with a `{page}` placeholder. */
    renderingPage: string
    /** Template with a `{page}` placeholder. */
    pageRenderFailed: string
    /** Template with a `{page}` placeholder. */
    pageAlt: string
    selectionRegion: string
    selectionText: string
  }
  errors: {
    sourceNotAllowed: string
    fetchFailed: string
    invalidPdf: string
    /** Template with `{actual}` and `{allowed}` placeholders. */
    tooLarge: string
    aborted: string
    encrypted: string
    openFailed: string
    unknown: string
  }
  runtime: {
    /** Template with a `{label}` placeholder. */
    cancelled: string
    /** Template with a `{label}` placeholder. */
    timedOut: string
    workerStartup: string
    textSearch: string
    textExtraction: string
    documentClose: string
    runtimeDestroy: string
  }
}

export interface OfficeMessages {
  common: OfficeCommonMessages
  docx: OfficeDocxMessages
  xlsx: OfficeXlsxMessages
  pptx: OfficePptxMessages
  pdf: OfficePdfMessages
  vueUi: {
    regionSelector: {
      label: string
      instructions: string
      defaultAriaLabel: string
      /** Template with `{left}`, `{top}`, `{width}` and `{height}` placeholders. */
      description: string
      noSelection: string
    }
    objectOutline: {
      defaultAriaLabel: string
      noObjects: string
      noCandidates: string
      /** Template with `{label}`, `{kind}` and `{reliability}` placeholders. */
      itemDescription: string
    }
    boundingBox: {
      noFields: string
      imageAlt: string
    }
    fileUpload: {
      prompt: string
    }
    layoutBlocks: {
      noBlocks: string
      imageAlt: string
    }
    signaturePad: {
      clear: string
    }
    fileThumbnail: {
      unknownType: string
    }
  }
}

/** Injection key shared by the viewer packages and the shared UI package. */
export const OFFICE_I18N_INJECTION_KEY: unique symbol = Symbol.for(
  "@arcships/office-ui.i18n",
)

/** Injection key for the resolved built-in locale (`en` or `zh-CN`). */
export const OFFICE_LOCALE_INJECTION_KEY: unique symbol = Symbol.for(
  "@arcships/office-ui.locale",
)

export const officeMessagesEn: OfficeMessages = {
  common: {
    loading: "Loading…",
    readOnly: "Read only",
    selectedText: "Selected text",
    search: {
      placeholder: "Find",
      previous: "Previous result",
      next: "Next result",
      close: "Close search",
      searching: "Searching…",
      error: "Search failed",
      noResults: "No results",
      resultOf: "{current} of {total}",
    },
    zoom: {
      in: "Zoom in",
      out: "Zoom out",
      level: "Zoom level",
      reset: "Fit to window",
    },
    page: {
      label: "Page",
      current: "Current page",
      navigation: "Page navigation",
      previous: "Previous page",
      next: "Next page",
      first: "First page",
      last: "Last page",
      of: "of",
      indicator: "Page {current} of {total}",
    },
    reference: {
      confirmFailed: "Unable to confirm {format} reference.",
      createFailed: "Unable to create {format} text reference.",
      captureUnsupported: "{format} preview capture is not provided by the Surface; use the host's capture pipeline.",
      describeAborted: "{format} reference description aborted.",
      hitTestFailed: "{format} hit test failed.",
    },
    descriptors: {
      document: "Document",
      page: "Page {number}",
      image: "Image",
      table: "Table {number}",
      row: "Row {number}",
      cell: "Cell {number}",
      paragraph: "Paragraph {number}",
      column: "Column {column}",
      workbook: "Workbook",
      presentation: "Presentation",
      slide: "Slide {number}",
      pdf: "PDF",
      textRange: "Text range",
      selectedRegion: "Selected region",
      regionOnPage: "Region on page {number}",
      regionOnSheet: "Region on {name}",
      regionOnSlide: "Region on slide {number}",
      comment: "Comment",
      commentById: "Comment {id}",
      commentBy: "Comment by {author}",
      commentText: "Comment: {text}",
      chart: "Chart",
      chartName: "Chart {name}",
      group: "Group {number}",
      shape: "Shape {id}",
      object: "Object {id}",
      textBox: "Text box {id}",
      imageId: "Image {id}",
      imageText: "Image: {text}",
      textLabel: "Text: {text}",
    },
  },
  docx: {
    viewer: {
      loadingTitle: "Opening document",
      loadingDetail: "Loading DOCX…",
      errorTitle: "Unable to open this document",
      chooseAnother: "Choose another Word file",
      emptyTitle: "No document open",
      emptyDetail: "Upload or provide a DOCX file to start.",
      chooseFile: "Choose Word file",
      parseError: "Unknown DOCX parse error",
      unsupportedFile: "Only .docx, .docm, .dotx and .dotm files are supported",
      fileTooLarge: "DOCX file size {actual} bytes exceeds the allowed maximum of {allowed} bytes.",
      live: {
        loading: "Opening DOCX document",
        error: "Unable to open document: {message}",
        ready: "Document ready",
        readyPages: "Document ready, page {page} of {total}",
      },
      pageLimitExceeded:
        "DOCX page count {count} exceeds the allowed maximum of {allowed}.",
    },
    toolbar: {
      sidebarToggle: "Toggle page thumbnails",
      pageThumbnails: "Page thumbnails",
      searchPlaceholder: "Search document",
      trackedChanges: "Tracked changes",
      comments: "Comments",
      trackedChangesShow: "Show {count} tracked changes",
      trackedChangesHide: "Hide {count} tracked changes",
      trackedChangesNone: "No tracked changes",
      commentsShow: "Show {count} comments",
      commentsHide: "Hide {count} comments",
      commentsNone: "No comments",
      lightThemeAction: "Use light document theme",
      darkThemeAction: "Use dark document theme",
      lightDocument: "Light document",
      darkDocument: "Dark document",
      upload: "Upload DOCX",
      download: "Download DOCX",
    },
    editor: {
      new: "New",
      newTitle: "New document",
      newDocument: "(new document)",
      importedDocument: "(imported document)",
      import: "Import",
      importTitle: "Import DOCX",
      export: "Export",
      exportTitle: "Export DOCX",
      edits: "Edits",
      toggleTheme: "Toggle document theme",
      undo: "Undo (⌘Z)",
      redo: "Redo (⇧⌘Z)",
      fontFamily: "Font family",
      font: "Font",
      fontSize: "Font size",
      size: "Size",
      bold: "Bold",
      italic: "Italic",
      underline: "Underline",
      strikethrough: "Strikethrough",
      superscript: "Superscript",
      subscript: "Subscript",
      textColor: "Text color",
      highlightColor: "Highlight color",
      placeholder: "Type here...",
      link: "Link",
      alignLeft: "Align left",
      alignCenter: "Align center",
      alignRight: "Align right",
      justify: "Justify",
      bulletList: "Bullet list",
      numberedList: "Numbered list",
      increaseIndent: "Increase indent",
      increaseIndentTitle: "Increase list indent (Tab)",
      decreaseIndent: "Decrease indent",
      decreaseIndentTitle: "Decrease list indent (Shift+Tab)",
      lineSpacing: "Line spacing",
      paragraphStyle: "Paragraph style",
      normal: "Normal",
      borders: "Borders",
      insertImage: "Insert image",
      insertTable: "Insert table",
      zoom: "Zoom",
      removeLink: 'Remove link "{url}"?',
      unsupportedLink: "Only http, https, mailto, internal #fragment, and relative links are allowed.",
      enterUrl: "Enter URL:",
      liveError: "Unable to display document: {message}",
      liveReady: "Document ready",
      acceptAll: "Accept all",
      acceptAllTitle: "Accept all tracked changes",
      rejectAll: "Reject all",
      rejectAllTitle: "Reject all tracked changes",
      status: {
        ready: "Ready",
        copied: "Copied to clipboard",
        pasted: "Pasted from clipboard",
        listParagraphNotFound: "Cannot change the list: paragraph not found",
        listAppliedBullet: "Applied bullet list",
        listAppliedNumbered: "Applied numbered list",
        maximumListLevel: "Maximum list level is {level}",
        discardChanges: "Discard unsaved changes?",
        loadingFile: "Loading {name}...",
        importFailed: "Failed to load file: {message}",
        unsupportedFile: "Only .docx, .docm, .dotx and .dotm files are supported",
        exportCancelledChanged: "Export cancelled: document changed",
        exportNotReady: "Wait for the document to finish loading before exporting",
        exportFailed: "Export failed: {message}",
        exportRequiredField: "Cannot export: required form field \"{label}\" is empty",
        exportRequiredFieldEmpty: "Cannot export: a required form field is empty",
        exporting: "Exporting DOCX...",
        exportingPhase: "Exporting DOCX: {phase}...",
        exported: "Exported DOCX",
        unknownImportError: "Unknown import error",
        unknownExportError: "Unknown export error",
        acceptedChanges: "Accepted {count} tracked changes",
        acceptedChange: "Accepted 1 tracked change",
        rejectedChanges: "Rejected {count} tracked changes",
        rejectedChange: "Rejected 1 tracked change",
        formattingCleared: " ({count} formatting revisions cleared without reverting; the model keeps no previous formatting)",
        formattingClearedOne: " (1 formatting revision cleared without reverting; the model keeps no previous formatting)",
        unsupportedFeature: "Unsupported: {feature}",
        imageTargetUnavailable: "Unsupported: image target is unavailable",
        imageSizePositive: "Unsupported: image size must be positive",
        resizeTargetUnavailable: "Unsupported: resize target is unavailable",
        lockedImageAnchorResize: "Unsupported: locked image anchor cannot be resized",
        imageOverlapDisabled: "Unsupported: image overlap is disabled for this anchor",
        syntheticTextBoxUnavailable: "Unsupported: synthetic text box editing is unavailable for this image",
        imageMoveTargetUnavailable: "Unsupported: image move target is unavailable",
        lockedImageAnchorMove: "Unsupported: locked image anchor cannot be moved",
        sectionImageTargetUnavailable: "Unsupported: section image target is unavailable",
        dropCapTargetUnavailable: "Unsupported: drop cap target is unavailable",
        dropCapFontSizeTargetUnavailable: "Unsupported: drop cap font size target is unavailable",
        dropCapTextTargetUnavailable: "Unsupported: drop cap text target is unavailable",
        imageMoveSourceUnavailable: "Unsupported: image move source is unavailable",
        selectImageFile: "Select an image file",
        imageInsertionCancelled: "Image insertion cancelled: document changed",
        insertedImage: "Inserted image: {name}",
        resizedImage: "Resized image to {width}×{height}",
        updatedTextBox: "Updated text box",
        imageWrapSet: "Set image wrap: {mode}",
        movedImage: "Moved image",
        movedSectionImage: "Moved section image",
        movedDropCap: "Moved drop cap",
        dropCapFontSize: "Set drop cap font size to {size}pt",
        updatedDropCapText: "Updated drop cap text",
        tableColumnWidthPositive: "Unsupported: table column widths must be positive",
        tableColumnWidthTargetUnavailable: "Unsupported: table column width target is unavailable",
        cannotInsertRowTableNotFound: "Cannot insert a row: table not found",
        cannotInsertTableRow: "Cannot insert a table row",
        cannotInsertColumnTableNotFound: "Cannot insert a column: table not found",
        cannotInsertTableColumn: "Cannot insert a table column",
        cannotDeleteRowTableNotFound: "Cannot delete a row: table not found",
        cannotDeleteOnlyTableRow: "Cannot delete the only table row",
        cannotDeleteColumnTableNotFound: "Cannot delete a column: table not found",
        cannotDeleteOnlyTableColumn: "Cannot delete the only table column",
        cannotDeleteTableColumn: "Cannot delete the table column",
        cannotDeleteTableNotFound: "Cannot delete the table: table not found",
        cannotClearCellsTableNotFound: "Cannot clear cells: table not found",
        cannotMergeCellsTableNotFound: "Cannot merge cells: table not found",
        cannotMergeSingleCell: "Cannot merge a single cell",
        cannotSplitCellTableNotFound: "Cannot split a cell: table not found",
        cannotSplitCellNotMerged: "Cannot split a cell that is not merged",
        cannotMoveTableNotFound: "Cannot move the table: table not found",
        cannotMoveEmbeddedHostTableNotFound: "Cannot move the embedded table: host table not found",
        cannotMoveEmbeddedHostCellNotFound: "Cannot move the embedded table: host cell not found",
        cannotMoveEmbeddedTableNotFound: "Cannot move the embedded table: table not found",
        insertedTable: "Inserted table",
        insertedTableRow: "Inserted table row {direction}",
        insertedTableColumn: "Inserted table column {direction}",
        deletedTableRow: "Deleted table row",
        deletedTableColumn: "Deleted table column",
        deletedTable: "Deleted table",
        clearedTableCells: "Cleared table cells",
        resizedTableColumns: "Resized table columns",
        mergedTableCells: "Merged table cells",
        splitTableCell: "Split table cell",
        movedTable: "Moved table",
        movedEmbeddedTable: "Moved embedded table",
        unsupportedDiagonalTableCellBorders: "Unsupported: diagonal table-cell borders ({preset})",
        unsupportedDiagonalParagraphBorders: "Unsupported: diagonal paragraph borders ({preset})",
        revisionLossOne: "Tracked changes in an edited paragraph were resolved with the edit",
        revisionLossMany: "Tracked changes in {count} edited paragraphs were resolved with the edit",
      },
      borderPresets: {
        bottom: "Bottom Border",
        top: "Top Border",
        left: "Left Border",
        right: "Right Border",
        none: "No Border",
        all: "All Borders",
        outside: "Outside Borders",
        inside: "Inside Borders",
        insideHorizontal: "Inside Horizontal",
        insideVertical: "Inside Vertical",
        diagonalDown: "Diagonal Down",
        diagonalUp: "Diagonal Up",
        horizontalLine: "Horizontal Line",
      },
    },
    contextMenu: {
      label: "Editor context menu",
      cut: "Cut",
      copy: "Copy",
      paste: "Paste",
      bold: "Bold",
      italic: "Italic",
      underline: "Underline",
      strikethrough: "Strikethrough",
      highlight: "Highlight",
      noHighlight: "No highlight",
      textColor: "Text Color",
      splitParagraph: "Split Paragraph",
      deleteParagraph: "Delete Paragraph",
      table: "Table",
      insertRowAbove: "Insert Row Above",
      insertRowBelow: "Insert Row Below",
      insertColumnLeft: "Insert Column Left",
      insertColumnRight: "Insert Column Right",
      deleteRow: "Delete Row",
      deleteColumn: "Delete Column",
      deleteTable: "Delete Table",
      mergeCells: "Merge cells",
      splitCell: "Split cell",
    },
    colors: {
      black: "Black",
      red: "Red",
      orange: "Orange",
      yellow: "Yellow",
      green: "Green",
      cyan: "Cyan",
      blue: "Blue",
      magenta: "Magenta",
      purple: "Purple",
    },
    thumbnail: {
      show: "Show thumbnails",
      hide: "Hide thumbnails",
      pages: "Pages",
      unavailable: "Thumbnail unavailable",
      unavailableReason: "This thumbnail API has no document source to render. Use useDocxPageThumbnails(editor) or DocxThumbnailPanel.",
    },
    image: {
      inline: "Inline",
      square: "Square",
      behind: "Behind",
      front: "Front",
      inlineTitle: "Inline with text",
      squareTitle: "Wrap text around image",
      behindTitle: "Behind text",
      frontTitle: "In front of text",
      alt: "Image",
      missing: "Missing image",
      wrapInline: "In Line with Text",
      wrapSquare: "Square",
      wrapTight: "Tight",
      wrapThrough: "Through",
      wrapTopAndBottom: "Top and Bottom",
      wrapBehind: "Behind Text",
      wrapFront: "In Front of Text",
      moveWithText: "Move with Text",
      fixPositionOnPage: "Fix Position on Page",
      editWrapBoundary: "Edit Wrap Boundary",
      moreLayoutOptions: "More Layout Options...",
    },
    formField: {
      placeholder: "Enter text...",
      selectDate: "Select a date",
      errors: {
        invalidValue: "Invalid value.",
        required: "This field is required.",
        noOptions: "This dropdown has no options.",
        chooseValue: "Choose a value from the list.",
        invalidDate: "Enter a valid date.",
        maxLength: "Use at most {max} characters.",
      },
    },
    notes: {
      footnote: "Footnotes",
      endnote: "Endnotes",
    },
    documentSurface: {
      region: "Choose document area",
    },
    trackedChange: {
      resolved: "Resolved",
      inserted: "Inserted",
      deleted: "Deleted",
      movedFrom: "Moved from",
      movedTo: "Moved to",
      format: "Formatted",
      paragraphFormatting: "Paragraph formatting",
      tableFormatting: "Table formatting",
      rowFormatting: "Row formatting",
      cellFormatting: "Cell formatting",
      comment: "Comment",
      annotation: "Annotation",
    },
    gutter: {
      accept: "Accept",
      acceptTitle: "Accept this revision",
      reject: "Reject",
      rejectTitle: "Reject this revision",
      hideComments: "Hide comments",
      showComments: "Show comments",
      comments: "Comments",
      commentMarker: "Comment {ordinal}",
      commentMarkerWithAuthor: "Comment {ordinal} — {author}",
    },
    dragOverlay: {
      drop: "Drop DOCX file to import",
    },
    diagnostics: {
      imageDecodeFailed: "DOCX image could not be decoded.",
      numberingImageDecodeFailed: "DOCX numbering image could not be decoded.",
    },
  },
  xlsx: {
    viewer: {
      loading: "Loading…",
      empty: "Open an XLSX file to get started",
      deferredUntitled: "Untitled workbook",
      deferredTitle: "Large file: parsing deferred",
      deferredDetail: "File size {size}. Resuming parsing uses more memory.",
      resumeParsing: "Resume parsing",
      cancel: "Cancel",
      live: {
        loading: "Loading workbook",
        error: "Unable to open workbook: {message}",
        sheetReady: "Sheet {name} is ready",
        chartsheetReady: "Chart sheet is ready",
        empty: "No workbook open",
      },
    },
    toolbar: {
      zoomGroup: "Workbook zoom",
      search: "Search workbook",
      searchPlaceholder: "Find values",
      upload: "Upload XLSX",
      download: "Download source",
      export: "Export XLSX",
    },
    ribbon: {
      undo: "Undo",
      redo: "Redo",
      cut: "Cut",
      copy: "Copy",
      paste: "Paste",
      pasteSpecial: "Paste special",
      insertRowAbove: "Insert row above",
      insertRowBelow: "Insert row below",
      deleteRow: "Delete row",
      insertColumnLeft: "Insert column left",
      insertColumnRight: "Insert column right",
      deleteColumn: "Delete column",
      numberFormatDialog: "Number format…",
      nameManager: "Name manager",
      fontFamily: "Font family",
      fontSize: "Font size",
      bold: "Bold",
      italic: "Italic",
      underline: "Underline",
      strikethrough: "Strikethrough",
      textColor: "Text color",
      fillColor: "Fill color",
      alignLeft: "Align left",
      alignCenter: "Align center",
      alignRight: "Align right",
      alignTop: "Align top",
      alignMiddle: "Align middle",
      alignBottom: "Align bottom",
      wrapText: "Wrap text",
      numberFormat: "Number format",
      currency: "Currency",
      percent: "Percent",
      borders: "Borders",
      allBorders: "All borders",
      outsideBorders: "Outside borders",
      leftBorder: "Left border",
      rightBorder: "Right border",
      topBorder: "Top border",
      bottomBorder: "Bottom border",
      noBorders: "No border",
      borderColor: "Border color",
      merge: "Merge",
      mergeAndCenter: "Merge & center",
      unmerge: "Unmerge",
      readOnlyTitle: "Read-only mode",
      numberFormats: {
        general: "General",
        number: "Number",
        currency: "Currency",
        percent: "Percent",
        date: "Date",
        text: "Text",
      },
    },
    sheetTabs: {
      previous: "Previous sheet",
      next: "Next sheet",
      add: "Add sheet",
      remove: "Delete sheet",
      defaultName: "Sheet",
    },
    contextMenu: {
      label: "Cell menu",
      copy: "Copy",
      cut: "Cut",
      paste: "Paste",
      pasteSpecial: "Paste special",
      pasteAll: "Paste all",
      pasteValues: "Values only",
      pasteFormats: "Formats only",
      pasteTranspose: "Transpose",
      editCell: "Edit cell",
      clear: "Clear contents",
      clearContents: "Clear contents",
      sortAscending: "Sort ascending",
      sortDescending: "Sort descending",
      undo: "Undo",
      redo: "Redo",
      changeChartType: "Change Chart Type",
      insertRowAbove: "Insert row above",
      insertRowBelow: "Insert row below",
      deleteRow: "Delete row",
      insertColumnLeft: "Insert column left",
      insertColumnRight: "Insert column right",
      deleteColumn: "Delete column",
      mergeAndCenter: "Merge & center",
      mergeAcross: "Merge across",
      mergeCells: "Merge cells",
      unmerge: "Unmerge cells",
      selectRow: "Select row",
      selectColumn: "Select column",
      selectAll: "Select all",
    },
    formulaBar: {
      activeCell: "Active cell",
      placeholder: "Enter a formula or value",
    },
    chart: {
      unsupported: "Unsupported chart type",
      loading: "Loading chart…",
      loadFailed: "Chart failed to load. Try again.",
      unavailable: "This chart type is not available yet.",
      noPayload: "This chart sheet does not expose an embedded chart payload.",
      ariaLabel: "Chart",
      unsupportedType: "Unsupported chart type: {type}",
      other: "Other",
      seriesName: "Series {number}",
      surfaceSeries: "Q{number}",
      root: "Root",
    },
    image: {
      alt: "Image",
    },
    table: {
      defaultName: "Table {number}",
      defaultColumn: "Column {number}",
    },
    authoring: {
      close: "Close",
      rangeLabel: "Range {range}",
      chart: {
        title: "Chart",
        insertMode: "New chart",
        editMode: "Edit chart",
        family: "Chart type",
        familyBar: "Column",
        familyLine: "Line",
        familyPie: "Pie",
        chartTitle: "Chart title",
        chartTitlePlaceholder: "Quarterly revenue",
        showDataLabels: "Show values on the chart",
        headerRow: "First row contains series names",
        headerColumn: "First column contains categories",
        insert: "Insert chart",
        update: "Update chart",
        delete: "Delete chart",
        selectedChart: "Editing {title}",
        noneSelected: "Select a chart first",
      },
      comment: {
        title: "Comment",
        text: "Comment",
        textPlaceholder: "Add a note for this cell",
        author: "Author",
        authorPlaceholder: "Your name",
        save: "Save comment",
        delete: "Delete comment",
      },
      hyperlink: {
        title: "Hyperlink",
        target: "Address",
        targetPlaceholder: "https://example.com or #Sheet1!A1",
        tooltip: "Tooltip",
        tooltipPlaceholder: "Shown on hover",
        save: "Save link",
        remove: "Remove link",
      },
      table: {
        title: "Table",
        name: "Table name",
        namePlaceholder: "Table1",
        hasHeaders: "First row contains headers",
        create: "Create table",
        remove: "Remove table",
        existing: "{name} covers {range}",
      },
      conditional: {
        title: "Conditional formatting",
        kind: "Rule type",
        kindColorScale: "Colour scale",
        kindDataBar: "Data bar",
        kindIconSet: "Icon set",
        thresholds: "Thresholds",
        typeMin: "Minimum",
        typeMax: "Maximum",
        typePercentile: "Percentile",
        typeNumber: "Number",
        typePercent: "Percent",
        value: "Value",
        color: "Colour",
        iconSet: "Icon set",
        showValue: "Show value",
        reverse: "Reverse icons",
        add: "Add rule",
        update: "Update rule",
        delete: "Delete rule",
        edit: "Edit rule",
        empty: "No conditional formats on this sheet",
        ruleSummary: "{kind} · {ranges}",
      },
      validation: {
        title: "Data validation",
        type: "Allow",
        typeList: "List",
        typeWhole: "Whole number",
        typeDecimal: "Decimal",
        typeDate: "Date",
        typeTextLength: "Text length",
        source: "List source",
        sourcePlaceholder: "Red,Green,Blue or A1:A5",
        value1: "Minimum",
        value2: "Maximum",
        allowBlank: "Allow blank",
        showDropdown: "Show dropdown",
        inputMessage: "Input message",
        errorMessage: "Error message",
        apply: "Apply validation",
        remove: "Remove validation",
        existing: "{type} on {ranges}",
      },
      freeze: {
        freezeHere: "Freeze panes here",
        unfreeze: "Unfreeze panes",
      },
      pasteSpecial: {
        title: "Paste special",
        mode: "Paste",
        modeAll: "All",
        modeValues: "Values only",
        modeFormats: "Formats only",
        transpose: "Transpose",
        paste: "Paste",
        empty: "Copy or cut a range first, then choose the paste options.",
      },
      numberFormat: {
        title: "Format cells",
        category: "Category",
        number: "Number",
        currency: "Currency",
        percent: "Percentage",
        date: "Date",
        text: "Text",
        custom: "Custom",
        decimals: "Decimal places",
        thousands: "Use thousands separator",
        pattern: "Format code",
        preview: "Preview",
        apply: "Apply",
      },
      names: {
        title: "Name manager",
        name: "Name",
        namePlaceholder: "QuarterlyRevenue",
        refersTo: "Refers to",
        refersToPlaceholder: "Sheet1!$A$1:$C$3",
        add: "Add",
        update: "Update",
        delete: "Delete",
        useSelection: "Use selection",
        empty: "No defined names yet.",
      },
      context: {
        insertComment: "Insert comment",
        editComment: "Edit comment",
        insertHyperlink: "Insert hyperlink",
        editHyperlink: "Edit hyperlink",
        conditional: "Conditional formatting…",
        createTable: "Create table",
        removeTable: "Remove table",
        insertChart: "Insert chart",
        validation: "Data validation…",
        freezeHere: "Freeze panes here",
        unfreeze: "Unfreeze panes",
      },
      status: {
        chartInserted: "Chart inserted",
        chartUpdated: "Chart updated",
        chartDeleted: "Chart deleted",
        commentSaved: "Comment saved",
        commentDeleted: "Comment deleted",
        hyperlinkSaved: "Hyperlink saved",
        hyperlinkRemoved: "Hyperlink removed",
        tableCreated: "Table created",
        tableRemoved: "Table removed",
        ruleAdded: "Conditional format added",
        ruleUpdated: "Conditional format updated",
        ruleRemoved: "Conditional format deleted",
        validationSet: "Data validation applied",
        validationRemoved: "Data validation removed",
        actionFailed: "Nothing to change: check the selected range and values",
      },
    },
    drawing: {
      button: "Button",
      dropdown: "Dropdown",
      editbox: "Text box",
      listbox: "List box",
      spinner: "Spinner",
      scrollbar: "Scroll bar",
      label: "Label",
      groupBox: "Group box",
      unknown: "Control",
    },
    selection: {
      current: "Current selection {address}",
      none: "No cells selected",
    },
    validation: {
      list: "The value is not in the allowed list.",
      whole: "Enter a valid whole number.",
      decimal: "Enter a valid number.",
      textLength: "The text length does not meet the requirement.",
      date: "Enter a valid date.",
      time: "Enter a valid time.",
      generic: "The value does not meet the data validation rule.",
    },
    calculation: {
      sheetMissing: "A formula references a sheet that no longer exists; cached values are shown.",
      snapshotUnavailable: "The workbook snapshot could not be saved, so recalculation was skipped to protect the current data; cached values are shown.",
      failed: "Formulas could not be recalculated; cached values are shown. Recalculate manually to retry.",
    },
    clipboard: {
      copyFailed: "Copy failed: this browser does not allow clipboard writes.",
      noDocument: "Download failed: no document is available in this environment.",
      invalidUrl: "The download URL is invalid.",
      unverifiedBlob: "Downloads must use a verified Blob URL.",
    },
    grid: {
      openValidation: "Open data validation dropdown",
      closeValidation: "Close validation message",
      worksheet: "Worksheet",
      worksheetNamed: "Worksheet {name}",
      blank: "Blank",
    },
    surface: {
      region: "Choose worksheet area",
    },
    diagnostics: {
      imageDecodeFailed: "XLSX image could not be decoded.",
      chartDataIncomplete: "XLSX Worker returned incomplete chart data.",
      parseTimeout: "XLSX parsing exceeded {ms} ms.",
      exportXmlEscapeFailed: "The XML escape repair before export failed; the exported workbook may contain double escaping.",
      macroLoss: "The source workbook is macro-enabled (.xlsm/.xltm); exporting to .xlsx drops VBA macros.",
      csvEncoding: "CSV must use UTF-8, or UTF-16 with a BOM.",
      csvParseFailed: "Unable to parse the CSV file.",
      xlsbParseFailed: "Unable to parse the XLSB file.",
      fileOrSrcRequired: "Either `file` or `src` must be provided.",
      sheetNameInvalid: "Worksheet names cannot contain these characters: \\ / : * ? [ ]",
    },
    errors: {
      sourceNotAllowed: "The spreadsheet URL is outside the allowed policy.",
      fetchFailed: "Unable to fetch the spreadsheet.",
      invalidWorkbook: "The file is not a valid XLSX workbook.",
      limitExceeded: "The spreadsheet exceeds the configured limit.",
      imageLimitExceeded: "The workbook contains too many images.",
      invalidImage: "The workbook contains an invalid image.",
      imageDecodeFailed: "An image could not be decoded.",
      timeout: "Loading the spreadsheet timed out.",
      workerUnavailable: "The workbook worker is unavailable.",
      aborted: "Loading the spreadsheet was cancelled.",
      unknown: "Unable to load the spreadsheet.",
    },
  },
  pptx: {
    viewer: {
      previousSlide: "Previous slide",
      nextSlide: "Next slide",
      hiddenSlide: "Hidden slide",
      searchPlaceholder: "Search slides",
      search: "Search",
      sidebar: "Slides",
      empty: "Open a PPTX file to get started",
      loading: "Loading PPTX…",
      waiting: "Waiting for the slide stage…",
      cancelled: "PPTX loading was cancelled.",
      reload: "Reload",
      retry: "Retry",
      skip: "Skip",
      slideRenderFailed: "Failed to render slide {page}: {message}",
      previousStep: "Previous step",
      nextStep: "Next step",
      replay: "Replay",
      play: "Play",
      pause: "Pause",
      resumeMedia: "Resume media",
      resumeMediaAction: "Resume media playback",
      fullscreen: "Fullscreen",
      capability: {
        strict: "Exact {count}/{total}",
        approximate: "Approximate {count}",
        static: "Static {count}",
        unparsed: "Unparsed {count}",
      },
      status: {
        idle: "Not started",
        ready: "Ready",
        transitioning: "Transitioning",
        running: "Playing",
        waiting: "Waiting for next step",
        paused: "Paused",
        blocked: "Waiting for media permission",
        ended: "Playback ended",
        disposed: "Closed",
      },
      error: {
        controllerNotReady: "The PPTX playback controller is not ready.",
        operationFailed: "The PPTX playback operation failed.",
        previewNotReady: "The PPTX preview session is not ready.",
        fullscreenUnsupported: "Fullscreen playback is not supported in this environment.",
        fullscreenEnterRejected: "The browser rejected fullscreen mode.",
        fullscreenExitRejected: "The browser rejected exiting fullscreen mode.",
        documentOperationFailed: "The PPTX operation failed.",
      },
      live: {
        empty: "No presentation open",
        loading: "Loading presentation",
        error: "Unable to open presentation: {message}",
        ready: "Slides ready, slide {current} of {total}",
        hiddenSuffix: ", hidden slide",
        playbackSuffix: ", {status}",
        unknownError: "Unknown error",
      },
    },
    documentErrors: {
      loadFailed: "Unable to load PPTX.",
      previewNotReady: "The PPTX preview session is not ready.",
      disposed: "The PPTX document has been disposed.",
      staleOpen: "The PPTX open request was replaced.",
      timeout: "The PPTX stage did not become ready in time.",
      timeoutMs: "PPTX loading did not finish within {ms} ms.",
      staleResult: "The PPTX load result is stale.",
      aborted: "The PPTX open request was cancelled.",
      slideRenderFailedWithMessage: "Slide render failed: {message}",
      slideRenderFailed: "Slide render failed.",
    },
    thumbnail: {
      label: "Slide {number}",
      hiddenLabel: "Slide {number}, hidden",
      hidden: "Hidden",
    },
    stage: {
      region: "Choose slide area",
    },
  },
  pdf: {
    viewer: {
      noPdf: "No PDF loaded",
      opening: "Opening PDF…",
      toggleThumbnails: "Show or hide thumbnails",
      pageInput: "Page number",
      rotateClockwise: "Rotate clockwise",
      searchLabel: "Search PDF text",
      searchPlaceholder: "Search text",
      search: "Search",
      download: "Download PDF",
      thumbnails: "Page thumbnails",
      thumbnailAlt: "Page {page} thumbnail",
      pageLabel: "Page {page}",
      renderingPage: "Rendering page {page}…",
      pageAlt: "PDF page {page}",
      searchSearching: "Searching…",
      searchNoResults: "No results",
      searchError: "Search failed",
      searchResult: "Result {current} / {total}, page {page}",
      renderFailed: "The PDF page could not be rendered.",
      invalidPdf: "The PDF file could not be parsed or rendered.",
      selectionLayer: "PDF page text selection layer",
    },
    surface: {
      noPdf: "No PDF loaded",
      opening: "Opening PDF…",
      renderingPage: "Rendering page {page}…",
      pageRenderFailed: "Failed to render page {page}",
      pageAlt: "Page {page}",
      selectionRegion: "Choose PDF page region",
      selectionText: "PDF page text selection layer",
    },
    errors: {
      sourceNotAllowed: "The PDF URL is outside the allowed policy.",
      fetchFailed: "Unable to fetch the PDF file.",
      invalidPdf: "The PDF file could not be parsed or rendered.",
      tooLarge: "PDF file size {actual} bytes exceeds the allowed maximum of {allowed} bytes.",
      aborted: "PDF loading was cancelled.",
      encrypted: "The PDF document is encrypted and requires a password.",
      openFailed: "Unable to open the PDF file.",
      unknown: "unknown",
    },
    runtime: {
      cancelled: "{label} cancelled",
      timedOut: "{label} timed out",
      workerStartup: "PDF Worker startup and document open",
      textSearch: "PDF text search",
      textExtraction: "PDF text extraction",
      documentClose: "PDF document close",
      runtimeDestroy: "PDF runtime destroy",
    },
  },
  vueUi: {
    regionSelector: {
      label: "Selected region",
      instructions: "Drag to select a region. Use the arrow keys to move it, Shift+arrow keys to resize, Enter to confirm, and Escape to cancel.",
      defaultAriaLabel: "Choose document region",
      description: "Region, left {left}%, top {top}%, width {width}%, height {height}%",
      noSelection: "No region selected",
    },
    objectOutline: {
      defaultAriaLabel: "Selectable document object",
      noObjects: "No objects available",
      noCandidates: "No candidate objects",
      itemDescription: "{label}, type {kind}, reliability {reliability}",
    },
    boundingBox: {
      noFields: "No fields to display",
      imageAlt: "Citation source preview",
    },
    fileUpload: {
      prompt: "Drag & drop files here, or click to select",
    },
    layoutBlocks: {
      noBlocks: "No layout blocks detected",
      imageAlt: "Layout preview",
    },
    signaturePad: {
      clear: "Clear",
    },
    fileThumbnail: {
      unknownType: "Unknown type",
    },
  },
}

export const officeMessagesZhCN: OfficeMessages = {
  common: {
    loading: "加载中…",
    readOnly: "只读",
    selectedText: "所选文本",
    search: {
      placeholder: "查找",
      previous: "上一个搜索结果",
      next: "下一个搜索结果",
      close: "关闭查找",
      searching: "正在搜索",
      error: "搜索失败",
      noResults: "没有结果",
      resultOf: "第 {current} 项，共 {total} 项",
    },
    zoom: {
      in: "放大",
      out: "缩小",
      level: "缩放比例",
      reset: "恢复适合窗口",
    },
    page: {
      label: "页",
      current: "当前页",
      navigation: "页面导航",
      previous: "上一页",
      next: "下一页",
      first: "第一页",
      last: "最后一页",
      of: "共",
      indicator: "第 {current} 页，共 {total} 页",
    },
    reference: {
      confirmFailed: "无法确认 {format} 引用。",
      createFailed: "无法创建 {format} 文本引用。",
      captureUnsupported: "Surface 不提供 {format} 预览捕获；请使用宿主的捕获流程。",
      describeAborted: "{format} 引用描述已中止。",
      hitTestFailed: "{format} 命中测试失败。",
    },
    descriptors: {
      document: "文档",
      page: "第 {number} 页",
      image: "图片",
      table: "表格 {number}",
      row: "第 {number} 行",
      cell: "第 {number} 个单元格",
      paragraph: "第 {number} 段",
      column: "{column} 列",
      workbook: "工作簿",
      presentation: "演示文稿",
      slide: "第 {number} 张幻灯片",
      pdf: "PDF",
      textRange: "文本范围",
      selectedRegion: "选定区域",
      regionOnPage: "第 {number} 页上的区域",
      regionOnSheet: "{name} 上的区域",
      regionOnSlide: "第 {number} 张幻灯片上的区域",
      comment: "批注",
      commentById: "批注 {id}",
      commentBy: "{author} 的批注",
      commentText: "批注：{text}",
      chart: "图表",
      chartName: "图表 {name}",
      group: "组 {number}",
      shape: "形状 {id}",
      object: "对象 {id}",
      textBox: "文本框 {id}",
      imageId: "图片 {id}",
      imageText: "图片：{text}",
      textLabel: "文本：{text}",
    },
  },
  docx: {
    viewer: {
      loadingTitle: "正在打开文档",
      loadingDetail: "正在加载 DOCX…",
      errorTitle: "无法打开此文档",
      chooseAnother: "选择其他 Word 文件",
      emptyTitle: "未打开文档",
      emptyDetail: "上传或提供 DOCX 文件开始查看。",
      chooseFile: "选择 Word 文件",
      parseError: "未知的 DOCX 解析错误",
      unsupportedFile: "仅支持 .docx、.docm、.dotx 和 .dotm 文件",
      fileTooLarge: "DOCX 文件大小 {actual} 字节超过允许值 {allowed} 字节。",
      live: {
        loading: "正在打开 DOCX 文档",
        error: "无法打开文档：{message}",
        ready: "文档已就绪",
        readyPages: "文档已就绪，第 {page} 页，共 {total} 页",
      },
      pageLimitExceeded: "DOCX 页数 {count} 超过允许值 {allowed}。",
    },
    toolbar: {
      sidebarToggle: "显示或隐藏页面缩略图",
      pageThumbnails: "页面缩略图",
      searchPlaceholder: "搜索文档",
      trackedChanges: "修订",
      comments: "批注",
      trackedChangesShow: "显示 {count} 处修订",
      trackedChangesHide: "隐藏 {count} 处修订",
      trackedChangesNone: "没有修订",
      commentsShow: "显示 {count} 条批注",
      commentsHide: "隐藏 {count} 条批注",
      commentsNone: "没有批注",
      lightThemeAction: "使用浅色文档主题",
      darkThemeAction: "使用深色文档主题",
      lightDocument: "浅色文档",
      darkDocument: "深色文档",
      upload: "上传 DOCX",
      download: "下载 DOCX",
    },
    editor: {
      new: "新建",
      newTitle: "新建文档",
      newDocument: "（新文档）",
      importedDocument: "（导入的文档）",
      import: "导入",
      importTitle: "导入 DOCX",
      export: "导出",
      exportTitle: "导出 DOCX",
      edits: "修订",
      toggleTheme: "切换文档主题",
      undo: "撤销 (⌘Z)",
      redo: "重做 (⇧⌘Z)",
      fontFamily: "字体",
      font: "字体",
      fontSize: "字号",
      size: "字号",
      bold: "加粗",
      italic: "斜体",
      underline: "下划线",
      strikethrough: "删除线",
      superscript: "上标",
      subscript: "下标",
      textColor: "文字颜色",
      highlightColor: "突出显示颜色",
      placeholder: "在此输入…",
      link: "链接",
      alignLeft: "左对齐",
      alignCenter: "居中对齐",
      alignRight: "右对齐",
      justify: "两端对齐",
      bulletList: "项目符号列表",
      numberedList: "编号列表",
      increaseIndent: "增加缩进",
      increaseIndentTitle: "增加列表缩进 (Tab)",
      decreaseIndent: "减少缩进",
      decreaseIndentTitle: "减少列表缩进 (Shift+Tab)",
      lineSpacing: "行距",
      paragraphStyle: "段落样式",
      normal: "正文",
      borders: "边框",
      insertImage: "插入图片",
      insertTable: "插入表格",
      zoom: "缩放",
      removeLink: "移除链接“{url}”？",
      unsupportedLink: "仅允许 http、https、mailto、内部 #片段和相对链接。",
      enterUrl: "输入网址：",
      liveError: "文档无法显示：{message}",
      liveReady: "文档已就绪",
      acceptAll: "全部接受",
      acceptAllTitle: "接受所有修订",
      rejectAll: "全部拒绝",
      rejectAllTitle: "拒绝所有修订",
      status: {
        ready: "就绪",
        copied: "已复制到剪贴板",
        pasted: "已从剪贴板粘贴",
        listParagraphNotFound: "无法更改列表：未找到段落",
        listAppliedBullet: "已应用项目符号列表",
        listAppliedNumbered: "已应用编号列表",
        maximumListLevel: "列表最大层级为 {level}",
        discardChanges: "放弃未保存的更改？",
        loadingFile: "正在加载 {name}…",
        importFailed: "文件加载失败：{message}",
        unsupportedFile: "仅支持 .docx、.docm、.dotx 和 .dotm 文件",
        exportCancelledChanged: "导出已取消：文档已更改",
        exportNotReady: "请等待文档加载完成后再导出",
        exportFailed: "导出失败：{message}",
        exportRequiredField: "无法导出：必填表单字段“{label}”为空",
        exportRequiredFieldEmpty: "无法导出：必填表单字段为空",
        exporting: "正在导出 DOCX…",
        exportingPhase: "正在导出 DOCX：{phase}…",
        exported: "已导出 DOCX",
        unknownImportError: "未知导入错误",
        unknownExportError: "未知导出错误",
        acceptedChanges: "已接受 {count} 处修订",
        acceptedChange: "已接受 1 处修订",
        rejectedChanges: "已拒绝 {count} 处修订",
        rejectedChange: "已拒绝 1 处修订",
        formattingCleared: "（已清除 {count} 处格式修订且未回退；模型未保留旧格式）",
        formattingClearedOne: "（已清除 1 处格式修订且未回退；模型未保留旧格式）",
        unsupportedFeature: "不支持：{feature}",
        imageTargetUnavailable: "不支持：图片目标不可用",
        imageSizePositive: "不支持：图片尺寸必须为正数",
        resizeTargetUnavailable: "不支持：调整大小的目标不可用",
        lockedImageAnchorResize: "不支持：锁定的图片锚点不能调整大小",
        imageOverlapDisabled: "不支持：此锚点不允许图片重叠",
        syntheticTextBoxUnavailable: "不支持：此图片无法编辑合成文本框",
        imageMoveTargetUnavailable: "不支持：图片移动目标不可用",
        lockedImageAnchorMove: "不支持：锁定的图片锚点不能移动",
        sectionImageTargetUnavailable: "不支持：节图片目标不可用",
        dropCapTargetUnavailable: "不支持：首字下沉目标不可用",
        dropCapFontSizeTargetUnavailable: "不支持：首字下沉字号目标不可用",
        dropCapTextTargetUnavailable: "不支持：首字下沉文本目标不可用",
        imageMoveSourceUnavailable: "不支持：图片移动源不可用",
        selectImageFile: "请选择图片文件",
        imageInsertionCancelled: "图片插入已取消：文档已更改",
        insertedImage: "已插入图片：{name}",
        resizedImage: "已将图片调整为 {width}×{height}",
        updatedTextBox: "已更新文本框",
        imageWrapSet: "已设置图片环绕方式：{mode}",
        movedImage: "已移动图片",
        movedSectionImage: "已移动节图片",
        movedDropCap: "已移动首字下沉",
        dropCapFontSize: "已将首字下沉字号设为 {size}pt",
        updatedDropCapText: "已更新首字下沉文本",
        tableColumnWidthPositive: "不支持：表格列宽必须为正数",
        tableColumnWidthTargetUnavailable: "不支持：表格列宽目标不可用",
        cannotInsertRowTableNotFound: "无法插入行：未找到表格",
        cannotInsertTableRow: "无法插入表格行",
        cannotInsertColumnTableNotFound: "无法插入列：未找到表格",
        cannotInsertTableColumn: "无法插入表格列",
        cannotDeleteRowTableNotFound: "无法删除行：未找到表格",
        cannotDeleteOnlyTableRow: "无法删除唯一的表格行",
        cannotDeleteColumnTableNotFound: "无法删除列：未找到表格",
        cannotDeleteOnlyTableColumn: "无法删除唯一的表格列",
        cannotDeleteTableColumn: "无法删除该表格列",
        cannotDeleteTableNotFound: "无法删除表格：未找到表格",
        cannotClearCellsTableNotFound: "无法清除单元格：未找到表格",
        cannotMergeCellsTableNotFound: "无法合并单元格：未找到表格",
        cannotMergeSingleCell: "无法合并单个单元格",
        cannotSplitCellTableNotFound: "无法拆分单元格：未找到表格",
        cannotSplitCellNotMerged: "无法拆分未合并的单元格",
        cannotMoveTableNotFound: "无法移动表格：未找到表格",
        cannotMoveEmbeddedHostTableNotFound: "无法移动嵌入表格：未找到宿主表格",
        cannotMoveEmbeddedHostCellNotFound: "无法移动嵌入表格：未找到宿主单元格",
        cannotMoveEmbeddedTableNotFound: "无法移动嵌入表格：未找到表格",
        insertedTable: "已插入表格",
        insertedTableRow: "已插入表格行（{direction}）",
        insertedTableColumn: "已插入表格列（{direction}）",
        deletedTableRow: "已删除表格行",
        deletedTableColumn: "已删除表格列",
        deletedTable: "已删除表格",
        clearedTableCells: "已清除表格单元格",
        resizedTableColumns: "已调整表格列宽",
        mergedTableCells: "已合并表格单元格",
        splitTableCell: "已拆分表格单元格",
        movedTable: "已移动表格",
        movedEmbeddedTable: "已移动嵌入表格",
        unsupportedDiagonalTableCellBorders: "不支持：对角线表格单元格边框（{preset}）",
        unsupportedDiagonalParagraphBorders: "不支持：对角线段落边框（{preset}）",
        revisionLossOne: "已编辑段落中的修订已随编辑一起处理",
        revisionLossMany: "已编辑的 {count} 个段落中的修订已随编辑一起处理",
      },
      borderPresets: {
        bottom: "下边框",
        top: "上边框",
        left: "左边框",
        right: "右边框",
        none: "无边框",
        all: "所有边框",
        outside: "外侧边框",
        inside: "内侧边框",
        insideHorizontal: "内部横边框",
        insideVertical: "内部竖边框",
        diagonalDown: "斜下边框",
        diagonalUp: "斜上边框",
        horizontalLine: "水平线",
      },
    },
    contextMenu: {
      label: "编辑器上下文菜单",
      cut: "剪切",
      copy: "复制",
      paste: "粘贴",
      bold: "加粗",
      italic: "斜体",
      underline: "下划线",
      strikethrough: "删除线",
      highlight: "突出显示",
      noHighlight: "无突出显示",
      textColor: "文字颜色",
      splitParagraph: "拆分段落",
      deleteParagraph: "删除段落",
      table: "表格",
      insertRowAbove: "在上方插入行",
      insertRowBelow: "在下方插入行",
      insertColumnLeft: "在左侧插入列",
      insertColumnRight: "在右侧插入列",
      deleteRow: "删除行",
      deleteColumn: "删除列",
      deleteTable: "删除表格",
      mergeCells: "合并单元格",
      splitCell: "拆分单元格",
    },
    colors: {
      black: "黑色",
      red: "红色",
      orange: "橙色",
      yellow: "黄色",
      green: "绿色",
      cyan: "青色",
      blue: "蓝色",
      magenta: "洋红色",
      purple: "紫色",
    },
    thumbnail: {
      show: "显示缩略图",
      hide: "隐藏缩略图",
      pages: "页面",
      unavailable: "缩略图暂不可用",
      unavailableReason: "此缩略图接口没有文档来源，无法渲染。请使用 useDocxPageThumbnails(editor) 或 DocxThumbnailPanel。",
    },
    image: {
      inline: "嵌入",
      square: "四周环绕",
      behind: "衬于文字下方",
      front: "浮于文字上方",
      inlineTitle: "嵌入文字",
      squareTitle: "文字环绕图片",
      behindTitle: "衬于文字下方",
      frontTitle: "浮于文字上方",
      alt: "图片",
      missing: "图片缺失",
      wrapInline: "嵌入文字",
      wrapSquare: "四周环绕",
      wrapTight: "紧密环绕",
      wrapThrough: "穿越环绕",
      wrapTopAndBottom: "上下环绕",
      wrapBehind: "衬于文字下方",
      wrapFront: "浮于文字上方",
      moveWithText: "随文字移动",
      fixPositionOnPage: "固定在页面位置",
      editWrapBoundary: "编辑环绕边界",
      moreLayoutOptions: "更多布局选项…",
    },
    formField: {
      placeholder: "输入文本…",
      selectDate: "选择日期",
      errors: {
        invalidValue: "值无效。",
        required: "此字段为必填项。",
        noOptions: "此下拉框没有选项。",
        chooseValue: "请从列表中选择一个值。",
        invalidDate: "请输入有效的日期。",
        maxLength: "最多使用 {max} 个字符。",
      },
    },
    notes: {
      footnote: "脚注",
      endnote: "尾注",
    },
    documentSurface: {
      region: "选择文档区域",
    },
    trackedChange: {
      resolved: "已解决",
      inserted: "插入",
      deleted: "删除",
      movedFrom: "移出",
      movedTo: "移入",
      format: "格式更改",
      paragraphFormatting: "段落格式",
      tableFormatting: "表格格式",
      rowFormatting: "行格式",
      cellFormatting: "单元格格式",
      comment: "批注",
      annotation: "标注",
    },
    gutter: {
      accept: "接受",
      acceptTitle: "接受此修订",
      reject: "拒绝",
      rejectTitle: "拒绝此修订",
      hideComments: "隐藏批注",
      showComments: "显示批注",
      comments: "批注",
      commentMarker: "批注 {ordinal}",
      commentMarkerWithAuthor: "批注 {ordinal} — {author}",
    },
    dragOverlay: {
      drop: "拖放 DOCX 文件以导入",
    },
    diagnostics: {
      imageDecodeFailed: "DOCX 图片无法解码。",
      numberingImageDecodeFailed: "DOCX 编号图片无法解码。",
    },
  },
  xlsx: {
    viewer: {
      loading: "加载中…",
      empty: "打开一个 XLSX 文件开始查看",
      deferredUntitled: "未命名工作簿",
      deferredTitle: "文件较大，已延迟解析",
      deferredDetail: "文件大小 {size}，继续解析会占用更多内存。",
      resumeParsing: "继续解析",
      cancel: "取消",
      live: {
        loading: "正在加载工作簿",
        error: "无法打开工作簿：{message}",
        sheetReady: "工作表 {name} 已就绪",
        chartsheetReady: "图表工作表已就绪",
        empty: "没有打开工作簿",
      },
    },
    toolbar: {
      zoomGroup: "工作簿缩放",
      search: "搜索工作簿",
      searchPlaceholder: "查找值",
      upload: "上传 XLSX",
      download: "下载源文件",
      export: "导出 XLSX",
    },
    ribbon: {
      undo: "撤销",
      redo: "重做",
      cut: "剪切",
      copy: "复制",
      paste: "粘贴",
      pasteSpecial: "选择性粘贴",
      insertRowAbove: "在上方插入行",
      insertRowBelow: "在下方插入行",
      deleteRow: "删除行",
      insertColumnLeft: "在左侧插入列",
      insertColumnRight: "在右侧插入列",
      deleteColumn: "删除列",
      numberFormatDialog: "数字格式…",
      nameManager: "名称管理器",
      fontFamily: "字体",
      fontSize: "字号",
      bold: "加粗",
      italic: "斜体",
      underline: "下划线",
      strikethrough: "删除线",
      textColor: "文字颜色",
      fillColor: "填充颜色",
      alignLeft: "左对齐",
      alignCenter: "居中对齐",
      alignRight: "右对齐",
      alignTop: "顶端对齐",
      alignMiddle: "垂直居中",
      alignBottom: "底端对齐",
      wrapText: "自动换行",
      numberFormat: "数字格式",
      currency: "货币",
      percent: "百分比",
      borders: "边框",
      allBorders: "所有边框",
      outsideBorders: "外侧边框",
      leftBorder: "左侧边框",
      rightBorder: "右侧边框",
      topBorder: "上边框",
      bottomBorder: "下边框",
      noBorders: "无边框",
      borderColor: "边框颜色",
      merge: "合并",
      mergeAndCenter: "合并并居中",
      unmerge: "取消合并",
      readOnlyTitle: "只读模式",
      numberFormats: {
        general: "常规",
        number: "数值",
        currency: "货币",
        percent: "百分比",
        date: "日期",
        text: "文本",
      },
    },
    sheetTabs: {
      previous: "上一个工作表",
      next: "下一个工作表",
      add: "添加工作表",
      remove: "删除当前工作表",
      defaultName: "工作表",
    },
    contextMenu: {
      label: "单元格菜单",
      copy: "复制",
      cut: "剪切",
      paste: "粘贴",
      pasteSpecial: "选择性粘贴",
      pasteAll: "全部",
      pasteValues: "仅值",
      pasteFormats: "仅格式",
      pasteTranspose: "转置",
      editCell: "编辑单元格",
      clear: "清除内容",
      clearContents: "清除内容",
      sortAscending: "按此列升序",
      sortDescending: "按此列降序",
      undo: "撤销",
      redo: "重做",
      changeChartType: "更改图表类型",
      insertRowAbove: "在上方插入行",
      insertRowBelow: "在下方插入行",
      deleteRow: "删除行",
      insertColumnLeft: "在左侧插入列",
      insertColumnRight: "在右侧插入列",
      deleteColumn: "删除列",
      mergeAndCenter: "合并并居中",
      mergeAcross: "跨列合并",
      mergeCells: "合并单元格",
      unmerge: "取消合并",
      selectRow: "选择整行",
      selectColumn: "选择整列",
      selectAll: "全选",
    },
    formulaBar: {
      activeCell: "活动单元格",
      placeholder: "输入公式或值",
    },
    chart: {
      unsupported: "不支持的图表类型",
      loading: "正在加载图表…",
      loadFailed: "图表加载失败，请重试。",
      unavailable: "此图表类型暂不可用。",
      noPayload: "此图表工作表未提供嵌入图表数据。",
      ariaLabel: "图表",
      unsupportedType: "不支持的图表类型：{type}",
      other: "其他",
      seriesName: "系列 {number}",
      surfaceSeries: "第 {number} 季度",
      root: "根节点",
    },
    image: {
      alt: "图片",
    },
    table: {
      defaultName: "表格 {number}",
      defaultColumn: "列 {number}",
    },
    authoring: {
      close: "关闭",
      rangeLabel: "区域 {range}",
      chart: {
        title: "图表",
        insertMode: "新建图表",
        editMode: "编辑图表",
        family: "图表类型",
        familyBar: "柱形图",
        familyLine: "折线图",
        familyPie: "饼图",
        chartTitle: "图表标题",
        chartTitlePlaceholder: "季度收入",
        showDataLabels: "在图表上显示数值",
        headerRow: "首行是系列名称",
        headerColumn: "首列是分类标签",
        insert: "插入图表",
        update: "更新图表",
        delete: "删除图表",
        selectedChart: "正在编辑 {title}",
        noneSelected: "请先选择图表",
      },
      comment: {
        title: "批注",
        text: "批注内容",
        textPlaceholder: "为这个单元格添加备注",
        author: "作者",
        authorPlaceholder: "你的名字",
        save: "保存批注",
        delete: "删除批注",
      },
      hyperlink: {
        title: "超链接",
        target: "地址",
        targetPlaceholder: "https://example.com 或 #Sheet1!A1",
        tooltip: "提示文字",
        tooltipPlaceholder: "悬停时显示",
        save: "保存链接",
        remove: "移除链接",
      },
      table: {
        title: "表格",
        name: "表格名称",
        namePlaceholder: "Table1",
        hasHeaders: "首行是标题行",
        create: "创建表格",
        remove: "删除表格",
        existing: "{name} 覆盖 {range}",
      },
      conditional: {
        title: "条件格式",
        kind: "规则类型",
        kindColorScale: "色阶",
        kindDataBar: "数据条",
        kindIconSet: "图标集",
        thresholds: "阈值",
        typeMin: "最小值",
        typeMax: "最大值",
        typePercentile: "百分位",
        typeNumber: "数值",
        typePercent: "百分比",
        value: "值",
        color: "颜色",
        iconSet: "图标集",
        showValue: "显示数值",
        reverse: "反转图标",
        add: "添加规则",
        update: "更新规则",
        delete: "删除规则",
        edit: "编辑规则",
        empty: "此工作表没有条件格式",
        ruleSummary: "{kind} · {ranges}",
      },
      validation: {
        title: "数据验证",
        type: "允许",
        typeList: "序列",
        typeWhole: "整数",
        typeDecimal: "小数",
        typeDate: "日期",
        typeTextLength: "文本长度",
        source: "序列来源",
        sourcePlaceholder: "红,绿,蓝 或 A1:A5",
        value1: "最小值",
        value2: "最大值",
        allowBlank: "允许空值",
        showDropdown: "显示下拉箭头",
        inputMessage: "输入提示",
        errorMessage: "错误信息",
        apply: "应用验证",
        remove: "移除验证",
        existing: "{type}，区域 {ranges}",
      },
      freeze: {
        freezeHere: "在此冻结窗格",
        unfreeze: "取消冻结窗格",
      },
      pasteSpecial: {
        title: "选择性粘贴",
        mode: "粘贴",
        modeAll: "全部",
        modeValues: "仅值",
        modeFormats: "仅格式",
        transpose: "转置",
        paste: "粘贴",
        empty: "请先复制或剪切一个区域，再选择粘贴方式。",
      },
      numberFormat: {
        title: "设置单元格格式",
        category: "分类",
        number: "数值",
        currency: "货币",
        percent: "百分比",
        date: "日期",
        text: "文本",
        custom: "自定义",
        decimals: "小数位数",
        thousands: "使用千位分隔符",
        pattern: "格式代码",
        preview: "预览",
        apply: "应用",
      },
      names: {
        title: "名称管理器",
        name: "名称",
        namePlaceholder: "季度收入",
        refersTo: "引用位置",
        refersToPlaceholder: "Sheet1!$A$1:$C$3",
        add: "添加",
        update: "更新",
        delete: "删除",
        useSelection: "使用当前选区",
        empty: "尚未定义名称。",
      },
      context: {
        insertComment: "插入批注",
        editComment: "编辑批注",
        insertHyperlink: "插入超链接",
        editHyperlink: "编辑超链接",
        conditional: "条件格式…",
        createTable: "创建表格",
        removeTable: "删除表格",
        insertChart: "插入图表",
        validation: "数据验证…",
        freezeHere: "在此冻结窗格",
        unfreeze: "取消冻结窗格",
      },
      status: {
        chartInserted: "已插入图表",
        chartUpdated: "已更新图表",
        chartDeleted: "已删除图表",
        commentSaved: "已保存批注",
        commentDeleted: "已删除批注",
        hyperlinkSaved: "已保存超链接",
        hyperlinkRemoved: "已移除超链接",
        tableCreated: "已创建表格",
        tableRemoved: "已删除表格",
        ruleAdded: "已添加条件格式",
        ruleUpdated: "已更新条件格式",
        ruleRemoved: "已删除条件格式",
        validationSet: "已应用数据验证",
        validationRemoved: "已移除数据验证",
        actionFailed: "没有可应用的更改：请检查组区域和输入内容",
      },
    },
    drawing: {
      button: "按钮",
      dropdown: "下拉框",
      editbox: "输入框",
      listbox: "列表框",
      spinner: "微调框",
      scrollbar: "滚动条",
      label: "标签",
      groupBox: "分组框",
      unknown: "控件",
    },
    selection: {
      current: "当前选择 {address}",
      none: "当前没有选择单元格",
    },
    validation: {
      list: "输入值不在允许的列表中。",
      whole: "请输入有效的整数。",
      decimal: "请输入有效的数字。",
      textLength: "文本长度不符合要求。",
      date: "请输入有效的日期。",
      time: "请输入有效的时间。",
      generic: "输入值不符合数据验证规则。",
    },
    calculation: {
      sheetMissing: "公式引用了不存在的工作表，当前显示缓存值。",
      snapshotUnavailable: "无法保存工作簿快照，已跳过公式重算以保护当前数据，当前显示缓存值。",
      failed: "公式无法重新计算，当前显示缓存值。可手动重新计算重试。",
    },
    clipboard: {
      copyFailed: "复制失败：当前浏览器不允许写入剪贴板。",
      noDocument: "下载失败：当前环境没有可用的 document。",
      invalidUrl: "下载地址无效。",
      unverifiedBlob: "下载只能使用已验证的 Blob 地址。",
    },
    grid: {
      openValidation: "打开数据验证下拉列表",
      closeValidation: "关闭验证提示",
      worksheet: "工作表",
      worksheetNamed: "工作表 {name}",
      blank: "空白",
    },
    surface: {
      region: "选择工作表区域",
    },
    diagnostics: {
      imageDecodeFailed: "XLSX 图片无法解码。",
      chartDataIncomplete: "XLSX Worker 返回的图表数据不完整。",
      parseTimeout: "XLSX 解析超过 {ms} 毫秒。",
      exportXmlEscapeFailed: "导出前的 XML 转义修复失败，导出的工作簿可能包含双重转义。",
      macroLoss: "源文件为启用宏的工作簿（.xlsm/.xltm），导出的 .xlsx 会丢失 VBA 宏。",
      csvEncoding: "CSV 必须使用 UTF-8，或带 BOM 的 UTF-16 编码。",
      csvParseFailed: "无法解析 CSV 文件。",
      xlsbParseFailed: "无法解析 XLSB 文件。",
      fileOrSrcRequired: "必须提供 `file` 或 `src`。",
      sheetNameInvalid: "工作表名称不能包含以下字符：\\ / : * ? [ ]",
    },
    errors: {
      sourceNotAllowed: "电子表格地址不在允许范围内。",
      fetchFailed: "无法获取电子表格。",
      invalidWorkbook: "该文件不是有效的 XLSX 工作簿。",
      limitExceeded: "电子表格超出配置的限制。",
      imageLimitExceeded: "工作簿包含过多图片。",
      invalidImage: "工作簿包含无效图片。",
      imageDecodeFailed: "图片无法解码。",
      timeout: "加载电子表格超时。",
      workerUnavailable: "工作簿 Worker 不可用。",
      aborted: "电子表格加载已取消。",
      unknown: "无法加载电子表格。",
    },
  },
  pptx: {
    viewer: {
      previousSlide: "上一页",
      nextSlide: "下一页",
      hiddenSlide: "隐藏页",
      searchPlaceholder: "搜索幻灯片",
      search: "搜索",
      sidebar: "幻灯片列表",
      empty: "打开一个 PPTX 文件开始查看",
      loading: "正在加载 PPTX…",
      waiting: "正在等待幻灯片舞台…",
      cancelled: "PPTX 加载已取消。",
      reload: "重新加载",
      retry: "重试",
      skip: "跳过",
      slideRenderFailed: "第 {page} 页渲染失败：{message}",
      previousStep: "上一步",
      nextStep: "下一步",
      replay: "重新播放",
      play: "播放",
      pause: "暂停",
      resumeMedia: "继续媒体",
      resumeMediaAction: "继续播放媒体",
      fullscreen: "全屏",
      capability: {
        strict: "精确 {count}/{total}",
        approximate: "近似 {count}",
        static: "静态 {count}",
        unparsed: "未解析 {count}",
      },
      status: {
        idle: "尚未开始",
        ready: "已就绪",
        transitioning: "正在切换",
        running: "正在播放",
        waiting: "等待下一步",
        paused: "已暂停",
        blocked: "等待媒体授权",
        ended: "播放结束",
        disposed: "已关闭",
      },
      error: {
        controllerNotReady: "PPTX 播放控制器尚未准备好。",
        operationFailed: "PPTX 播放操作失败。",
        previewNotReady: "PPTX 预览会话尚未准备好。",
        fullscreenUnsupported: "当前环境不支持全屏播放。",
        fullscreenEnterRejected: "浏览器拒绝进入全屏。",
        fullscreenExitRejected: "浏览器拒绝退出全屏。",
        documentOperationFailed: "PPTX 操作失败。",
      },
      live: {
        empty: "没有打开幻灯片文件",
        loading: "正在加载幻灯片文件",
        error: "无法打开幻灯片文件：{message}",
        ready: "幻灯片已就绪，第 {current} 张，共 {total} 张",
        hiddenSuffix: "，隐藏页",
        playbackSuffix: "，{status}",
        unknownError: "未知错误",
      },
    },
    documentErrors: {
      loadFailed: "无法加载 PPTX。",
      previewNotReady: "PPTX 预览会话尚未准备好。",
      disposed: "PPTX 文档已经销毁。",
      staleOpen: "PPTX 打开请求已经被替换。",
      timeout: "PPTX 舞台长时间未就绪。",
      timeoutMs: "PPTX 加载超过 {ms} 毫秒仍未完成。",
      staleResult: "PPTX 加载结果已经过期。",
      aborted: "PPTX 打开请求已经取消。",
      slideRenderFailedWithMessage: "页面渲染失败：{message}",
      slideRenderFailed: "页面渲染失败。",
    },
    thumbnail: {
      label: "第 {number} 页",
      hiddenLabel: "第 {number} 页，隐藏页",
      hidden: "隐藏",
    },
    stage: {
      region: "选择幻灯片区域",
    },
  },
  pdf: {
    viewer: {
      noPdf: "未加载 PDF",
      opening: "正在打开 PDF…",
      toggleThumbnails: "显示或隐藏缩略图",
      pageInput: "页码",
      rotateClockwise: "顺时针旋转",
      searchLabel: "搜索 PDF 文本",
      searchPlaceholder: "搜索文本",
      search: "搜索",
      download: "下载 PDF",
      thumbnails: "页面缩略图",
      thumbnailAlt: "第 {page} 页缩略图",
      pageLabel: "第 {page} 页",
      renderingPage: "正在渲染第 {page} 页…",
      pageAlt: "PDF 第 {page} 页",
      searchSearching: "正在搜索…",
      searchNoResults: "没有找到",
      searchError: "搜索失败",
      searchResult: "第 {current} / {total} 项，第 {page} 页",
      renderFailed: "PDF 页面渲染失败。",
      invalidPdf: "PDF 文件无法解析或渲染。",
      selectionLayer: "PDF 页面文本选择层",
    },
    surface: {
      noPdf: "未加载 PDF",
      opening: "正在打开 PDF…",
      renderingPage: "正在渲染第 {page} 页…",
      pageRenderFailed: "第 {page} 页渲染失败",
      pageAlt: "第 {page} 页",
      selectionRegion: "选择 PDF 页面区域",
      selectionText: "PDF 页面文本选择层",
    },
    errors: {
      sourceNotAllowed: "PDF 地址不在允许范围内。",
      fetchFailed: "无法获取 PDF 文件。",
      invalidPdf: "PDF 文件无法解析或渲染。",
      tooLarge: "PDF 文件大小 {actual} 字节超过允许值 {allowed} 字节。",
      aborted: "PDF 加载已取消。",
      encrypted: "PDF 文档已加密，需要密码。",
      openFailed: "无法打开 PDF 文件。",
      unknown: "未知",
    },
    runtime: {
      cancelled: "{label} 已取消",
      timedOut: "{label} 已超时",
      workerStartup: "PDF Worker 启动和文档打开",
      textSearch: "PDF 文本搜索",
      textExtraction: "PDF 文本提取",
      documentClose: "PDF 文档关闭",
      runtimeDestroy: "PDF 运行时销毁",
    },
  },
  vueUi: {
    regionSelector: {
      label: "选定区域",
      instructions: "拖动以框选区域。方向键移动区域，按住 Shift 配合方向键调整大小，Enter 确认，Escape 取消。",
      defaultAriaLabel: "选择文档区域",
      description: "区域，左侧 {left}%，顶部 {top}%，宽 {width}%，高 {height}%",
      noSelection: "尚未选择区域",
    },
    objectOutline: {
      defaultAriaLabel: "可选择的文档对象",
      noObjects: "没有可用对象",
      noCandidates: "没有候选对象",
      itemDescription: "{label}，类型 {kind}，识别可靠度 {reliability}",
    },
    boundingBox: {
      noFields: "没有可显示的字段",
      imageAlt: "引用来源预览",
    },
    fileUpload: {
      prompt: "将文件拖放到此处，或点击选择",
    },
    layoutBlocks: {
      noBlocks: "未检测到布局块",
      imageAlt: "布局预览",
    },
    signaturePad: {
      clear: "清除",
    },
    fileThumbnail: {
      unknownType: "未知类型",
    },
  },
}

/**
 * Resolve a host-provided locale to a built-in catalog. Chinese locales map
 * to `zh-CN`; everything else falls back to English.
 */
export function resolveOfficeLocale(locale?: string | null): OfficeBuiltInLocale {
  if (!locale) return "en"
  const normalized = locale.trim().toLowerCase().replace(/_/gu, "-")
  if (
    normalized === "zh"
    || normalized === "zh-cn"
    || normalized === "zh-hans"
    || normalized === "zh-hans-cn"
  ) return "zh-CN"
  return "en"
}

function mergeMessages<T>(base: T, overrides: DeepPartial<T> | undefined): T {
  if (!overrides) return base
  const result: Record<string, unknown> = { ...(base as Record<string, unknown>) }
  for (const [key, value] of Object.entries(overrides as Record<string, unknown>)) {
    if (value === undefined) continue
    const current = result[key]
    if (
      value !== null
      && typeof value === "object"
      && !Array.isArray(value)
      && current !== null
      && typeof current === "object"
      && !Array.isArray(current)
    ) {
      result[key] = mergeMessages(current, value as DeepPartial<unknown>)
    } else {
      result[key] = value
    }
  }
  return result as T
}

/**
 * Resolve the active message catalog, applying host overrides on top of the
 * built-in locale defaults.
 */
export function resolveOfficeMessages(
  locale?: string | null,
  messages?: OfficeMessagesOverride,
): OfficeMessages {
  const base = resolveOfficeLocale(locale) === "zh-CN" ? officeMessagesZhCN : officeMessagesEn
  return mergeMessages(base, messages)
}

/** Replace `{name}` placeholders in a catalog message. */
export function formatOfficeMessage(
  template: string,
  values: Record<string, string | number>,
): string {
  return template.replace(/\{([^{}]+)\}/gu, (match, key: string) =>
    Object.prototype.hasOwnProperty.call(values, key) ? String(values[key]) : match,
  )
}

/** Format a date using the viewer locale (falls back to the host locale). */
export function formatOfficeDate(
  value: Date | string | number,
  locale?: string | null,
  options?: Intl.DateTimeFormatOptions,
): string {
  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) return ""
  return new Intl.DateTimeFormat(
    resolveOfficeLocale(locale),
    options ?? { dateStyle: "medium" },
  ).format(date)
}

/** Format a number using the viewer locale. */
export function formatOfficeNumber(
  value: number,
  locale?: string | null,
  options?: Intl.NumberFormatOptions,
): string {
  return new Intl.NumberFormat(resolveOfficeLocale(locale), options).format(value)
}
