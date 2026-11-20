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
  /** Toggle label: match case exactly. */
  caseSensitive: string
  /** Toggle label: match whole words only. */
  wholeWord: string
  /** Toggle label: treat the query as a regular expression. */
  regex: string
  /** Template with `{current}` and `{total}` placeholders. */
  resultOf: string
  /** Placeholder of the replace field. */
  replacePlaceholder: string
  /** Replace the active result. */
  replace: string
  /** Replace every result of the query. */
  replaceAll: string
  /** Template with a `{count}` placeholder: summary after a replace. */
  replaced: string
  /** Toggle label: match the whole cell instead of a substring. */
  wholeCell: string
  /** Toggle label: include formulas in the search and replace. */
  formulas: string
}

export interface OfficeZoomMessages {
  in: string
  out: string
  level: string
  reset: string
  /** Zoom-select preset: scale the page to the available width. */
  fitWidth: string
  /** Zoom-select preset: scale the page to fit the viewport height. */
  fitPage: string
  /** Zoom-select preset: 100% without fit scaling. */
  actualSize: string
  /** Percent label template with a `{value}` placeholder. */
  percent: string
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

/** Shared labels for the Office viewer/editor chrome (issue #158). */
export interface OfficeChromeMessages {
  /** Accessible name of `OfficeToolbar`; hosts can override with `label`. */
  toolbar: string
  /** Accessible name of `DocumentViewerSidebar`. */
  sidebar: string
  /** Accessible name of the sidebar toggle rendered by hosts. */
  sidebarToggle: string
  /** Accessible name of `OfficeStatusBar`. */
  statusBar: string
  /** Label prefix shown before a document file name. */
  fileName: string
  error: string
  warning: string
  retry: string
  emptyTitle: string
  emptyDescription: string
  errorTitle: string
  errorDescription: string
  loadingTitle: string
  loadingDescription: string
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
    heading: string
    listItem: string
    textBlock: string
    link: string
    formField: string
    /** Template with a `{name}` placeholder. */
    formFieldName: string
    annotation: string
    /** Template with a `{text}` placeholder. */
    annotationText: string
    /** Template with a `{id}` placeholder. */
    annotationById: string
    /** Template with a `{target}` placeholder. */
    actionTarget: string
    /** Template with a `{name}` placeholder. */
    tableByName: string
    /** Template with a `{name}` placeholder. */
    chartSeries: string
    /** Template with a `{number}` placeholder. */
    chartPoint: string
    /** Template with a `{number}` placeholder. */
    legendEntry: string
    /** Template with a `{name}` placeholder. */
    styleName: string
    /** Template with a `{name}` placeholder. */
    bookmarkName: string
    /** Template with a `{name}` placeholder. */
    fieldName: string
    /** Template with a `{number}` placeholder. */
    section: string
    /** Template with a `{author}` placeholder. */
    trackedChange: string
    /** Template with a `{state}` placeholder. */
    commentState: string
    /** Template with a `{name}` placeholder. */
    crossReference: string
    /** Template with a `{formula}` placeholder. */
    formulaLabel: string
    /** Template with a `{number}` placeholder. */
    conditionalFormat: string
    /** Template with a `{number}` placeholder. */
    dataValidation: string
    /** Template with a `{name}` placeholder. */
    namedRange: string
    protection: string
    /** Template with a `{number}` placeholder. */
    animation: string
    transition: string
    action: string
    master: string
    layout: string
    note: string
    /** Template with a `{name}` placeholder. */
    attachment: string
    /** Template with a `{name}` placeholder. */
    layer: string
    /** Template with a `{name}` placeholder. */
    signature: string
    permission: string
  }
  chrome: OfficeChromeMessages
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
    /** Shown for OpenDocument text, which has no DOCX writer. */
    viewOnlyReason: string
    live: {
      loading: string
      error: string
      ready: string
      /** Template with `{page}` and `{total}` placeholders. */
      readyPages: string
    }
    /** Template with `{count}` and `{allowed}` placeholders. */
    pageLimitExceeded: string
    /** Template with a `{name}` placeholder for an unresolved internal link. */
    bookmarkMissing: string
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
    /** Label for the inline/balloons/none revision display control. */
    revisionDisplay: string
    revisionDisplayInline: string
    revisionDisplayBalloons: string
    revisionDisplayNone: string
    /** Page layout control: continuous scroll, one page, or spreads. */
    viewMode: string
    viewModeContinuous: string
    viewModeSingle: string
    viewModeSpread: string
    /** Fit-to-viewport toggles. */
    fitWidth: string
    fitPage: string
    /** Trigger for the compact toolbar overflow menu. */
    moreActions: string
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
    find: string
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
    columns: string
    columnsOne: string
    columnsTwo: string
    columnsThree: string
    /** Template with a `{count}` placeholder. */
    columnsMore: string
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
    /** Tooltip of the drag handle that floats a table. */
    moveTableHandle: string
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
    addComment: string
    addCommentTitle: string
    /** Default author name stored on comments created without one. */
    commentAuthor: string
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
      /** Shown when an OpenDocument file is opened (view-only). */
      viewOnly: string
      /** Template with a `{name}` placeholder for an OpenDocument import. */
      loadedViewOnly: string
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
      commentAdded: string
      commentReplied: string
      commentResolved: string
      commentReopened: string
      commentDeleted: string
      /** The selection or comment a comment command targets is gone. */
      commentTargetMissing: string
      commentTextRequired: string
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
      /** The image is not floating, so it has no layer order. */
      imageLayerOrderRequiresFloating: string
      imageLayerOrderChanged: string
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
      resizedTableRow: string
      tableRowHeightPositive: string
      tableRowHeightTargetUnavailable: string
      floatedTable: string
      inlinedTable: string
      cannotFloatTableNotFound: string
      movedBodyNode: string
      cannotMoveBodyNodeNotFound: string
      updatedFormField: string
      formFieldTargetUnavailable: string
      formFieldOptionsRequired: string
      /** Template with a `{preset}` placeholder. */
      unsupportedDiagonalTableCellBorders: string
      /** Template with a `{preset}` placeholder. */
      unsupportedDiagonalParagraphBorders: string
      revisionLossOne: string
      /** Template with a `{count}` placeholder. */
      revisionLossMany: string
      pageMarginsUpdated: string
      pageSizeUpdated: string
      sectionColumnsUpdated: string
      headerAdded: string
      footerAdded: string
      /** Template with a `{count}` placeholder. */
      replacedMatches: string
      replaceTargetUnavailable: string
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
  pageSetup: {
    title: string
    open: string
    /** Template with an `{index}` placeholder. */
    sectionOption: string
    sectionLabel: string
    paperSize: string
    paperSizeCustom: string
    presetLetter: string
    presetLegal: string
    presetTabloid: string
    presetA3: string
    presetA4: string
    presetA5: string
    presetB5: string
    orientation: string
    orientationPortrait: string
    orientationLandscape: string
    width: string
    height: string
    margins: string
    marginTop: string
    marginRight: string
    marginBottom: string
    marginLeft: string
    units: string
    apply: string
    cancel: string
    close: string
    invalidValue: string
  }
  headerFooter: {
    header: string
    footer: string
    editHeader: string
    editFooter: string
    addHeader: string
    addFooter: string
    closeEditing: string
    editingHeader: string
    editingFooter: string
    placeholder: string
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
    selectTable: string
    clearContents: string
    layer: string
    bringToFront: string
    bringForward: string
    sendBackward: string
    sendToBack: string
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
  formFieldDialog: {
    title: string
    fieldType: string
    name: string
    namePlaceholder: string
    value: string
    placeholder: string
    required: string
    checked: string
    options: string
    option: string
    optionValue: string
    addOption: string
    removeOption: string
    maxLength: string
    save: string
    cancel: string
    close: string
    noOptions: string
    types: {
      text: string
      checkbox: string
      dropdown: string
      date: string
    }
  }
  linkCard: {
    label: string
    open: string
    edit: string
    remove: string
    save: string
    cancel: string
    url: string
    urlPlaceholder: string
    internal: string
  }
  stylePreview: {
    label: string
    sample: string
    apply: string
    normal: string
    heading: string
  }
  nodeDrag: {
    handle: string
    handleTitle: string
    moveUp: string
    moveDown: string
    dropBefore: string
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
    reply: string
    replyTitle: string
    resolve: string
    resolveTitle: string
    reopen: string
    reopenTitle: string
    delete: string
    deleteTitle: string
    composerTitle: string
    composerReplyTitle: string
    composerPlaceholder: string
    composerReplyPlaceholder: string
    composerSubmit: string
    composerCancel: string
  }
  dragOverlay: {
    drop: string
  }
  diagnostics: {
    imageDecodeFailed: string
    numberingImageDecodeFailed: string
    /** Template with a `{feature}` placeholder. */
    unsupportedImageFeature: string
  }
  /**
   * Labels for drawing objects the model cannot paint itself (shapes, charts,
   * SmartArt, embedded objects, 3D models) and the equations it shows as text.
   */
  drawing: {
    textBox: string
    shape: string
    wordArt: string
    group: string
    smartArt: string
    chart: string
    oleObject: string
    model3d: string
    equation: string
    /** Template with a `{progId}` placeholder. */
    embeddedObject: string
    /** Template with a `{count}` placeholder. */
    chartSeries: string
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
    /** Rejected drop: the file type is outside the workbook accept list. */
    unsupportedDrop: string
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
    printPreview: string
    pageBreaks: string
  }
  statusBar: {
    /** Accessible label of the selected cell/range address. */
    range: string
    /** Shown when nothing is selected. */
    ready: string
    count: string
    numericalCount: string
    sum: string
    average: string
    min: string
    max: string
    /** Multi-cell selection without filtered rows. */
    selected: string
    /** The selection contains rows hidden by the active AutoFilter. */
    filtered: string
    /** Template with a `{count}` placeholder. */
    truncated: string
  }
  print: {
    title: string
    close: string
    print: string
    noPages: string
    /** Template with `{count}`, `{paper}`, `{orientation}` and `{scale}` placeholders. */
    summary: string
    orientationPortrait: string
    orientationLandscape: string
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
    rename: string
    renameLabel: string
    renameEmpty: string
    renameTooLong: string
    renameDuplicate: string
    renameApostrophe: string
    renameFailed: string
    moveFailed: string
    menuLabel: string
    hide: string
    hideFailed: string
    unhide: string
    unhideFailed: string
    hiddenTab: string
    noHiddenSheets: string
    tabListLabel: string
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
    duplicate: string
    delete: string
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
  autoFilter: {
    label: string
    apply: string
    blanks: string
    clear: string
    filterByValue: string
    noValues: string
    openMenu: string
    selectAll: string
  }
  formulaBar: {
    activeCell: string
    /** Tooltip/accessible label of the name box that navigates to a name. */
    goToName: string
    /** Accessible label of the defined-name list. */
    definedNames: string
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
    /** Template with a `{name}` placeholder for a slicer card. */
    slicerHint: string
  }
  pivot: {
    /** Template with a `{name}` placeholder; the cached output is static. */
    staticPreview: string
  }
  selection: {
    /** Template with an `{address}` placeholder. */
    current: string
    none: string
    /** Template with an `{address}` placeholder for a pending cut. */
    cutPending: string
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
    /** Handle at the selection's bottom-right that drag-fills cells. */
    fillHandle: string
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
    /** Accessible name of the viewer toolbar. */
    toolbar: string
    /** Browse layout control: single slide list, slide-sorter grid or present. */
    viewMode: string
    viewModeSingle: string
    viewModeSorter: string
    viewModePresent: string
    /** Leaves presentation mode and returns to the single browse layout. */
    viewModeBrowse: string
    /** Toolbar toggle for the speaker-notes / presenter rail. */
    notesToggle: string
    /** Accessible name of the slide-sorter grid. */
    sorter: string
    searchPlaceholder: string
    search: string
    sidebar: string
    empty: string
    loading: string
    waiting: string
    cancelled: string
    /** Password-protected or legacy binary (OLE/CFB) presentations. */
    encryptedTitle: string
    encryptedDescription: string
    /** Rejected drop: the file type is not a PowerPoint presentation. */
    unsupportedDrop: string
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
    /** Built-in right-click link menu and its two entries. */
    link: {
      /** Accessible name of the menu. */
      menu: string
      open: string
      copy: string
      /** Template with a `{number}` placeholder. */
      openSlide: string
    }
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
    notes: {
      /** Accessible name of the speaker-notes panel in browse mode. */
      panel: string
      /** Accessible name of the presenter rail in presentation mode. */
      presenter: string
      title: string
      /** Shown when the current slide has no notes. */
      empty: string
      /** Heading above the next-slide preview. */
      next: string
      /** Shown when the current slide is the last one. */
      noNext: string
      /** Accessible name of the elapsed-time readout. */
      timer: string
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
    /** Accessible name of the thumbnail loading skeleton. */
    loading: string
    /** Shown when the thumbnail renderer rejects. */
    failed: string
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
    /** Page layout control label and its three modes. */
    pageLayout: string
    pageLayoutContinuous: string
    pageLayoutSingle: string
    pageLayoutSpread: string
    pageInput: string
    rotateClockwise: string
    searchLabel: string
    searchPlaceholder: string
    search: string
    download: string
    /** Print the whole document through the browser print dialog. */
    print: string
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
    /** Rejected drop: the file type is not a PDF document. */
    unsupportedDrop: string
    selectionLayer: string
    /** A password is required to unlock an encrypted document. */
    passwordRequired: string
    /** Accessible name of the floating text-selection toolbar. */
    selectionActions: string
    copySelection: string
    copiedSelection: string
    /** Accessible name of the link hover card. */
    linkActions: string
    openLink: string
    /** Title of the annotation details dialog. */
    annotationDetails: string
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
    /** Accessible name of the bookmark/outline navigation panel. */
    outline: string
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
    outline: string
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
    regionAnnotator: {
      defaultAriaLabel: string
      toolbarLabel: string
      arrow: string
      rectangle: string
      text: string
      highlight: string
      color: string
      remove: string
      clear: string
      empty: string
      imageAlt: string
      textPlaceholder: string
      textSave: string
      textCancel: string
      hint: string
      /** Template with a `{count}` placeholder. */
      count: string
    }
    boundingBox: {
      noFields: string
      imageAlt: string
    }
    fileUpload: {
      prompt: string
      title: string
      browse: string
      drop: string
      ready: string
      rejected: string
      unknownType: string
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
      caseSensitive: "Match case",
      wholeWord: "Match whole word",
      regex: "Use regular expression",
      resultOf: "{current} of {total}",
      replacePlaceholder: "Replace with",
      replace: "Replace",
      replaceAll: "Replace all",
      replaced: "Replaced {count}",
      wholeCell: "Match entire cell",
      formulas: "Include formulas",
    },
    zoom: {
      in: "Zoom in",
      out: "Zoom out",
      level: "Zoom level",
      reset: "Fit to window",
      fitWidth: "Fit width",
      fitPage: "Fit page",
      actualSize: "Actual size",
      percent: "{value}%",
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
      heading: "Heading",
      listItem: "List item",
      textBlock: "Text block",
      link: "Link",
      formField: "Form field",
      formFieldName: "Form field: {name}",
      annotation: "Annotation",
      annotationText: "Annotation: {text}",
      annotationById: "Annotation {id}",
      actionTarget: "Action: {target}",
      tableByName: "Table {name}",
      chartSeries: "Series: {name}",
      chartPoint: "Point {number}",
      legendEntry: "Legend entry {number}",
      styleName: "Style: {name}",
      bookmarkName: "Bookmark: {name}",
      fieldName: "Field: {name}",
      section: "Section {number}",
      trackedChange: "Revision by {author}",
      commentState: "Comment state: {state}",
      crossReference: "Cross-reference: {name}",
      formulaLabel: "Formula: {formula}",
      conditionalFormat: "Conditional format {number}",
      dataValidation: "Data validation {number}",
      namedRange: "Named range: {name}",
      protection: "Sheet protection",
      animation: "Animation {number}",
      transition: "Slide transition",
      action: "Action",
      master: "Slide master",
      layout: "Slide layout",
      note: "Speaker notes",
      attachment: "Attachment: {name}",
      layer: "Optional content: {name}",
      signature: "Signature: {name}",
      permission: "Permissions",
    },
    chrome: {
      toolbar: "Document toolbar",
      sidebar: "Document pages",
      sidebarToggle: "Toggle document pages",
      statusBar: "Document status",
      fileName: "File",
      error: "Error",
      warning: "Warning",
      retry: "Try again",
      emptyTitle: "No document open",
      emptyDescription: "Open a file to get started.",
      errorTitle: "Unable to open this document",
      errorDescription: "Check the file and try again.",
      loadingTitle: "Opening document",
      loadingDescription: "Preparing the document…",
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
      unsupportedFile: "Only Word (.docx, .docm, .dotx, .dotm) and OpenDocument text (.odt, .ott, .odm, .otm, .oth, .fodt) files are supported",
      fileTooLarge: "DOCX file size {actual} bytes exceeds the allowed maximum of {allowed} bytes.",
      viewOnlyReason: "OpenDocument files are view-only: they cannot be edited or saved as DOCX.",
      live: {
        loading: "Opening DOCX document",
        error: "Unable to open document: {message}",
        ready: "Document ready",
        readyPages: "Document ready, page {page} of {total}",
      },
      pageLimitExceeded:
        "DOCX page count {count} exceeds the allowed maximum of {allowed}.",
      bookmarkMissing: "Bookmark \"{name}\" was not found in this document.",
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
      revisionDisplay: "Display for review",
      revisionDisplayInline: "All markup",
      revisionDisplayBalloons: "Balloons",
      revisionDisplayNone: "No markup",
      viewMode: "Page layout",
      viewModeContinuous: "Continuous scroll",
      viewModeSingle: "Single page",
      viewModeSpread: "Two-page spread",
      fitWidth: "Fit width",
      fitPage: "Fit page",
      moreActions: "More toolbar actions",
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
      find: "Find and replace",
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
      columns: "Columns",
      columnsOne: "1 column",
      columnsTwo: "2 columns",
      columnsThree: "3 columns",
      columnsMore: "{count} columns",
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
      moveTableHandle: "Drag to move or float this table",
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
      addComment: "New comment",
      addCommentTitle: "Comment on the selection",
      commentAuthor: "Author",
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
        unsupportedFile: "Only Word (.docx, .docm, .dotx, .dotm) and OpenDocument text (.odt, .ott, .odm, .otm, .oth, .fodt) files are supported",
        viewOnly: "OpenDocument files are view-only: they cannot be edited or saved as DOCX.",
        loadedViewOnly: "Loaded {name} (view-only)",
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
        commentAdded: "Comment added",
        commentReplied: "Reply added",
        commentResolved: "Comment thread resolved",
        commentReopened: "Comment thread reopened",
        commentDeleted: "Comment deleted",
        commentTargetMissing: "Unsupported: the comment or selection is no longer available",
        commentTextRequired: "Unsupported: a comment needs text",
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
        imageLayerOrderRequiresFloating: "Unsupported: layer order applies to floating images only",
        imageLayerOrderChanged: "Changed image layer order",
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
        resizedTableRow: "Resized table row",
        tableRowHeightPositive: "Unsupported: row height must be positive",
        tableRowHeightTargetUnavailable: "Unsupported: table row target unavailable",
        floatedTable: "Floated table",
        inlinedTable: "Table back in the text flow",
        cannotFloatTableNotFound: "Cannot float: table not found",
        movedBodyNode: "Moved paragraph or table",
        cannotMoveBodyNodeNotFound: "Cannot move: paragraph or table not found",
        updatedFormField: "Updated form field",
        formFieldTargetUnavailable: "Unsupported: form field target unavailable",
        formFieldOptionsRequired: "Unsupported: a dropdown needs at least one option",
        unsupportedDiagonalTableCellBorders: "Unsupported: diagonal table-cell borders ({preset})",
        unsupportedDiagonalParagraphBorders: "Unsupported: diagonal paragraph borders ({preset})",
        revisionLossOne: "Tracked changes in an edited paragraph were resolved with the edit",
        revisionLossMany: "Tracked changes in {count} edited paragraphs were resolved with the edit",
        pageMarginsUpdated: "Page margins updated",
        pageSizeUpdated: "Paper size and orientation updated",
        sectionColumnsUpdated: "Section columns updated",
        headerAdded: "Header added",
        footerAdded: "Footer added",
        replacedMatches: "Replaced {count}",
        replaceTargetUnavailable: "Unsupported: this result is in a header, footer or note and cannot be replaced",
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
    pageSetup: {
      title: "Page setup",
      open: "Page setup",
      sectionOption: "Section {index}",
      sectionLabel: "Apply to",
      paperSize: "Paper size",
      paperSizeCustom: "Custom",
      presetLetter: "Letter (8.5 × 11 in)",
      presetLegal: "Legal (8.5 × 14 in)",
      presetTabloid: "Tabloid (11 × 17 in)",
      presetA3: "A3 (297 × 420 mm)",
      presetA4: "A4 (210 × 297 mm)",
      presetA5: "A5 (148 × 210 mm)",
      presetB5: "B5 (176 × 250 mm)",
      orientation: "Orientation",
      orientationPortrait: "Portrait",
      orientationLandscape: "Landscape",
      width: "Paper width",
      height: "Paper height",
      margins: "Margins",
      marginTop: "Top margin",
      marginRight: "Right margin",
      marginBottom: "Bottom margin",
      marginLeft: "Left margin",
      units: "Values are in inches",
      apply: "Apply",
      cancel: "Cancel",
      close: "Close page setup",
      invalidValue: "Enter a non-negative number",
    },
    headerFooter: {
      header: "Header",
      footer: "Footer",
      editHeader: "Edit header",
      editFooter: "Edit footer",
      addHeader: "Add header",
      addFooter: "Add footer",
      closeEditing: "Done editing",
      editingHeader: "Editing header",
      editingFooter: "Editing footer",
      placeholder: "Type header text",
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
      selectTable: "Select table",
      clearContents: "Clear contents",
      layer: "Layer",
      bringToFront: "Bring to Front",
      bringForward: "Bring Forward",
      sendBackward: "Send Backward",
      sendToBack: "Send to Back",
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
    formFieldDialog: {
      title: "Form field properties",
      fieldType: "Type",
      name: "Name",
      namePlaceholder: "Field name",
      value: "Value",
      placeholder: "Placeholder",
      required: "Required",
      checked: "Checked",
      options: "Options",
      option: "Display text",
      optionValue: "Value",
      addOption: "Add option",
      removeOption: "Remove option",
      maxLength: "Maximum length",
      save: "Save",
      cancel: "Cancel",
      close: "Close form field properties",
      noOptions: "This dropdown has no options yet.",
      types: {
        text: "Text",
        checkbox: "Check box",
        dropdown: "Dropdown",
        date: "Date",
      },
    },
    linkCard: {
      label: "Link actions",
      open: "Open link",
      edit: "Edit link",
      remove: "Remove link",
      save: "Save link",
      cancel: "Cancel link edit",
      url: "Address",
      urlPlaceholder: "https://example.com",
      internal: "Links to a place in this document",
    },
    stylePreview: {
      label: "Paragraph styles",
      sample: "The quick brown fox jumps over the lazy dog",
      apply: "Apply {style}",
      normal: "Normal",
      heading: "Heading",
    },
    nodeDrag: {
      handle: "Reorder",
      handleTitle: "Drag to move this paragraph or table",
      moveUp: "Move up",
      moveDown: "Move down",
      dropBefore: "Drop here",
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
      reply: "Reply",
      replyTitle: "Reply to this comment",
      resolve: "Resolve",
      resolveTitle: "Resolve this comment thread",
      reopen: "Reopen",
      reopenTitle: "Reopen this comment thread",
      delete: "Delete",
      deleteTitle: "Delete this comment",
      composerTitle: "New comment",
      composerReplyTitle: "Reply",
      composerPlaceholder: "Write a comment…",
      composerReplyPlaceholder: "Write a reply…",
      composerSubmit: "Comment",
      composerCancel: "Cancel",
    },
    dragOverlay: {
      drop: "Drop DOCX file to import",
    },
    diagnostics: {
      imageDecodeFailed: "DOCX image could not be decoded.",
      numberingImageDecodeFailed: "DOCX numbering image could not be decoded.",
      unsupportedImageFeature: "DOCX SVG picture uses unsupported {feature} content; showing a placeholder.",
    },
    drawing: {
      textBox: "Text box",
      shape: "Shape",
      wordArt: "WordArt",
      group: "Grouped shapes",
      smartArt: "SmartArt",
      chart: "Chart",
      oleObject: "Embedded object",
      model3d: "3D model",
      equation: "Equation",
      embeddedObject: "Embedded object ({progId})",
      chartSeries: "{count} series",
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
      unsupportedDrop: "Only .xlsx, .xls, .xlsm, .xlsb, .xltx, .xltm and .csv files are supported",
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
      printPreview: "Print preview",
      pageBreaks: "Page break preview",
    },
    statusBar: {
      range: "Selected range",
      ready: "Ready",
      count: "Count",
      numericalCount: "Numerical count",
      sum: "Sum",
      average: "Average",
      min: "Min",
      max: "Max",
      selected: "Selected",
      filtered: "Filtered",
      truncated: "Showing the first {count} cells",
    },
    print: {
      title: "Print preview",
      close: "Close print preview",
      print: "Print",
      noPages: "This sheet has no printable pages",
      summary: "{count} pages · {paper} {orientation} · {scale}%",
      orientationPortrait: "portrait",
      orientationLandscape: "landscape",
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
      rename: "Rename",
      renameLabel: "Sheet name",
      renameEmpty: "Enter a sheet name.",
      renameTooLong: "Sheet names cannot be longer than 31 characters.",
      renameDuplicate: "A sheet with that name already exists.",
      renameApostrophe: "Sheet names cannot start or end with an apostrophe.",
      renameFailed: "The sheet could not be renamed.",
      moveFailed: "The sheet could not be moved.",
      menuLabel: "Sheet menu",
      hide: "Hide",
      hideFailed: "The sheet could not be hidden.",
      unhide: "Unhide…",
      unhideFailed: "The sheet could not be unhidden.",
      hiddenTab: "Hidden sheet",
      noHiddenSheets: "No hidden sheets.",
      tabListLabel: "Sheets",
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
      duplicate: "Duplicate",
      delete: "Delete",
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
    autoFilter: {
      label: "Filter",
      apply: "Apply",
      blanks: "(Blanks)",
      clear: "Clear filter",
      filterByValue: "Filter by value",
      noValues: "No values to filter.",
      openMenu: "Filter column {column}",
      selectAll: "Select all",
    },
    formulaBar: {
      activeCell: "Active cell",
      goToName: "Go to a cell, range or defined name",
      definedNames: "Defined names",
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
      slicerHint: "Slicer \u201c{name}\u201d: static preview. Change the selection in Excel.",
    },
    pivot: {
      staticPreview: "Pivot table \u201c{name}\u201d: static cached values. Refresh it in Excel.",
    },
    selection: {
      current: "Current selection {address}",
      none: "No cells selected",
      cutPending: "Cut {address}, press Escape to cancel",
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
      fillHandle: "Fill handle: drag to fill the selection's pattern into more cells",
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
      toolbar: "Presentation toolbar",
      viewMode: "Browse layout",
      viewModeSingle: "Single slide view",
      viewModeSorter: "Slide sorter",
      viewModePresent: "Present",
      viewModeBrowse: "Exit presentation",
      notesToggle: "Speaker notes",
      sorter: "Slide sorter",
      searchPlaceholder: "Search slides",
      search: "Search",
      sidebar: "Slides",
      empty: "Open a PPTX file to get started",
      loading: "Loading PPTX…",
      waiting: "Waiting for the slide stage…",
      cancelled: "PPTX loading was cancelled.",
      encryptedTitle: "Presentation is protected",
      encryptedDescription: "This presentation is password-protected or saved in the legacy binary format. Remove the password or save it as .pptx and try again.",
      unsupportedDrop: "Only .pptx, .pptm, .ppsx, .ppsm, .potx and .potm files are supported",
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
      link: {
        menu: "Link actions",
        open: "Open link",
        copy: "Copy link",
        openSlide: "Go to slide {number}",
      },
      capability: {
        strict: "Exact {count}/{total}",
        approximate: "Approximate {count}",
        static: "Static {count}",
        unparsed: "Unparsed {count}",
      },
      notes: {
        panel: "Speaker notes",
        presenter: "Presenter view",
        title: "Notes",
        empty: "No speaker notes for this slide.",
        next: "Next slide",
        noNext: "Last slide",
        timer: "Elapsed time",
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
      loading: "Loading thumbnail",
      failed: "Thumbnail unavailable",
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
      pageLayout: "Page layout",
      pageLayoutContinuous: "Continuous scroll",
      pageLayoutSingle: "Single page",
      pageLayoutSpread: "Two-page spread",
      pageInput: "Page number",
      rotateClockwise: "Rotate clockwise",
      searchLabel: "Search PDF text",
      searchPlaceholder: "Search text",
      search: "Search",
      download: "Download PDF",
      print: "Print PDF",
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
      unsupportedDrop: "Only PDF files are supported",
      selectionLayer: "PDF page text selection layer",
      passwordRequired: "Enter the document password to unlock this PDF.",
      selectionActions: "Selection actions",
      copySelection: "Copy selection",
      copiedSelection: "Copied",
      linkActions: "Link preview",
      openLink: "Open link",
      annotationDetails: "Annotation details",
    },
    surface: {
      noPdf: "No PDF loaded",
      opening: "Opening PDF…",
      renderingPage: "Rendering page {page}…",
      pageRenderFailed: "Failed to render page {page}",
      pageAlt: "Page {page}",
      selectionRegion: "Choose PDF page region",
      selectionText: "PDF page text selection layer",
      outline: "Document outline",
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
      outline: "PDF outline",
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
    regionAnnotator: {
      defaultAriaLabel: "Region annotation editor",
      toolbarLabel: "Annotation tools",
      arrow: "Arrow",
      rectangle: "Rectangle",
      text: "Text",
      highlight: "Highlight",
      color: "Annotation color",
      remove: "Remove selected annotation",
      clear: "Clear annotations",
      empty: "Capture a region to annotate it.",
      imageAlt: "Captured region preview",
      textPlaceholder: "Annotation text",
      textSave: "Add text",
      textCancel: "Cancel text",
      hint: "Drag on the image to draw. Select an annotation and press Delete to remove it.",
      count: "{count} annotations",
    },
    boundingBox: {
      noFields: "No fields to display",
      imageAlt: "Citation source preview",
    },
    fileUpload: {
      prompt: "Drag & drop files here, or click to select",
      title: "Click to upload or drop files",
      browse: "Browse files",
      drop: "Drop to add",
      ready: "Ready",
      rejected: "This file type is not supported here.",
      unknownType: "Unknown type",
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
      caseSensitive: "区分大小写",
      wholeWord: "全字匹配",
      regex: "使用正则表达式",
      resultOf: "第 {current} 项，共 {total} 项",
      replacePlaceholder: "替换为",
      replace: "替换",
      replaceAll: "全部替换",
      replaced: "已替换 {count} 处",
      wholeCell: "匹配整个单元格",
      formulas: "包含公式",
    },
    zoom: {
      in: "放大",
      out: "缩小",
      level: "缩放比例",
      reset: "恢复适合窗口",
      fitWidth: "适合宽度",
      fitPage: "适合页面",
      actualSize: "实际大小",
      percent: "{value}%",
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
      heading: "标题",
      listItem: "列表项",
      textBlock: "文本块",
      link: "链接",
      formField: "表单控件",
      formFieldName: "表单控件：{name}",
      annotation: "批注标记",
      annotationText: "批注标记：{text}",
      annotationById: "批注标记 {id}",
      actionTarget: "动作：{target}",
      tableByName: "表格 {name}",
      chartSeries: "系列：{name}",
      chartPoint: "数据点 {number}",
      legendEntry: "图例项 {number}",
      styleName: "样式：{name}",
      bookmarkName: "书签：{name}",
      fieldName: "域：{name}",
      section: "第 {number} 节",
      trackedChange: "{author} 的修订",
      commentState: "批注状态：{state}",
      crossReference: "交叉引用：{name}",
      formulaLabel: "公式：{formula}",
      conditionalFormat: "条件格式 {number}",
      dataValidation: "数据验证 {number}",
      namedRange: "命名区域：{name}",
      protection: "工作表保护",
      animation: "动画 {number}",
      transition: "幻灯片切换",
      action: "动作",
      master: "幻灯片母版",
      layout: "幻灯片版式",
      note: "演讲者备注",
      attachment: "附件：{name}",
      layer: "可选内容组：{name}",
      signature: "签名：{name}",
      permission: "权限",
    },
    chrome: {
      toolbar: "文档工具栏",
      sidebar: "文档页面",
      sidebarToggle: "切换文档页面",
      statusBar: "文档状态",
      fileName: "文件",
      error: "错误",
      warning: "警告",
      retry: "重试",
      emptyTitle: "未打开文档",
      emptyDescription: "打开一个文件开始。",
      errorTitle: "无法打开此文档",
      errorDescription: "请检查文件后重试。",
      loadingTitle: "正在打开文档",
      loadingDescription: "正在准备文档…",
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
      unsupportedFile: "仅支持 Word（.docx、.docm、.dotx、.dotm）和 OpenDocument 文本（.odt、.ott、.odm、.otm、.oth、.fodt）文件",
      fileTooLarge: "DOCX 文件大小 {actual} 字节超过允许值 {allowed} 字节。",
      viewOnlyReason: "OpenDocument 文件为只读：无法编辑，也无法另存为 DOCX。",
      live: {
        loading: "正在打开 DOCX 文档",
        error: "无法打开文档：{message}",
        ready: "文档已就绪",
        readyPages: "文档已就绪，第 {page} 页，共 {total} 页",
      },
      pageLimitExceeded: "DOCX 页数 {count} 超过允许值 {allowed}。",
      bookmarkMissing: "未找到书签“{name}”。",
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
      revisionDisplay: "审阅显示方式",
      revisionDisplayInline: "所有标记",
      revisionDisplayBalloons: "批注框",
      revisionDisplayNone: "无标记",
      viewMode: "页面布局",
      viewModeContinuous: "连续滚动",
      viewModeSingle: "单页",
      viewModeSpread: "双页",
      fitWidth: "适应宽度",
      fitPage: "适应页面",
      moreActions: "更多工具栏操作",
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
      find: "查找和替换",
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
      columns: "分栏",
      columnsOne: "1 栏",
      columnsTwo: "2 栏",
      columnsThree: "3 栏",
      columnsMore: "{count} 栏",
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
      moveTableHandle: "拖动以移动或浮动表格",
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
      addComment: "新建批注",
      addCommentTitle: "为所选内容添加批注",
      commentAuthor: "批注人",
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
        unsupportedFile: "仅支持 Word（.docx、.docm、.dotx、.dotm）和 OpenDocument 文本（.odt、.ott、.odm、.otm、.oth、.fodt）文件",
        viewOnly: "OpenDocument 文件为只读：无法编辑，也无法另存为 DOCX。",
        loadedViewOnly: "已加载 {name}（只读）",
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
        commentAdded: "已添加批注",
        commentReplied: "已添加回复",
        commentResolved: "批注已标记为已解决",
        commentReopened: "批注已重新打开",
        commentDeleted: "已删除批注",
        commentTargetMissing: "不支持：批注或选区已不存在",
        commentTextRequired: "不支持：批注内容不能为空",
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
        imageLayerOrderRequiresFloating: "不支持：图层顺序仅适用于浮动图片",
        imageLayerOrderChanged: "已更改图片图层顺序",
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
        resizedTableRow: "已调整表格行高",
        tableRowHeightPositive: "不支持：表格行高必须为正数",
        tableRowHeightTargetUnavailable: "不支持：表格行目标不可用",
        floatedTable: "已将表格浮动",
        inlinedTable: "表格已回到文字流",
        cannotFloatTableNotFound: "无法浮动：未找到表格",
        movedBodyNode: "已移动段落或表格",
        cannotMoveBodyNodeNotFound: "无法移动：未找到段落或表格",
        updatedFormField: "已更新表单域",
        formFieldTargetUnavailable: "不支持：表单域目标不可用",
        formFieldOptionsRequired: "不支持：下拉框至少需要一个选项",
        unsupportedDiagonalTableCellBorders: "不支持：对角线表格单元格边框（{preset}）",
        unsupportedDiagonalParagraphBorders: "不支持：对角线段落边框（{preset}）",
        revisionLossOne: "已编辑段落中的修订已随编辑一起处理",
        revisionLossMany: "已编辑的 {count} 个段落中的修订已随编辑一起处理",
        pageMarginsUpdated: "已更新页边距",
        pageSizeUpdated: "已更新纸张大小与方向",
        sectionColumnsUpdated: "已更新分栏",
        headerAdded: "已添加页眉",
        footerAdded: "已添加页脚",
        replacedMatches: "已替换 {count} 处",
        replaceTargetUnavailable: "不支持：该结果位于页眉、页脚或注释中，无法替换",
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
    pageSetup: {
      title: "页面设置",
      open: "页面设置",
      sectionOption: "第 {index} 节",
      sectionLabel: "应用于",
      paperSize: "纸张大小",
      paperSizeCustom: "自定义",
      presetLetter: "Letter（8.5 × 11 英寸）",
      presetLegal: "Legal（8.5 × 14 英寸）",
      presetTabloid: "Tabloid（11 × 17 英寸）",
      presetA3: "A3（297 × 420 毫米）",
      presetA4: "A4（210 × 297 毫米）",
      presetA5: "A5（148 × 210 毫米）",
      presetB5: "B5（176 × 250 毫米）",
      orientation: "方向",
      orientationPortrait: "纵向",
      orientationLandscape: "横向",
      width: "纸张宽度",
      height: "纸张高度",
      margins: "页边距",
      marginTop: "上边距",
      marginRight: "右边距",
      marginBottom: "下边距",
      marginLeft: "左边距",
      units: "数值单位为英寸",
      apply: "应用",
      cancel: "取消",
      close: "关闭页面设置",
      invalidValue: "请输入非负数",
    },
    headerFooter: {
      header: "页眉",
      footer: "页脚",
      editHeader: "编辑页眉",
      editFooter: "编辑页脚",
      addHeader: "添加页眉",
      addFooter: "添加页脚",
      closeEditing: "完成编辑",
      editingHeader: "正在编辑页眉",
      editingFooter: "正在编辑页脚",
      placeholder: "输入页眉文字",
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
      selectTable: "选择整个表格",
      clearContents: "清除内容",
      layer: "图层",
      bringToFront: "置于顶层",
      bringForward: "上移一层",
      sendBackward: "下移一层",
      sendToBack: "置于底层",
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
    formFieldDialog: {
      title: "表单域属性",
      fieldType: "类型",
      name: "名称",
      namePlaceholder: "字段名称",
      value: "值",
      placeholder: "占位提示",
      required: "必填",
      checked: "已选中",
      options: "选项",
      option: "显示文本",
      optionValue: "值",
      addOption: "添加选项",
      removeOption: "删除选项",
      maxLength: "最大长度",
      save: "保存",
      cancel: "取消",
      close: "关闭表单域属性",
      noOptions: "此下拉框还没有选项。",
      types: {
        text: "文本",
        checkbox: "复选框",
        dropdown: "下拉框",
        date: "日期",
      },
    },
    linkCard: {
      label: "链接操作",
      open: "打开链接",
      edit: "编辑链接",
      remove: "删除链接",
      save: "保存链接",
      cancel: "取消编辑链接",
      url: "地址",
      urlPlaceholder: "https://example.com",
      internal: "链接到本文档中的位置",
    },
    stylePreview: {
      label: "段落样式",
      sample: "敏捷的棕色狐狸跳过懒狗",
      apply: "应用{style}",
      normal: "正文",
      heading: "标题",
    },
    nodeDrag: {
      handle: "重新排序",
      handleTitle: "拖动以移动此段落或表格",
      moveUp: "上移",
      moveDown: "下移",
      dropBefore: "放置到此处",
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
      reply: "回复",
      replyTitle: "回复此批注",
      resolve: "标记为已解决",
      resolveTitle: "将此批注标记为已解决",
      reopen: "重新打开",
      reopenTitle: "重新打开此批注",
      delete: "删除",
      deleteTitle: "删除此批注",
      composerTitle: "新建批注",
      composerReplyTitle: "回复",
      composerPlaceholder: "输入批注…",
      composerReplyPlaceholder: "输入回复…",
      composerSubmit: "批注",
      composerCancel: "取消",
    },
    dragOverlay: {
      drop: "拖放 DOCX 文件以导入",
    },
    diagnostics: {
      imageDecodeFailed: "DOCX 图片无法解码。",
      numberingImageDecodeFailed: "DOCX 编号图片无法解码。",
      unsupportedImageFeature: "DOCX SVG 图片包含不支持的 {feature} 内容，已显示占位符。",
    },
    drawing: {
      textBox: "文本框",
      shape: "形状",
      wordArt: "艺术字",
      group: "组合形状",
      smartArt: "SmartArt",
      chart: "图表",
      oleObject: "嵌入对象",
      model3d: "3D 模型",
      equation: "公式",
      embeddedObject: "嵌入对象（{progId}）",
      chartSeries: "{count} 个系列",
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
      unsupportedDrop: "仅支持 .xlsx、.xls、.xlsm、.xlsb、.xltx、.xltm 和 .csv 文件",
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
      printPreview: "打印预览",
      pageBreaks: "分页预览",
    },
    statusBar: {
      range: "选定区域",
      ready: "就绪",
      count: "计数",
      numericalCount: "数值计数",
      sum: "求和",
      average: "平均值",
      min: "最小值",
      max: "最大值",
      selected: "已选定",
      filtered: "已筛选",
      truncated: "仅显示前 {count} 个单元格",
    },
    print: {
      title: "打印预览",
      close: "关闭打印预览",
      print: "打印",
      noPages: "此工作表没有可打印的页面",
      summary: "{count} 页 · {paper} {orientation} · {scale}%",
      orientationPortrait: "纵向",
      orientationLandscape: "横向",
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
      rename: "重命名",
      renameLabel: "工作表名称",
      renameEmpty: "请输入工作表名称。",
      renameTooLong: "工作表名称不能超过 31 个字符。",
      renameDuplicate: "已存在同名工作表。",
      renameApostrophe: "工作表名称不能以撇号开头或结尾。",
      renameFailed: "无法重命名该工作表。",
      moveFailed: "无法移动该工作表。",
      menuLabel: "工作表菜单",
      hide: "隐藏",
      hideFailed: "无法隐藏该工作表。",
      unhide: "取消隐藏…",
      unhideFailed: "无法取消隐藏该工作表。",
      hiddenTab: "隐藏的工作表",
      noHiddenSheets: "没有隐藏的工作表。",
      tabListLabel: "工作表标签",
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
      duplicate: "创建副本",
      delete: "删除",
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
    autoFilter: {
      label: "筛选",
      apply: "应用",
      blanks: "（空白）",
      clear: "清除筛选",
      filterByValue: "按值筛选",
      noValues: "没有可筛选的值。",
      openMenu: "筛选 {column} 列",
      selectAll: "全选",
    },
    formulaBar: {
      activeCell: "活动单元格",
      goToName: "跳转到单元格、区域或定义名称",
      definedNames: "定义名称",
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
      slicerHint: "切片器“{name}”：静态预览，请在 Excel 中更改选择。",
    },
    pivot: {
      staticPreview: "透视表“{name}”：静态缓存值，请在 Excel 中刷新。",
    },
    selection: {
      current: "当前选择 {address}",
      none: "当前没有选择单元格",
      cutPending: "已剪切 {address}，按 Esc 取消",
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
      fillHandle: "填充柄：拖动以将选区内容填充到更多单元格",
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
      toolbar: "幻灯片工具栏",
      viewMode: "浏览布局",
      viewModeSingle: "单页视图",
      viewModeSorter: "幻灯片浏览",
      viewModePresent: "演示",
      viewModeBrowse: "退出演示",
      notesToggle: "演讲者备注",
      sorter: "幻灯片浏览",
      searchPlaceholder: "搜索幻灯片",
      search: "搜索",
      sidebar: "幻灯片列表",
      empty: "打开一个 PPTX 文件开始查看",
      loading: "正在加载 PPTX…",
      waiting: "正在等待幻灯片舞台…",
      cancelled: "PPTX 加载已取消。",
      encryptedTitle: "演示文稿受保护",
      encryptedDescription: "该演示文稿已加密，或保存为旧版二进制格式。请取消密码保护或另存为 .pptx 后重试。",
      unsupportedDrop: "仅支持 .pptx、.pptm、.ppsx、.ppsm、.potx 和 .potm 文件",
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
      link: {
        menu: "链接操作",
        open: "打开链接",
        copy: "复制链接",
        openSlide: "跳到第 {number} 页",
      },
      capability: {
        strict: "精确 {count}/{total}",
        approximate: "近似 {count}",
        static: "静态 {count}",
        unparsed: "未解析 {count}",
      },
      notes: {
        panel: "演讲者备注",
        presenter: "演讲者视图",
        title: "备注",
        empty: "当前幻灯片没有演讲者备注。",
        next: "下一页",
        noNext: "已是最后一页",
        timer: "已用时间",
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
      loading: "正在加载缩略图",
      failed: "缩略图不可用",
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
      pageLayout: "页面布局",
      pageLayoutContinuous: "连续滚动",
      pageLayoutSingle: "单页",
      pageLayoutSpread: "双页",
      pageInput: "页码",
      rotateClockwise: "顺时针旋转",
      searchLabel: "搜索 PDF 文本",
      searchPlaceholder: "搜索文本",
      search: "搜索",
      download: "下载 PDF",
      print: "打印 PDF",
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
      unsupportedDrop: "仅支持 PDF 文件",
      selectionLayer: "PDF 页面文本选择层",
      passwordRequired: "输入文档密码以解锁此 PDF。",
      selectionActions: "选中内容操作",
      copySelection: "复制选中内容",
      copiedSelection: "已复制",
      linkActions: "链接预览",
      openLink: "打开链接",
      annotationDetails: "批注详情",
    },
    surface: {
      noPdf: "未加载 PDF",
      opening: "正在打开 PDF…",
      renderingPage: "正在渲染第 {page} 页…",
      pageRenderFailed: "第 {page} 页渲染失败",
      pageAlt: "第 {page} 页",
      selectionRegion: "选择 PDF 页面区域",
      selectionText: "PDF 页面文本选择层",
      outline: "文档大纲",
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
      outline: "PDF 大纲",
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
    regionAnnotator: {
      defaultAriaLabel: "区域标注编辑器",
      toolbarLabel: "标注工具",
      arrow: "箭头",
      rectangle: "矩形",
      text: "文字",
      highlight: "高亮",
      color: "标注颜色",
      remove: "删除选中的标注",
      clear: "清除所有标注",
      empty: "先框选并截图，然后在此添加标注。",
      imageAlt: "区域截图预览",
      textPlaceholder: "标注文字",
      textSave: "添加文字",
      textCancel: "取消文字",
      hint: "在图片上拖动即可绘制。选中标注后按 Delete 删除。",
      count: "{count} 条标注",
    },
    boundingBox: {
      noFields: "没有可显示的字段",
      imageAlt: "引用来源预览",
    },
    fileUpload: {
      prompt: "将文件拖放到此处，或点击选择",
      title: "点击上传或拖放文件",
      browse: "浏览文件",
      drop: "拖放以添加",
      ready: "就绪",
      rejected: "不支持此文件类型。",
      unknownType: "未知类型",
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
