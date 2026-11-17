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
  }
  formField: {
    placeholder: string
    selectDate: string
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
  }
  dragOverlay: {
    drop: string
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
    allBorders: string
    merge: string
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
  }
  contextMenu: {
    label: string
    copy: string
    paste: string
    clear: string
    sortAscending: string
    sortDescending: string
    undo: string
    redo: string
  }
  formulaBar: {
    activeCell: string
    placeholder: string
  }
  chart: {
    unsupported: string
    loading: string
    unavailable: string
    noPayload: string
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
}

export interface OfficeMessages {
  common: OfficeCommonMessages
  docx: OfficeDocxMessages
  xlsx: OfficeXlsxMessages
  pptx: OfficePptxMessages
  pdf: OfficePdfMessages
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
    },
    formField: {
      placeholder: "Enter text...",
      selectDate: "Select a date",
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
    },
    dragOverlay: {
      drop: "Drop DOCX file to import",
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
      allBorders: "All borders",
      merge: "Merge",
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
    },
    contextMenu: {
      label: "Cell menu",
      copy: "Copy",
      paste: "Paste",
      clear: "Clear contents",
      sortAscending: "Sort ascending",
      sortDescending: "Sort descending",
      undo: "Undo",
      redo: "Redo",
    },
    formulaBar: {
      activeCell: "Active cell",
      placeholder: "Enter a formula or value",
    },
    chart: {
      unsupported: "Unsupported chart type",
      loading: "Loading chart…",
      unavailable: "This chart type is not available yet.",
      noPayload: "This chart sheet does not expose an embedded chart payload.",
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
    },
    formField: {
      placeholder: "输入文本…",
      selectDate: "选择日期",
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
    },
    dragOverlay: {
      drop: "拖放 DOCX 文件以导入",
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
      allBorders: "所有边框",
      merge: "合并",
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
    },
    contextMenu: {
      label: "单元格菜单",
      copy: "复制",
      paste: "粘贴",
      clear: "清除内容",
      sortAscending: "按此列升序",
      sortDescending: "按此列降序",
      undo: "撤销",
      redo: "重做",
    },
    formulaBar: {
      activeCell: "活动单元格",
      placeholder: "输入公式或值",
    },
    chart: {
      unsupported: "不支持的图表类型",
      loading: "正在加载图表…",
      unavailable: "此图表类型暂不可用。",
      noPayload: "此图表工作表未提供嵌入图表数据。",
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
