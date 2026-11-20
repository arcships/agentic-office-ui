/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/xlsx-viewer.tsx (WorkbookSheetTabs,
 * WorkbookSheetTabsInner) plus apps/v4/registry/new-york-v4/ui/tabs.tsx and
 * apps/v4/components/ui/scroll-area.tsx class strings.
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (#160): the workbook engine is not vendored, so
 * `useXlsxViewer()` (sheets/activeSheetIndex) becomes props, and the thumbnail
 * preview (`useXlsxViewerThumbnails` + the portalled preview card) is dropped —
 * it only renders on hover, which the matrix never captures. The Base UI
 * `Tabs`/`ScrollArea` primitives are replaced by static elements with the same
 * class strings: `data-has-overflow-x` is measured on the viewport so the
 * upstream `has-[...]:h-[50px]` variant still decides the strip height, and the
 * sliding indicator is painted on the active trigger's own background (same
 * border-box, radius and shadow as the absolute indicator it replaces).
 */
import * as React from "react";
import { cn } from "../../lib/cn";

export interface XlsxSheetTabsProps {
  activeIndex: number;
  sheets: readonly string[];
}

const TABS_LIST_CLASSNAME =
  "relative z-0 flex w-fit items-center justify-center gap-x-0.5 text-muted-foreground rounded-lg bg-muted p-0.5 text-muted-foreground/72";

const TABS_TRIGGER_CLASSNAME =
  "relative flex h-9 shrink-0 grow cursor-pointer items-center justify-center gap-1.5 rounded-md border border-transparent px-[calc(--spacing(2.5)-1px)] text-base font-medium whitespace-nowrap transition-[color,background-color,box-shadow] outline-none hover:text-muted-foreground focus-visible:ring-2 focus-visible:ring-ring sm:h-8 sm:text-sm";

const ACTIVE_TRIGGER_CLASSNAME =
  "bg-background text-foreground shadow-sm/5 dark:bg-input";

export function XlsxSheetTabs({
  activeIndex,
  sheets,
}: XlsxSheetTabsProps): React.ReactElement {
  const viewportRef = React.useRef<HTMLDivElement | null>(null);
  const [overflowing, setOverflowing] = React.useState(false);

  React.useEffect(() => {
    const element = viewportRef.current;
    if (!element) return;

    const measure = () => {
      setOverflowing(element.scrollWidth - element.clientWidth > 1);
    };
    measure();

    if (typeof ResizeObserver === "undefined") return;
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    return () => observer.disconnect();
  }, [sheets.length]);

  return (
    <div
      className="border-t bg-muted/40 px-3 py-2"
      data-parity-xlsx-sheet-tabs=""
    >
      <div className="flex flex-col gap-0" data-slot="tabs">
        <div
          className="size-full min-h-0 h-10 w-full has-[[data-slot=scroll-area-viewport][data-has-overflow-x]]:h-[50px]"
          data-slot="scroll-area"
        >
          <div
            ref={viewportRef}
            className="transition-shadows h-full rounded-[inherit] outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-1 focus-visible:ring-offset-background pb-3.5 overflow-y-hidden"
            data-slot="scroll-area-viewport"
            data-has-overflow-x={overflowing ? "" : undefined}
          >
            <div className="flex h-full items-center">
              <div
                className={cn(TABS_LIST_CLASSNAME, "shrink-0")}
                data-slot="tabs-list"
                role="tablist"
              >
                {sheets.map((name, index) => (
                  <button
                    key={`${index}-${name}`}
                    type="button"
                    role="tab"
                    aria-selected={index === activeIndex}
                    className={cn(
                      TABS_TRIGGER_CLASSNAME,
                      "max-w-48 flex-none",
                      index === activeIndex && ACTIVE_TRIGGER_CLASSNAME,
                    )}
                    data-slot="tabs-tab"
                  >
                    <span className="truncate">{name}</span>
                  </button>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
