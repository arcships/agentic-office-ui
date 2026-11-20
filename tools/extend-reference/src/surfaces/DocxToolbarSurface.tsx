import * as React from "react";
import {
  ArrowLeft01Icon,
  ArrowRight01Icon,
  ArrowHorizontalIcon,
  ArrowVerticalIcon,
  Clock01Icon,
  ColumnInsertIcon,
  Comment01Icon,
  Download01Icon,
  File01Icon,
  LeftToRightListBulletIcon,
  MinusSignCircleIcon,
  Moon02Icon,
  MoreHorizontalIcon,
  PlusSignCircleIcon,
  Search01Icon,
  SidebarLeftIcon,
  Sun03Icon,
  Upload01Icon,
} from "@hugeicons/core-free-icons";
import { HugeiconsIcon } from "@hugeicons/react";

import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { Button } from "../vendor/ui/button";
import { Input } from "../vendor/ui/input";
import { SelectTrigger, SelectValue } from "../vendor/ui/select";
import { Separator } from "../vendor/ui/separator";

const toolbar = fixtures.docxToolbar;
const fileName = "legal-contract.docx";
const COMPACT_BREAKPOINT = 1360;

function ToolbarIcon({ icon }: { icon: React.ComponentProps<typeof HugeiconsIcon>["icon"] }) {
  return <HugeiconsIcon icon={icon} className="size-4" />;
}

function Group({ children, label }: { children: React.ReactNode; label?: string }) {
  return (
    <div
      role="group"
      aria-label={label}
      className="flex flex-none items-center gap-1"
    >
      {children}
    </div>
  );
}

function SearchControls() {
  return (
    <Group label="Search document">
      <Input
        size="sm"
        className="w-[130px] flex-none"
        placeholder="Search document"
        aria-label="Search document"
      />
      <Button type="button" size="icon-xs" variant="outline" aria-label="Match case">
        Aa
      </Button>
      <Button type="button" size="icon-xs" variant="outline" aria-label="Match whole word">
        ab|
      </Button>
      <Button type="button" size="icon-xs" variant="outline" aria-label="Use regular expression">
        .*
      </Button>
      <span className="min-w-8 text-center text-xs tabular-nums text-muted-foreground" />
      <Button type="button" size="icon-xs" variant="outline" aria-label="Previous match">
        <ToolbarIcon icon={ArrowLeft01Icon} />
      </Button>
      <Button type="button" size="icon-xs" variant="outline" aria-label="Next match">
        <ToolbarIcon icon={ArrowRight01Icon} />
      </Button>
    </Group>
  );
}

function ZoomControls() {
  return (
    <Group label="Zoom level">
      <Button type="button" variant="ghost" size="icon-sm" aria-label="Zoom out">
        <ToolbarIcon icon={MinusSignCircleIcon} />
      </Button>
      <SelectTrigger size="sm" className="w-[104px] min-w-[104px]" aria-label="Zoom level">
        <SelectValue>{Math.round(toolbar.zoomPercent)}%</SelectValue>
      </SelectTrigger>
      <Button type="button" variant="ghost" size="icon-sm" aria-label="Zoom in">
        <ToolbarIcon icon={PlusSignCircleIcon} />
      </Button>
    </Group>
  );
}

function FullControls({ disabled }: { disabled: boolean }) {
  return (
    <>
      <ZoomControls />
      <SearchControls />
      <Group label="Page layout">
        <Button
          type="button"
          variant="ghost"
          size="icon-sm"
          aria-label="Continuous scroll"
          aria-pressed="true"
          className="bg-accent text-foreground"
          disabled={disabled}
        >
          <ToolbarIcon icon={LeftToRightListBulletIcon} />
        </Button>
        <Button type="button" variant="ghost" size="icon-sm" aria-label="Single page" disabled={disabled}>
          <ToolbarIcon icon={File01Icon} />
        </Button>
        <Button type="button" variant="ghost" size="icon-sm" aria-label="Two-page spread" disabled={disabled}>
          <ToolbarIcon icon={ColumnInsertIcon} />
        </Button>
      </Group>
      <Group label="Fit">
        <Button type="button" variant="ghost" size="icon-sm" aria-label="Fit width" disabled={disabled}>
          <ToolbarIcon icon={ArrowHorizontalIcon} />
        </Button>
        <Button type="button" variant="ghost" size="icon-sm" aria-label="Fit page" disabled={disabled}>
          <ToolbarIcon icon={ArrowVerticalIcon} />
        </Button>
      </Group>
      <Separator orientation="vertical" className="mx-1 h-4 self-center" />
      <Button type="button" variant="outline" size="sm" aria-label="Hide 2 tracked changes" disabled={disabled}>
        <ToolbarIcon icon={Clock01Icon} />
        <span className="inline-flex h-4 min-w-4 items-center justify-center rounded-full bg-foreground px-1 text-[10px] text-background">2</span>
      </Button>
      <SelectTrigger size="sm" className="h-6 w-[112px] min-w-[112px]" aria-label="Display for review" disabled={disabled}>
        <SelectValue>Balloons</SelectValue>
      </SelectTrigger>
      <Button type="button" variant="outline" size="sm" aria-label="Hide 3 comments" disabled={disabled}>
        <ToolbarIcon icon={Comment01Icon} />
        <span className="inline-flex h-4 min-w-4 items-center justify-center rounded-full bg-foreground px-1 text-[10px] text-background">3</span>
      </Button>
      <Separator orientation="vertical" className="mx-1 h-4 self-center" />
      <Button type="button" variant="ghost" size="icon-sm" aria-label="Use dark document theme" disabled={disabled}>
        <ToolbarIcon icon={Moon02Icon} />
      </Button>
      <Button type="button" variant="ghost" size="icon-sm" aria-label="Upload DOCX">
        <ToolbarIcon icon={Upload01Icon} />
      </Button>
      <Button type="button" variant="ghost" size="icon-sm" aria-label="Download DOCX" disabled={disabled || !toolbar.showDownload}>
        <ToolbarIcon icon={Download01Icon} />
      </Button>
    </>
  );
}

function CompactControls({ disabled }: { disabled: boolean }) {
  return (
    <>
      <ZoomControls />
      <Button type="button" variant="ghost" size="icon-sm" aria-label="Search document">
        <ToolbarIcon icon={Search01Icon} />
      </Button>
      <Separator orientation="vertical" className="mx-1 h-4 self-center" />
      <Button type="button" variant="ghost" size="icon-sm" aria-label="More toolbar actions" disabled={disabled}>
        <ToolbarIcon icon={MoreHorizontalIcon} />
      </Button>
    </>
  );
}

function DocxToolbar({ disabled }: { disabled: boolean }) {
  const rootRef = React.useRef<HTMLDivElement | null>(null);
  const [width, setWidth] = React.useState(0);

  React.useEffect(() => {
    const element = rootRef.current;
    if (!element) return;
    const update = () => setWidth(element.getBoundingClientRect().width);
    update();
    const observer = new ResizeObserver(update);
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  const compact = width > 0 && width < COMPACT_BREAKPOINT;

  return (
    <div
      ref={rootRef}
      data-parity-docx-toolbar=""
      className="flex min-h-12 flex-wrap items-center justify-between gap-2 border-b bg-background px-3 py-2"
    >
      <div className="flex min-w-0 flex-wrap items-center gap-2">
        <Button
          type="button"
          variant="ghost"
          size="icon-sm"
          aria-label="Toggle page thumbnails"
          disabled={disabled}
        >
          <ToolbarIcon icon={SidebarLeftIcon} />
        </Button>
        {!compact && (
          <span className="max-w-[13.75rem] truncate text-sm font-semibold" title={fileName}>
            {fileName}
          </span>
        )}
        <div className="ms-1 flex items-center text-sm whitespace-nowrap text-primary">
          <span>Page</span>
          <Input
            size="sm"
            readOnly
            value={toolbar.activePage}
            aria-label={`Current page ${toolbar.activePage}. Edit page number`}
            className="h-6 w-[52px] flex-none"
            disabled={disabled || !toolbar.pageCount}
          />
          <span>of {toolbar.pageCount || "-"}</span>
        </div>
      </div>
      <div className="ml-auto flex min-w-0 flex-wrap items-center justify-end gap-1">
        {compact ? <CompactControls disabled={disabled} /> : <FullControls disabled={disabled} />}
      </div>
    </div>
  );
}

const states = [
  { state: "default", matrixState: "default", disabled: false },
  { state: "hover", matrixState: "hover", disabled: false },
  { state: "focus", matrixState: "focus", disabled: false },
  { state: "disabled", matrixState: "empty", disabled: true },
];

export function DocxToolbarSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {states.map((entry) => (
        <ParityCell
          key={entry.state}
          matrixState={entry.matrixState}
          state={entry.state}
          width={cellWidth}
        >
          <DocxToolbar disabled={entry.disabled} />
        </ParityCell>
      ))}
    </>
  );
}
