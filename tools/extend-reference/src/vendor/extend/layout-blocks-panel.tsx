/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/layout-blocks-panel.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (parity harness only):
 * - `@tanstack/react-virtual` is replaced by a flow list. The fixture has four
 *   blocks and the panel is 420px tall, so the visible rows, classes and
 *   measured card boxes are unchanged.
 * - `IconPlaceholder` resolves Hugeicons directly.
 * - The Radix/Base UI ScrollArea is replaced by a plain scroll viewport.
 * - Loading / error / markdown-empty states are harness extensions with the
 *   same class contract as the Vue subject.
 */
import * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  AlignBoxBottomCenterIcon,
  Heading01Icon,
  ImageCompositionIcon,
  LeftToRightListBulletIcon,
  ParagraphIcon,
  Table01Icon,
  TextCenterlineCenterTopIcon,
  TextNumberSignIcon,
} from "@hugeicons/core-free-icons";

import { cn } from "../../lib/cn";

export type LayoutBlockType =
  | "heading"
  | "paragraph"
  | "list"
  | "table"
  | "figure"
  | "header"
  | "footer"
  | "page_number";

export type LayoutBlockItem = {
  id: string;
  type: LayoutBlockType;
  label?: string;
  text: string;
  page?: number;
  confidence?: number;
};

export const layoutBlockStyles: Record<
  LayoutBlockType,
  {
    label: string;
    icon: typeof Heading01Icon;
    ring: string;
    badge: string;
  }
> = {
  heading: {
    label: "Heading",
    icon: Heading01Icon,
    ring: "border-violet-500/60 bg-violet-500/5",
    badge:
      "bg-violet-50 text-violet-600 dark:bg-violet-300/10 dark:text-violet-300",
  },
  paragraph: {
    label: "Paragraph",
    icon: ParagraphIcon,
    ring: "border-blue-500/60 bg-blue-500/5",
    badge: "bg-blue-50 text-blue-600 dark:bg-blue-300/10 dark:text-blue-300",
  },
  list: {
    label: "List",
    icon: LeftToRightListBulletIcon,
    ring: "border-emerald-500/60 bg-emerald-500/5",
    badge:
      "bg-emerald-50 text-emerald-600 dark:bg-emerald-300/10 dark:text-emerald-300",
  },
  table: {
    label: "Table",
    icon: Table01Icon,
    ring: "border-amber-500/60 bg-amber-500/5",
    badge:
      "bg-amber-50 text-amber-600 dark:bg-amber-300/10 dark:text-amber-300",
  },
  figure: {
    label: "Figure",
    icon: ImageCompositionIcon,
    ring: "border-rose-500/60 bg-rose-500/5",
    badge: "bg-rose-50 text-rose-600 dark:bg-rose-300/10 dark:text-rose-300",
  },
  header: {
    label: "Header",
    icon: TextCenterlineCenterTopIcon,
    ring: "border-cyan-500/60 bg-cyan-500/5",
    badge: "bg-cyan-50 text-cyan-600 dark:bg-cyan-300/10 dark:text-cyan-300",
  },
  footer: {
    label: "Footer",
    icon: AlignBoxBottomCenterIcon,
    ring: "border-slate-500/60 bg-slate-500/5",
    badge:
      "bg-slate-50 text-slate-600 dark:bg-slate-300/10 dark:text-slate-300",
  },
  page_number: {
    label: "Page number",
    icon: TextNumberSignIcon,
    ring: "border-zinc-500/60 bg-zinc-500/5",
    badge: "bg-zinc-50 text-zinc-600 dark:bg-zinc-300/10 dark:text-zinc-300",
  },
};

export function LayoutBlockCard({
  block,
  isActive,
  onFocusBlock,
  renderText,
}: {
  block: LayoutBlockItem;
  isActive: boolean;
  onFocusBlock: (block: LayoutBlockItem) => void;
  renderText?: (text: string) => React.ReactNode;
}) {
  const style = layoutBlockStyles[block.type];

  return (
    <button
      type="button"
      aria-pressed={isActive}
      data-layout-block-id={block.id}
      data-block-type={block.type}
      onClick={() => onFocusBlock(block)}
      onMouseEnter={() => onFocusBlock(block)}
      onFocus={() => onFocusBlock(block)}
      className={cn(
        "w-full rounded-lg border bg-background p-3 text-left hover:bg-muted/40 focus-visible:ring-2 focus-visible:ring-ring focus-visible:outline-none",
        isActive && cn("shadow-[0_0_0_1px_rgb(0_0_0_/_4%)]", style.ring)
      )}
    >
      <div className="min-w-0">
        <div className="flex items-center justify-between gap-2">
          <div className="flex min-w-0 items-center gap-2">
            <div
              className={cn(
                "inline-flex shrink-0 items-center gap-1 rounded-full px-2 py-0.5 text-xs font-medium",
                style.badge
              )}
            >
              <HugeiconsIcon icon={style.icon} className="size-3.5" />
              {block.label ?? style.label}
            </div>
            {block.confidence !== undefined && (
              <div className="truncate text-xs text-muted-foreground">
                {Math.round(block.confidence * 100)}%
              </div>
            )}
          </div>
          {block.page !== undefined && (
            <div className="shrink-0 rounded-full bg-muted px-2 py-0.5 text-xs text-muted-foreground">
              p. {block.page}
            </div>
          )}
        </div>
        <div className="mt-2 text-sm text-foreground/90">
          {renderText ? (
            renderText(block.text)
          ) : (
            <p className="line-clamp-3 text-[13px] leading-5 wrap-anywhere whitespace-pre-line">
              {block.text}
            </p>
          )}
        </div>
      </div>
    </button>
  );
}

export function LayoutBlocksPanel({
  activeBlockId,
  blocks,
  className,
  heading,
  emptyMessage = "No layout blocks found.",
  loading = false,
  error,
  markdown = false,
  onBlockFocus,
}: {
  activeBlockId?: string;
  blocks: readonly LayoutBlockItem[];
  className?: string;
  heading?: string;
  emptyMessage?: string;
  loading?: boolean;
  error?: string;
  markdown?: boolean;
  onBlockFocus?: (block: LayoutBlockItem) => void;
}) {
  const [localActiveId, setLocalActiveId] = React.useState<string>();
  const activeId = activeBlockId ?? localActiveId;
  const focusedBlockId = blocks.some((block) => block.id === activeId)
    ? activeId
    : blocks[0]?.id;

  return (
    <aside
      className={cn("flex h-[420px] min-h-0 flex-col bg-background", className)}
      data-layout-panel
    >
      {heading && (
        <div className="flex items-center justify-between border-b px-4 py-3 text-sm font-medium">
          {heading}
          <span className="text-xs text-muted-foreground">{blocks.length}</span>
        </div>
      )}
      <div className="min-h-0 flex-1 overflow-y-auto">
        {error ? (
          <div className="p-3">
            <div
              role="alert"
              className="rounded-lg border border-destructive/45 bg-destructive/8 p-3 text-sm text-destructive"
            >
              {error}
            </div>
          </div>
        ) : loading ? (
          <div className="p-3">
            <div
              role="status"
              className="rounded-lg border bg-muted/30 p-3 text-sm text-muted-foreground"
            >
              Loading layout blocks…
            </div>
          </div>
        ) : blocks.length ? (
          <div className="flex flex-col gap-2 p-3">
            {blocks.map((block) => (
              <LayoutBlockCard
                key={block.id}
                block={block}
                isActive={block.id === focusedBlockId}
                onFocusBlock={(next) => {
                  setLocalActiveId(next.id);
                  onBlockFocus?.(next);
                }}
                renderText={
                  markdown
                    ? (text) => (
                        <p
                          className="whitespace-normal text-[13px] leading-5"
                          data-layout-markdown
                          dangerouslySetInnerHTML={{ __html: text }}
                        />
                      )
                    : undefined
                }
              />
            ))}
          </div>
        ) : (
          <div className="p-3">
            <div
              role="status"
              className="rounded-lg border bg-muted/30 p-3 text-sm text-muted-foreground"
            >
              {emptyMessage}
            </div>
          </div>
        )}
      </div>
    </aside>
  );
}
