/**
 * Vendored from extend-hq/ui @ 70fe70c6f1e1eee53954c6923f3d9f583557a57e
 * Source: apps/v4/components/extend/bounding-box-citations.tsx
 * Copyright (c) 2026 CrowdView Inc, dba Extend. Portions Copyright (c) 2023 shadcn.
 * Licensed under the MIT License; see ../NOTICE.md.
 *
 * Adaptation (parity harness only):
 * - Only the `HumanReviewPanel` Form tab and scalar field cards are vendored.
 *   The JSON tab uses `@pierre/diffs` + glide-data-grid and is not reproducible
 *   as static markup; array/object field editors are likewise out of scope.
 * - Tooltips around the revert / set-NULL buttons are dropped; the trigger
 *   buttons render the same classes and glyphs.
 * - A small state wrapper keeps hover/focus/active selection behaviour, so the
 *   `hover`, `focus` and `active` cells exercise the same visual transition as
 *   the Vue subject.
 * - Loading / error / empty states are harness extensions (the upstream block
 *   only ships the populated state); they use the same class contract as the
 *   Vue subject.
 */
import * as React from "react";
import { HugeiconsIcon } from "@hugeicons/react";
import {
  CancelCircleIcon,
  HashIcon,
  InputTextIcon,
  SecondBracketIcon,
  SourceCodeSquareIcon,
  TextCheckIcon,
  Undo02Icon,
} from "@hugeicons/core-free-icons";

import { cn } from "../../lib/cn";
import { Button } from "../ui/button";
import { Input } from "../ui/input";

export type JsonValue =
  | string
  | number
  | boolean
  | null
  | JsonValue[]
  | { [key: string]: JsonValue };

export type ReviewFieldSchema = {
  type: "string" | "number" | "integer" | "boolean" | "object" | "array";
  title?: string;
  description?: string;
};

export type ReviewLocation = {
  page: number;
  area: { left: number; top: number; width: number; height: number };
};

export type ReviewField = {
  key: string;
  schema: ReviewFieldSchema;
  actual: JsonValue;
  expected: JsonValue;
  location?: ReviewLocation;
  metadataPath?: string;
};

function formatValue(value: JsonValue): string {
  if (value === null) return "null";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return JSON.stringify(value);
}

function fieldIcon(type: ReviewFieldSchema["type"]) {
  if (type === "number" || type === "integer") return HashIcon;
  if (type === "boolean") return TextCheckIcon;
  if (type === "array") return SecondBracketIcon;
  if (type === "object") return SourceCodeSquareIcon;
  return InputTextIcon;
}

function valuesEqual(left: JsonValue, right: JsonValue): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function HumanReviewFieldCard({
  field,
  value,
  active,
  showExpected,
  onFocusField,
}: {
  field: ReviewField;
  value: JsonValue;
  active: boolean;
  showExpected: boolean;
  onFocusField: (field: ReviewField) => void;
}) {
  const modified = showExpected && !valuesEqual(value, field.expected);
  const Icon = fieldIcon(field.schema.type);

  return (
    <div
      tabIndex={0}
      data-review-field-key={field.key}
      onFocusCapture={() => onFocusField(field)}
      onMouseEnter={() => onFocusField(field)}
      className={cn(
        "rounded-lg border bg-background p-3 transition-[border-color,background-color,box-shadow] focus-within:border-blue-500/50 focus-within:bg-blue-500/5 hover:border-blue-500/50 hover:bg-blue-500/5 focus-visible:ring-2 focus-visible:ring-blue-500/20 focus-visible:outline-none",
        active &&
          "border-blue-500/60 bg-blue-500/5 shadow-[0_0_0_1px_rgb(59_130_246_/_8%)]"
      )}
    >
      <div className="mb-3 flex min-h-8 items-start justify-between gap-3">
        <div className="min-w-0">
          <div className="flex min-w-0 items-center gap-2">
            <div className="min-w-0">
              <div className="truncate text-sm font-medium">
                {field.schema.title ?? field.key}
              </div>
            </div>
            <span
              className={cn(
                "size-2 shrink-0 rounded-full bg-amber-400",
                !modified && "opacity-0"
              )}
            />
          </div>
          <div className="truncate text-xs text-muted-foreground">
            {field.key}
          </div>
        </div>
        <div className="flex shrink-0 items-center gap-1">
          {showExpected && modified ? (
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              className="text-muted-foreground"
              aria-label={`Undo ${field.key}`}
            >
              <HugeiconsIcon icon={Undo02Icon} className="size-4" />
            </Button>
          ) : null}
          {showExpected ? (
            <Button
              type="button"
              variant="ghost"
              size="icon-sm"
              className="text-muted-foreground"
              aria-label={`Set ${field.key} to null`}
            >
              <HugeiconsIcon icon={CancelCircleIcon} className="size-4" />
            </Button>
          ) : null}
          <div className="flex h-6 items-center gap-1 rounded-md border bg-muted/50 px-1.5 text-xs text-muted-foreground">
            <HugeiconsIcon icon={Icon} className="size-3.5" />
            {field.schema.type}
          </div>
        </div>
      </div>
      <div className={cn("grid gap-2", showExpected && "sm:grid-cols-2")}>
        <div className="rounded-md border bg-muted/30 p-2">
          {field.schema.description ? (
            <p className="mb-2 text-xs text-muted-foreground">
              {field.schema.description}
            </p>
          ) : null}
          <div className="mb-1 text-[11px] font-medium text-muted-foreground">
            Actual
          </div>
          <div className="min-h-7 rounded-md bg-background px-2 py-1.5 text-sm">
            {formatValue(field.actual)}
          </div>
        </div>
        {showExpected ? (
          <div className="rounded-md border bg-muted/30 p-2">
            <div className="mb-1 text-[11px] font-medium text-muted-foreground">
              Expected
            </div>
            <Input
              nativeInput
              readOnly
              type={field.schema.type === "number" ? "number" : "text"}
              value={value === null ? "" : String(value)}
              onChange={() => {}}
            />
          </div>
        ) : null}
      </div>
    </div>
  );
}

export function HumanReviewPanel({
  fields,
  activeFieldKey,
  showExpected = true,
  loading = false,
  error,
  emptyMessage = "No fields to display",
  modifiedFieldKey,
  modifiedExpected,
  className,
}: {
  fields: ReviewField[];
  activeFieldKey?: string;
  showExpected?: boolean;
  loading?: boolean;
  error?: string;
  emptyMessage?: string;
  modifiedFieldKey?: string;
  modifiedExpected?: JsonValue;
  className?: string;
}) {
  const [hoveredKey, setHoveredKey] = React.useState<string | undefined>(
    undefined
  );
  const activeKey = hoveredKey ?? activeFieldKey ?? fields[0]?.key;

  return (
    <div
      className={cn(
        "flex h-[560px] flex-col gap-0 bg-background",
        className
      )}
      data-review-panel
    >
      <div className="flex min-h-12 items-center justify-between gap-3 border-b px-3">
        <div className="relative z-0 flex h-8 w-fit items-center justify-center gap-x-0.5 rounded-lg bg-muted p-0.5 text-xs text-muted-foreground/72 sm:h-7">
          <button
            type="button"
            className="relative flex h-7 shrink-0 grow cursor-pointer items-center justify-center gap-1.5 rounded-md border border-transparent bg-background px-[calc(--spacing(2.5)-1px)] text-sm font-medium whitespace-nowrap text-foreground shadow-sm/5 transition-[color,background-color,box-shadow] outline-none sm:h-6"
            role="tab"
            aria-selected="true"
          >
            <HugeiconsIcon icon={TextCheckIcon} className="size-4" />
            Form
          </button>
          <button
            type="button"
            className="relative flex h-7 shrink-0 grow cursor-pointer items-center justify-center gap-1.5 rounded-md border border-transparent px-[calc(--spacing(2.5)-1px)] text-sm font-medium whitespace-nowrap transition-[color,background-color,box-shadow] outline-none hover:text-muted-foreground sm:h-6"
            role="tab"
            aria-selected="false"
          >
            <HugeiconsIcon icon={SourceCodeSquareIcon} className="size-4" />
            JSON
          </button>
        </div>
      </div>
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
              Loading review fields…
            </div>
          </div>
        ) : fields.length === 0 ? (
          <div className="p-3">
            <div
              role="status"
              className="rounded-lg border bg-muted/30 p-3 text-sm text-muted-foreground"
            >
              {emptyMessage}
            </div>
          </div>
        ) : (
          <div className="space-y-3 p-3">
            {fields.map((field) => (
              <HumanReviewFieldCard
                key={field.key}
                field={field}
                value={
                  field.key === modifiedFieldKey && modifiedExpected !== undefined
                    ? modifiedExpected
                    : field.expected
                }
                active={field.key === activeKey}
                showExpected={showExpected}
                onFocusField={(next) => setHoveredKey(next.key)}
              />
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
