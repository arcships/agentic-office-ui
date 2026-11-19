import { HugeiconsIcon } from "@hugeicons/react";
import { MoreHorizontalIcon } from "@hugeicons/core-free-icons";

import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { Button } from "../vendor/ui/button";

const copy = fixtures.copy.buttons;

const cells: Array<{
  state: string;
  matrixState: string;
  render: () => React.ReactElement;
}> = [
  { state: "default", matrixState: "default", render: () => <Button>{copy.label}</Button> },
  { state: "hover", matrixState: "hover", render: () => <Button>{copy.label}</Button> },
  { state: "focus", matrixState: "focus", render: () => <Button>{copy.label}</Button> },
  { state: "active", matrixState: "active", render: () => <Button>{copy.label}</Button> },
  {
    state: "disabled",
    matrixState: "disabled",
    render: () => <Button disabled>{copy.label}</Button>,
  },
  {
    state: "loading",
    matrixState: "loading",
    render: () => <Button loading>{copy.label}</Button>,
  },
  {
    state: "outline",
    matrixState: "default",
    render: () => <Button variant="outline">{copy.outline}</Button>,
  },
  {
    state: "secondary",
    matrixState: "default",
    render: () => <Button variant="secondary">{copy.secondary}</Button>,
  },
  {
    state: "ghost",
    matrixState: "default",
    render: () => <Button variant="ghost">{copy.ghost}</Button>,
  },
  {
    state: "destructive",
    matrixState: "default",
    render: () => <Button variant="destructive">{copy.destructive}</Button>,
  },
  {
    state: "link",
    matrixState: "default",
    render: () => <Button variant="link">{copy.link}</Button>,
  },
  {
    state: "size-sm",
    matrixState: "default",
    render: () => <Button size="sm">{copy.label}</Button>,
  },
  {
    state: "size-lg",
    matrixState: "default",
    render: () => <Button size="lg">{copy.label}</Button>,
  },
  {
    state: "icon-sm",
    matrixState: "default",
    render: () => (
      <Button size="icon-sm" aria-label={copy.iconAlt}>
        <HugeiconsIcon icon={MoreHorizontalIcon} className="size-4" />
      </Button>
    ),
  },
];

export function ButtonSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <>
      {cells.map((cell) => (
        <ParityCell
          key={cell.state}
          matrixState={cell.matrixState}
          state={cell.state}
          width={cellWidth}
        >
          {cell.render()}
        </ParityCell>
      ))}
    </>
  );
}
