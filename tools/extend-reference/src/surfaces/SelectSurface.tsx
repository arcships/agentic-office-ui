import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { SelectTrigger, SelectValue } from "../vendor/ui/select";

const copy = fixtures.copy.select;

const cells: Array<{
  state: string;
  matrixState: string;
  interaction?: { type: string; selector?: string };
  render: () => React.ReactElement;
}> = [
  {
    state: "default",
    matrixState: "default",
    render: () => (
      <SelectTrigger aria-label={copy.label}>
        <SelectValue>{copy.label}</SelectValue>
      </SelectTrigger>
    ),
  },
  {
    state: "focus",
    matrixState: "focus",
    interaction: { type: "focus", selector: "[data-slot=select-trigger]" },
    render: () => (
      <SelectTrigger aria-label={copy.label}>
        <SelectValue>{copy.label}</SelectValue>
      </SelectTrigger>
    ),
  },
  {
    state: "disabled",
    matrixState: "disabled",
    render: () => (
      <SelectTrigger aria-label={copy.label} disabled>
        <SelectValue>{copy.label}</SelectValue>
      </SelectTrigger>
    ),
  },
  {
    state: "placeholder",
    matrixState: "default",
    render: () => (
      <SelectTrigger aria-label={copy.label}>
        <SelectValue data-placeholder>{copy.placeholder}</SelectValue>
      </SelectTrigger>
    ),
  },
];

export function SelectSurface({ cellWidth }: { cellWidth: number }) {
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
