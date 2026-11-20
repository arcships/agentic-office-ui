import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { Input } from "../vendor/ui/input";

const copy = fixtures.copy.input;

const cells: Array<{
  state: string;
  matrixState: string;
  interaction?: { type: string; selector?: string };
  render: () => React.ReactElement;
}> = [
  {
    state: "default",
    matrixState: "default",
    render: () => <Input defaultValue={copy.value} aria-label={copy.label} />,
  },
  {
    state: "focus",
    matrixState: "focus",
    interaction: { type: "focus", selector: "input" },
    render: () => <Input defaultValue={copy.value} aria-label={copy.label} />,
  },
  {
    state: "disabled",
    matrixState: "disabled",
    render: () => <Input defaultValue={copy.disabled} disabled aria-label={copy.label} />,
  },
  {
    state: "invalid",
    matrixState: "error",
    render: () => (
      <Input
        aria-invalid="true"
        aria-label={copy.label}
        defaultValue={copy.invalid}
      />
    ),
  },
  {
    state: "size-sm",
    matrixState: "default",
    render: () => <Input size="sm" defaultValue={copy.value} aria-label={copy.label} />,
  },
];

export function InputSurface({ cellWidth }: { cellWidth: number }) {
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
