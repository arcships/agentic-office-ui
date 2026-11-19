import { fixtures } from "../fixtures";
import { ParityCell } from "../parity/ParitySurface";
import { Button } from "../vendor/ui/button";
import {
  DialogBackdrop,
  DialogClose,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogPanel,
  DialogPopup,
  DialogTitle,
  DialogViewport,
} from "../vendor/ui/dialog";

const copy = fixtures.copy.dialog;

function CloseIcon(): React.ReactElement {
  return (
    <svg
      aria-hidden="true"
      className="size-4.5 sm:size-4"
      fill="none"
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="2"
      viewBox="0 0 24 24"
    >
      <path d="M18 6 6 18" />
      <path d="m6 6 12 12" />
    </svg>
  );
}

export function DialogSurface({ cellWidth }: { cellWidth: number }) {
  return (
    <ParityCell matrixState="default" state="open" width={cellWidth}>
      <DialogBackdrop />
      <DialogViewport className="max-sm:items-end max-sm:p-0 max-sm:pt-12">
        <DialogPopup>
          <DialogHeader>
            <DialogTitle>{copy.title}</DialogTitle>
            <DialogDescription>{copy.description}</DialogDescription>
          </DialogHeader>
          <DialogPanel className="p-6 pt-1" />
          <DialogFooter>
            <Button variant="ghost">{copy.cancel}</Button>
            <Button>{copy.save}</Button>
          </DialogFooter>
          <DialogClose>
            <Button aria-label={copy.close} size="icon" variant="ghost">
              <CloseIcon />
            </Button>
          </DialogClose>
        </DialogPopup>
      </DialogViewport>
    </ParityCell>
  );
}
