import { useEffect, useRef, useState } from "react";
import {
  clearHistoryImages,
  closeDialog,
  copyHistoryUrl,
  deleteHistoryImage,
  onHistoryActionResult,
  openHistoryUrl,
  revealHistoryFile,
  saveHistoryImage,
  type HistoryActionResult,
  type HistoryData,
  type HistoryEntry,
} from "./lib/bridge";
import { Button } from "./components/ui/button";

interface Props {
  history: HistoryData;
}

function formatBytes(bytes: number) {
  if (bytes >= 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  if (bytes >= 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${bytes} B`;
}

function formatCapturedAt(ms: number) {
  if (!ms) return "—";
  const date = new Date(ms);
  const time = date.toLocaleTimeString([], { hour12: false });
  const sameDay = date.toDateString() === new Date().toDateString();
  return sameDay ? time : `${date.toLocaleDateString()} ${time}`;
}

function HistoryRow({ entry }: { entry: HistoryEntry }) {
  return (
    <li className="flex items-center gap-3 px-3 py-2.5 hover:bg-neutral-50">
      <button
        type="button"
        title="Open in default browser"
        onClick={() => openHistoryUrl(entry.token)}
        className="relative flex h-20 w-32 shrink-0 cursor-pointer items-center justify-center overflow-hidden rounded border border-neutral-200 bg-neutral-100 transition-shadow hover:ring-2 hover:ring-neutral-300 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-neutral-400"
      >
        {entry.thumb ? (
          <>
            {/* Blurred cover copy fills the letterbox area behind the image */}
            <img
              src={entry.thumb}
              alt=""
              aria-hidden
              className="absolute inset-0 h-full w-full scale-110 object-cover blur-md brightness-90"
              draggable={false}
            />
            <img
              src={entry.thumb}
              alt={`${entry.width} × ${entry.height} preview`}
              className="relative max-h-full max-w-full object-contain"
              draggable={false}
            />
          </>
        ) : (
          <span className="text-[10px] text-neutral-400">No preview</span>
        )}
      </button>

      <div className="min-w-0 flex-1 space-y-0.5">
        <div className="flex items-center gap-2 text-xs font-medium text-neutral-900">
          <span className="tabular-nums">
            {entry.width} × {entry.height}
          </span>
          {entry.current && (
            <span className="rounded-full bg-emerald-100 px-2 py-0.5 text-[10px] font-medium text-emerald-700">
              Current
            </span>
          )}
        </div>
        <div className="text-[11px] tabular-nums text-neutral-500">
          {formatBytes(entry.bytes)} · {formatCapturedAt(entry.capturedAt)}
        </div>
        <button
          type="button"
          title={`Open ${entry.url} in default browser`}
          onClick={() => openHistoryUrl(entry.token)}
          className="block w-full cursor-pointer truncate text-left font-mono text-[10px] text-neutral-400 hover:text-neutral-600 hover:underline focus-visible:underline focus-visible:outline-none"
        >
          {entry.url}
        </button>
        {entry.storage === "disk" && entry.path ? (
          <button
            type="button"
            title="Show this file in File Explorer"
            onClick={() => revealHistoryFile(entry.token)}
            className="block w-full cursor-pointer truncate text-left text-[10px] text-neutral-400 hover:text-neutral-600 hover:underline focus-visible:underline focus-visible:outline-none"
          >
            Disk · <span className="font-mono">{entry.path}</span>
          </button>
        ) : (
          <div className="text-[10px] text-neutral-400">Stored in memory</div>
        )}
      </div>

      <div className="flex shrink-0 flex-col gap-1">
        <Button
          variant="outline"
          size="sm"
          className="h-7 justify-center px-2.5 text-[11px]"
          onClick={() => copyHistoryUrl(entry.token)}
        >
          Copy URL
        </Button>
        <Button
          variant="outline"
          size="sm"
          className="h-7 justify-center px-2.5 text-[11px]"
          onClick={() => saveHistoryImage(entry.token)}
        >
          Save…
        </Button>
        <Button
          variant="outline"
          size="sm"
          className="h-7 justify-center px-2.5 text-[11px] text-red-600 hover:bg-red-50 hover:text-red-700"
          onClick={() => deleteHistoryImage(entry.token)}
        >
          Delete
        </Button>
      </div>
    </li>
  );
}

export default function HistoryView({ history }: Props) {
  const [status, setStatus] = useState<HistoryActionResult | null>(null);
  const [confirmClear, setConfirmClear] = useState(false);
  const statusTimer = useRef<number | null>(null);

  useEffect(() => {
    const removeListener = onHistoryActionResult((result) => {
      setStatus(result);
      if (statusTimer.current !== null) window.clearTimeout(statusTimer.current);
      statusTimer.current = window.setTimeout(() => setStatus(null), 5000);
    });
    return () => {
      removeListener();
      if (statusTimer.current !== null) window.clearTimeout(statusTimer.current);
    };
  }, []);

  const entries = history.entries;

  const handleClearAll = () => {
    setConfirmClear(false);
    clearHistoryImages();
  };

  return (
    <div className="p-5 space-y-4">
      <div className="flex items-baseline justify-between gap-3">
        <div>
          <div className="text-xs font-medium">Retained Images</div>
          <p className="mt-1 text-[11px] text-neutral-500">
            Clipboard images retained by the built-in HTTP server, in memory
            or on disk. Everything here is discarded when ImagePaster exits.
          </p>
        </div>
        <span className="whitespace-nowrap text-[11px] tabular-nums text-neutral-400">
          {history.total} image{history.total === 1 ? "" : "s"} ·{" "}
          {formatBytes(history.totalBytes)}
        </span>
      </div>

      <div
        className="overflow-y-auto rounded-md border border-neutral-200"
        style={{ maxHeight: "430px" }}
      >
        {entries.length === 0 ? (
          <div className="p-6 text-center text-sm text-neutral-400">
            No images are currently retained in memory.
          </div>
        ) : (
          <ul className="divide-y divide-neutral-100">
            {entries.map((entry) => (
              <HistoryRow key={entry.token} entry={entry} />
            ))}
          </ul>
        )}
      </div>

      {history.shown < history.total && (
        <p className="text-[11px] text-neutral-500">
          Showing the newest {history.shown} of {history.total} images.
        </p>
      )}

      {status && (
        <div
          className={
            status.ok
              ? "rounded-md bg-emerald-50 px-3 py-2 text-[11px] text-emerald-700"
              : "rounded-md bg-red-50 px-3 py-2 text-[11px] text-red-700"
          }
        >
          {status.message}
        </div>
      )}

      <div className="flex items-center justify-between gap-3 pt-1">
        <Button
          variant="outline"
          size="sm"
          className="text-red-600 hover:bg-red-50 hover:text-red-700"
          disabled={entries.length === 0}
          onClick={() => setConfirmClear(true)}
        >
          Clear All
        </Button>
        <Button size="sm" className="min-w-[5rem]" onClick={closeDialog}>
          Close
        </Button>
      </div>

      {confirmClear && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/35 p-4">
          <div
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="clear-history-title"
            aria-describedby="clear-history-message"
            className="w-full max-w-sm space-y-3 rounded-lg border border-neutral-200 bg-white p-4 shadow-xl"
          >
            <div className="space-y-1">
              <h2 id="clear-history-title" className="text-sm font-semibold">
                Clear all retained images?
              </h2>
              <p
                id="clear-history-message"
                className="text-xs leading-relaxed text-neutral-600"
              >
                All {history.total} image{history.total === 1 ? "" : "s"} will
                be removed, including the current paste target. Disk-stored
                files are deleted, all URLs stop working immediately, and this
                cannot be undone.
              </p>
            </div>
            <div className="flex justify-end gap-2">
              <Button
                variant="outline"
                size="sm"
                autoFocus
                onClick={() => setConfirmClear(false)}
              >
                Cancel
              </Button>
              <Button variant="destructive" size="sm" onClick={handleClearAll}>
                Clear All
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
