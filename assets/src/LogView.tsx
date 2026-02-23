import { useEffect, useRef, useState } from "react";
import { onLogUpdate, clearLog, closeDialog, type LogEntry } from "./lib/bridge";
import { Button } from "./components/ui/button";

interface Props {
  initialLog: LogEntry[];
}

export default function LogView({ initialLog }: Props) {
  const [entries, setEntries] = useState<LogEntry[]>(initialLog);
  const scrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    onLogUpdate((entry) => {
      setEntries((prev) => [...prev, entry]);
    });
  }, []);

  useEffect(() => {
    const el = scrollRef.current;
    if (el) el.scrollTop = el.scrollHeight;
  }, [entries]);

  const handleClear = () => {
    setEntries([]);
    clearLog();
  };

  return (
    <div className="p-4 flex flex-col gap-3" style={{ minHeight: "100%" }}>
      <div
        ref={scrollRef}
        className="flex-1 overflow-y-auto border border-neutral-200 rounded-md"
        style={{ maxHeight: "400px" }}
      >
        {entries.length === 0 ? (
          <div className="p-6 text-center text-sm text-neutral-400">
            No activity recorded yet.
          </div>
        ) : (
          <table className="w-full text-xs">
            <thead className="bg-neutral-50 sticky top-0">
              <tr className="border-b border-neutral-200">
                <th className="text-left px-3 py-2 font-medium text-neutral-600">
                  Time
                </th>
                <th className="text-left px-3 py-2 font-medium text-neutral-600">
                  Message
                </th>
              </tr>
            </thead>
            <tbody>
              {entries.map((entry, i) => (
                <tr
                  key={i}
                  className="border-b border-neutral-100 hover:bg-neutral-50"
                >
                  <td className="px-3 py-1.5 whitespace-nowrap">{entry.time}</td>
                  <td className="px-3 py-1.5">{entry.message}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div className="flex justify-end gap-2">
        <Button variant="outline" size="sm" onClick={handleClear}>
          Clear
        </Button>
        <Button size="sm" onClick={() => closeDialog()}>
          Close
        </Button>
      </div>
    </div>
  );
}
