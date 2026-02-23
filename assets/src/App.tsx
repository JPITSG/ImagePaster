import { useEffect, useRef, useState } from "react";
import { onInit, getInit, reportHeight, type InitData } from "./lib/bridge";
import ConfigView from "./ConfigView";
import LogView from "./LogView";

export default function App() {
  const containerRef = useRef<HTMLDivElement>(null);
  const [initData, setInitData] = useState<InitData | null>(null);

  useEffect(() => {
    onInit((data) => setInitData(data));
    getInit();
  }, []);

  useEffect(() => {
    const el = containerRef.current;
    if (!el || !initData) return;

    const report = () => {
      reportHeight(Math.ceil(el.scrollHeight));
    };

    requestAnimationFrame(report);

    const observer = new ResizeObserver(report);
    observer.observe(el);
    return () => observer.disconnect();
  }, [initData]);

  if (!initData) return null;

  return (
    <div ref={containerRef}>
      {initData.view === "config" ? (
        <ConfigView config={initData.config!} />
      ) : (
        <LogView initialLog={initData.log ?? []} />
      )}
    </div>
  );
}
