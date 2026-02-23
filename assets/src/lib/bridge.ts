export interface ConfigData {
  titleMatch: string;
}

export interface LogEntry {
  time: string;
  message: string;
}

export interface InitData {
  view: "config" | "log";
  config?: ConfigData;
  log?: LogEntry[];
}

type InitCallback = (data: InitData) => void;
type LogUpdateCallback = (entry: LogEntry) => void;

let initCallback: InitCallback | null = null;
let logUpdateCallback: LogUpdateCallback | null = null;

declare global {
  interface Window {
    onInit: (data: InitData) => void;
    onLogUpdate: (entry: LogEntry) => void;
    chrome?: {
      webview?: {
        postMessage: (s: string) => void;
      };
    };
  }
}

window.onInit = (data: InitData) => {
  initCallback?.(data);
};

window.onLogUpdate = (entry: LogEntry) => {
  logUpdateCallback?.(entry);
};

export function onInit(cb: InitCallback) {
  initCallback = cb;
}

export function onLogUpdate(cb: LogUpdateCallback) {
  logUpdateCallback = cb;
}

function postMessage(msg: Record<string, unknown>) {
  try {
    window.chrome?.webview?.postMessage(JSON.stringify(msg));
  } catch {
    console.log("postMessage (no WebView2):", msg);
  }
}

export function getInit() {
  postMessage({ action: "getInit" });
}

export function saveSettings(config: ConfigData) {
  postMessage({ action: "saveSettings", titleMatch: config.titleMatch });
}

export function clearLog() {
  postMessage({ action: "clearLog" });
}

export function closeDialog() {
  postMessage({ action: "close" });
}

export function reportHeight(height: number) {
  postMessage({ action: "resize", height });
}
