export type PasteMethod = "base64" | "http";

export interface ConfigData {
  titleMatch: string;
  pasteMethod: PasteMethod;
  bindIp: string;
  httpPort: number;
  jpegQuality: number;
  shiftInsertPaste: boolean;
  availableIps: string[];
  bindIpAvailable: boolean;
  serverStatus: string;
  version: string;
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
type SaveResultCallback = (result: { ok: boolean; message?: string }) => void;

let initCallback: InitCallback | null = null;
let logUpdateCallback: LogUpdateCallback | null = null;
let saveResultCallback: SaveResultCallback | null = null;

declare global {
  interface Window {
    onInit: (data: InitData) => void;
    onLogUpdate: (entry: LogEntry) => void;
    onSaveResult: (result: { ok: boolean; message?: string }) => void;
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

window.onSaveResult = (result) => {
  saveResultCallback?.(result);
};

export function onInit(cb: InitCallback) {
  initCallback = cb;
}

export function onLogUpdate(cb: LogUpdateCallback) {
  logUpdateCallback = cb;
}

export function onSaveResult(cb: SaveResultCallback) {
  saveResultCallback = cb;
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
  postMessage({
    action: "saveSettings",
    titleMatch: config.titleMatch,
    pasteMethod: config.pasteMethod,
    bindIp: config.bindIp,
    httpPort: config.httpPort,
    jpegQuality: config.jpegQuality,
    shiftInsertPaste: config.shiftInsertPaste ? 1 : 0,
  });
}

export function clearLog() {
  postMessage({ action: "clearLog" });
}

export function copyLog() {
  postMessage({ action: "copyLog" });
}

export function closeDialog() {
  postMessage({ action: "close" });
}

export function reportHeight(height: number) {
  postMessage({ action: "resize", height });
}
