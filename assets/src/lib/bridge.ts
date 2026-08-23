export type PasteMethod = "base64" | "http";

export interface ConfigData {
  titleMatch: string;
  pasteMethod: PasteMethod;
  httpMessageTemplate: string;
  bindIp: string;
  httpPort: number;
  jpegQuality: number;
  imageHistoryLimit: number;
  compatibilityPaste: boolean;
  screenCaptureEnabled: boolean;
  autoCheckForUpdates: boolean;
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
  updateCompletedVersion?: string;
}

export interface UpdateResult {
  status:
    | "newer"
    | "same"
    | "older"
    | "cancelled"
    | "error"
    | "completed";
  title: string;
  message: string;
  currentVersion: string;
  remoteVersion: string;
}

export interface UpdateProgress {
  kilobytesPerSecond: number;
}

type InitCallback = (data: InitData) => void;
type LogUpdateCallback = (entry: LogEntry) => void;
type SaveResultCallback = (result: { ok: boolean; message?: string }) => void;

let initCallback: InitCallback | null = null;
let logUpdateCallback: LogUpdateCallback | null = null;
let saveResultCallback: SaveResultCallback | null = null;
let updateResultCallback: ((result: UpdateResult) => void) | null = null;
let updateProgressCallback: ((progress: UpdateProgress) => void) | null = null;

declare global {
  interface Window {
    onInit: (data: InitData) => void;
    onLogUpdate: (entry: LogEntry) => void;
    onSaveResult: (result: { ok: boolean; message?: string }) => void;
    onUpdateResult: (result: UpdateResult) => void;
    onUpdateProgress: (progress: UpdateProgress) => void;
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

window.onUpdateResult = (result) => {
  updateResultCallback?.(result);
};

window.onUpdateProgress = (progress) => {
  updateProgressCallback?.(progress);
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

export function onUpdateResult(cb: (result: UpdateResult) => void) {
  updateResultCallback = cb;
  return () => {
    if (updateResultCallback === cb) updateResultCallback = null;
  };
}

export function onUpdateProgress(cb: (progress: UpdateProgress) => void) {
  updateProgressCallback = cb;
  return () => {
    if (updateProgressCallback === cb) updateProgressCallback = null;
  };
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
    httpMessageTemplate: config.httpMessageTemplate,
    bindIp: config.bindIp,
    httpPort: config.httpPort,
    jpegQuality: config.jpegQuality,
    imageHistoryLimit: config.imageHistoryLimit,
    compatibilityPaste: config.compatibilityPaste ? 1 : 0,
    screenCaptureEnabled: config.screenCaptureEnabled ? 1 : 0,
    autoCheckForUpdates: config.autoCheckForUpdates ? 1 : 0,
  });
}

export function checkForUpdate() {
  postMessage({ action: "checkUpdate" });
}

export function cancelUpdateCheck() {
  postMessage({ action: "cancelUpdateCheck" });
}

export function installUpdate() {
  postMessage({ action: "installUpdate" });
}

export function dismissUpdate() {
  postMessage({ action: "dismissUpdate" });
}

export function dismissUpdateConfirmation() {
  postMessage({ action: "dismissUpdateConfirmation" });
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
