export type PasteMethod = "base64" | "http";
export type CaptureGapFill = "white" | "black" | "blur";
export type ImageStorage = "memory" | "disk";

export interface ConfigData {
  titleMatch: string;
  pasteMethod: PasteMethod;
  httpMessageTemplate: string;
  bindIp: string;
  httpPort: number;
  httpAllowList: string;
  jpegQuality: number;
  imageHistoryLimit: number;
  imageStorage: ImageStorage;
  imageStorageDir: string;
  compatibilityPaste: boolean;
  screenCaptureEnabled: boolean;
  captureGapFill: CaptureGapFill;
  autoCheckForUpdates: boolean;
  updateCheckPending: boolean;
  updatePromptPending: boolean;
  availableIps: string[];
  bindIpAvailable: boolean;
  serverStatus: string;
  version: string;
}

export interface LogEntry {
  time: string;
  message: string;
}

export interface HistoryEntry {
  token: string;
  current: boolean;
  width: number;
  height: number;
  bytes: number;
  capturedAt: number;
  storage: ImageStorage;
  path: string;
  url: string;
}

/** One page of retained images, newest first; thumbnails arrive separately. */
export interface HistoryData {
  pasteMethod: PasteMethod;
  historyLimit: number;
  total: number;
  totalBytes: number;
  /** Zero-based page index. */
  page: number;
  pageSize: number;
  entries: HistoryEntry[];
}

/** An empty thumb means no preview could be made for that image. */
export interface HistoryThumb {
  token: string;
  thumb: string;
}

export interface HistoryActionResult {
  ok: boolean;
  message?: string;
}

export interface InitData {
  view: "config" | "log" | "history";
  config?: ConfigData;
  log?: LogEntry[];
  history?: HistoryData;
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
  automatic: boolean;
}

export interface UpdateProgress {
  percentComplete: number;
}

type InitCallback = (data: InitData) => void;
type LogUpdateCallback = (entry: LogEntry) => void;
type SaveResultCallback = (result: { ok: boolean; message?: string }) => void;

let initCallback: InitCallback | null = null;
let logUpdateCallback: LogUpdateCallback | null = null;
let saveResultCallback: SaveResultCallback | null = null;
let updateResultCallback: ((result: UpdateResult) => void) | null = null;
let updateProgressCallback: ((progress: UpdateProgress) => void) | null = null;
let historyActionResultCallback:
  | ((result: HistoryActionResult) => void)
  | null = null;
let historyDataCallback: ((history: HistoryData) => void) | null = null;
let historyThumbsCallback: ((thumbs: HistoryThumb[]) => void) | null = null;

declare global {
  interface Window {
    onInit: (data: InitData) => void;
    onLogUpdate: (entry: LogEntry) => void;
    onSaveResult: (result: { ok: boolean; message?: string }) => void;
    onUpdateResult: (result: UpdateResult) => void;
    onUpdateProgress: (progress: UpdateProgress) => void;
    onHistoryActionResult: (result: HistoryActionResult) => void;
    onHistoryData: (history: HistoryData) => void;
    onHistoryThumbs: (thumbs: HistoryThumb[]) => void;
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

window.onHistoryActionResult = (result) => {
  historyActionResultCallback?.(result);
};

window.onHistoryData = (history) => {
  historyDataCallback?.(history);
};

window.onHistoryThumbs = (thumbs) => {
  historyThumbsCallback?.(thumbs);
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
    httpAllowList: config.httpAllowList,
    jpegQuality: config.jpegQuality,
    imageHistoryLimit: config.imageHistoryLimit,
    imageStorage: config.imageStorage,
    compatibilityPaste: config.compatibilityPaste ? 1 : 0,
    screenCaptureEnabled: config.screenCaptureEnabled ? 1 : 0,
    captureGapFill: config.captureGapFill,
    autoCheckForUpdates: config.autoCheckForUpdates ? 1 : 0,
  });
}

export function configReady(checkAutomatically = false) {
  postMessage({
    action: "configReady",
    checkAutomatically: checkAutomatically ? 1 : 0,
  });
}

export function checkForUpdate(automatic = false) {
  postMessage({ action: "checkUpdate", automatic: automatic ? 1 : 0 });
}

export function cancelUpdateCheck() {
  postMessage({ action: "cancelUpdateCheck" });
}

export function installUpdate(reopenSettingsAfterUpdate: boolean) {
  postMessage({
    action: "installUpdate",
    reopenSettingsAfterUpdate: reopenSettingsAfterUpdate ? 1 : 0,
  });
}

export function dismissUpdate() {
  postMessage({ action: "dismissUpdate" });
}

export function ignoreUpdateVersion(version: string) {
  postMessage({ action: "ignoreUpdateVersion", version });
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

export function onHistoryActionResult(
  cb: (result: HistoryActionResult) => void,
) {
  historyActionResultCallback = cb;
  return () => {
    if (historyActionResultCallback === cb) historyActionResultCallback = null;
  };
}

export function onHistoryData(cb: (history: HistoryData) => void) {
  historyDataCallback = cb;
  return () => {
    if (historyDataCallback === cb) historyDataCallback = null;
  };
}

export function onHistoryThumbs(cb: (thumbs: HistoryThumb[]) => void) {
  historyThumbsCallback = cb;
  return () => {
    if (historyThumbsCallback === cb) historyThumbsCallback = null;
  };
}

export function requestHistoryPage(page: number) {
  postMessage({ action: "historyPage", page });
}

/** Replaces any earlier request: rows no longer listed are not generated. */
export function requestHistoryThumbs(tokens: string[]) {
  postMessage({ action: "historyThumbs", tokens: tokens.join(",") });
}

export function copyHistoryUrl(token: string) {
  postMessage({ action: "historyCopyUrl", token });
}

export function openHistoryUrl(token: string) {
  postMessage({ action: "historyOpenUrl", token });
}

export function revealHistoryFile(token: string) {
  postMessage({ action: "historyRevealFile", token });
}

export function saveHistoryImage(token: string) {
  postMessage({ action: "historySave", token });
}

export function deleteHistoryImage(token: string) {
  postMessage({ action: "historyDelete", token });
}

export function clearHistoryImages() {
  postMessage({ action: "historyClearAll" });
}

export function closeDialog() {
  postMessage({ action: "close" });
}

export function reportHeight(height: number) {
  postMessage({ action: "resize", height });
}

export function reportSize(height: number, width: number) {
  postMessage({ action: "resize", height, width });
}
