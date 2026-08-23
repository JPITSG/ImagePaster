import { useEffect, useMemo, useRef, useState } from "react";
import {
  type UpdateResult,
  saveSettings,
  closeDialog,
  onSaveResult,
  checkForUpdate,
  cancelUpdateCheck,
  installUpdate,
  dismissUpdate,
  dismissUpdateConfirmation,
  onUpdateResult,
  onUpdateProgress,
  type ConfigData,
  type PasteMethod,
} from "./lib/bridge";
import { Button } from "./components/ui/button";
import { Checkbox } from "./components/ui/checkbox";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";

interface Props {
  config: ConfigData;
  webView2Version: string;
  updateCompletedVersion: string;
}

const selectClass =
  "flex h-8 w-full rounded-md border border-neutral-300 bg-white px-3 py-1 text-xs shadow-sm focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-neutral-400";

function isValidIpv4(value: string) {
  const parts = value.split(".");
  return (
    parts.length === 4 &&
    parts.every((part) => {
      if (!/^\d{1,3}$/.test(part)) return false;
      const number = Number(part);
      return number >= 0 && number <= 255 && String(number) === part;
    })
  );
}

export default function ConfigView({
  config,
  webView2Version,
  updateCompletedVersion,
}: Props) {
  const configuredIpIsDetected = config.availableIps.includes(config.bindIp);
  const [titleMatch, setTitleMatch] = useState(config.titleMatch);
  const [pasteMethod, setPasteMethod] = useState<PasteMethod>(config.pasteMethod);
  const [ipChoice, setIpChoice] = useState(
    configuredIpIsDetected ? config.bindIp : "other",
  );
  const [customIp, setCustomIp] = useState(
    configuredIpIsDetected ? "" : config.bindIp,
  );
  const [httpPort, setHttpPort] = useState(String(config.httpPort));
  const [jpegQuality, setJpegQuality] = useState(String(config.jpegQuality));
  const [shiftInsertPaste, setShiftInsertPaste] = useState(
    config.shiftInsertPaste,
  );
  const [autoCheckForUpdates, setAutoCheckForUpdates] = useState(
    config.autoCheckForUpdates ?? true,
  );
  const [error, setError] = useState("");
  const [updateChecking, setUpdateChecking] = useState(false);
  const [updateCancelling, setUpdateCancelling] = useState(false);
  const [updateSpeedKbps, setUpdateSpeedKbps] = useState<number | null>(null);
  const [updateAlert, setUpdateAlert] = useState<UpdateResult | null>(() =>
    updateCompletedVersion
      ? {
          status: "completed",
          title: "Update complete",
          message: `ImagePaster has been updated to version ${updateCompletedVersion}.`,
          currentVersion: "",
          remoteVersion: "",
        }
      : null,
  );
  const updateRequestMode = useRef<"automatic" | "manual" | null>(null);
  const automaticUpdateStarted = useRef(false);

  const selectedBindIp = useMemo(
    () => (ipChoice === "other" ? customIp.trim() : ipChoice),
    [customIp, ipChoice],
  );

  useEffect(() => {
    onSaveResult((result) => {
      if (!result.ok) setError(result.message ?? "Could not save the settings.");
    });
  }, []);

  useEffect(() => {
    const removeResultListener = onUpdateResult((result) => {
      const wasAutomatic = updateRequestMode.current === "automatic";
      updateRequestMode.current = null;
      setUpdateChecking(false);
      setUpdateCancelling(false);
      setUpdateSpeedKbps(null);
      if (result.status === "cancelled") {
        setUpdateAlert(null);
      } else if (wasAutomatic && result.status !== "newer") {
        dismissUpdate();
        setUpdateAlert(null);
      } else {
        setUpdateAlert(result);
      }
    });
    const removeProgressListener = onUpdateProgress((progress) => {
      setUpdateSpeedKbps(Math.max(0, Math.round(progress.kilobytesPerSecond)));
    });

    if (
      config.autoCheckForUpdates &&
      !updateCompletedVersion &&
      !automaticUpdateStarted.current
    ) {
      automaticUpdateStarted.current = true;
      updateRequestMode.current = "automatic";
      setUpdateChecking(true);
      checkForUpdate();
    }

    return () => {
      removeResultListener();
      removeProgressListener();
    };
  }, [config.autoCheckForUpdates, updateCompletedVersion]);

  function handleUpdate() {
    if (updateChecking) {
      setUpdateCancelling(true);
      cancelUpdateCheck();
      return;
    }
    setUpdateAlert(null);
    setUpdateChecking(true);
    setUpdateCancelling(false);
    setUpdateSpeedKbps(null);
    updateRequestMode.current = "manual";
    checkForUpdate();
  }

  function handleInstallUpdate() {
    setUpdateChecking(true);
    setUpdateCancelling(false);
    setUpdateSpeedKbps(null);
    installUpdate();
  }

  function handleDismissUpdate() {
    if (updateAlert?.status === "completed") {
      dismissUpdateConfirmation();
    } else {
      dismissUpdate();
    }
    setUpdateAlert(null);
  }

  const handleSave = () => {
    const port = Number(httpPort);
    const quality = Number(jpegQuality);
    setError("");

    if (!isValidIpv4(selectedBindIp)) {
      setError("Enter a valid IPv4 bind address.");
      return;
    }
    if (!Number.isInteger(port) || port < 1 || port > 65535) {
      setError("HTTP port must be between 1 and 65535.");
      return;
    }
    if (!Number.isInteger(quality) || quality < 0 || quality > 100) {
      setError("JPEG quality must be between 0 and 100.");
      return;
    }

    saveSettings({
      ...config,
      titleMatch: titleMatch.trim(),
      pasteMethod,
      bindIp: selectedBindIp,
      httpPort: port,
      jpegQuality: quality,
      shiftInsertPaste,
      autoCheckForUpdates,
    });
  };

  return (
    <div className="p-5 space-y-4">
      <div className="space-y-1.5">
        <Label htmlFor="pasteMethod">Paste Method</Label>
        <select
          id="pasteMethod"
          className={selectClass}
          value={pasteMethod}
          onChange={(event) => setPasteMethod(event.target.value as PasteMethod)}
        >
          <option value="base64">Base64-encoded PNG</option>
          <option value="http">HTTP image URL</option>
        </select>
        <p className="text-[11px] text-neutral-500">
          HTTP mode pastes a short instruction containing the current image URL.
        </p>
      </div>

      <div className="space-y-1.5">
        <Label htmlFor="titleMatch">Application Title Match</Label>
        <Input
          id="titleMatch"
          value={titleMatch}
          onChange={(event) => setTitleMatch(event.target.value)}
          placeholder="xshell, putty, terminal"
        />
        <p className="text-[11px] text-neutral-500">
          Comma-separated, case-insensitive window-title keywords.
        </p>
      </div>

      <div className="rounded-md border border-neutral-200 px-3 py-2.5">
        <label
          htmlFor="shiftInsertPaste"
          className="flex cursor-pointer items-center gap-2 text-xs font-medium"
        >
          <Checkbox
            id="shiftInsertPaste"
            checked={shiftInsertPaste}
            onChange={(event) => setShiftInsertPaste(event.target.checked)}
          />
          Use Shift+Insert for generated text
        </label>
        <p className="mt-1.5 pl-6 text-[11px] text-neutral-500">
          Xshell compatibility workaround. Disable it for applications that expect
          the generic Ctrl+V shortcut.
        </p>
      </div>

      <div className="border-t border-neutral-200 pt-4 space-y-3">
        <div>
          <div className="text-xs font-medium">Image HTTP Server</div>
          <p className="mt-1 text-[11px] text-neutral-500">
            The server stays active in both paste modes so the latest clipboard image
            is always ready.
          </p>
        </div>

        <div className="grid grid-cols-[1fr_120px] gap-3">
          <div className="space-y-1.5">
            <Label htmlFor="bindIp">Bind Address</Label>
            <select
              id="bindIp"
              className={selectClass}
              value={ipChoice}
              onChange={(event) => setIpChoice(event.target.value)}
            >
              {config.availableIps.map((ip) => (
                <option key={ip} value={ip}>
                  {ip}
                </option>
              ))}
              <option value="other">Other…</option>
            </select>
          </div>

          <div className="space-y-1.5">
            <Label htmlFor="httpPort">Port</Label>
            <Input
              id="httpPort"
              type="number"
              min={1}
              max={65535}
              value={httpPort}
              onChange={(event) => setHttpPort(event.target.value)}
            />
          </div>
        </div>

        {ipChoice === "other" && (
          <div className="space-y-1.5">
            <Label htmlFor="customIp">Custom IPv4 Address</Label>
            <Input
              id="customIp"
              value={customIp}
              onChange={(event) => setCustomIp(event.target.value)}
              placeholder="192.168.1.100"
            />
            {!config.bindIpAvailable && customIp === config.bindIp && (
              <p className="text-[11px] text-amber-700">
                This saved address is currently unavailable. It will remain saved and
                the server will start automatically if the address returns.
              </p>
            )}
          </div>
        )}

        <div className="space-y-1.5">
          <Label htmlFor="jpegQuality">JPEG Quality (%)</Label>
          <Input
            id="jpegQuality"
            type="number"
            min={0}
            max={100}
            value={jpegQuality}
            onChange={(event) => setJpegQuality(event.target.value)}
          />
        </div>

        <div className="rounded-md bg-neutral-50 px-3 py-2 text-[11px] text-neutral-600">
          Server status: {config.serverStatus}
        </div>
      </div>

      <div className="flex items-start gap-2 pt-1">
        <Checkbox
          id="autoCheckForUpdates"
          className="mt-0.5"
          checked={autoCheckForUpdates}
          onChange={(event) => setAutoCheckForUpdates(event.target.checked)}
        />
        <div className="space-y-0.5">
          <Label htmlFor="autoCheckForUpdates" className="cursor-pointer">
            Automatically check for updates
          </Label>
          <p className="text-neutral-500 text-[11px] leading-snug">
            Checks whenever this dialog opens and prompts only when a newer
            version is available.
          </p>
        </div>
      </div>

      {error && (
        <div className="rounded-md bg-red-50 px-3 py-2 text-[11px] text-red-700">
          {error}
        </div>
      )}

      <div className="flex items-center justify-between gap-3 pt-1">
        <span
          className="select-none whitespace-nowrap text-[11px] leading-none tabular-nums text-neutral-400"
          title="Application version / WebView2 version"
        >
          v{config.version} / {webView2Version}
        </span>
        <div className="flex items-center gap-2">
          <Button
            variant={updateChecking ? "destructive" : "outline"}
            size="sm"
            className="min-w-[5rem]"
            disabled={updateCancelling}
            aria-label={
              updateChecking ? "Stop update check and download" : undefined
            }
            title={
              updateChecking ? "Stop update check and download" : undefined
            }
            onClick={handleUpdate}
          >
            {updateCancelling
              ? "Stopping..."
              : updateChecking
                ? updateSpeedKbps === null
                  ? "Checking..."
                  : `Checking (${updateSpeedKbps.toLocaleString()} KB/s)...`
                : "Update"}
          </Button>
          <Button
            variant="outline"
            size="sm"
            className="min-w-[5rem]"
            onClick={closeDialog}
          >
            Cancel
          </Button>
          <Button size="sm" className="min-w-[5rem]" onClick={handleSave}>
            Save
          </Button>
        </div>
      </div>

      {updateAlert && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/35 p-4">
          <div
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="update-alert-title"
            aria-describedby="update-alert-message"
            className="w-full max-w-sm space-y-3 rounded-lg border border-neutral-200 bg-white p-4 shadow-xl"
          >
            <div className="space-y-1">
              <h2 id="update-alert-title" className="text-sm font-semibold">
                {updateAlert.title}
              </h2>
              <p
                id="update-alert-message"
                className="text-xs leading-relaxed text-neutral-600"
              >
                {updateAlert.message}
              </p>
            </div>
            {updateAlert.currentVersion && updateAlert.remoteVersion && (
              <dl className="grid grid-cols-[1fr_auto] gap-x-4 gap-y-1 rounded-md border border-neutral-200 bg-neutral-50 px-3 py-2 text-xs">
                <dt className="text-neutral-500">Current version</dt>
                <dd className="font-medium tabular-nums text-neutral-900">
                  {updateAlert.currentVersion}
                </dd>
                <dt className="text-neutral-500">Remote version</dt>
                <dd className="font-medium tabular-nums text-neutral-900">
                  {updateAlert.remoteVersion}
                </dd>
              </dl>
            )}
            <div className="flex justify-end gap-2">
              {(updateAlert.status === "newer" ||
                updateAlert.status === "same") && (
                <Button
                  variant="outline"
                  size="sm"
                  autoFocus
                  disabled={updateChecking}
                  onClick={handleDismissUpdate}
                >
                  Cancel
                </Button>
              )}
              <Button
                size="sm"
                autoFocus={
                  updateAlert.status !== "newer" &&
                  updateAlert.status !== "same"
                }
                disabled={updateChecking}
                onClick={
                  updateAlert.status === "newer" ||
                  updateAlert.status === "same"
                    ? handleInstallUpdate
                    : handleDismissUpdate
                }
              >
                {updateChecking
                  ? "Starting..."
                  : updateAlert.status === "same"
                    ? "Force update"
                    : updateAlert.status === "newer"
                      ? "Update"
                      : "OK"}
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
