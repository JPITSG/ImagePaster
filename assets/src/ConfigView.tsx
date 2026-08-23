import { useEffect, useMemo, useState } from "react";
import {
  saveSettings,
  closeDialog,
  onSaveResult,
  type ConfigData,
  type PasteMethod,
} from "./lib/bridge";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";

interface Props {
  config: ConfigData;
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

export default function ConfigView({ config }: Props) {
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
  const [error, setError] = useState("");

  const selectedBindIp = useMemo(
    () => (ipChoice === "other" ? customIp.trim() : ipChoice),
    [customIp, ipChoice],
  );

  useEffect(() => {
    onSaveResult((result) => {
      if (!result.ok) setError(result.message ?? "Could not save the settings.");
    });
  }, []);

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
          <input
            id="shiftInsertPaste"
            type="checkbox"
            checked={shiftInsertPaste}
            onChange={(event) => setShiftInsertPaste(event.target.checked)}
            className="h-4 w-4 rounded border-neutral-300"
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

      {error && (
        <div className="rounded-md bg-red-50 px-3 py-2 text-[11px] text-red-700">
          {error}
        </div>
      )}

      <div className="flex items-center justify-between gap-2 pt-1">
        <span className="text-[10px] text-neutral-400">v{config.version}</span>
        <div className="flex gap-2">
          <Button variant="outline" size="sm" className="w-20" onClick={closeDialog}>
            Cancel
          </Button>
          <Button size="sm" className="w-20" onClick={handleSave}>
            Save
          </Button>
        </div>
      </div>
    </div>
  );
}
