import { useState } from "react";
import { saveSettings, closeDialog, type ConfigData } from "./lib/bridge";
import { Button } from "./components/ui/button";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";

interface Props {
  config: ConfigData;
}

export default function ConfigView({ config }: Props) {
  const [titleMatch, setTitleMatch] = useState(config.titleMatch);

  const handleSave = () => {
    saveSettings({ titleMatch: titleMatch.trim() });
  };

  const handleCancel = () => {
    closeDialog();
  };

  return (
    <div className="p-5 space-y-4">
      <div className="space-y-1.5">
        <Label htmlFor="titleMatch">Application Title Match</Label>
        <p className="text-[11px] text-neutral-500 font-normal">
          When you press Ctrl+V with an image on the clipboard, it will be converted to base64 text if the focused window's title contains any of these keywords.
        </p>
        <Input
          id="titleMatch"
          value={titleMatch}
          onChange={(e) => setTitleMatch(e.target.value)}
          placeholder="xshell, putty, terminal"
        />
      </div>

      <div className="flex justify-end gap-2 pt-2">
        <Button variant="outline" size="sm" className="w-20" onClick={handleCancel}>
          Cancel
        </Button>
        <Button size="sm" className="w-20" onClick={handleSave}>
          Save
        </Button>
      </div>
    </div>
  );
}
