# ImagePaster

A Windows system tray utility that intercepts `Ctrl+V` and converts clipboard images to base64-encoded PNG text before pasting. Designed for terminal applications like XShell, PuTTY, and SSH clients that don't support native image pasting.

## Features

- Intercepts `Ctrl+V` when a matching window is focused and the clipboard contains an image
- Converts the image to a base64-encoded PNG string and pastes that instead
- Configurable window title matching (comma-separated keywords)
- Modern WebView2-based configuration and activity log dialogs (React + Tailwind CSS)
- In-memory activity log with live updates (500-entry ring buffer)
- Configuration stored in the Windows registry (`HKCU\SOFTWARE\JPIT\ImagePaster`)
- System tray icon with context menu
- Single-instance enforcement

## Requirements

- **Windows 7+**
- **Microsoft Edge WebView2 Runtime** — required for the configuration and activity log dialogs. Usually pre-installed on Windows 10/11; can be downloaded from [Microsoft](https://developer.microsoft.com/en-us/microsoft-edge/webview2/).

## How It Works

1. A low-level keyboard hook monitors for `Ctrl+V` globally
2. When detected, it checks if the focused window's title contains any configured keyword
3. If a match is found and the clipboard contains an image (`CF_DIB`):
   - The image is extracted from the clipboard
   - Encoded to PNG using GDI+
   - Base64-encoded
   - Placed back on the clipboard as plain text
   - `Ctrl+V` is re-injected so the application receives the base64 string

## Building

Requires MinGW-w64 cross-compiler and Node.js (for the frontend build).

```sh
make
```

This builds the React frontend (`assets/dist/index.html`), compiles resources, and outputs `release/ImagePaster.exe`.

To build only the frontend:

```sh
make assets
```

To clean all build artifacts:

```sh
make clean
```

## Configuration

Right-click the tray icon and select **Configuration** to open the settings dialog.

| Setting | Registry Value | Type | Default |
|---------|---------------|------|---------|
| Title Match | `TitleMatch` | REG_SZ | `xshell` |

The title match field accepts comma-separated keywords (e.g. `xshell, putty, terminal`). Matching is case-insensitive and checks for substring presence in the focused window's title.

Settings are stored under `HKEY_CURRENT_USER\SOFTWARE\JPIT\ImagePaster`.

## Project Structure

```
├── main.c              # Application source (tray icon, keyboard hook, WebView2 integration)
├── resource.h          # Resource IDs
├── resources.rc        # Resource definitions (icon, HTML, DLL)
├── Makefile            # Cross-compilation build system
├── assets/
│   ├── src/
│   │   ├── App.tsx           # Root component (view router, resize reporting)
│   │   ├── ConfigView.tsx    # Configuration dialog
│   │   ├── LogView.tsx       # Activity log table
│   │   ├── lib/
│   │   │   ├── bridge.ts     # C <-> JS communication bridge
│   │   │   └── utils.ts      # Tailwind merge utility
│   │   └── components/ui/    # Reusable UI components (button, input, label)
│   ├── icon.ico              # Application icon (multi-size)
│   ├── WebView2Loader.dll    # Embedded WebView2 loader
│   ├── package.json          # Frontend dependencies
│   ├── vite.config.ts        # Vite + single-file plugin config
│   └── tailwind.config.ts    # Tailwind CSS config
└── release/
    └── ImagePaster.exe       # Built executable
```

## License

[MIT](LICENSE)
