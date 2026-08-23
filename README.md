# ImagePaster 1.0.1

A Windows system tray utility that makes clipboard images usable in terminal applications such as Xshell, PuTTY, and other SSH clients that cannot forward the Windows image clipboard to a remote CLI.

## Features

- Intercepts `Ctrl+V` when a matching window is focused and the latest user clipboard item is an image
- Selectable paste method:
  - raw base64-encoded PNG text (legacy behavior)
  - a short instruction containing a temporary HTTP JPEG URL
- Watches the clipboard continuously and keeps the latest image encoded in memory for both methods
- Clears the cached image when the user copies non-image content
- Built-in IPv4 HTTP server with configurable bind address, port, and JPEG quality
- Detects active local IPv4 addresses for selection in the configuration dialog
- Keeps an unavailable saved address and automatically starts listening if that address returns
- Uses a random 256-bit image identifier; superseded URLs return `410 Gone`
- Configurable window title matching (comma-separated keywords)
- Optional Shift+Insert text-paste compatibility mode for terminals such as Xshell
- Modern WebView2-based configuration and activity log dialogs (React + Tailwind CSS)
- In-memory activity log with live updates and one-click clipboard copying (500-entry ring buffer)
- Configuration stored in the Windows registry (`HKCU\SOFTWARE\JPIT\ImagePaster`)
- System tray icon with context menu
- Single-instance enforcement

## Requirements

- **Windows 7+**
- **Microsoft Edge WebView2 Runtime** — required for the configuration and activity log dialogs. Usually pre-installed on Windows 10/11; can be downloaded from [Microsoft](https://developer.microsoft.com/en-us/microsoft-edge/webview2/).

## How It Works

1. A low-level keyboard hook monitors for `Ctrl+V` globally
2. When detected, it checks if the focused window's title contains any configured keyword
3. A clipboard listener keeps an in-memory representation of the latest image as:
   - base64-encoded PNG for the legacy method
   - JPEG at the configured quality for HTTP delivery
4. If a matching window receives `Ctrl+V` while an image is current:
   - Base64 mode places the encoded PNG text on the clipboard.
   - HTTP mode places text similar to the following on the clipboard:

     ```text
     [ image available at http://192.168.1.100:10444/<random-id>.jpg - if you feel this image will be useful later on be sure to save it to /tmp or a temp location for later use ]
     ```

5. The configured text-paste shortcut is injected so the target application receives the selected representation. The default Shift+Insert compatibility mode avoids forwarding a second Ctrl+V into Xshell; it can be disabled to use generic Ctrl+V re-injection.

The HTTP server runs in both modes. A URL for the current image returns `200 OK` with `image/jpeg`. A previously issued URL returns `410 Gone` with a plain-language response body and header. Unknown or malformed image paths return `404 Not Found`.

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
| Paste Method | `PasteMethod` | REG_DWORD | Base64 |
| HTTP Bind Address | `BindIp` | REG_SZ | `127.0.0.1` |
| HTTP Port | `HttpPort` | REG_DWORD | `10444` |
| JPEG Quality | `JpegQuality` | REG_DWORD | `80` |
| Shift+Insert Paste | `ShiftInsertPaste` | REG_DWORD | Enabled |

The title match field accepts comma-separated keywords (e.g. `xshell, putty, terminal`). Matching is case-insensitive and checks for substring presence in the focused window's title.

The bind-address menu lists IPv4 addresses on active adapters and includes an **Other** option. If a saved address disappears, such as after a laptop changes networks, ImagePaster retains it, stops the unavailable listener safely, and retries periodically. Selecting a non-loopback address may require a Windows Firewall rule, and the remote machine must be able to route to that address.

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
