# ImagePaster 1.0.24

A Windows system tray utility that makes clipboard images usable in terminal applications such as Xshell, PuTTY, and other SSH clients that cannot forward the Windows image clipboard to a remote CLI.

## Features

- Intercepts `Ctrl+V` when a matching window is focused and the latest user clipboard item is an image
- Selectable paste method:
  - raw base64-encoded PNG text (legacy behavior)
  - a short instruction containing a temporary HTTP JPEG URL
- Watches the clipboard continuously and keeps the latest image encoded and ready for both methods
- Retains a configurable JPEG history (1–1000 total images or Unlimited) in memory or on disk until exit
- Selectable image storage: memory (default) or disk under `%LOCALAPPDATA%\ImagePaster`, with graceful fallback to memory when the folder or a file cannot be written and automatic file cleanup on eviction, exit, and startup
- History dialog with live-updating thumbnails, capture time and size details, per-image save to disk, URL copying, deletion, and one-click clear-all; clicking a thumbnail or URL opens the image in the default browser, and each entry shows whether it is stored in memory or on disk with a click-to-reveal path that opens File Explorer with the file selected
- Clears the current paste target when the user copies non-image content while preserving configured HTTP history
- Built-in IPv4 HTTP server with configurable bind address, port, and JPEG quality
- Optional HTTP client allowlist restricting image downloads to specific IPv4 addresses and CIDR subnets
- Customizable HTTP paste message with case-sensitive `{URL}` substitution
- Detects active local IPv4 addresses for selection in the configuration dialog
- Keeps an unavailable saved address and automatically starts listening if that address returns
- Uses a random 256-bit image identifier; retained URLs stay available and evicted URLs return `410 Gone`
- Configurable window title matching (comma-separated keywords)
- Optional, application-neutral compatibility paste mode using Shift+Insert
- Optional interactive Print Screen capture across the full multi-monitor desktop
- Dimmed capture overlay with cross-monitor selection and a glass-style toolbar on every display
- Shift-additive multi-region capture with pixel-accurate spacing and configurable white, black, or heavily blurred gaps
- Double-buffered, anti-aliased capture rendering for smooth selection feedback
- Self update with embedded-version comparison, cancellation, safe replacement, and restart
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
3. A clipboard listener keeps the latest image ready as:
   - a base64-encoded PNG, always held in memory, for the legacy method
   - a JPEG at the configured quality for HTTP delivery, kept in memory or on
     disk according to the **Image Storage** setting
4. If a matching window receives `Ctrl+V` while an image is current:
   - Base64 mode places the encoded PNG text on the clipboard.
   - HTTP mode replaces `{URL}` in the configured message with the current image
     URL and places the resulting text on the clipboard. The default is:

     ```text
     [ image available at http://192.168.1.100:10444/<random-id>.jpg - if you feel this image will be useful later on be sure to save it to /tmp or a temp location for later use ]
     ```

5. The configured text-paste shortcut is injected so the target application receives the selected representation. Compatibility mode uses Shift+Insert for applications that handle Ctrl+V themselves; disabling it uses standard Ctrl+V re-injection.

The HTTP server runs in both modes. URLs for the current image and retained history return `200 OK` with `image/jpeg`. When a retained image exceeds the configured limit, its URL returns `410 Gone` with a plain-language response body and header. Unknown or malformed image paths return `404 Not Found`. History is discarded when ImagePaster exits regardless of the storage setting.

**Image Storage** chooses where retained JPEGs live. Memory (the default) keeps everything in RAM, exactly as before. Disk writes each new image to `%LOCALAPPDATA%\ImagePaster\<image-id>.jpg` and serves HTTP requests, thumbnails, and saves from that file; the base64 PNG for the current image always stays in memory so base64 pasting is unaffected. The setting applies to newly copied images — existing entries keep their location until evicted. Failure handling is deliberately graceful: if the folder cannot be created or a file cannot be written, the failure is logged to the Activity Log and that image is kept in memory instead; nothing crashes and pasting keeps working. Disk files are deleted when their image is evicted, deleted, cleared, or when ImagePaster exits, and stale `<image-id>.jpg` files left by an unclean shutdown are removed at the next startup.

If the **Allowed Clients** list is non-empty, only connections from the listed IPv4 addresses and CIDR subnets are served; everything else is dropped before the request is read and the rejection is recorded in the Activity Log. An empty list allows every client (equivalent to `0.0.0.0/0`).

### Interactive Print Screen Capture

When **Enable interactive Print Screen capture** is enabled, ImagePaster replaces
the normal Print Screen action with a multi-monitor capture workflow:

1. Pressing `Print Screen` freezes and gently darkens the full Windows virtual
   desktop. A compact Clip, Copy, and Cancel toolbar appears near the bottom of
   every monitor.
2. The Clip tool is selected initially. Drag anywhere across the virtual desktop,
   including across monitor boundaries, to reveal a bright selection with a
   dashed outline and live dimensions. Hold `Shift` while dragging additional
   regions to retain the existing selections; dragging without `Shift` starts a
   new selection set. Existing boxes can be picked up and moved by dragging
   inside them, and removed with the ✕ pill at their top-right corner;
   `Shift`+drag inside a box draws a new region across it instead.
3. Copy places the selection on the Windows clipboard. If multiple regions are
   selected, their leftmost, topmost, rightmost, and bottommost edges define the
   output bounds. Selected pixels keep their exact relative positions. The
   configurable **Multi-region Gap Fill** makes all unselected space inside those
   bounds white, black, or a heavy blur of the real underlying desktop. If no
   selection exists, Copy captures the complete virtual desktop. The overlay then
   closes and the image enters the normal ImagePaster cache and history pipeline.
4. Pressing `Print Screen` again while the overlay is open immediately copies the
   complete virtual desktop. Pressing `Esc` or choosing Cancel closes the overlay
   without changing the clipboard.

The overlay and its controls are rendered from a pre-overlay snapshot, so neither
the dimming nor the toolbar is included in the copied image.

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
| HTTP Paste Message | `HttpMessageTemplate` | REG_SZ | `[ image available at {URL} - ... ]` |
| HTTP Bind Address | `BindIp` | REG_SZ | `127.0.0.1` |
| HTTP Port | `HttpPort` | REG_DWORD | `10444` |
| Allowed Clients | `HttpAllowList` | REG_SZ | Empty (allow all) |
| JPEG Quality | `JpegQuality` | REG_DWORD | `80` |
| Image History Limit | `ImageHistoryLimit` | REG_DWORD | `1` |
| Image Storage | `ImageStorage` | REG_DWORD | Memory |
| Compatibility Paste | `CompatibilityPaste` | REG_DWORD | Enabled |
| Interactive Print Screen Capture | `ScreenCaptureEnabled` | REG_DWORD | Disabled |
| Multi-region Gap Fill | `CaptureGapFill` | REG_DWORD | White |
| Automatically Check for Updates | `AutoCheckForUpdates` | REG_DWORD | Enabled |
| Ignored Update Version | `IgnoredUpdateVersion` | REG_SZ | Empty |

The title match field accepts comma-separated keywords (e.g. `xshell, putty, terminal`). Matching is case-insensitive and checks for substring presence in the focused window's title.

The HTTP paste message is used when HTTP mode intercepts `Ctrl+V`. It must
contain the case-sensitive `{URL}` placeholder; every occurrence is replaced
with the current image URL. Quotes, backslashes, line breaks, and Unicode text
are preserved.

The bind-address menu lists IPv4 addresses on active adapters and includes an **Other** option. If a saved address disappears, such as after a laptop changes networks, ImagePaster retains it, stops the unavailable listener safely, and retries periodically. Selecting a non-loopback address may require a Windows Firewall rule, and the remote machine must be able to route to that address.

Settings are stored under `HKEY_CURRENT_USER\SOFTWARE\JPIT\ImagePaster`.

The history limit counts the current image: `1` keeps only the current image,
`2` keeps it plus one historical image, and so on through `1000`. Selecting
**Unlimited** stores `0` in the registry and retains every image until the
application exits.

**Multi-region Gap Fill** affects only the unselected pixels inside the combined
bounds of two or more capture boxes. Its registry values are `0` for White, `1`
for Black, and `2` for Heavy blur.

The footer displays the application version as `v<application version>`.

When **Automatically check for updates** is enabled, ImagePaster checks at
startup, whenever the Configuration dialog opens, and once every 60 minutes
using a single low-frequency Windows timer. A newer build opens the
Configuration dialog and its update prompt; matching or older builds and failed
automatic checks are silently discarded. An automatically opened prompt offers
**Ignore this version**, which suppresses that version during later automatic
checks, including after restart. The manual **Update** button still displays
every result and can install an ignored version. Checks use the repository's
[`release/ImagePaster.exe`](release/ImagePaster.exe).

The update check downloads the executable to the user's temporary directory
and compares its embedded Windows file version with the running executable's
version. While downloading, the button displays the current transfer speed and
can be clicked again to stop the check and remove the partial download. The
result dialog displays both version numbers. A newer build can be installed
normally, while a matching build offers a **Force update** action to reinstall
it; an older repository build is never installed. Installation requests
standard Windows UAC approval, safely replaces the current executable, and
restarts ImagePaster. After a successful update, the restarted application
opens an HTML confirmation with the newly installed version. Dismissing that
confirmation leaves the configuration dialog open. Cancelling the download,
result dialog, or UAC prompt leaves the current version running. File size is
used only to validate the download and enforce its safety limit.

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
