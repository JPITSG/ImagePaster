# ImagePaster 1.0.42

A Windows system tray utility that makes clipboard images usable in terminal applications such as Xshell, PuTTY, and other SSH clients that cannot forward the Windows image clipboard to a remote CLI.

## Features

- Intercepts `Ctrl+V` when a matching window is focused and the latest user clipboard item is an image
- Selectable paste method:
  - raw base64-encoded PNG text (legacy behavior)
  - a short instruction containing a temporary HTTP JPEG URL
- Watches the clipboard continuously and keeps the latest image encoded and ready for both methods
- Retains a configurable JPEG history (1–1000 total images or Unlimited) in memory for the current session or on disk across application restarts
- Selectable image storage: memory (default) or persistent disk storage under `%LOCALAPPDATA%\ImagePaster`, with graceful fallback to memory when an image cannot be written and JPEG-based recovery when the durable history index is missing or invalid
- History dialog that opens instantly and pages through any number of images (50 per page, newest first), with thumbnails loaded lazily in the background as rows scroll into view, live updates, capture time and size details, per-image save to disk, URL copying, deletion, and one-click clear-all; clicking a thumbnail or URL opens the image in the default browser, and each entry shows whether it is stored in memory or on disk with a click-to-reveal path that opens File Explorer with the file selected
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
- Optional start with Windows at sign-in (per user, no administrator rights needed)
- System tray icon with a live status tooltip and context menu
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

The keyboard hook runs on a dedicated, above-normal-priority input thread,
separate from image encoding, disk storage, and dialogs. Ordinary keys return
immediately; Print Screen/Esc only read atomic capture state and enqueue work.
They never inspect window titles, query the clipboard, allocate, log, or acquire
application locks. At most one Print Screen action and one cancel action can be
pending until the UI finishes them, so repeated taps cannot flood a busy UI.
Identical pending image-paste requests are also coalesced; if the target or
clipboard changes while that slot is busy, a new paste passes through normally.
The Ctrl+V path uses non-waiting lock attempts, skips title checks for non-image
pastes, and bypasses the app's own injected paste events. ImagePaster's own
dialogs use normal clipboard handling.

The input thread renews the hook every 5 seconds and retries failed installations
without discarding an existing hook. While interactive capture is enabled (or a
manually opened overlay is active), it registers an unmodified Print Screen
hotkey, and the capture starts from that hotkey. Windows generates a hotkey only
after every program's low-level keyboard hook has passed the key on, so a tool
that shares this keyboard with another computer decides first where Print Screen
goes, and ImagePaster captures only when the key stays on this computer. The
hook merely notes each plain Print Screen press and passes it on (with Shift,
Control, Alt or a Windows key held, the hook still captures itself); a hotkey for
a press the hook never saw means Windows removed the hook, which is then
reinstalled at once. If another program holds Print Screen, the conflict is logged and retried,
the hook captures instead (ahead of keyboard-sharing tools), and the settings
show the conflict in red under the capture option. Keyboards that expose only
the Print Screen key-up transition are always handled by the hook. The
registration is released when capture is disabled and no overlay is active.

Windows can still remove a timed-out hook, and OS scheduling, another app's hook,
or an unavailable hotkey cannot be controlled by ImagePaster. These measures
minimize application-side stalls and provide recovery, not an absolute delivery
guarantee. A blocked UI can still delay the overlay even while input handling
remains responsive. No system timeout registry changes or realtime priority are
used.

The HTTP server runs in both modes. URLs for the current image and retained history return `200 OK` with `image/jpeg`. When a retained image exceeds the configured limit, its URL returns `410 Gone` with a plain-language response body and header. Unknown or malformed image paths return `404 Not Found`. Memory-backed history is discarded when ImagePaster exits; disk-backed history, metadata, ordering, and URLs are restored on the next launch when Disk storage is selected.

**Image Storage** chooses where retained JPEGs live. Memory (the default) keeps everything in RAM for the current session. Disk writes each new image to `%LOCALAPPDATA%\ImagePaster\<image-id>.jpg` and maintains an atomic history index in the same folder. On restart, indexed JPEGs are validated and restored with their original metadata, ordering, and URL tokens; existing token-named JPEGs from versions without an index are imported automatically. The base64 PNG for the current clipboard image remains in memory, so base64 pasting is unaffected. The setting applies to newly copied images — existing entries keep their location until evicted. If the folder, image file, or index cannot be written, the failure is logged and image capture continues without crashing. Disk files and their index entries are removed when images are evicted, deleted, or cleared, while indexed files present at exit remain available in the next session.

The **History** dialog opens with metadata only (sizes, times, URLs, and
storage), 50 images per page, newest first; the **« ‹ Page N of M › »** controls
reach older images. Thumbnails are requested only for rows on screen or about
to scroll into view. A background thread decodes each JPEG with the Windows
Imaging Component, which lets the JPEG decoder shrink a large screenshot while
decoding it instead of decoding it at full size (GDI+ remains as a fallback),
and 256-pixel previews appear as they finish. Rows scrolled past, or pages left
before their previews are ready, are skipped. The 256 most recently used
previews stay cached for the session, so reopening the dialog or returning to a
page is immediate. Disk-backed images are read without holding the image cache
lock, so browsing History never delays clipboard handling. Each preview is shown
whole and at its true proportions on a neutral mat, with a hairline edge and a
soft shadow so dark and white screenshots both stand out; small images are not
enlarged, and the loading placeholder already has the image's shape.

If the **Allowed Clients** list is non-empty, only connections from the listed IPv4 addresses and CIDR subnets are served; everything else is dropped before the request is read and the rejection is recorded in the Activity Log. An empty list allows every client (equivalent to `0.0.0.0/0`).

### Interactive Print Screen Capture

When **Enable interactive Print Screen capture** is enabled, ImagePaster replaces
the normal Print Screen action with a multi-monitor capture workflow:

1. Pressing `Print Screen` (or choosing **Capture** from the tray menu, which
   works even while Print Screen interception is disabled) freezes and gently
   darkens the full Windows virtual desktop. A compact Clip, Copy, and Cancel
   toolbar appears near the bottom of every monitor; it fades to half opacity
   while a box is being drawn and disappears entirely while the drag cursor
   passes over it.
2. The Clip tool is selected initially. Drag anywhere across the virtual desktop,
   including across monitor boundaries, to reveal a bright selection with a
   dashed outline and live dimensions. Hold `Shift` while dragging additional
   regions to retain the existing selections; dragging without `Shift` starts a
   new selection set. Existing boxes can be picked up and moved by dragging
   inside them, resized by dragging any of their four walls or corners, and
   removed with the ✕ pill at their top-right corner; `Shift`+drag inside a
   box draws a new region across it instead.
3. Copy places the selection on the Windows clipboard. After creating a box,
   pressing `Enter` performs the same action as **Copy Selection**. If multiple
   regions are selected, their leftmost, topmost, rightmost, and bottommost edges
   define the output bounds. Selected pixels keep their exact relative positions.
   The configurable **Multi-region Gap Fill** makes all unselected space inside
   those bounds white, black, or a heavy blur of the real underlying desktop. If
   no selection exists, Copy captures the complete virtual desktop. The overlay
   then closes and the image enters the normal ImagePaster cache and history
   pipeline.
4. Pressing `Print Screen` again while the overlay is open immediately copies the
   complete virtual desktop. Pressing `Esc` or choosing Cancel closes the overlay
   without changing the clipboard.

The overlay and its controls are rendered from a pre-overlay snapshot, so neither
the dimming nor the toolbar is included in the copied image.

When Print Screen is claimed by another program, a red notice under **Enable
interactive Print Screen capture** says so: capture still works, but tools that
share the keyboard with another computer cannot send Print Screen there. The
notice follows the key live while settings are open, and appears as soon as the
option is ticked.

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

Right-click the tray icon and select **Configure** to open the settings dialog.
Selecting **Configure** again restores and focuses the existing dialog.
Double-clicking the tray icon starts interactive capture when that feature is
enabled; otherwise it opens or refocuses **Configure**.

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
| Start with Windows | `ImagePaster` (see below) | REG_SZ | Off |
| Automatically Check for Updates | `AutoCheckForUpdates` | REG_DWORD | Enabled |
| Ignored Update Version | `IgnoredUpdateVersion` | REG_SZ | Empty |

The title match field accepts comma-separated keywords (e.g. `xshell, putty, terminal`). Matching is case-insensitive and checks for substring presence in the focused window's title.

The HTTP paste message is used when HTTP mode intercepts `Ctrl+V`. It must
contain the case-sensitive `{URL}` placeholder; every occurrence is replaced
with the current image URL. Quotes, backslashes, line breaks, and Unicode text
are preserved.

The bind-address menu lists IPv4 addresses on active adapters and includes an **Other** option. If a saved address disappears, such as after a laptop changes networks, ImagePaster retains it, stops the unavailable listener safely, and retries periodically. Selecting a non-loopback address may require a Windows Firewall rule, and the remote machine must be able to route to that address.

Settings are stored under `HKEY_CURRENT_USER\SOFTWARE\JPIT\ImagePaster`, except
**Start with Windows**: it adds or removes an `ImagePaster` value under
`HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run` that launches
this executable when you sign in. An entry disabled in Task Manager's startup
apps shows as off; turning the toggle on re-enables it.

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
version. While downloading, the red button displays the share of the file
received so far as a whole percentage, for example **Checking (42%)...**; it
reads 100% only once the last byte has arrived. Clicking the red button again
stops the check immediately, even when the connection has stalled: the button
returns to **Update** at once, the background download is abandoned, and its
partial file is removed as soon as the pending network call returns. Its late
result is never shown, and an **Update** click made in the meantime starts a
fresh check as soon as that cleanup finishes. The
result dialog displays both version numbers. Installable results also offer an
unchecked-by-default
**Reopen settings after update** checkbox. This choice applies only to that
installation and is not saved as a preference. A newer build can be installed
normally, while a matching build offers a **Force update** action to reinstall
it; an older repository build is never installed. Installation requests
standard Windows UAC approval, safely replaces the current executable, and
restarts ImagePaster. After a successful update, settings stay closed unless
the checkbox was checked. If checked, the restarted application opens settings
with an HTML confirmation of the newly installed version; dismissing the
confirmation leaves settings open. Cancelling the download,
result dialog, or UAC prompt leaves the current version running. File size is
used only to validate the download and enforce its safety limit.

Native regression checks can run without launching the Windows application:
`python3 -B -m unittest discover -s tests -v` compiles production hook, updater
progress and stop handling, History paging and thumbnail, Start with Windows,
and command-line parsing functions with inert operating-system boundaries
(requires a host C compiler). Start with Windows tests run the Run entry against
an in-memory registry: the quoted path, Task Manager's disabled marker, an entry
left by another copy, case-insensitive paths, and logged failures. Updater tests cover
stopping mid-download or just as a result arrives, stale progress, clicks while
a stopped check unwinds, closing settings, and lost result messages. History
tests parse every pushed page and preview as JSON and cover page order and
clamping, visible-only requests that replace earlier ones, cached and failed
previews, rows re-requested mid-decode, batching, eviction, and that nothing
decodes or calls into the page while holding a lock. Hook tests cover blocked UI/contended locks, millions
of unrelated key events, bounded capture queues, held keys across renewal,
Print Screen passed on to the hotkey (so a hook behind ImagePaster can send it to
another computer), hotkey conflicts/fallback, silent removal, installation
failures, and shutdown.
These are deterministic fault-injection tests, not Windows latency measurements.
On Windows, smoke-test capture on/off, manual tray capture, holding Print Screen
across a renewal, rapid taps during large image processing, Ctrl+V in a matching
window, hotkey conflicts, Print Screen with a keyboard-sharing tool pointed at
another computer, and sleep/resume. Actual input delivery and long-running
reliability still require a Windows desktop.

After `make`, `python3 tests/check_update_ui.py`,
`python3 tests/check_history_ui.py`, `python3 tests/check_capture_ui.py` and
`python3 tests/check_startup_ui.py` check the built UI with a recording WebView
bridge (requires Python Playwright and Chromium). The startup check covers the
Startup section's wording, its place above Updates in both layouts, and the
saved choice. The capture check covers the Print Screen conflict notice: its
wording, red text, place under the capture option, and live updates. The History check covers
visible-only thumbnail requests, scrolling, failed previews, preview
proportions without blur or layout shift, paging, late previews, and live
refreshes. These checks do not exercise Windows UAC,
replacement, relaunch, or real WIC decoding.

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
