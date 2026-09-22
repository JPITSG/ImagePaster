/*
 * ImagePaster - main.c
 *
 * System tray utility that intercepts Ctrl+V when a matching window is focused
 * and the clipboard contains an image. It can paste either a raw base64-encoded
 * PNG string or a short URL served by the built-in HTTP image server.
 *
 * Features:
 *   - Shared clipboard image cache (in memory or on disk) with configurable
 *     JPEG history
 *   - Configurable base64 or HTTP URL paste mode
 *   - Optional multi-monitor Print Screen capture with multi-region selection
 *   - Configurable title matching and HTTP bind settings (registry-persisted)
 *   - WebView2-based configuration and activity log modals
 *   - System tray icon with context menu
 *   - In-memory log ring buffer pushed live to the Activity Log view
 *
 * Cross-compiled with MinGW-w64 using GDI+ flat C API.
 */

#define UNICODE
#define _UNICODE
#define _WIN32_WINNT 0x0600
#define WIN32_LEAN_AND_MEAN
#define COBJMACROS

#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#include <iphlpapi.h>
#include <wincrypt.h>
#include <commctrl.h>
#include <commdlg.h>
#include <objbase.h>
#include <wincodec.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <sddl.h>
#include <bcrypt.h>
#include <winhttp.h>
#include <winver.h>
#include <userenv.h>

#ifndef WM_DPICHANGED
#define WM_DPICHANGED 0x02E0
#endif
#ifndef MOD_NOREPEAT
#define MOD_NOREPEAT 0x4000
#endif

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "resource.h"

/* ── GDI+ flat API declarations ─────────────────────────────────────────── */

typedef int GpStatus;
typedef void GpBitmap;
typedef void GpImage;
typedef void GpBrush;
typedef void GpSolidFill;
typedef void GpFont;
typedef void GpFontCollection;
typedef void GpFontFamily;
typedef void GpGraphics;
typedef void GpLineGradient;
typedef void GpPath;
typedef void GpPen;
typedef void GpStringFormat;

typedef struct {
    float X;
    float Y;
    float Width;
    float Height;
} CaptureGpRectF;

#pragma pack(push, 8)
typedef struct {
    UINT32 GdiplusVersion;
    void *DebugEventCallback;
    BOOL SuppressBackgroundThread;
    BOOL SuppressExternalCodecs;
} GdiplusStartupInput;

typedef struct {
    CLSID Clsid;
    GUID  FormatID;
    const WCHAR *CodecName;
    const WCHAR *DllName;
    const WCHAR *FormatDescription;
    const WCHAR *FilenameExtension;
    const WCHAR *MimeType;
    DWORD Flags;
    DWORD Version;
    DWORD SigCount;
    DWORD SigSize;
    const BYTE *SigPattern;
    const BYTE *SigMask;
} ImageCodecInfo;

typedef struct {
    GUID Guid;
    ULONG NumberOfValues;
    ULONG Type;
    VOID *Value;
} EncoderParameter;

typedef struct {
    UINT Count;
    EncoderParameter Parameter[1];
} EncoderParameters;
#pragma pack(pop)

/* GDI+ flat API imports */
GpStatus __stdcall GdiplusStartup(ULONG_PTR *token, const GdiplusStartupInput *input, void *output);
void     __stdcall GdiplusShutdown(ULONG_PTR token);
GpStatus __stdcall GdipCreateBitmapFromGdiDib(const BITMAPINFO *gdiBitmapInfo, void *gdiBitmapData, GpBitmap **bitmap);
GpStatus __stdcall GdipCreateBitmapFromStream(IStream *stream, GpBitmap **bitmap);
GpStatus __stdcall GdipLoadImageFromFile(const WCHAR *filename, GpImage **image);
GpStatus __stdcall GdipCreateBitmapFromScan0(INT width, INT height, INT stride,
                                             INT format, BYTE *scan0,
                                             GpBitmap **bitmap);
GpStatus __stdcall GdipGetImageGraphicsContext(GpImage *image,
                                               GpGraphics **graphics);
GpStatus __stdcall GdipSetInterpolationMode(GpGraphics *graphics, int mode);
GpStatus __stdcall GdipDrawImageRectI(GpGraphics *graphics, GpImage *image,
                                      INT x, INT y, INT width, INT height);
GpStatus __stdcall GdipGetImageEncodersSize(UINT *numEncoders, UINT *size);
GpStatus __stdcall GdipGetImageEncoders(UINT numEncoders, UINT size, ImageCodecInfo *encoders);
GpStatus __stdcall GdipSaveImageToStream(GpImage *image, IStream *stream, const CLSID *clsidEncoder, const void *encoderParams);
GpStatus __stdcall GdipDisposeImage(GpImage *image);
GpStatus __stdcall GdipGetImageWidth(GpImage *image, UINT *width);
GpStatus __stdcall GdipGetImageHeight(GpImage *image, UINT *height);
GpStatus __stdcall GdipCreateFromHDC(HDC dc, GpGraphics **graphics);
GpStatus __stdcall GdipDeleteGraphics(GpGraphics *graphics);
GpStatus __stdcall GdipSetCompositingQuality(GpGraphics *graphics, int quality);
GpStatus __stdcall GdipSetSmoothingMode(GpGraphics *graphics, int mode);
GpStatus __stdcall GdipSetPixelOffsetMode(GpGraphics *graphics, int mode);
GpStatus __stdcall GdipSetTextRenderingHint(GpGraphics *graphics, int hint);
GpStatus __stdcall GdipCreatePen1(DWORD color, float width, int unit,
                                  GpPen **pen);
GpStatus __stdcall GdipDeletePen(GpPen *pen);
GpStatus __stdcall GdipSetPenStartCap(GpPen *pen, int cap);
GpStatus __stdcall GdipSetPenEndCap(GpPen *pen, int cap);
GpStatus __stdcall GdipSetPenDashCap197819(GpPen *pen, int cap);
GpStatus __stdcall GdipSetPenLineJoin(GpPen *pen, int join);
GpStatus __stdcall GdipSetPenDashStyle(GpPen *pen, int style);
GpStatus __stdcall GdipDrawLine(GpGraphics *graphics, GpPen *pen,
                                float x1, float y1, float x2, float y2);
GpStatus __stdcall GdipCreatePath(int fillMode, GpPath **path);
GpStatus __stdcall GdipDeletePath(GpPath *path);
GpStatus __stdcall GdipAddPathArc(GpPath *path, float x, float y,
                                  float width, float height,
                                  float startAngle, float sweepAngle);
GpStatus __stdcall GdipAddPathLine(GpPath *path, float x1, float y1,
                                   float x2, float y2);
GpStatus __stdcall GdipClosePathFigure(GpPath *path);
GpStatus __stdcall GdipDrawPath(GpGraphics *graphics, GpPen *pen,
                                GpPath *path);
GpStatus __stdcall GdipFillPath(GpGraphics *graphics, GpBrush *brush,
                                GpPath *path);
GpStatus __stdcall GdipCreateSolidFill(DWORD color, GpSolidFill **brush);
GpStatus __stdcall GdipCreateLineBrushFromRect(
    const CaptureGpRectF *rect, DWORD color1, DWORD color2,
    int mode, int wrapMode, GpLineGradient **lineGradient);
GpStatus __stdcall GdipDeleteBrush(GpBrush *brush);
GpStatus __stdcall GdipCreateFontFamilyFromName(
    const WCHAR *name, GpFontCollection *collection, GpFontFamily **family);
GpStatus __stdcall GdipGetGenericFontFamilySansSerif(GpFontFamily **family);
GpStatus __stdcall GdipDeleteFontFamily(GpFontFamily *family);
GpStatus __stdcall GdipCreateFont(const GpFontFamily *family, float size,
                                  int style, int unit, GpFont **font);
GpStatus __stdcall GdipDeleteFont(GpFont *font);
GpStatus __stdcall GdipCreateStringFormat(int flags, LANGID language,
                                          GpStringFormat **format);
GpStatus __stdcall GdipSetStringFormatFlags(GpStringFormat *format, int flags);
GpStatus __stdcall GdipSetStringFormatAlign(GpStringFormat *format,
                                            int alignment);
GpStatus __stdcall GdipSetStringFormatLineAlign(GpStringFormat *format,
                                                int alignment);
GpStatus __stdcall GdipDeleteStringFormat(GpStringFormat *format);
GpStatus __stdcall GdipDrawString(GpGraphics *graphics, const WCHAR *text,
                                  int length, const GpFont *font,
                                  const CaptureGpRectF *layout,
                                  const GpStringFormat *format,
                                  const GpBrush *brush);
GpStatus __stdcall GdipMeasureString(GpGraphics *graphics, const WCHAR *text,
                                     int length, const GpFont *font,
                                     const CaptureGpRectF *layout,
                                     const GpStringFormat *format,
                                     CaptureGpRectF *boundingBox,
                                     int *codepointsFitted, int *linesFilled);

/* ── Constants ──────────────────────────────────────────────────────────── */

#define APP_NAME          L"ImagePaster"
#define APP_VERSION_A     "1.0.40"
#define APP_VERSION_W     L"1.0.40"
#define MUTEX_NAME        L"ImagePaster_SingleInstance"
#define WM_TRAYICON       (WM_USER + 1)
#define WM_DO_PASTE       (WM_APP + 1)
#define WM_HTTP_EVENT      (WM_APP + 2)
#define WM_APP_UPDATE_RESULT   (WM_APP + 3)
#define WM_APP_UPDATE_PROGRESS (WM_APP + 4)
#define WM_SCREEN_CAPTURE_BEGIN  (WM_APP + 5)
#define WM_SCREEN_CAPTURE_COPY   (WM_APP + 6)
#define WM_SCREEN_CAPTURE_CANCEL (WM_APP + 7)
#define WM_KEYBOARD_HOOK_STATUS   (WM_APP + 8)
#define WM_KEYBOARD_HOTKEY_STATUS (WM_APP + 9)
#define WM_HISTORY_THUMBS_READY   (WM_APP + 10)
#define KEYBOARD_HOOK_RENEW_MS    5000u
#define ID_HOTKEY_PRINT_SCREEN   1
#define HOOK_CAPTURE_ENABLED     1
#define HOOK_CAPTURE_ACTIVE      2
#define HOOK_CAPTURE_MESSAGE     1
#define PASTE_INPUT_TAG          ((ULONG_PTR)0x49505053u)
#define ID_TRAY_LOG       1001
#define ID_TRAY_CONFIGURE 1002
#define ID_TRAY_EXIT      1003
#define ID_TRAY_HISTORY   1004
#define ID_TRAY_CAPTURE   1005
#define ID_TIMER_WEBVIEW_SHOW_FALLBACK 1006
#define ID_TIMER_HTTP_RECONCILE         1007
#define ID_TIMER_CLIPBOARD_RETRY        1008
#define ID_TIMER_DEFERRED_PASTE         1009
#define ID_TIMER_AUTO_UPDATE             1010
#define WEBVIEW_SHOW_FALLBACK_DELAY_MS 350
#define HTTP_RECONCILE_INTERVAL_MS      3000
#define CLIPBOARD_RETRY_DELAY_MS        150
#define DEFERRED_PASTE_DELAY_MS          10
#define AUTO_UPDATE_INTERVAL_MS          (60u * 60u * 1000u)

#define UPDATE_URL L"https://github.com/JPITSG/ImagePaster/raw/refs/heads/main/release/ImagePaster.exe"
#define UPDATE_MAX_BYTES (100ULL * 1024ULL * 1024ULL)
#define UPDATE_HELPER_READY_MS 10000
#define UPDATE_HELPER_WAIT_MS 120000

#define REG_KEY_PATH       "SOFTWARE\\JPIT\\ImagePaster"
#define REG_VALUE_TITLE    "TitleMatch"
#define REG_VALUE_METHOD   "PasteMethod"
#define REG_VALUE_HTTP_MESSAGE_TEMPLATE "HttpMessageTemplate"
#define REG_VALUE_BIND_IP  "BindIp"
#define REG_VALUE_PORT     "HttpPort"
#define REG_VALUE_QUALITY  "JpegQuality"
#define REG_VALUE_HISTORY_LIMIT "ImageHistoryLimit"
#define REG_VALUE_COMPATIBILITY_PASTE "CompatibilityPaste"
#define REG_VALUE_LEGACY_COMPATIBILITY_PASTE "ShiftInsertPaste"
#define REG_VALUE_SCREEN_CAPTURE "ScreenCaptureEnabled"
#define REG_VALUE_CAPTURE_GAP_FILL "CaptureGapFill"
#define REG_VALUE_AUTO_UPDATE "AutoCheckForUpdates"
#define REG_VALUE_IGNORED_UPDATE_VERSION "IgnoredUpdateVersion"
#define REG_VALUE_HTTP_ALLOW_LIST "HttpAllowList"
#define REG_VALUE_IMAGE_STORAGE "ImageStorage"

#define LOG_RING_CAPACITY  500
#define MAX_KEYWORDS       64
#define MAX_DETECTED_IPS    64
#define MAX_CAPTURE_MONITORS 32
#define IMAGE_TOKEN_BYTES   32
#define IMAGE_TOKEN_HEX_LEN (IMAGE_TOKEN_BYTES * 2)
#define DEFAULT_HTTP_PORT   10444
#define DEFAULT_JPEG_QUALITY 80
#define DEFAULT_IMAGE_HISTORY_LIMIT 1
#define MAX_IMAGE_HISTORY_LIMIT 1000
#define HTTP_URL_PLACEHOLDER "{URL}"
#define DEFAULT_HTTP_MESSAGE_TEMPLATE \
    "[ image available at " HTTP_URL_PLACEHOLDER \
    " - if you feel this image will be useful later on be sure to save it " \
    "to /tmp or a temp location for later use ]"
#define MAX_HTTP_MESSAGE_TEMPLATE_BYTES 8192
#define PASTE_METHOD_BASE64 0
#define PASTE_METHOD_HTTP   1
#define IMAGE_STORAGE_MEMORY 0
#define IMAGE_STORAGE_DISK   1
#define IMAGE_STORAGE_DIR_NAME L"ImagePaster"
#define IMAGE_HISTORY_INDEX_NAME L"history-v1.dat"
#define IMAGE_HISTORY_INDEX_TEMP_NAME L"history-v1.tmp"
#define IMAGE_HISTORY_INDEX_MAGIC "IPIDX001"
#define MAX_DISK_HISTORY_INDEX_RECORDS 1000000u
#define HTTP_EVENT_SERVED   200
#define HTTP_EVENT_GONE     410
#define HTTP_EVENT_NOT_FOUND 404
#define HTTP_EVENT_DENIED   403
#define MAX_HTTP_ALLOW_LIST_BYTES 512
#define MAX_HTTP_ALLOW_RULES 64
#define CAPTURE_TOOL_CLIP    0
#define CAPTURE_TOOL_COPY    1
#define CAPTURE_TOOL_CANCEL  2
#define CAPTURE_TOOL_COUNT   3
#define CAPTURE_GAP_FILL_WHITE 0
#define CAPTURE_GAP_FILL_BLACK 1
#define CAPTURE_GAP_FILL_BLUR  2
#define CAPTURE_PANEL_WIDTH  292
#define CAPTURE_PANEL_HEIGHT 74
#define CAPTURE_BUTTON_WIDTH 70
/* Copy is wider so its label can swap between "Copy All" and
   "Copy Selection" without clipping. */
#define CAPTURE_COPY_BUTTON_WIDTH 98
#define CAPTURE_BUTTON_HEIGHT 54
#define CAPTURE_BUTTON_GAP    8
#define CAPTURE_SEPARATOR_GAP 14
#define CAPTURE_PANEL_BOTTOM_MARGIN 40

/* History dialog: pages carry metadata only. Thumbnails are small JPEG data
   URIs made on a background thread when rows scroll into view, then cached. */
#define HISTORY_VIEW_PAGE_SIZE      50
#define HISTORY_THUMB_CACHE_ENTRIES 256
#define HISTORY_THUMB_QUEUE_MAX     HISTORY_VIEW_PAGE_SIZE
#define HISTORY_THUMB_PUSH_BATCH    8
#define HISTORY_THUMB_MAX_DIM      256
#define HISTORY_THUMB_JPEG_QUALITY 82
#define HISTORY_THUMB_PIXEL_FORMAT 0x0026200A /* PixelFormat32bppARGB */
#define GDIP_INTERPOLATION_HIGH_BICUBIC 7

/* ── Log ring buffer ───────────────────────────────────────────────────── */

typedef struct {
    char time[24];     /* HH:MM:SS.mmm */
    char message[512];
} LogEntry;

static LogEntry g_logRing[LOG_RING_CAPACITY];
static int g_logHead  = 0;   /* next write position */
static int g_logCount = 0;   /* total entries (capped at capacity) */

/* ── Globals ────────────────────────────────────────────────────────────── */

static HINSTANCE g_hInstance;
static HWND      g_hWndMain;
/* Only the dedicated hook thread owns the hook handle and key-down state. */
static HHOOK     g_hHook;
static HANDLE   g_keyboardHookThread;
static HANDLE   g_keyboardHookStopEvent;
static HANDLE   g_keyboardHookWakeEvent;
static volatile LONG g_hookCaptureState = 0;
static volatile LONG g_hookHasKeywords = FALSE;
static volatile LONG g_hookPrintPending = FALSE;
static volatile LONG g_hookCancelPending = FALSE;
static volatile LONG g_hookPastePending = FALSE;
static HWND g_hookPasteTarget = NULL; /* input thread only */
static DWORD g_hookPasteSequence = 0;
static volatile LONG g_hookRenewRequested = FALSE;
static BOOL g_printHotkeyRegistered = FALSE; /* input thread only */
static DWORD g_printHotkeyError = 0;
static DWORD g_hookInstallError = 0;
static ULONG_PTR g_gdipToken;
static HANDLE    g_hMutex;
static HICON     g_hAppIcon;
static NOTIFYICONDATAW g_nid;
static HMENU     g_hMenu;

static BOOL g_writingClipboardText = FALSE;
static DWORD g_ownClipboardSequence = 0;
static volatile LONG g_lastClipboardSequence = 0;
typedef BOOL (WINAPI *PFN_AddClipboardFormatListener)(HWND);
typedef BOOL (WINAPI *PFN_RemoveClipboardFormatListener)(HWND);
static PFN_AddClipboardFormatListener fnAddClipboardFormatListener = NULL;
static PFN_RemoveClipboardFormatListener fnRemoveClipboardFormatListener = NULL;

/* Persisted configuration */
static char g_configTitleMatch[2048] = "xshell";
static int  g_configPasteMethod = PASTE_METHOD_BASE64;
static char g_configHttpMessageTemplate[MAX_HTTP_MESSAGE_TEMPLATE_BYTES] =
    DEFAULT_HTTP_MESSAGE_TEMPLATE;
static char g_configBindIp[INET_ADDRSTRLEN] = "127.0.0.1";
static int  g_configHttpPort = DEFAULT_HTTP_PORT;
static int  g_configJpegQuality = DEFAULT_JPEG_QUALITY;
/* Zero means unlimited; finite limits include the current image. */
static int  g_configImageHistoryLimit = DEFAULT_IMAGE_HISTORY_LIMIT;
/* Where newly cached JPEGs are kept; memory is the default. */
static int  g_configImageStorage = IMAGE_STORAGE_MEMORY;
static BOOL g_configCompatibilityPaste = TRUE;
static BOOL g_configScreenCaptureEnabled = FALSE;
static int  g_configCaptureGapFill = CAPTURE_GAP_FILL_WHITE;
static BOOL g_configAutoCheckForUpdates = TRUE;
static char g_ignoredUpdateVersion[32] = "";
static BOOL g_pasteDeferred = FALSE;
static HWND g_deferredPasteTarget = NULL;
static DWORD g_deferredPasteSequence = 0;
static WCHAR g_keywords[MAX_KEYWORDS][128];
static int   g_keywordCount = 0;
static SRWLOCK g_keywordLock = SRWLOCK_INIT;

/* Clipboard image cache shared with the HTTP worker thread. */
typedef struct {
    BYTE *jpegData;  /* NULL when the JPEG lives on disk instead */
    DWORD jpegSize;  /* JPEG size in bytes regardless of location */
    char *base64Data;
    DWORD base64Len;
    char token[IMAGE_TOKEN_HEX_LEN + 1];
    UINT width;
    UINT height;
    ULONGLONG capturedAtMs; /* Unix epoch milliseconds, UTC */
    WCHAR *diskPath; /* malloc'd file path when stored on disk, else NULL */
} CachedImage;

#pragma pack(push, 1)
typedef struct {
    char magic[8];
    DWORD recordCount;
} DiskHistoryIndexHeader;

typedef struct {
    char token[IMAGE_TOKEN_HEX_LEN + 1];
    DWORD width;
    DWORD height;
    DWORD jpegSize;
    ULONGLONG capturedAtMs;
    BYTE isCurrent;
    BYTE reserved[7];
} DiskHistoryIndexRecord;
#pragma pack(pop)

typedef struct {
    CachedImage image;
    BOOL isCurrent;
} LoadedDiskImage;

static CachedImage g_cachedImage = {0};
static CachedImage *g_imageHistory = NULL; /* oldest to newest */
static size_t g_imageHistoryCount = 0;
static size_t g_imageHistoryCapacity = 0;
static SRWLOCK g_imageLock = SRWLOCK_INIT;
static BOOL g_diskHistoryCatalogLoaded = FALSE;
static BOOL g_restoredDiskCurrentNeedsClipboardCheck = FALSE;
static char (*g_goneTokens)[IMAGE_TOKEN_HEX_LEN + 1] = NULL;
static size_t g_goneTokenCount = 0;
static size_t g_goneTokenCapacity = 0;

/* Interactive Print Screen capture overlay. Coordinates are relative to the
   top-left of the Windows virtual desktop, which can have a negative origin. */
typedef struct {
    RECT monitorRect;
    RECT panelRect;
    RECT buttonRects[CAPTURE_TOOL_COUNT];
    UINT dpi;
} CapturePanel;

typedef struct {
    int first;
    int second;
    unsigned int fraction;
} CaptureBlurSample;

static HWND g_captureOverlayHwnd = NULL;
static HBITMAP g_captureOriginalBitmap = NULL;
static HBITMAP g_captureFrameBitmap = NULL;
static DWORD *g_captureOriginalPixels = NULL;
static int g_captureVirtualX = 0;
static int g_captureVirtualY = 0;
static int g_captureWidth = 0;
static int g_captureHeight = 0;
static CapturePanel g_capturePanels[MAX_CAPTURE_MONITORS];
static int g_capturePanelCount = 0;
static int g_captureSelectedTool = CAPTURE_TOOL_CLIP;
static int g_captureHoveredPanel = -1;
static int g_captureHoveredTool = -1;
static int g_capturePressedPanel = -1;
static int g_capturePressedTool = -1;
static BOOL g_captureDragging = FALSE;
static BOOL g_captureCopyShowsSelection = FALSE;
static int g_captureHoveredRemoval = -1; /* stored-selection index or -1 */
static int g_captureMovingIndex = -1;    /* box being dragged around, or -1 */
static POINT g_captureMoveGrabOffset = {0}; /* cursor minus box top-left */

enum {
    CAPTURE_EDGE_NONE = 0,
    CAPTURE_EDGE_LEFT,
    CAPTURE_EDGE_TOP,
    CAPTURE_EDGE_RIGHT,
    CAPTURE_EDGE_BOTTOM,
    CAPTURE_EDGE_TOPLEFT,
    CAPTURE_EDGE_TOPRIGHT,
    CAPTURE_EDGE_BOTTOMLEFT,
    CAPTURE_EDGE_BOTTOMRIGHT
};
static int g_captureResizingIndex = -1;  /* box being resized, or -1 */
static int g_captureResizingEdge = CAPTURE_EDGE_NONE;
/* Panel hidden because the drag cursor is currently over it, or -1. */
static int g_captureVanishedPanel = -1;
static POINT g_captureDragStart = {0};
static POINT g_captureDragCurrent = {0};
static RECT *g_captureSelections = NULL;
static size_t g_captureSelectionCount = 0;
static size_t g_captureSelectionCapacity = 0;
static HANDLE g_capturePreviousDpiContext = NULL;
static BOOL g_printScreenKeyDown = FALSE;
static BOOL g_escapeKeyDown = FALSE;

/* HTTP client allowlist. Empty text allows every client; otherwise only
   matching source addresses may connect. The parsed rules are shared with
   the HTTP worker thread. */
static char g_configHttpAllowList[MAX_HTTP_ALLOW_LIST_BYTES] = "";

typedef struct {
    DWORD network; /* host byte order, already masked */
    DWORD mask;    /* host byte order */
} HttpAllowRule;

static HttpAllowRule g_httpAllowRules[MAX_HTTP_ALLOW_RULES];
static int g_httpAllowRuleCount = 0;
static BOOL g_httpAllowRestrictive = FALSE;
static SRWLOCK g_httpAllowLock = SRWLOCK_INIT;

/* HTTP listener state. */
static BOOL g_winsockReady = FALSE;
static SOCKET g_httpListenSocket = INVALID_SOCKET;
static SOCKET g_httpClientSocket = INVALID_SOCKET;
static HANDLE g_httpThread = NULL;
static volatile LONG g_httpStopRequested = 0;
static CRITICAL_SECTION g_httpSocketLock;
static BOOL g_httpSocketLockReady = FALSE;
static char g_httpStatus[256] = "Not started";
static int g_httpState = 0; /* 0 stopped, 1 listening, 2 waiting, 3 error */

/* ── WebView2 COM interface definitions (minimal vtable approach) ─────── */

DEFINE_GUID(IID_ICoreWebView2Environment, 0xb96d755e,0x0319,0x4e92,0xa2,0x96,0x23,0x43,0x6f,0x46,0xa1,0xfc);
DEFINE_GUID(IID_ICoreWebView2Controller, 0x4d00c0d1,0x9583,0x4f38,0x8e,0x50,0xa9,0xa6,0xb3,0x44,0x78,0xcd);
DEFINE_GUID(IID_ICoreWebView2, 0x76eceacb,0x0462,0x4d94,0xac,0x83,0x42,0x3a,0x67,0x93,0x77,0x5e);
DEFINE_GUID(IID_ICoreWebView2Settings, 0xe562e4f0,0xd7fa,0x43ac,0x8d,0x71,0xc0,0x51,0x50,0x49,0x9f,0x00);

typedef struct EventRegistrationToken { __int64 value; } EventRegistrationToken;

typedef struct ICoreWebView2Environment ICoreWebView2Environment;
typedef struct ICoreWebView2Controller ICoreWebView2Controller;
typedef struct ICoreWebView2 ICoreWebView2;
typedef struct ICoreWebView2Settings ICoreWebView2Settings;
typedef struct ICoreWebView2WebMessageReceivedEventArgs ICoreWebView2WebMessageReceivedEventArgs;
typedef struct ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler;
typedef struct ICoreWebView2CreateCoreWebView2ControllerCompletedHandler ICoreWebView2CreateCoreWebView2ControllerCompletedHandler;
typedef struct ICoreWebView2WebMessageReceivedEventHandler ICoreWebView2WebMessageReceivedEventHandler;

/* ICoreWebView2Environment vtable */
typedef struct ICoreWebView2EnvironmentVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2Environment*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2Environment*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2Environment*);
    HRESULT (STDMETHODCALLTYPE *CreateCoreWebView2Controller)(ICoreWebView2Environment*, HWND, ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*);
    HRESULT (STDMETHODCALLTYPE *CreateWebResourceResponse)(ICoreWebView2Environment*, void*, int, LPCWSTR, LPCWSTR, void**);
    HRESULT (STDMETHODCALLTYPE *get_BrowserVersionString)(ICoreWebView2Environment*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *add_NewBrowserVersionAvailable)(ICoreWebView2Environment*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NewBrowserVersionAvailable)(ICoreWebView2Environment*, EventRegistrationToken);
} ICoreWebView2EnvironmentVtbl;
struct ICoreWebView2Environment { const ICoreWebView2EnvironmentVtbl *lpVtbl; };

/* ICoreWebView2Controller vtable */
typedef struct ICoreWebView2ControllerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2Controller*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2Controller*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2Controller*);
    HRESULT (STDMETHODCALLTYPE *get_IsVisible)(ICoreWebView2Controller*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsVisible)(ICoreWebView2Controller*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_Bounds)(ICoreWebView2Controller*, RECT*);
    HRESULT (STDMETHODCALLTYPE *put_Bounds)(ICoreWebView2Controller*, RECT);
    HRESULT (STDMETHODCALLTYPE *get_ZoomFactor)(ICoreWebView2Controller*, double*);
    HRESULT (STDMETHODCALLTYPE *put_ZoomFactor)(ICoreWebView2Controller*, double);
    HRESULT (STDMETHODCALLTYPE *add_ZoomFactorChanged)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ZoomFactorChanged)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *SetBoundsAndZoomFactor)(ICoreWebView2Controller*, RECT, double);
    HRESULT (STDMETHODCALLTYPE *MoveFocus)(ICoreWebView2Controller*, int);
    HRESULT (STDMETHODCALLTYPE *add_MoveFocusRequested)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_MoveFocusRequested)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_GotFocus)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_GotFocus)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_LostFocus)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_LostFocus)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_AcceleratorKeyPressed)(ICoreWebView2Controller*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_AcceleratorKeyPressed)(ICoreWebView2Controller*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *get_ParentWindow)(ICoreWebView2Controller*, HWND*);
    HRESULT (STDMETHODCALLTYPE *put_ParentWindow)(ICoreWebView2Controller*, HWND);
    HRESULT (STDMETHODCALLTYPE *NotifyParentWindowPositionChanged)(ICoreWebView2Controller*);
    HRESULT (STDMETHODCALLTYPE *Close)(ICoreWebView2Controller*);
    HRESULT (STDMETHODCALLTYPE *get_CoreWebView2)(ICoreWebView2Controller*, ICoreWebView2**);
} ICoreWebView2ControllerVtbl;
struct ICoreWebView2Controller { const ICoreWebView2ControllerVtbl *lpVtbl; };

/* ICoreWebView2 vtable */
typedef struct ICoreWebView2Vtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *get_Settings)(ICoreWebView2*, ICoreWebView2Settings**);
    HRESULT (STDMETHODCALLTYPE *get_Source)(ICoreWebView2*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *Navigate)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *NavigateToString)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *add_NavigationStarting)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NavigationStarting)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_ContentLoading)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ContentLoading)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_SourceChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_SourceChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_HistoryChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_HistoryChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_NavigationCompleted)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NavigationCompleted)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_FrameNavigationStarting)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_FrameNavigationStarting)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_FrameNavigationCompleted)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_FrameNavigationCompleted)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_ScriptDialogOpening)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ScriptDialogOpening)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_PermissionRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_PermissionRequested)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_ProcessFailed)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ProcessFailed)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *AddScriptToExecuteOnDocumentCreated)(ICoreWebView2*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *RemoveScriptToExecuteOnDocumentCreated)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *ExecuteScript)(ICoreWebView2*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *CapturePreview)(ICoreWebView2*, int, void*, void*);
    HRESULT (STDMETHODCALLTYPE *Reload)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *PostWebMessageAsJson)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *PostWebMessageAsString)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *add_WebMessageReceived)(ICoreWebView2*, ICoreWebView2WebMessageReceivedEventHandler*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_WebMessageReceived)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *CallDevToolsProtocolMethod)(ICoreWebView2*, LPCWSTR, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *get_BrowserProcessId)(ICoreWebView2*, UINT32*);
    HRESULT (STDMETHODCALLTYPE *get_CanGoBack)(ICoreWebView2*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *get_CanGoForward)(ICoreWebView2*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *GoBack)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *GoForward)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *GetDevToolsProtocolEventReceiver)(ICoreWebView2*, LPCWSTR, void**);
    HRESULT (STDMETHODCALLTYPE *Stop)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *add_NewWindowRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_NewWindowRequested)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *add_DocumentTitleChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_DocumentTitleChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *get_DocumentTitle)(ICoreWebView2*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *AddHostObjectToScript)(ICoreWebView2*, LPCWSTR, void*);
    HRESULT (STDMETHODCALLTYPE *RemoveHostObjectFromScript)(ICoreWebView2*, LPCWSTR);
    HRESULT (STDMETHODCALLTYPE *OpenDevToolsWindow)(ICoreWebView2*);
    HRESULT (STDMETHODCALLTYPE *add_ContainsFullScreenElementChanged)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_ContainsFullScreenElementChanged)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *get_ContainsFullScreenElement)(ICoreWebView2*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *add_WebResourceRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_WebResourceRequested)(ICoreWebView2*, EventRegistrationToken);
    HRESULT (STDMETHODCALLTYPE *AddWebResourceRequestedFilter)(ICoreWebView2*, LPCWSTR, int);
    HRESULT (STDMETHODCALLTYPE *RemoveWebResourceRequestedFilter)(ICoreWebView2*, LPCWSTR, int);
    HRESULT (STDMETHODCALLTYPE *add_WindowCloseRequested)(ICoreWebView2*, void*, EventRegistrationToken*);
    HRESULT (STDMETHODCALLTYPE *remove_WindowCloseRequested)(ICoreWebView2*, EventRegistrationToken);
} ICoreWebView2Vtbl;
struct ICoreWebView2 { const ICoreWebView2Vtbl *lpVtbl; };

/* ICoreWebView2Settings vtable */
typedef struct ICoreWebView2SettingsVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2Settings*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2Settings*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2Settings*);
    HRESULT (STDMETHODCALLTYPE *get_IsScriptEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsScriptEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsWebMessageEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsWebMessageEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreDefaultScriptDialogsEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreDefaultScriptDialogsEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsStatusBarEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsStatusBarEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreDevToolsEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreDevToolsEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreDefaultContextMenusEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreDefaultContextMenusEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_AreHostObjectsAllowed)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_AreHostObjectsAllowed)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsZoomControlEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsZoomControlEnabled)(ICoreWebView2Settings*, BOOL);
    HRESULT (STDMETHODCALLTYPE *get_IsBuiltInErrorPageEnabled)(ICoreWebView2Settings*, BOOL*);
    HRESULT (STDMETHODCALLTYPE *put_IsBuiltInErrorPageEnabled)(ICoreWebView2Settings*, BOOL);
} ICoreWebView2SettingsVtbl;
struct ICoreWebView2Settings { const ICoreWebView2SettingsVtbl *lpVtbl; };

/* ICoreWebView2WebMessageReceivedEventArgs vtable */
typedef struct ICoreWebView2WebMessageReceivedEventArgsVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2WebMessageReceivedEventArgs*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2WebMessageReceivedEventArgs*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2WebMessageReceivedEventArgs*);
    HRESULT (STDMETHODCALLTYPE *get_Source)(ICoreWebView2WebMessageReceivedEventArgs*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *get_WebMessageAsJson)(ICoreWebView2WebMessageReceivedEventArgs*, LPWSTR*);
    HRESULT (STDMETHODCALLTYPE *TryGetWebMessageAsString)(ICoreWebView2WebMessageReceivedEventArgs*, LPWSTR*);
} ICoreWebView2WebMessageReceivedEventArgsVtbl;
struct ICoreWebView2WebMessageReceivedEventArgs { const ICoreWebView2WebMessageReceivedEventArgsVtbl *lpVtbl; };

/* ── COM callback handler types ──────────────────────────────────────────── */

typedef struct EnvironmentCompletedHandlerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);
    HRESULT (STDMETHODCALLTYPE *Invoke)(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*, HRESULT, ICoreWebView2Environment*);
} EnvironmentCompletedHandlerVtbl;

struct ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler {
    const EnvironmentCompletedHandlerVtbl *lpVtbl;
    ULONG refCount;
};

typedef struct ControllerCompletedHandlerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*);
    HRESULT (STDMETHODCALLTYPE *Invoke)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, HRESULT, ICoreWebView2Controller*);
} ControllerCompletedHandlerVtbl;

struct ICoreWebView2CreateCoreWebView2ControllerCompletedHandler {
    const ControllerCompletedHandlerVtbl *lpVtbl;
    ULONG refCount;
};

typedef struct WebMessageReceivedHandlerVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(ICoreWebView2WebMessageReceivedEventHandler*, REFIID, void**);
    ULONG   (STDMETHODCALLTYPE *AddRef)(ICoreWebView2WebMessageReceivedEventHandler*);
    ULONG   (STDMETHODCALLTYPE *Release)(ICoreWebView2WebMessageReceivedEventHandler*);
    HRESULT (STDMETHODCALLTYPE *Invoke)(ICoreWebView2WebMessageReceivedEventHandler*, ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs*);
} WebMessageReceivedHandlerVtbl;

struct ICoreWebView2WebMessageReceivedEventHandler {
    const WebMessageReceivedHandlerVtbl *lpVtbl;
    ULONG refCount;
};

/* ── WebView2 globals ──────────────────────────────────────────────────── */

static HWND g_webviewHwnd = NULL;
static ICoreWebView2Environment *g_webviewEnv = NULL;
static ICoreWebView2Controller *g_webviewController = NULL;
static ICoreWebView2 *g_webviewView = NULL;
static char g_pendingView[16] = "";
static BOOL g_webviewWindowShown = FALSE;
static BOOL g_configViewReady = FALSE;
static BOOL g_updateConfirmationPending = FALSE;
static volatile LONG g_updateCheckPending = FALSE;
static volatile LONG g_updateCheckAutomatic = FALSE;
static BOOL g_updateInstallReady = FALSE;
static volatile LONG g_updateRequestSequence = 0;
static HANDLE g_updateCancelEvent = NULL;
static volatile LONG g_updateProgressPercent = 0;
static volatile LONG g_updateProgressPosted = FALSE;
/* UI thread only. A stopped check can remain inside a blocking network call
   until its timeout, so settings are told at once and its late result is
   discarded. A check requested meanwhile starts when that worker exits. */
static BOOL g_updateCheckAbandoned = FALSE;
static BOOL g_updateQueuedCheck = FALSE;
static BOOL g_updateQueuedAutomatic = FALSE;

typedef struct {
    WORD major;
    WORD minor;
    WORD patch;
    WORD build;
} ExecutableVersion;

typedef enum {
    UPDATE_CHECK_SAME = 1,
    UPDATE_CHECK_NEWER,
    UPDATE_CHECK_OLDER,
    UPDATE_CHECK_CANCELLED,
    UPDATE_CHECK_ERROR
} UpdateCheckKind;

typedef struct {
    HWND targetWindow;
    BOOL automatic;
    UpdateCheckKind kind;
    ULONGLONG cacheBuster;
    ExecutableVersion runningVersion;
    ExecutableVersion availableVersion;
    wchar_t message[512];
    wchar_t targetPath[MAX_PATH];
    wchar_t stagedPath[MAX_PATH];
} UpdateCheckTask;

static UpdateCheckTask* volatile g_updatePostedResult = NULL;
static UpdateCheckTask* g_updateNoticeTask = NULL;
static UpdateCheckTask* g_updateReadyTask = NULL;

typedef HRESULT (STDAPICALLTYPE *PFN_CreateCoreWebView2EnvironmentWithOptions)(
    LPCWSTR browserExecutableFolder, LPCWSTR userDataFolder, void* options,
    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler* handler);

static PFN_CreateCoreWebView2EnvironmentWithOptions fnCreateEnvironment = NULL;
static WCHAR g_extractedDllPath[MAX_PATH] = {0};

/* ── Forward declarations ──────────────────────────────────────────────── */

static void LogMessage(const char *fmt, ...);
static void ParseKeywords(void);
static BOOL LoadConfigFromRegistry(void);
static void SaveConfigToRegistry(void);
static void ShowWebViewDialog(const char* view, int width, int height);
static void UpdateTooltip(void);
static void ReconcileHttpServer(void);
static void StopHttpServer(void);
static BOOL RefreshClipboardImageCache(void);
static void ClearCurrentImage(const char *reason);
static void DestroyImageCache(void);
static void PersistDiskHistoryIndex(void);
static size_t LoadPersistedDiskHistory(void);
static BOOL BeginScreenCapture(void);
static BOOL CompleteScreenCapture(BOOL forceFullDesktop);
static void CancelScreenCapture(const char *reason);
static void StartUpdateCheck(BOOL automatic);
static void CancelUpdateCheck(void);
static void InstallPreparedUpdate(BOOL reopenSettings);
static void DiscardPreparedUpdate(void);

/* ── Base64 encoder ─────────────────────────────────────────────────────── */

static const char b64_table[] =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

static char *Base64Encode(const BYTE *data, DWORD len, DWORD *outLen)
{
    DWORD i, j;
    DWORD encLen = 4 * ((len + 2) / 3);
    char *out = (char *)malloc(encLen + 1);
    if (!out) return NULL;

    for (i = 0, j = 0; i < len; ) {
        DWORD a = i < len ? data[i++] : 0;
        DWORD b = i < len ? data[i++] : 0;
        DWORD c = i < len ? data[i++] : 0;
        DWORD triple = (a << 16) | (b << 8) | c;

        out[j++] = b64_table[(triple >> 18) & 0x3F];
        out[j++] = b64_table[(triple >> 12) & 0x3F];
        out[j++] = b64_table[(triple >>  6) & 0x3F];
        out[j++] = b64_table[ triple        & 0x3F];
    }

    /* padding */
    if (len % 3 >= 1) out[encLen - 1] = '=';
    if (len % 3 == 1) out[encLen - 2] = '=';

    out[encLen] = '\0';
    if (outLen) *outLen = encLen;
    return out;
}

/* ── Logging (in-memory ring buffer) ───────────────────────────────────── */

static void webview_execute_script(const wchar_t* script);
static void NotifyHistoryViewChanged(void);

static void LogMessage(const char *fmt, ...)
{
    char buf[512];
    SYSTEMTIME st;

    va_list args;
    va_start(args, fmt);
    vsnprintf(buf, sizeof(buf), fmt, args);
    va_end(args);
    buf[sizeof(buf) - 1] = '\0';

    GetLocalTime(&st);

    /* Write into ring buffer */
    LogEntry *entry = &g_logRing[g_logHead];
    wsprintfA(entry->time, "%02d:%02d:%02d.%03d",
              st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
    strncpy(entry->message, buf, sizeof(entry->message) - 1);
    entry->message[sizeof(entry->message) - 1] = '\0';

    g_logHead = (g_logHead + 1) % LOG_RING_CAPACITY;
    if (g_logCount < LOG_RING_CAPACITY) g_logCount++;

    /* If Activity Log WebView is open, push the new entry live */
    if (g_webviewView && strcmp(g_pendingView, "log") == 0) {
        wchar_t wTime[32], wMsg[1024], script[2048];
        MultiByteToWideChar(CP_UTF8, 0, entry->time, -1, wTime, 32);

        /* Escape the message for JSON embedding */
        size_t j = 0;
        for (size_t i = 0; entry->message[i] && j < 1020; i++) {
            char c = entry->message[i];
            if (c == '"' || c == '\\') {
                wMsg[j++] = L'\\';
                wMsg[j++] = (wchar_t)c;
            } else if (c == '\n') {
                wMsg[j++] = L'\\'; wMsg[j++] = L'n';
            } else if (c == '\r') {
                wMsg[j++] = L'\\'; wMsg[j++] = L'r';
            } else {
                wMsg[j++] = (wchar_t)(unsigned char)c;
            }
        }
        wMsg[j] = L'\0';

        swprintf(script, 2048,
            L"window.onLogUpdate && window.onLogUpdate({\"time\":\"%s\",\"message\":\"%s\"})",
            wTime, wMsg);
        webview_execute_script(script);
    }
}

/* ── Clipboard image encoding and shared cache ─────────────────────────── */

static const GUID g_encoderQualityGuid =
    {0x1d5be4b5, 0xfa4a, 0x452d, {0x9c, 0xdd, 0x5d, 0xb3, 0x51, 0x05, 0xe7, 0xeb}};

static BOOL GetEncoderClsid(const WCHAR *mimeType, CLSID *pClsid)
{
    UINT num = 0, size = 0;
    ImageCodecInfo *codecs = NULL;

    if (GdipGetImageEncodersSize(&num, &size) != 0 || size == 0) return FALSE;
    codecs = (ImageCodecInfo *)malloc(size);
    if (!codecs) return FALSE;
    if (GdipGetImageEncoders(num, size, codecs) != 0) {
        free(codecs);
        return FALSE;
    }

    for (UINT i = 0; i < num; i++) {
        if (codecs[i].MimeType && wcscmp(codecs[i].MimeType, mimeType) == 0) {
            *pClsid = codecs[i].Clsid;
            free(codecs);
            return TRUE;
        }
    }

    free(codecs);
    return FALSE;
}

static BOOL EncodeImageToMemory(GpImage *image, const WCHAR *mimeType, int quality,
                                BYTE **outData, DWORD *outSize)
{
    CLSID encoderClsid;
    IStream *stream = NULL;
    EncoderParameters params;
    EncoderParameters *paramsPtr = NULL;
    ULONG qualityValue = (ULONG)quality;
    STATSTG stat;
    LARGE_INTEGER zero;
    ULONG bytesRead = 0;
    BYTE *data = NULL;

    *outData = NULL;
    *outSize = 0;

    if (!GetEncoderClsid(mimeType, &encoderClsid)) return FALSE;
    if (CreateStreamOnHGlobal(NULL, TRUE, &stream) != S_OK) return FALSE;

    if (wcscmp(mimeType, L"image/jpeg") == 0) {
        ZeroMemory(&params, sizeof(params));
        params.Count = 1;
        params.Parameter[0].Guid = g_encoderQualityGuid;
        params.Parameter[0].NumberOfValues = 1;
        params.Parameter[0].Type = 4; /* EncoderParameterValueTypeLong */
        params.Parameter[0].Value = &qualityValue;
        paramsPtr = &params;
    }

    if (GdipSaveImageToStream(image, stream, &encoderClsid, paramsPtr) != 0) {
        IStream_Release(stream);
        return FALSE;
    }

    ZeroMemory(&stat, sizeof(stat));
    if (IStream_Stat(stream, &stat, STATFLAG_NONAME) != S_OK ||
        stat.cbSize.QuadPart == 0 || stat.cbSize.QuadPart > 0xffffffffULL) {
        IStream_Release(stream);
        return FALSE;
    }

    data = (BYTE *)malloc((size_t)stat.cbSize.QuadPart);
    if (!data) {
        IStream_Release(stream);
        return FALSE;
    }

    zero.QuadPart = 0;
    if (IStream_Seek(stream, zero, STREAM_SEEK_SET, NULL) != S_OK ||
        IStream_Read(stream, data, (ULONG)stat.cbSize.QuadPart, &bytesRead) != S_OK ||
        bytesRead != (ULONG)stat.cbSize.QuadPart) {
        free(data);
        IStream_Release(stream);
        return FALSE;
    }

    IStream_Release(stream);
    *outData = data;
    *outSize = bytesRead;
    return TRUE;
}

static BOOL GenerateImageToken(char token[IMAGE_TOKEN_HEX_LEN + 1])
{
    BYTE randomBytes[IMAGE_TOKEN_BYTES];
    HCRYPTPROV provider = 0;
    BOOL randomOk = FALSE;
    static const char hex[] = "0123456789abcdef";

    if (CryptAcquireContextA(&provider, NULL, NULL, PROV_RSA_FULL,
                             CRYPT_VERIFYCONTEXT | CRYPT_SILENT)) {
        randomOk = CryptGenRandom(provider, sizeof(randomBytes), randomBytes);
        CryptReleaseContext(provider, 0);
    }

    if (!randomOk) {
        GUID ids[2];
        if (CoCreateGuid(&ids[0]) != S_OK || CoCreateGuid(&ids[1]) != S_OK) {
            return FALSE;
        }
        memcpy(randomBytes, ids, sizeof(randomBytes));
    }

    for (int i = 0; i < IMAGE_TOKEN_BYTES; i++) {
        token[i * 2] = hex[randomBytes[i] >> 4];
        token[i * 2 + 1] = hex[randomBytes[i] & 0x0f];
    }
    token[IMAGE_TOKEN_HEX_LEN] = '\0';
    return TRUE;
}

/* Caller must hold g_imageLock exclusively. */
static void RememberGoneTokenLocked(const char *token)
{
    if (!token || !token[0]) return;
    if (g_goneTokenCount == g_goneTokenCapacity) {
        size_t newCapacity = g_goneTokenCapacity ? g_goneTokenCapacity * 2 : 64;
        void *expanded = realloc(g_goneTokens,
                                 newCapacity * sizeof(*g_goneTokens));
        if (!expanded) return;
        g_goneTokens = expanded;
        g_goneTokenCapacity = newCapacity;
    }
    strncpy(g_goneTokens[g_goneTokenCount], token, IMAGE_TOKEN_HEX_LEN + 1);
    g_goneTokenCount++;
}

/* ── Disk image storage (%LOCALAPPDATA%\ImagePaster) ───────────────────── */

static BOOL BuildImageStorageDirectoryPath(WCHAR directory[MAX_PATH])
{
    WCHAR base[MAX_PATH];
    int written;
    if (FAILED(SHGetFolderPathW(NULL, CSIDL_LOCAL_APPDATA, NULL,
                                SHGFP_TYPE_CURRENT, base))) {
        return FALSE;
    }
    written = swprintf(directory, MAX_PATH, L"%s\\%s", base,
                       IMAGE_STORAGE_DIR_NAME);
    return written > 0 && written < MAX_PATH;
}

static BOOL EnsureImageStorageDirectory(WCHAR directory[MAX_PATH])
{
    if (!BuildImageStorageDirectoryPath(directory)) {
        LogMessage("WARNING: Could not resolve %%LOCALAPPDATA%% for disk image storage");
        return FALSE;
    }
    if (!CreateDirectoryW(directory, NULL) &&
        GetLastError() != ERROR_ALREADY_EXISTS) {
        LogMessage("WARNING: Could not create disk image storage directory (error %lu)",
                   GetLastError());
        return FALSE;
    }
    return TRUE;
}

static BOOL BuildImageStorageFilePath(const WCHAR *directory,
                                      const WCHAR *fileName,
                                      WCHAR path[MAX_PATH])
{
    int written;
    if (!directory || !directory[0] || !fileName || !fileName[0]) {
        return FALSE;
    }
    written = swprintf(path, MAX_PATH, L"%s\\%s", directory, fileName);
    return written > 0 && written < MAX_PATH;
}

static BOOL IsValidImageToken(const char *token)
{
    if (!token || strlen(token) != IMAGE_TOKEN_HEX_LEN) return FALSE;
    for (int i = 0; i < IMAGE_TOKEN_HEX_LEN; i++) {
        char c = token[i];
        if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f'))) {
            return FALSE;
        }
    }
    return TRUE;
}

static BOOL ParseDiskImageFileName(
    const WCHAR *fileName, char token[IMAGE_TOKEN_HEX_LEN + 1])
{
    size_t length;
    if (!fileName || !token) return FALSE;
    length = wcslen(fileName);
    if (length != IMAGE_TOKEN_HEX_LEN + 4 ||
        _wcsicmp(fileName + IMAGE_TOKEN_HEX_LEN, L".jpg") != 0) {
        return FALSE;
    }
    for (int i = 0; i < IMAGE_TOKEN_HEX_LEN; i++) {
        WCHAR c = fileName[i];
        if (c >= L'A' && c <= L'F') c = c - L'A' + L'a';
        if (!((c >= L'0' && c <= L'9') || (c >= L'a' && c <= L'f'))) {
            return FALSE;
        }
        token[i] = (char)c;
    }
    token[IMAGE_TOKEN_HEX_LEN] = '\0';
    return TRUE;
}

static ULONGLONG FileTimeToUnixTimeMs(const FILETIME *fileTime)
{
    ULARGE_INTEGER value;
    if (!fileTime) return 0;
    value.LowPart = fileTime->dwLowDateTime;
    value.HighPart = fileTime->dwHighDateTime;
    if (value.QuadPart < 116444736000000000ULL) return 0;
    return (value.QuadPart - 116444736000000000ULL) / 10000ULL;
}

static BOOL GetDiskImageMetadata(const WCHAR *path, DWORD *jpegSize,
                                 UINT *width, UINT *height,
                                 ULONGLONG *capturedAtMs)
{
    WIN32_FILE_ATTRIBUTE_DATA attributes;
    GpImage *image = NULL;
    UINT imageWidth = 0;
    UINT imageHeight = 0;
    const FILETIME *captureTime;

    if (!path || !jpegSize || !width || !height || !capturedAtMs ||
        !GetFileAttributesExW(path, GetFileExInfoStandard, &attributes) ||
        (attributes.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ||
        attributes.nFileSizeHigh != 0 || attributes.nFileSizeLow == 0) {
        return FALSE;
    }
    if (GdipLoadImageFromFile(path, &image) != 0 || !image ||
        GdipGetImageWidth(image, &imageWidth) != 0 ||
        GdipGetImageHeight(image, &imageHeight) != 0 ||
        imageWidth == 0 || imageHeight == 0) {
        if (image) GdipDisposeImage(image);
        return FALSE;
    }
    GdipDisposeImage(image);

    captureTime = (attributes.ftCreationTime.dwLowDateTime ||
                   attributes.ftCreationTime.dwHighDateTime)
        ? &attributes.ftCreationTime : &attributes.ftLastWriteTime;
    *jpegSize = attributes.nFileSizeLow;
    *width = imageWidth;
    *height = imageHeight;
    *capturedAtMs = FileTimeToUnixTimeMs(captureTime);
    return TRUE;
}

static BOOL ReadFileFully(HANDLE file, void *buffer, DWORD bytes)
{
    BYTE *next = (BYTE *)buffer;
    DWORD total = 0;
    while (total < bytes) {
        DWORD readNow = 0;
        if (!ReadFile(file, next + total, bytes - total, &readNow, NULL) ||
            readNow == 0) {
            return FALSE;
        }
        total += readNow;
    }
    return TRUE;
}

static BOOL WriteFileFully(HANDLE file, const void *buffer, DWORD bytes)
{
    const BYTE *next = (const BYTE *)buffer;
    DWORD total = 0;
    while (total < bytes) {
        DWORD writtenNow = 0;
        if (!WriteFile(file, next + total, bytes - total,
                       &writtenNow, NULL) || writtenNow == 0) {
            return FALSE;
        }
        total += writtenNow;
    }
    return TRUE;
}

/* Writes a cached JPEG to disk storage. Returns the malloc'd file path, or
   NULL after logging; the caller then keeps the image in memory instead. */
static WCHAR *StoreImageToDisk(const char *token, const BYTE *jpegData,
                               DWORD jpegSize)
{
    WCHAR directory[MAX_PATH];
    WCHAR *path;
    size_t pathChars;
    HANDLE file;
    DWORD written = 0;
    BOOL wrote;

    if (!token || !token[0] || !jpegData || jpegSize == 0) return NULL;
    if (!EnsureImageStorageDirectory(directory)) return NULL;

    pathChars = wcslen(directory) + 1 + IMAGE_TOKEN_HEX_LEN + 4 + 1;
    path = (WCHAR *)malloc(pathChars * sizeof(WCHAR));
    if (!path) return NULL;
    if (swprintf(path, pathChars, L"%s\\%hs.jpg", directory, token) <= 0) {
        free(path);
        return NULL;
    }

    file = CreateFileW(path, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
                       FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        LogMessage("WARNING: Could not create image file in disk storage (error %lu); keeping image in memory",
                   GetLastError());
        free(path);
        return NULL;
    }
    wrote = WriteFile(file, jpegData, jpegSize, &written, NULL) &&
            written == jpegSize;
    CloseHandle(file);
    if (!wrote) {
        LogMessage("WARNING: Could not write image file to disk storage (error %lu); keeping image in memory",
                   GetLastError());
        DeleteFileW(path);
        free(path);
        return NULL;
    }
    return path;
}

/* Frees an image's resources. In-session removal (eviction, delete, clear)
   deletes a disk-stored file; app exit keeps the files so disk storage
   doubles as a lasting record. */
static void FreeCachedImageDataEx(CachedImage *image, BOOL deleteFile)
{
    if (image->diskPath) {
        if (deleteFile) DeleteFileW(image->diskPath);
        free(image->diskPath);
    }
    free(image->jpegData);
    free(image->base64Data);
    ZeroMemory(image, sizeof(*image));
}

static void FreeCachedImageData(CachedImage *image)
{
    FreeCachedImageDataEx(image, TRUE);
}

/* Returns a malloc'd copy of a disk-stored JPEG. Sharing delete access lets
   a reader outside g_imageLock never make an eviction's DeleteFileW fail;
   the file then disappears once this handle closes. Must not log. */
static BYTE *ReadRetainedImageFile(const WCHAR *path, DWORD *size)
{
    HANDLE file = CreateFileW(path, GENERIC_READ,
                              FILE_SHARE_READ | FILE_SHARE_DELETE, NULL,
                              OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    LARGE_INTEGER fileSize;
    BYTE *data;
    DWORD readTotal = 0;

    *size = 0;
    if (file == INVALID_HANDLE_VALUE) return NULL;
    if (!GetFileSizeEx(file, &fileSize) || fileSize.QuadPart <= 0 ||
        fileSize.QuadPart > 0x7fffffffLL) {
        CloseHandle(file);
        return NULL;
    }
    data = (BYTE *)malloc((size_t)fileSize.QuadPart);
    if (!data) {
        CloseHandle(file);
        return NULL;
    }
    while (readTotal < (DWORD)fileSize.QuadPart) {
        DWORD readNow = 0;
        if (!ReadFile(file, data + readTotal,
                      (DWORD)fileSize.QuadPart - readTotal,
                      &readNow, NULL) || readNow == 0) {
            break;
        }
        readTotal += readNow;
    }
    CloseHandle(file);
    if (readTotal != (DWORD)fileSize.QuadPart) {
        free(data);
        return NULL;
    }
    *size = readTotal;
    return data;
}

/* Returns a malloc'd copy of the image's JPEG bytes from memory or disk.
   Caller must hold g_imageLock (shared is enough); safe on the HTTP worker
   thread, so it must not call LogMessage. */
static BYTE *LoadCachedImageBytesLocked(const CachedImage *image, DWORD *size)
{
    *size = 0;
    if (image->jpegData && image->jpegSize > 0) {
        BYTE *copy = (BYTE *)malloc(image->jpegSize);
        if (!copy) return NULL;
        memcpy(copy, image->jpegData, image->jpegSize);
        *size = image->jpegSize;
        return copy;
    }
    if (image->diskPath) return ReadRetainedImageFile(image->diskPath, size);
    return NULL;
}

/* Caller must hold g_imageLock exclusively. A finite setting reserves one
   slot for the current image, so 1 intentionally disables history. */
static size_t HistoricalImageLimitLocked(void)
{
    if (g_configImageHistoryLimit == 0) return (size_t)-1;
    return (size_t)(g_configImageHistoryLimit - 1);
}

/* Caller must hold g_imageLock exclusively. Transfers image ownership into
   the oldest-to-newest history array without enforcing its configured limit. */
static BOOL AppendHistoricalImageLocked(CachedImage *image)
{
    size_t newCapacity;
    void *expanded;

    if (!image || !image->token[0] ||
        (!image->jpegData && !image->diskPath)) {
        return FALSE;
    }
    if (g_imageHistoryCount == g_imageHistoryCapacity) {
        newCapacity = g_imageHistoryCapacity
            ? g_imageHistoryCapacity * 2 : 8;
        if (newCapacity < g_imageHistoryCapacity ||
            newCapacity > (size_t)-1 / sizeof(*g_imageHistory)) {
            return FALSE;
        }
        expanded = realloc(g_imageHistory,
                           newCapacity * sizeof(*g_imageHistory));
        if (!expanded) return FALSE;
        g_imageHistory = expanded;
        g_imageHistoryCapacity = newCapacity;
    }
    g_imageHistory[g_imageHistoryCount++] = *image;
    ZeroMemory(image, sizeof(*image));
    return TRUE;
}

/* Caller must hold g_imageLock exclusively. */
static void EvictOldestHistoricalImageLockedEx(BOOL deleteFile)
{
    if (g_imageHistoryCount == 0) return;
    RememberGoneTokenLocked(g_imageHistory[0].token);
    FreeCachedImageDataEx(&g_imageHistory[0], deleteFile);
    g_imageHistoryCount--;
    if (g_imageHistoryCount > 0) {
        memmove(&g_imageHistory[0], &g_imageHistory[1],
                g_imageHistoryCount * sizeof(*g_imageHistory));
    }
    ZeroMemory(&g_imageHistory[g_imageHistoryCount], sizeof(*g_imageHistory));
}

static void EvictOldestHistoricalImageLocked(void)
{
    EvictOldestHistoricalImageLockedEx(TRUE);
}

/* Caller must hold g_imageLock exclusively. */
static size_t EnforceImageHistoryLimitLockedEx(BOOL deleteFiles)
{
    size_t evicted = 0;
    size_t limit = HistoricalImageLimitLocked();
    while (g_imageHistoryCount > limit) {
        EvictOldestHistoricalImageLockedEx(deleteFiles);
        evicted++;
    }
    return evicted;
}

static size_t EnforceImageHistoryLimitLocked(void)
{
    return EnforceImageHistoryLimitLockedEx(TRUE);
}

/* Caller must hold g_imageLock exclusively. Transfers the current JPEG into
   history without retaining its base64 representation. */
static void ArchiveCurrentImageLocked(void)
{
    CachedImage archived;

    if (!g_cachedImage.token[0]) {
        FreeCachedImageData(&g_cachedImage);
        return;
    }

    archived = g_cachedImage;
    archived.base64Data = NULL;
    archived.base64Len = 0;
    free(g_cachedImage.base64Data);
    ZeroMemory(&g_cachedImage, sizeof(g_cachedImage));

    if (HistoricalImageLimitLocked() == 0 ||
        (!archived.jpegData && !archived.diskPath)) {
        RememberGoneTokenLocked(archived.token);
        FreeCachedImageData(&archived);
        return;
    }

    if (!AppendHistoricalImageLocked(&archived)) {
        RememberGoneTokenLocked(archived.token);
        FreeCachedImageData(&archived);
        return;
    }
    EnforceImageHistoryLimitLocked();
}

static ULONGLONG GetUnixTimeMs(void)
{
    FILETIME fileTime;
    ULARGE_INTEGER value;
    GetSystemTimeAsFileTime(&fileTime);
    value.LowPart = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    return (value.QuadPart - 116444736000000000ULL) / 10000ULL;
}

static BOOL BuildTokenImagePath(
    const WCHAR *directory, const char *token, WCHAR path[MAX_PATH])
{
    WCHAR fileName[IMAGE_TOKEN_HEX_LEN + 5];
    if (!IsValidImageToken(token) ||
        swprintf(fileName, IMAGE_TOKEN_HEX_LEN + 5,
                 L"%hs.jpg", token) <= 0) {
        return FALSE;
    }
    return BuildImageStorageFilePath(directory, fileName, path);
}

static int CompareCachedImagesByCaptureTime(const void *left,
                                            const void *right)
{
    const CachedImage *a = (const CachedImage *)left;
    const CachedImage *b = (const CachedImage *)right;
    if (a->capturedAtMs < b->capturedAtMs) return -1;
    if (a->capturedAtMs > b->capturedAtMs) return 1;
    return strcmp(a->token, b->token);
}

static int CompareLoadedDiskImagesByCaptureTime(const void *left,
                                                const void *right)
{
    const LoadedDiskImage *a = (const LoadedDiskImage *)left;
    const LoadedDiskImage *b = (const LoadedDiskImage *)right;
    return CompareCachedImagesByCaptureTime(&a->image, &b->image);
}

static BOOL RetainedImageTokenExistsLocked(const char *token)
{
    if (g_cachedImage.token[0] &&
        strcmp(g_cachedImage.token, token) == 0) {
        return TRUE;
    }
    for (size_t i = 0; i < g_imageHistoryCount; i++) {
        if (strcmp(g_imageHistory[i].token, token) == 0) return TRUE;
    }
    return FALSE;
}

static BOOL LoadedDiskImageTokenExists(const LoadedDiskImage *images,
                                       size_t count, const char *token)
{
    for (size_t i = 0; i < count; i++) {
        if (strcmp(images[i].image.token, token) == 0) return TRUE;
    }
    return FALSE;
}

static BOOL AppendLoadedDiskImage(LoadedDiskImage **images, size_t *count,
                                  size_t *capacity,
                                  LoadedDiskImage *candidate)
{
    size_t newCapacity;
    void *expanded;
    if (*count == *capacity) {
        newCapacity = *capacity ? *capacity * 2 : 16;
        if (newCapacity < *capacity ||
            newCapacity > (size_t)-1 / sizeof(**images)) {
            return FALSE;
        }
        expanded = realloc(*images, newCapacity * sizeof(**images));
        if (!expanded) return FALSE;
        *images = (LoadedDiskImage *)expanded;
        *capacity = newCapacity;
    }
    (*images)[(*count)++] = *candidate;
    ZeroMemory(candidate, sizeof(*candidate));
    return TRUE;
}

static void FreeLoadedDiskImages(LoadedDiskImage *images, size_t count)
{
    if (!images) return;
    for (size_t i = 0; i < count; i++) {
        FreeCachedImageDataEx(&images[i].image, FALSE);
    }
    free(images);
}

static BOOL CreateLoadedDiskImage(const WCHAR *directory, const char *token,
                                  ULONGLONG indexedCaptureTime,
                                  BOOL isCurrent,
                                  LoadedDiskImage *loaded)
{
    WCHAR path[MAX_PATH];
    DWORD jpegSize = 0;
    UINT width = 0;
    UINT height = 0;
    ULONGLONG fileCaptureTime = 0;

    ZeroMemory(loaded, sizeof(*loaded));
    if (!BuildTokenImagePath(directory, token, path) ||
        !GetDiskImageMetadata(path, &jpegSize, &width, &height,
                              &fileCaptureTime)) {
        return FALSE;
    }
    loaded->image.diskPath = _wcsdup(path);
    if (!loaded->image.diskPath) return FALSE;
    strncpy(loaded->image.token, token, sizeof(loaded->image.token) - 1);
    loaded->image.jpegSize = jpegSize;
    loaded->image.width = width;
    loaded->image.height = height;
    loaded->image.capturedAtMs = indexedCaptureTime
        ? indexedCaptureTime : fileCaptureTime;
    loaded->isCurrent = isCurrent;
    return TRUE;
}

static BOOL ReadDiskHistoryIndex(const WCHAR *directory,
                                 LoadedDiskImage **images,
                                 size_t *count, BOOL *indexFound)
{
    WCHAR indexPath[MAX_PATH];
    HANDLE file = INVALID_HANDLE_VALUE;
    LARGE_INTEGER fileSize;
    DiskHistoryIndexHeader header;
    size_t capacity = 0;
    BOOL valid = FALSE;

    *images = NULL;
    *count = 0;
    *indexFound = FALSE;
    if (!BuildImageStorageFilePath(directory, IMAGE_HISTORY_INDEX_NAME,
                                   indexPath)) {
        return FALSE;
    }
    file = CreateFileW(indexPath, GENERIC_READ, FILE_SHARE_READ, NULL,
                       OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) return FALSE;
    *indexFound = TRUE;

    if (!GetFileSizeEx(file, &fileSize) ||
        fileSize.QuadPart < (LONGLONG)sizeof(header) ||
        !ReadFileFully(file, &header, sizeof(header)) ||
        memcmp(header.magic, IMAGE_HISTORY_INDEX_MAGIC,
               sizeof(header.magic)) != 0 ||
        header.recordCount > MAX_DISK_HISTORY_INDEX_RECORDS ||
        fileSize.QuadPart !=
            (LONGLONG)sizeof(header) +
            (LONGLONG)header.recordCount *
                (LONGLONG)sizeof(DiskHistoryIndexRecord)) {
        goto cleanup;
    }

    for (DWORD i = 0; i < header.recordCount; i++) {
        DiskHistoryIndexRecord record;
        LoadedDiskImage loaded;
        if (!ReadFileFully(file, &record, sizeof(record))) goto cleanup;
        if (record.token[IMAGE_TOKEN_HEX_LEN] != '\0' ||
            !IsValidImageToken(record.token) || record.isCurrent > 1 ||
            LoadedDiskImageTokenExists(*images, *count, record.token)) {
            continue;
        }
        if (!CreateLoadedDiskImage(directory, record.token,
                                   record.capturedAtMs,
                                   record.isCurrent != 0, &loaded)) {
            continue;
        }
        if (!AppendLoadedDiskImage(images, count, &capacity, &loaded)) {
            FreeCachedImageDataEx(&loaded.image, FALSE);
            goto cleanup;
        }
    }
    valid = TRUE;

cleanup:
    CloseHandle(file);
    if (!valid) {
        FreeLoadedDiskImages(*images, *count);
        *images = NULL;
        *count = 0;
    }
    return valid;
}

static BOOL ScanLegacyDiskHistory(const WCHAR *directory,
                                  LoadedDiskImage **images, size_t *count)
{
    WCHAR pattern[MAX_PATH];
    WIN32_FIND_DATAW found;
    HANDLE search;
    size_t capacity = 0;

    *images = NULL;
    *count = 0;
    if (!BuildImageStorageFilePath(directory, L"*.jpg", pattern)) {
        return FALSE;
    }
    search = FindFirstFileW(pattern, &found);
    if (search == INVALID_HANDLE_VALUE) {
        DWORD errorCode = GetLastError();
        return errorCode == ERROR_FILE_NOT_FOUND ||
               errorCode == ERROR_PATH_NOT_FOUND;
    }
    do {
        char token[IMAGE_TOKEN_HEX_LEN + 1];
        LoadedDiskImage loaded;
        if ((found.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ||
            !ParseDiskImageFileName(found.cFileName, token) ||
            LoadedDiskImageTokenExists(*images, *count, token) ||
            !CreateLoadedDiskImage(directory, token, 0, FALSE, &loaded)) {
            continue;
        }
        if (!AppendLoadedDiskImage(images, count, &capacity, &loaded)) {
            FreeCachedImageDataEx(&loaded.image, FALSE);
            FindClose(search);
            FreeLoadedDiskImages(*images, *count);
            *images = NULL;
            *count = 0;
            return FALSE;
        }
    } while (FindNextFileW(search, &found));
    FindClose(search);

    if (*count > 1) {
        qsort(*images, *count, sizeof(**images),
              CompareLoadedDiskImagesByCaptureTime);
    }
    if (*count > 0) (*images)[*count - 1].isCurrent = TRUE;
    return TRUE;
}

static void FillDiskHistoryRecord(DiskHistoryIndexRecord *record,
                                  const CachedImage *image, BOOL isCurrent)
{
    ZeroMemory(record, sizeof(*record));
    strncpy(record->token, image->token, sizeof(record->token) - 1);
    record->width = image->width;
    record->height = image->height;
    record->jpegSize = image->jpegSize;
    record->capturedAtMs = image->capturedAtMs;
    record->isCurrent = isCurrent ? 1 : 0;
}

/* A stale index is worse than no index: without one, the next launch safely
   reconstructs the catalog from the token-named JPEG files. */
static void InvalidateDiskHistoryIndex(const WCHAR *indexPath)
{
    if (!DeleteFileW(indexPath) && GetLastError() != ERROR_FILE_NOT_FOUND) {
        LogMessage("WARNING: Could not invalidate the stale disk image history index (error %lu)",
                   GetLastError());
    }
}

static void PersistDiskHistoryIndex(void)
{
    WCHAR directory[MAX_PATH];
    WCHAR indexPath[MAX_PATH];
    WCHAR tempPath[MAX_PATH];
    DiskHistoryIndexHeader header;
    DiskHistoryIndexRecord *records = NULL;
    size_t recordCount = 0;
    size_t position = 0;
    HANDLE file;
    BOOL wrote;
    DWORD errorCode;

    if (!g_diskHistoryCatalogLoaded) return;

    AcquireSRWLockShared(&g_imageLock);
    for (size_t i = 0; i < g_imageHistoryCount; i++) {
        if (g_imageHistory[i].diskPath &&
            IsValidImageToken(g_imageHistory[i].token)) {
            recordCount++;
        }
    }
    if (g_cachedImage.diskPath && IsValidImageToken(g_cachedImage.token)) {
        recordCount++;
    }
    if (recordCount > MAX_DISK_HISTORY_INDEX_RECORDS) {
        ReleaseSRWLockShared(&g_imageLock);
        LogMessage("WARNING: Disk image history is too large to persist its index");
        return;
    }
    if (recordCount > 0) {
        records = (DiskHistoryIndexRecord *)calloc(
            recordCount, sizeof(*records));
        if (!records) {
            ReleaseSRWLockShared(&g_imageLock);
            LogMessage("WARNING: Could not allocate the disk image history index");
            return;
        }
    }
    for (size_t i = 0; i < g_imageHistoryCount; i++) {
        if (!g_imageHistory[i].diskPath ||
            !IsValidImageToken(g_imageHistory[i].token)) {
            continue;
        }
        FillDiskHistoryRecord(&records[position++], &g_imageHistory[i], FALSE);
    }
    if (g_cachedImage.diskPath && IsValidImageToken(g_cachedImage.token)) {
        FillDiskHistoryRecord(&records[position++], &g_cachedImage, TRUE);
    }
    ReleaseSRWLockShared(&g_imageLock);

    if (!EnsureImageStorageDirectory(directory) ||
        !BuildImageStorageFilePath(directory, IMAGE_HISTORY_INDEX_NAME,
                                   indexPath) ||
        !BuildImageStorageFilePath(directory, IMAGE_HISTORY_INDEX_TEMP_NAME,
                                   tempPath)) {
        free(records);
        return;
    }

    file = CreateFileW(tempPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
                       FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        LogMessage("WARNING: Could not create the disk image history index (error %lu)",
                   GetLastError());
        InvalidateDiskHistoryIndex(indexPath);
        free(records);
        return;
    }
    ZeroMemory(&header, sizeof(header));
    memcpy(header.magic, IMAGE_HISTORY_INDEX_MAGIC, sizeof(header.magic));
    header.recordCount = (DWORD)recordCount;
    wrote = WriteFileFully(file, &header, sizeof(header)) &&
            (recordCount == 0 ||
             WriteFileFully(file, records,
                            (DWORD)(recordCount * sizeof(*records)))) &&
            FlushFileBuffers(file);
    errorCode = wrote ? ERROR_SUCCESS : GetLastError();
    CloseHandle(file);
    free(records);

    if (!wrote ||
        !MoveFileExW(tempPath, indexPath,
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        if (wrote) errorCode = GetLastError();
        DeleteFileW(tempPath);
        InvalidateDiskHistoryIndex(indexPath);
        LogMessage("WARNING: Could not update the disk image history index (error %lu)",
                   errorCode);
    }
}

static size_t LoadPersistedDiskHistory(void)
{
    WCHAR directory[MAX_PATH];
    LoadedDiskImage *loaded = NULL;
    size_t loadedCount = 0;
    size_t restoredCount = 0;
    size_t evicted = 0;
    size_t currentIndex = (size_t)-1;
    BOOL indexFound = FALSE;
    BOOL indexValid;
    BOOL recoveredLegacy = FALSE;

    if (g_diskHistoryCatalogLoaded ||
        g_configImageStorage != IMAGE_STORAGE_DISK) {
        return 0;
    }
    g_diskHistoryCatalogLoaded = TRUE;
    if (!BuildImageStorageDirectoryPath(directory)) {
        LogMessage("WARNING: Could not resolve disk image history location");
        return 0;
    }

    indexValid = ReadDiskHistoryIndex(directory, &loaded, &loadedCount,
                                      &indexFound);
    if (!indexValid) {
        if (indexFound) {
            LogMessage("WARNING: Disk image history index is invalid; rebuilding it from JPEG files");
        }
        if (!ScanLegacyDiskHistory(directory, &loaded, &loadedCount)) {
            LogMessage("WARNING: Could not scan the disk image history folder");
            return 0;
        }
        recoveredLegacy = loadedCount > 0;
    }

    for (size_t i = 0; i < loadedCount; i++) {
        if (loaded[i].isCurrent) currentIndex = i;
    }
    if (loadedCount > 0 && currentIndex == (size_t)-1) {
        /* If a formerly-current JPEG vanished, promote the newest surviving
           indexed entry so the total-history limit still behaves normally. */
        currentIndex = loadedCount - 1;
    }

    AcquireSRWLockExclusive(&g_imageLock);
    for (size_t i = 0; i < loadedCount; i++) {
        if (RetainedImageTokenExistsLocked(loaded[i].image.token)) continue;
        if (i == currentIndex && !g_cachedImage.token[0]) {
            g_cachedImage = loaded[i].image;
            ZeroMemory(&loaded[i].image, sizeof(loaded[i].image));
            g_restoredDiskCurrentNeedsClipboardCheck = TRUE;
            restoredCount++;
        } else if (AppendHistoricalImageLocked(&loaded[i].image)) {
            restoredCount++;
        }
    }
    if (g_imageHistoryCount > 1) {
        qsort(g_imageHistory, g_imageHistoryCount,
              sizeof(*g_imageHistory), CompareCachedImagesByCaptureTime);
    }
    /* A legacy scan can discover files that earlier releases deliberately
       left as a lasting record. Keep excess migrated files on disk even when
       the configured live-history limit omits them from the rebuilt index. */
    evicted = EnforceImageHistoryLimitLockedEx(!recoveredLegacy);
    ReleaseSRWLockExclusive(&g_imageLock);
    FreeLoadedDiskImages(loaded, loadedCount);

    PersistDiskHistoryIndex();
    if (restoredCount > 0) {
        LogMessage("Disk image history restored: %llu image(s)%s",
                   (unsigned long long)restoredCount,
                   recoveredLegacy ? " (migrated existing JPEG files)" : "");
    }
    if (evicted > 0) {
        LogMessage("Disk image history limit evicted %llu restored image(s)",
                   (unsigned long long)evicted);
    }
    NotifyHistoryViewChanged();
    return restoredCount;
}

/* When the startup clipboard still contains the last persisted current image,
   reuse its token and file instead of adding a duplicate on every restart. */
static BOOL AttachClipboardDataToRestoredCurrent(
    const BYTE *jpegData, DWORD jpegSize, char *base64Data, DWORD base64Len,
    UINT width, UINT height, char token[IMAGE_TOKEN_HEX_LEN + 1])
{
    BYTE *storedBytes = NULL;
    DWORD storedSize = 0;
    BOOL matched = FALSE;

    if (!g_restoredDiskCurrentNeedsClipboardCheck) return FALSE;
    AcquireSRWLockExclusive(&g_imageLock);
    g_restoredDiskCurrentNeedsClipboardCheck = FALSE;
    if (g_cachedImage.diskPath && !g_cachedImage.base64Data &&
        g_cachedImage.jpegSize == jpegSize) {
        storedBytes = LoadCachedImageBytesLocked(&g_cachedImage, &storedSize);
        if (storedBytes && storedSize == jpegSize &&
            memcmp(storedBytes, jpegData, jpegSize) == 0) {
            g_cachedImage.base64Data = base64Data;
            g_cachedImage.base64Len = base64Len;
            g_cachedImage.width = width;
            g_cachedImage.height = height;
            strncpy(token, g_cachedImage.token, IMAGE_TOKEN_HEX_LEN + 1);
            matched = TRUE;
        }
    }
    ReleaseSRWLockExclusive(&g_imageLock);
    free(storedBytes);
    if (matched) NotifyHistoryViewChanged();
    return matched;
}

/* Takes ownership of jpegData (may be NULL when diskPath is set), base64Data,
   and diskPath. */
static void ReplaceCurrentImage(BYTE *jpegData, DWORD jpegSize,
                                char *base64Data, DWORD base64Len,
                                const char *token, UINT width, UINT height,
                                WCHAR *diskPath)
{
    AcquireSRWLockExclusive(&g_imageLock);
    ArchiveCurrentImageLocked();
    g_cachedImage.jpegData = jpegData;
    g_cachedImage.jpegSize = jpegSize;
    g_cachedImage.base64Data = base64Data;
    g_cachedImage.base64Len = base64Len;
    strncpy(g_cachedImage.token, token, sizeof(g_cachedImage.token) - 1);
    g_cachedImage.width = width;
    g_cachedImage.height = height;
    g_cachedImage.capturedAtMs = GetUnixTimeMs();
    g_cachedImage.diskPath = diskPath;
    ReleaseSRWLockExclusive(&g_imageLock);
    PersistDiskHistoryIndex();
    NotifyHistoryViewChanged();
}

static void ClearCurrentImage(const char *reason)
{
    BOOL hadImage;
    size_t retainedCount;

    AcquireSRWLockExclusive(&g_imageLock);
    hadImage = g_cachedImage.token[0] != '\0';
    ArchiveCurrentImageLocked();
    retainedCount = g_imageHistoryCount;
    ReleaseSRWLockExclusive(&g_imageLock);

    if (hadImage) {
        PersistDiskHistoryIndex();
        LogMessage("Current clipboard image cleared: %s (%llu historical retained)",
                   reason, (unsigned long long)retainedCount);
        NotifyHistoryViewChanged();
    }
}

static size_t SetImageHistoryLimit(int limit)
{
    size_t evicted;
    AcquireSRWLockExclusive(&g_imageLock);
    g_configImageHistoryLimit = limit;
    evicted = EnforceImageHistoryLimitLocked();
    ReleaseSRWLockExclusive(&g_imageLock);
    if (evicted > 0) {
        PersistDiskHistoryIndex();
        NotifyHistoryViewChanged();
    }
    return evicted;
}

static void DestroyImageCache(void)
{
    PersistDiskHistoryIndex();
    AcquireSRWLockExclusive(&g_imageLock);
    /* Exit keeps indexed disk-stored files on disk for the next session. */
    FreeCachedImageDataEx(&g_cachedImage, FALSE);
    for (size_t i = 0; i < g_imageHistoryCount; i++) {
        FreeCachedImageDataEx(&g_imageHistory[i], FALSE);
    }
    free(g_imageHistory);
    g_imageHistory = NULL;
    g_imageHistoryCount = 0;
    g_imageHistoryCapacity = 0;
    free(g_goneTokens);
    g_goneTokens = NULL;
    g_goneTokenCount = 0;
    g_goneTokenCapacity = 0;
    ReleaseSRWLockExclusive(&g_imageLock);
}

static DWORD GetCachedClipboardSequence(void)
{
    return (DWORD)InterlockedCompareExchange(&g_lastClipboardSequence, 0, 0);
}

static void SetCachedClipboardSequence(DWORD sequence)
{
    InterlockedExchange(&g_lastClipboardSequence, (LONG)sequence);
}

/* The hook must never wait for image encoding, disk I/O, or HTTP readers. */
static BOOL TryHasCachedImage(void)
{
    BOOL available;
    if (!TryAcquireSRWLockShared(&g_imageLock)) return FALSE;
    available = g_cachedImage.token[0] != '\0' &&
                (g_cachedImage.jpegData != NULL ||
                 g_cachedImage.diskPath != NULL) &&
                g_cachedImage.base64Data != NULL;
    ReleaseSRWLockShared(&g_imageLock);
    return available;
}

static BOOL AnyRetainedImages(void)
{
    BOOL any;
    AcquireSRWLockShared(&g_imageLock);
    any = (g_cachedImage.token[0] != '\0' &&
           (g_cachedImage.jpegData != NULL ||
            g_cachedImage.diskPath != NULL)) ||
          g_imageHistoryCount > 0;
    ReleaseSRWLockShared(&g_imageLock);
    return any;
}

/* Copies the JPEG bytes of a retained image (current or historical) so they
   can be used outside g_imageLock. The caller frees *jpegCopy. */
static BOOL CopyRetainedImageByToken(const char *token, BYTE **jpegCopy,
                                     DWORD *jpegSize, ULONGLONG *capturedAtMs)
{
    const CachedImage *found = NULL;
    BOOL success = FALSE;

    *jpegCopy = NULL;
    *jpegSize = 0;
    if (capturedAtMs) *capturedAtMs = 0;
    if (!token || !token[0]) return FALSE;

    AcquireSRWLockShared(&g_imageLock);
    if (g_cachedImage.token[0] && strcmp(g_cachedImage.token, token) == 0) {
        found = &g_cachedImage;
    } else {
        for (size_t i = 0; i < g_imageHistoryCount; i++) {
            if (strcmp(g_imageHistory[i].token, token) == 0) {
                found = &g_imageHistory[i];
                break;
            }
        }
    }
    if (found) {
        *jpegCopy = LoadCachedImageBytesLocked(found, jpegSize);
        if (*jpegCopy) {
            if (capturedAtMs) *capturedAtMs = found->capturedAtMs;
            success = TRUE;
        }
    }
    ReleaseSRWLockShared(&g_imageLock);
    return success;
}

/* Returns a malloc'd copy of the disk path for a retained image, or NULL if
   the image is gone or stored in memory. */
static WCHAR *GetRetainedImageDiskPathCopy(const char *token)
{
    WCHAR *path = NULL;
    const CachedImage *found = NULL;

    if (!token || !token[0]) return NULL;
    AcquireSRWLockShared(&g_imageLock);
    if (g_cachedImage.token[0] && strcmp(g_cachedImage.token, token) == 0) {
        found = &g_cachedImage;
    } else {
        for (size_t i = 0; i < g_imageHistoryCount; i++) {
            if (strcmp(g_imageHistory[i].token, token) == 0) {
                found = &g_imageHistory[i];
                break;
            }
        }
    }
    if (found && found->diskPath) path = _wcsdup(found->diskPath);
    ReleaseSRWLockShared(&g_imageLock);
    return path;
}

/* Removes a retained image; its URL starts answering 410 Gone. Deleting the
   current image intentionally leaves nothing to paste until the next copy. */
static BOOL DeleteRetainedImageByToken(const char *token)
{
    BOOL removed = FALSE;

    if (!token || !token[0]) return FALSE;
    AcquireSRWLockExclusive(&g_imageLock);
    if (g_cachedImage.token[0] && strcmp(g_cachedImage.token, token) == 0) {
        RememberGoneTokenLocked(g_cachedImage.token);
        FreeCachedImageData(&g_cachedImage);
        removed = TRUE;
    } else {
        for (size_t i = 0; i < g_imageHistoryCount; i++) {
            if (strcmp(g_imageHistory[i].token, token) != 0) continue;
            RememberGoneTokenLocked(g_imageHistory[i].token);
            FreeCachedImageData(&g_imageHistory[i]);
            g_imageHistoryCount--;
            if (i < g_imageHistoryCount) {
                memmove(&g_imageHistory[i], &g_imageHistory[i + 1],
                        (g_imageHistoryCount - i) * sizeof(*g_imageHistory));
            }
            ZeroMemory(&g_imageHistory[g_imageHistoryCount],
                       sizeof(*g_imageHistory));
            removed = TRUE;
            break;
        }
    }
    ReleaseSRWLockExclusive(&g_imageLock);
    if (removed) PersistDiskHistoryIndex();
    return removed;
}

static size_t ClearAllRetainedImages(void)
{
    size_t removed = 0;

    AcquireSRWLockExclusive(&g_imageLock);
    if (g_cachedImage.token[0]) {
        RememberGoneTokenLocked(g_cachedImage.token);
        FreeCachedImageData(&g_cachedImage);
        removed++;
    }
    for (size_t i = 0; i < g_imageHistoryCount; i++) {
        RememberGoneTokenLocked(g_imageHistory[i].token);
        FreeCachedImageData(&g_imageHistory[i]);
        removed++;
    }
    g_imageHistoryCount = 0;
    ReleaseSRWLockExclusive(&g_imageLock);
    PersistDiskHistoryIndex();
    return removed;
}

static BOOL RefreshClipboardImageCache(void)
{
    DWORD sequence = GetClipboardSequenceNumber();
    HANDLE hDib = NULL;
    BITMAPINFOHEADER *pBih = NULL;
    BYTE *pBits = NULL;
    GpBitmap *bitmap = NULL;
    BYTE *pngData = NULL;
    BYTE *jpegData = NULL;
    DWORD pngSize = 0;
    DWORD jpegSize = 0;
    char *base64Data = NULL;
    DWORD base64Len = 0;
    char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
    UINT width = 0, height = 0;
    BOOL success = FALSE;

    if (!IsClipboardFormatAvailable(CF_DIB)) {
        g_restoredDiskCurrentNeedsClipboardCheck = FALSE;
        SetCachedClipboardSequence(sequence);
        KillTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY);
        ClearCurrentImage("clipboard now contains non-image data");
        return TRUE;
    }

    if (!OpenClipboard(g_hWndMain)) {
        LogMessage("Clipboard image is temporarily unavailable; retrying");
        SetTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY, CLIPBOARD_RETRY_DELAY_MS, NULL);
        return FALSE;
    }

    hDib = GetClipboardData(CF_DIB);
    if (!hDib || GlobalSize(hDib) < sizeof(BITMAPINFOHEADER)) goto cleanup_clipboard;
    pBih = (BITMAPINFOHEADER *)GlobalLock(hDib);
    if (!pBih || pBih->biSize < sizeof(BITMAPINFOHEADER)) goto cleanup_clipboard;

    {
        SIZE_T dibSize = GlobalSize(hDib);
        SIZE_T colorTableSize = 0;
        SIZE_T pixelOffset;
        if (pBih->biBitCount <= 8) {
            DWORD colorCount = pBih->biClrUsed ? pBih->biClrUsed : (1u << pBih->biBitCount);
            colorTableSize = (SIZE_T)colorCount * sizeof(RGBQUAD);
        } else if (pBih->biSize == sizeof(BITMAPINFOHEADER) &&
                   pBih->biCompression == BI_BITFIELDS) {
            colorTableSize = 3 * sizeof(DWORD);
        }
        pixelOffset = (SIZE_T)pBih->biSize + colorTableSize;
        if (pixelOffset >= dibSize) goto cleanup_clipboard;
        pBits = (BYTE *)pBih + pixelOffset;
    }

    if (GdipCreateBitmapFromGdiDib((const BITMAPINFO *)pBih, pBits, &bitmap) != 0) {
        goto cleanup_clipboard;
    }
    GdipGetImageWidth((GpImage *)bitmap, &width);
    GdipGetImageHeight((GpImage *)bitmap, &height);

cleanup_clipboard:
    if (pBih) GlobalUnlock(hDib);
    CloseClipboard();

    if (!bitmap) {
        LogMessage("ERROR: Failed to read clipboard image");
        SetTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY, CLIPBOARD_RETRY_DELAY_MS, NULL);
        return FALSE;
    }

    if (!EncodeImageToMemory((GpImage *)bitmap, L"image/png", 0, &pngData, &pngSize)) {
        LogMessage("ERROR: Failed to encode clipboard image as PNG");
        goto cleanup;
    }
    if (!EncodeImageToMemory((GpImage *)bitmap, L"image/jpeg", g_configJpegQuality,
                             &jpegData, &jpegSize)) {
        LogMessage("ERROR: Failed to encode clipboard image as JPEG");
        goto cleanup;
    }
    base64Data = Base64Encode(pngData, pngSize, &base64Len);
    if (!base64Data) {
        LogMessage("ERROR: Failed to prepare clipboard image cache");
        goto cleanup;
    }

    /* Do not publish stale data if the clipboard changed during compression. */
    if (GetClipboardSequenceNumber() != sequence) {
        LogMessage("Clipboard changed during image encoding; retrying latest contents");
        SetTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY, CLIPBOARD_RETRY_DELAY_MS, NULL);
        goto cleanup;
    }

    if (AttachClipboardDataToRestoredCurrent(
            jpegData, jpegSize, base64Data, base64Len,
            width, height, token)) {
        free(jpegData);
        jpegData = NULL;
        base64Data = NULL;
        LogMessage("Clipboard image matched restored disk history: %ux%u, JPEG=%lu bytes, id=%.12s...",
                   width, height, jpegSize, token);
    } else {
        WCHAR *diskPath = NULL;
        if (!GenerateImageToken(token)) {
            LogMessage("ERROR: Failed to prepare clipboard image cache");
            goto cleanup;
        }
        if (g_configImageStorage == IMAGE_STORAGE_DISK) {
            diskPath = StoreImageToDisk(token, jpegData, jpegSize);
            if (diskPath) {
                free(jpegData);
                jpegData = NULL;
            }
        }
        ReplaceCurrentImage(jpegData, jpegSize, base64Data, base64Len,
                            token, width, height, diskPath);
        LogMessage("Clipboard image cached (%s): %ux%u, JPEG=%lu bytes, quality=%d%%, id=%.12s...",
                   diskPath ? "disk" : "memory", width, height, jpegSize,
                   g_configJpegQuality, token);
    }
    jpegData = NULL;
    base64Data = NULL;
    SetCachedClipboardSequence(sequence);
    KillTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY);
    success = TRUE;

cleanup:
    free(pngData);
    free(jpegData);
    free(base64Data);
    GdipDisposeImage((GpImage *)bitmap);
    return success;
}

/* ── IPv4 discovery and micro HTTP server ─────────────────────────────── */

static BOOL AddDetectedIp(char ips[][INET_ADDRSTRLEN], int *count, int maxCount,
                          const char *candidate)
{
    for (int i = 0; i < *count; i++) {
        if (strcmp(ips[i], candidate) == 0) return TRUE;
    }
    if (*count >= maxCount) return FALSE;
    strncpy(ips[*count], candidate, INET_ADDRSTRLEN - 1);
    ips[*count][INET_ADDRSTRLEN - 1] = '\0';
    (*count)++;
    return TRUE;
}

static int EnumerateDetectedIpv4Addresses(char ips[][INET_ADDRSTRLEN], int maxCount)
{
    ULONG size = 15000;
    IP_ADAPTER_ADDRESSES *addresses = NULL;
    ULONG result;
    int count = 0;

    AddDetectedIp(ips, &count, maxCount, "127.0.0.1");
    for (int attempt = 0; attempt < 3; attempt++) {
        addresses = (IP_ADAPTER_ADDRESSES *)malloc(size);
        if (!addresses) return count;
        result = GetAdaptersAddresses(AF_INET,
            GAA_FLAG_SKIP_ANYCAST | GAA_FLAG_SKIP_MULTICAST | GAA_FLAG_SKIP_DNS_SERVER,
            NULL, addresses, &size);
        if (result != ERROR_BUFFER_OVERFLOW) break;
        free(addresses);
        addresses = NULL;
    }

    if (addresses && result == NO_ERROR) {
        for (IP_ADAPTER_ADDRESSES *adapter = addresses; adapter;
             adapter = adapter->Next) {
            if (adapter->OperStatus != IfOperStatusUp) continue;
            for (IP_ADAPTER_UNICAST_ADDRESS *unicast = adapter->FirstUnicastAddress;
                 unicast; unicast = unicast->Next) {
                char text[INET_ADDRSTRLEN];
                SOCKADDR *address = unicast->Address.lpSockaddr;
                if (!address || address->sa_family != AF_INET) continue;
                if (InetNtopA(AF_INET,
                              &((SOCKADDR_IN *)address)->sin_addr,
                              text, sizeof(text))) {
                    AddDetectedIp(ips, &count, maxCount, text);
                }
            }
        }
    }
    free(addresses);
    return count;
}

static BOOL IsConfiguredBindAddressPresent(void)
{
    char ips[MAX_DETECTED_IPS][INET_ADDRSTRLEN];
    int count;

    if (strcmp(g_configBindIp, "0.0.0.0") == 0) return TRUE;
    count = EnumerateDetectedIpv4Addresses(ips, MAX_DETECTED_IPS);
    for (int i = 0; i < count; i++) {
        if (strcmp(ips[i], g_configBindIp) == 0) return TRUE;
    }
    return FALSE;
}

static void SetHttpStatus(int state, const char *fmt, ...)
{
    char status[sizeof(g_httpStatus)];
    va_list args;
    va_start(args, fmt);
    vsnprintf(status, sizeof(status), fmt, args);
    va_end(args);
    status[sizeof(status) - 1] = '\0';

    if (state != g_httpState || strcmp(status, g_httpStatus) != 0) {
        g_httpState = state;
        strncpy(g_httpStatus, status, sizeof(g_httpStatus) - 1);
        g_httpStatus[sizeof(g_httpStatus) - 1] = '\0';
        LogMessage("HTTP server: %s", g_httpStatus);
        UpdateTooltip();
    }
}

static BOOL SocketSendAll(SOCKET client, const BYTE *data, DWORD size)
{
    DWORD sentTotal = 0;
    while (sentTotal < size) {
        int chunk = (size - sentTotal > 1024 * 1024)
            ? 1024 * 1024 : (int)(size - sentTotal);
        int sent = send(client, (const char *)data + sentTotal, chunk, 0);
        if (sent == SOCKET_ERROR || sent == 0) return FALSE;
        sentTotal += (DWORD)sent;
    }
    return TRUE;
}

static int FindImageForToken(const char *token, BYTE **jpegCopy, DWORD *jpegSize)
{
    int status = HTTP_EVENT_NOT_FOUND;
    *jpegCopy = NULL;
    *jpegSize = 0;

    AcquireSRWLockShared(&g_imageLock);
    if (g_cachedImage.token[0] && strcmp(g_cachedImage.token, token) == 0) {
        *jpegCopy = LoadCachedImageBytesLocked(&g_cachedImage, jpegSize);
        status = *jpegCopy ? HTTP_EVENT_SERVED : 500;
    } else {
        for (size_t i = g_imageHistoryCount; i > 0; i--) {
            CachedImage *historical = &g_imageHistory[i - 1];
            if (strcmp(historical->token, token) == 0) {
                *jpegCopy = LoadCachedImageBytesLocked(historical, jpegSize);
                status = *jpegCopy ? HTTP_EVENT_SERVED : 500;
                break;
            }
        }
        if (status == HTTP_EVENT_NOT_FOUND) {
            for (size_t i = 0; i < g_goneTokenCount; i++) {
                if (strcmp(g_goneTokens[i], token) == 0) {
                    status = HTTP_EVENT_GONE;
                    break;
                }
            }
        }
    }
    ReleaseSRWLockShared(&g_imageLock);
    return status;
}

static BOOL ParseImageTokenFromPath(const char *path,
                                    char token[IMAGE_TOKEN_HEX_LEN + 1])
{
    static const char prefix[] = "/";
    const char *value;
    size_t pathLen = strlen(path);
    size_t expectedLen = strlen(prefix) + IMAGE_TOKEN_HEX_LEN + 4;

    if (pathLen != expectedLen || strncmp(path, prefix, strlen(prefix)) != 0 ||
        strcmp(path + pathLen - 4, ".jpg") != 0) return FALSE;
    value = path + strlen(prefix);
    for (int i = 0; i < IMAGE_TOKEN_HEX_LEN; i++) {
        char c = value[i];
        if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f'))) return FALSE;
        token[i] = c;
    }
    token[IMAGE_TOKEN_HEX_LEN] = '\0';
    return TRUE;
}

static void SendHttpResponse(SOCKET client, int status, BOOL headOnly,
                             const BYTE *imageData, DWORD imageSize)
{
    const char *reason;
    const char *contentType;
    const char *messageHeader;
    const char *body;
    char headers[1024];
    DWORD bodySize;

    switch (status) {
    case HTTP_EVENT_SERVED:
        reason = "OK";
        contentType = "image/jpeg";
        messageHeader = "Retained clipboard image";
        body = NULL;
        bodySize = imageSize;
        break;
    case HTTP_EVENT_GONE:
        reason = "Gone";
        contentType = "text/plain; charset=utf-8";
        messageHeader = "This image was evicted from the image history";
        body = "This image is no longer available because it was evicted from ImagePaster's image history.\n";
        bodySize = (DWORD)strlen(body);
        break;
    case 405:
        reason = "Method Not Allowed";
        contentType = "text/plain; charset=utf-8";
        messageHeader = "Only GET and HEAD are supported";
        body = "Only GET and HEAD requests are supported.\n";
        bodySize = (DWORD)strlen(body);
        break;
    case 500:
        reason = "Internal Server Error";
        contentType = "text/plain; charset=utf-8";
        messageHeader = "The image could not be prepared";
        body = "The image exists but could not be prepared for this request.\n";
        bodySize = (DWORD)strlen(body);
        break;
    default:
        status = HTTP_EVENT_NOT_FOUND;
        reason = "Not Found";
        contentType = "text/plain; charset=utf-8";
        messageHeader = "No image was issued for this URL";
        body = "No image was issued for this URL. Check that the complete image URL was used.\n";
        bodySize = (DWORD)strlen(body);
        break;
    }

    snprintf(headers, sizeof(headers),
        "HTTP/1.1 %d %s\r\n"
        "Content-Type: %s\r\n"
        "Content-Length: %lu\r\n"
        "Cache-Control: no-store\r\n"
        "X-Content-Type-Options: nosniff\r\n"
        "X-ImagePaster-Message: %s\r\n"
        "%s"
        "Connection: close\r\n\r\n",
        status, reason, contentType, bodySize, messageHeader,
        status == 405 ? "Allow: GET, HEAD\r\n" : "");
    SocketSendAll(client, (const BYTE *)headers, (DWORD)strlen(headers));
    if (!headOnly) {
        if (status == HTTP_EVENT_SERVED) {
            SocketSendAll(client, imageData, imageSize);
        } else if (body) {
            SocketSendAll(client, (const BYTE *)body, bodySize);
        }
    }
}

/* Parses one allowlist entry: "a.b.c.d" (treated as /32) or "a.b.c.d/n". */
static BOOL ParseHttpAllowEntry(const char *entry, HttpAllowRule *rule)
{
    char addressText[64];
    const char *slash = strchr(entry, '/');
    size_t addressLen = slash ? (size_t)(slash - entry) : strlen(entry);
    IN_ADDR parsed;
    int prefix = 32;

    if (addressLen == 0 || addressLen >= sizeof(addressText)) return FALSE;
    memcpy(addressText, entry, addressLen);
    addressText[addressLen] = '\0';
    if (InetPtonA(AF_INET, addressText, &parsed) != 1) return FALSE;
    if (slash) {
        const char *digits = slash + 1;
        size_t digitCount = strlen(digits);
        if (digitCount == 0 || digitCount > 2) return FALSE;
        for (size_t i = 0; i < digitCount; i++) {
            if (digits[i] < '0' || digits[i] > '9') return FALSE;
        }
        prefix = atoi(digits);
        if (prefix > 32) return FALSE;
    }
    rule->mask = prefix == 0 ? 0 : 0xffffffffu << (32 - prefix);
    rule->network = ntohl(parsed.s_addr) & rule->mask;
    return TRUE;
}

/* Parses a comma-separated allowlist. Blank segments are skipped; any
   malformed entry or rule overflow fails the whole list. */
static BOOL ParseHttpAllowList(const char *text, HttpAllowRule *rules,
                               int maxRules, int *count)
{
    char copy[MAX_HTTP_ALLOW_LIST_BYTES];
    char *entry;
    int parsed = 0;

    *count = 0;
    strncpy(copy, text, sizeof(copy) - 1);
    copy[sizeof(copy) - 1] = '\0';

    entry = strtok(copy, ",");
    while (entry) {
        char *start = entry;
        char *end;
        while (*start == ' ' || *start == '\t') start++;
        end = start + strlen(start);
        while (end > start && (end[-1] == ' ' || end[-1] == '\t')) end--;
        *end = '\0';
        if (*start) {
            if (parsed >= maxRules) return FALSE;
            if (!ParseHttpAllowEntry(start, &rules[parsed])) return FALSE;
            parsed++;
        }
        entry = strtok(NULL, ",");
    }
    *count = parsed;
    return TRUE;
}

/* Publishes the parsed form of g_configHttpAllowList for the HTTP worker.
   An unparseable list fails closed: every client is rejected until the
   configuration is corrected. */
static void ApplyHttpAllowList(void)
{
    HttpAllowRule rules[MAX_HTTP_ALLOW_RULES];
    int count = 0;
    BOOL restrictive;

    if (ParseHttpAllowList(g_configHttpAllowList, rules,
                           MAX_HTTP_ALLOW_RULES, &count)) {
        restrictive = count > 0;
    } else {
        LogMessage("WARNING: HTTP allowlist is invalid; rejecting all clients until it is corrected");
        count = 0;
        restrictive = TRUE;
    }

    AcquireSRWLockExclusive(&g_httpAllowLock);
    if (count > 0) {
        memcpy(g_httpAllowRules, rules, (size_t)count * sizeof(rules[0]));
    }
    g_httpAllowRuleCount = count;
    g_httpAllowRestrictive = restrictive;
    ReleaseSRWLockExclusive(&g_httpAllowLock);
}

/* Called from the HTTP worker thread for every accepted connection. */
static BOOL IsHttpClientAllowed(DWORD addressHostOrder)
{
    BOOL allowed = FALSE;

    AcquireSRWLockShared(&g_httpAllowLock);
    if (!g_httpAllowRestrictive) {
        allowed = TRUE;
    } else {
        for (int i = 0; i < g_httpAllowRuleCount; i++) {
            if ((addressHostOrder & g_httpAllowRules[i].mask) ==
                g_httpAllowRules[i].network) {
                allowed = TRUE;
                break;
            }
        }
    }
    ReleaseSRWLockShared(&g_httpAllowLock);
    return allowed;
}

static void HandleHttpClient(SOCKET client)
{
    char request[4096];
    int used = 0;
    char method[16] = {0};
    char path[512] = {0};
    char protocol[16] = {0};
    char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
    BYTE *jpegCopy = NULL;
    DWORD jpegSize = 0;
    int status;
    BOOL headOnly = FALSE;

    while (used < (int)sizeof(request) - 1) {
        int received = recv(client, request + used, (int)sizeof(request) - 1 - used, 0);
        if (received <= 0) return;
        used += received;
        request[used] = '\0';
        if (strstr(request, "\r\n\r\n")) break;
    }

    if (sscanf(request, "%15s %511s %15s", method, path, protocol) != 3) {
        SendHttpResponse(client, HTTP_EVENT_NOT_FOUND, FALSE, NULL, 0);
        return;
    }
    if (strcmp(method, "GET") != 0 && strcmp(method, "HEAD") != 0) {
        SendHttpResponse(client, 405, FALSE, NULL, 0);
        return;
    }
    headOnly = strcmp(method, "HEAD") == 0;
    if (!ParseImageTokenFromPath(path, token)) {
        status = HTTP_EVENT_NOT_FOUND;
    } else {
        status = FindImageForToken(token, &jpegCopy, &jpegSize);
    }

    SendHttpResponse(client, status, headOnly, jpegCopy, jpegSize);
    free(jpegCopy);
    if (status == HTTP_EVENT_SERVED || status == HTTP_EVENT_GONE ||
        status == HTTP_EVENT_NOT_FOUND || status == 500) {
        PostMessage(g_hWndMain, WM_HTTP_EVENT, (WPARAM)status, 0);
    }
}

static DWORD WINAPI HttpServerThreadProc(LPVOID parameter)
{
    SOCKET listenSocket = (SOCKET)(UINT_PTR)parameter;

    while (InterlockedCompareExchange(&g_httpStopRequested, 0, 0) == 0) {
        fd_set readSet;
        struct timeval timeout;
        FD_ZERO(&readSet);
        FD_SET(listenSocket, &readSet);
        timeout.tv_sec = 0;
        timeout.tv_usec = 500000;
        int ready = select(0, &readSet, NULL, NULL, &timeout);
        if (ready == SOCKET_ERROR) break;
        if (ready == 0) continue;

        struct sockaddr_in clientAddress;
        int addressLength = (int)sizeof(clientAddress);
        ZeroMemory(&clientAddress, sizeof(clientAddress));
        SOCKET client = accept(listenSocket,
                               (struct sockaddr *)&clientAddress,
                               &addressLength);
        if (client == INVALID_SOCKET) {
            if (InterlockedCompareExchange(&g_httpStopRequested, 0, 0) != 0) break;
            continue;
        }

        if (!IsHttpClientAllowed(ntohl(clientAddress.sin_addr.s_addr))) {
            /* Drop before reading anything; the request never gets a
               response, matching firewall-style filtering. */
            closesocket(client);
            PostMessage(g_hWndMain, WM_HTTP_EVENT, (WPARAM)HTTP_EVENT_DENIED,
                        (LPARAM)clientAddress.sin_addr.s_addr);
            continue;
        }

        {
            DWORD timeoutMs = 2000;
            setsockopt(client, SOL_SOCKET, SO_RCVTIMEO,
                       (const char *)&timeoutMs, sizeof(timeoutMs));
            setsockopt(client, SOL_SOCKET, SO_SNDTIMEO,
                       (const char *)&timeoutMs, sizeof(timeoutMs));
        }

        EnterCriticalSection(&g_httpSocketLock);
        g_httpClientSocket = client;
        LeaveCriticalSection(&g_httpSocketLock);

        HandleHttpClient(client);

        EnterCriticalSection(&g_httpSocketLock);
        if (g_httpClientSocket == client) {
            shutdown(client, SD_BOTH);
            closesocket(client);
            g_httpClientSocket = INVALID_SOCKET;
        }
        LeaveCriticalSection(&g_httpSocketLock);
    }
    return 0;
}

static void StopHttpServer(void)
{
    InterlockedExchange(&g_httpStopRequested, 1);
    if (g_httpListenSocket != INVALID_SOCKET) {
        shutdown(g_httpListenSocket, SD_BOTH);
        closesocket(g_httpListenSocket);
        g_httpListenSocket = INVALID_SOCKET;
    }
    if (g_httpSocketLockReady) {
        EnterCriticalSection(&g_httpSocketLock);
        if (g_httpClientSocket != INVALID_SOCKET) {
            shutdown(g_httpClientSocket, SD_BOTH);
            closesocket(g_httpClientSocket);
            g_httpClientSocket = INVALID_SOCKET;
        }
        LeaveCriticalSection(&g_httpSocketLock);
    }
    if (g_httpThread) {
        DWORD waitResult = WaitForSingleObject(g_httpThread, 5000);
        if (waitResult == WAIT_OBJECT_0) {
            CloseHandle(g_httpThread);
            g_httpThread = NULL;
        } else {
            SetHttpStatus(3, "HTTP worker is still shutting down");
        }
    }
}

static BOOL StartHttpServer(void)
{
    SOCKET listenSocket;
    SOCKADDR_IN address;
    BOOL exclusive = TRUE;

    if (!g_winsockReady) {
        SetHttpStatus(3, "unavailable because Winsock failed to initialize");
        return FALSE;
    }
    if (!IsConfiguredBindAddressPresent()) {
        SetHttpStatus(2, "waiting for %s to become available", g_configBindIp);
        return FALSE;
    }

    listenSocket = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (listenSocket == INVALID_SOCKET) {
        SetHttpStatus(3, "socket creation failed (%d)", WSAGetLastError());
        return FALSE;
    }
    setsockopt(listenSocket, SOL_SOCKET, SO_EXCLUSIVEADDRUSE,
               (const char *)&exclusive, sizeof(exclusive));

    ZeroMemory(&address, sizeof(address));
    address.sin_family = AF_INET;
    address.sin_port = htons((u_short)g_configHttpPort);
    if (InetPtonA(AF_INET, g_configBindIp, &address.sin_addr) != 1) {
        closesocket(listenSocket);
        SetHttpStatus(3, "configured bind address is invalid: %s", g_configBindIp);
        return FALSE;
    }
    if (bind(listenSocket, (SOCKADDR *)&address, sizeof(address)) == SOCKET_ERROR) {
        int error = WSAGetLastError();
        closesocket(listenSocket);
        SetHttpStatus(3, "cannot bind %s:%d (Winsock error %d)",
                      g_configBindIp, g_configHttpPort, error);
        return FALSE;
    }
    if (listen(listenSocket, SOMAXCONN) == SOCKET_ERROR) {
        int error = WSAGetLastError();
        closesocket(listenSocket);
        SetHttpStatus(3, "cannot listen on %s:%d (Winsock error %d)",
                      g_configBindIp, g_configHttpPort, error);
        return FALSE;
    }

    InterlockedExchange(&g_httpStopRequested, 0);
    g_httpListenSocket = listenSocket;
    g_httpThread = CreateThread(NULL, 0, HttpServerThreadProc,
                                (LPVOID)(UINT_PTR)listenSocket, 0, NULL);
    if (!g_httpThread) {
        closesocket(listenSocket);
        g_httpListenSocket = INVALID_SOCKET;
        SetHttpStatus(3, "could not start the HTTP worker thread");
        return FALSE;
    }

    SetHttpStatus(1, "listening on http://%s:%d", g_configBindIp, g_configHttpPort);
    return TRUE;
}

static void ReconcileHttpServer(void)
{
    BOOL addressPresent = IsConfiguredBindAddressPresent();

    if (g_httpThread && WaitForSingleObject(g_httpThread, 0) == WAIT_OBJECT_0) {
        CloseHandle(g_httpThread);
        g_httpThread = NULL;
        if (g_httpListenSocket != INVALID_SOCKET) {
            closesocket(g_httpListenSocket);
            g_httpListenSocket = INVALID_SOCKET;
        }
        SetHttpStatus(3, "listener stopped unexpectedly; retrying");
    }

    if (!addressPresent) {
        if (g_httpThread || g_httpListenSocket != INVALID_SOCKET) StopHttpServer();
        SetHttpStatus(2, "waiting for %s to become available", g_configBindIp);
        return;
    }
    if (!g_httpThread) StartHttpServer();
}

/* ── Paste re-injection ─────────────────────────────────────────────────── */

static void SimulateStandardTextPaste(void)
{
    INPUT inputs[4];
    UINT sent;
    ZeroMemory(inputs, sizeof(inputs));

    /* Ctrl key down */
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_CONTROL;

    /* V key down */
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = 'V';

    /* V key up */
    inputs[2].type = INPUT_KEYBOARD;
    inputs[2].ki.wVk = 'V';
    inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;

    /* Ctrl key up */
    inputs[3].type = INPUT_KEYBOARD;
    inputs[3].ki.wVk = VK_CONTROL;
    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

    /* Identify our injected paste without a cross-thread skip-next flag. */
    for (size_t i = 0; i < sizeof(inputs) / sizeof(inputs[0]); i++) {
        inputs[i].ki.dwExtraInfo = PASTE_INPUT_TAG;
    }

    sent = SendInput(4, inputs, sizeof(INPUT));
    if (sent == 4) {
        LogMessage("Simulated Ctrl+V (generic text paste)");
    } else {
        LogMessage("ERROR: Ctrl+V re-injection sent %u of 4 events (%lu)",
                   sent, GetLastError());
    }
}

static void SimulateCompatibilityTextPaste(void)
{
    INPUT inputs[4];
    UINT sent;
    ZeroMemory(inputs, sizeof(inputs));

    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_SHIFT;

    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = VK_INSERT;
    inputs[1].ki.dwFlags = KEYEVENTF_EXTENDEDKEY;

    inputs[2].type = INPUT_KEYBOARD;
    inputs[2].ki.wVk = VK_INSERT;
    inputs[2].ki.dwFlags = KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP;

    inputs[3].type = INPUT_KEYBOARD;
    inputs[3].ki.wVk = VK_SHIFT;
    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

    sent = SendInput(4, inputs, sizeof(INPUT));
    if (sent == 4) {
        LogMessage("Simulated Shift+Insert (compatibility text paste)");
    } else {
        LogMessage("ERROR: Shift+Insert re-injection sent %u of 4 events (%lu)",
                   sent, GetLastError());
    }
}

static void SimulateConfiguredTextPaste(void)
{
    if (g_configCompatibilityPaste) {
        SimulateCompatibilityTextPaste();
    } else {
        SimulateStandardTextPaste();
    }
}

static BOOL PlaceUtf8TextOnClipboard(const char *text, BOOL preserveCachedImage)
{
    int wideLen;
    HGLOBAL memory;
    WCHAR *wideText;

    wideLen = MultiByteToWideChar(CP_UTF8, 0, text, -1, NULL, 0);
    if (wideLen <= 0) return FALSE;
    memory = GlobalAlloc(GMEM_MOVEABLE, (SIZE_T)wideLen * sizeof(WCHAR));
    if (!memory) return FALSE;
    wideText = (WCHAR *)GlobalLock(memory);
    if (!wideText) {
        GlobalFree(memory);
        return FALSE;
    }
    MultiByteToWideChar(CP_UTF8, 0, text, -1, wideText, wideLen);
    GlobalUnlock(memory);

    g_writingClipboardText = TRUE;
    if (!OpenClipboard(g_hWndMain)) {
        g_writingClipboardText = FALSE;
        GlobalFree(memory);
        LogMessage("ERROR: OpenClipboard for text paste failed (%lu)", GetLastError());
        return FALSE;
    }
    EmptyClipboard();
    if (!SetClipboardData(CF_UNICODETEXT, memory)) {
        LogMessage("ERROR: SetClipboardData(CF_UNICODETEXT) failed (%lu)", GetLastError());
        CloseClipboard();
        g_writingClipboardText = FALSE;
        GlobalFree(memory);
        return FALSE;
    }
    CloseClipboard();

    SetCachedClipboardSequence(GetClipboardSequenceNumber());
    g_ownClipboardSequence = preserveCachedImage ? GetCachedClipboardSequence() : 0;
    g_writingClipboardText = FALSE;
    if (!preserveCachedImage) {
        ClearCurrentImage("activity log was copied as text");
    }
    return TRUE;
}

static BOOL CopyActivityLogToClipboard(void)
{
    int entryCount = g_logCount;
    size_t bufferSize = (size_t)entryCount *
                        (sizeof(g_logRing[0].time) + sizeof(g_logRing[0].message) + 4) + 1;
    char *textBuffer = (char *)malloc(bufferSize);
    size_t position = 0;

    if (!textBuffer) {
        LogMessage("ERROR: Could not allocate activity log clipboard text");
        return FALSE;
    }
    textBuffer[0] = '\0';

    for (int i = 0; i < entryCount; i++) {
        int bufferIndex = entryCount < LOG_RING_CAPACITY
            ? i : (g_logHead + i) % LOG_RING_CAPACITY;
        LogEntry *entry = &g_logRing[bufferIndex];
        int written = snprintf(textBuffer + position, bufferSize - position,
                               "%s  %s\r\n", entry->time, entry->message);
        if (written < 0 || (size_t)written >= bufferSize - position) {
            free(textBuffer);
            LogMessage("ERROR: Could not format activity log clipboard text");
            return FALSE;
        }
        position += (size_t)written;
    }

    if (!PlaceUtf8TextOnClipboard(textBuffer, FALSE)) {
        free(textBuffer);
        return FALSE;
    }

    free(textBuffer);
    LogMessage("Activity log copied to clipboard (%d entries)", entryCount);
    return TRUE;
}

static char *BuildHttpPasteText(const char *token, DWORD *textLen)
{
    char imageUrl[128];
    const char *messageTemplate = g_configHttpMessageTemplate;
    const size_t placeholderLen = sizeof(HTTP_URL_PLACEHOLDER) - 1;
    size_t templateLen;
    size_t urlLen;
    size_t placeholderCount = 0;
    size_t outputLen;
    const char *scan;
    const char *match;
    char *result;
    char *write;
    int urlWritten;

    if (!token || !*token || !textLen) return NULL;
    *textLen = 0;

    urlWritten = snprintf(imageUrl, sizeof(imageUrl), "http://%s:%d/%s.jpg",
                          g_configBindIp, g_configHttpPort, token);
    if (urlWritten <= 0 || urlWritten >= (int)sizeof(imageUrl)) return NULL;

    templateLen = strlen(messageTemplate);
    urlLen = (size_t)urlWritten;
    scan = messageTemplate;
    while ((match = strstr(scan, HTTP_URL_PLACEHOLDER)) != NULL) {
        placeholderCount++;
        scan = match + placeholderLen;
    }
    if (placeholderCount == 0) return NULL;

    outputLen = templateLen;
    if (urlLen >= placeholderLen) {
        outputLen += placeholderCount * (urlLen - placeholderLen);
    } else {
        outputLen -= placeholderCount * (placeholderLen - urlLen);
    }
    if (outputLen > MAXDWORD) return NULL;

    result = (char *)malloc(outputLen + 1);
    if (!result) return NULL;

    scan = messageTemplate;
    write = result;
    while ((match = strstr(scan, HTTP_URL_PLACEHOLDER)) != NULL) {
        size_t prefixLen = (size_t)(match - scan);
        memcpy(write, scan, prefixLen);
        write += prefixLen;
        memcpy(write, imageUrl, urlLen);
        write += urlLen;
        scan = match + placeholderLen;
    }
    strcpy(write, scan);
    *textLen = (DWORD)outputLen;
    return result;
}

static BOOL PasteCachedImage(void)
{
    char *pasteText = NULL;
    DWORD textLen = 0;
    char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
    BOOL success = FALSE;

    AcquireSRWLockShared(&g_imageLock);
    if (g_cachedImage.token[0]) {
        strncpy(token, g_cachedImage.token, sizeof(token) - 1);
        if (g_configPasteMethod == PASTE_METHOD_BASE64 && g_cachedImage.base64Data) {
            textLen = g_cachedImage.base64Len;
            pasteText = (char *)malloc((size_t)textLen + 1);
            if (pasteText) memcpy(pasteText, g_cachedImage.base64Data, (size_t)textLen + 1);
        }
    }
    ReleaseSRWLockShared(&g_imageLock);

    if (!token[0]) {
        LogMessage("Paste cancelled: no current clipboard image is cached");
        return FALSE;
    }

    if (g_configPasteMethod == PASTE_METHOD_HTTP) {
        pasteText = BuildHttpPasteText(token, &textLen);
    }

    if (!pasteText) {
        LogMessage("ERROR: Could not prepare paste text");
        return FALSE;
    }

    if (PlaceUtf8TextOnClipboard(pasteText, TRUE)) {
        LogMessage("Prepared %s paste (%lu characters, image id %.12s...)",
                   g_configPasteMethod == PASTE_METHOD_HTTP ? "HTTP URL" : "base64",
                   textLen, token);
        SimulateConfiguredTextPaste();
        success = TRUE;
    }
    free(pasteText);
    return success;
}

/* ── Keyword parsing ───────────────────────────────────────────────────── */

/* The callback reads a single atomic word, never UI-owned window/config data.
   Wake the input thread to reconcile the independent hotkey immediately. */
static void SetHookCaptureFlag(LONG flag, BOOL enabled)
{
    if (enabled) InterlockedOr(&g_hookCaptureState, flag);
    else InterlockedAnd(&g_hookCaptureState, ~flag);
    if (g_keyboardHookWakeEvent) SetEvent(g_keyboardHookWakeEvent);
}

static void ParseKeywords(void)
{
    AcquireSRWLockExclusive(&g_keywordLock);
    g_keywordCount = 0;
    char copy[2048];
    strncpy(copy, g_configTitleMatch, sizeof(copy) - 1);
    copy[sizeof(copy) - 1] = '\0';

    char *token = strtok(copy, ",");
    while (token && g_keywordCount < MAX_KEYWORDS) {
        /* Trim leading/trailing whitespace */
        while (*token == ' ' || *token == '\t') token++;
        char *end = token + strlen(token) - 1;
        while (end > token && (*end == ' ' || *end == '\t')) *end-- = '\0';

        if (*token) {
            /* Convert to lowercase wide string */
            MultiByteToWideChar(CP_UTF8, 0, token, -1,
                                g_keywords[g_keywordCount], 128);
            /* Lowercase it */
            for (WCHAR *p = g_keywords[g_keywordCount]; *p; p++) {
                if (*p >= L'A' && *p <= L'Z')
                    *p = *p - L'A' + L'a';
            }
            g_keywordCount++;
        }
        token = strtok(NULL, ",");
    }
    InterlockedExchange(&g_hookHasKeywords, g_keywordCount != 0);
    ReleaseSRWLockExclusive(&g_keywordLock);
    SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, g_configScreenCaptureEnabled);
}

/* ── Registry configuration ──────────────────────────────────────────── */

static BOOL LoadConfigFromRegistry(void)
{
    HKEY hKey;
    LONG result = RegOpenKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, KEY_READ, &hKey);
    if (result != ERROR_SUCCESS) return FALSE;

    DWORD type, size, value;
    size = sizeof(g_configTitleMatch);
    if (RegQueryValueExA(hKey, REG_VALUE_TITLE, NULL, &type,
                         (LPBYTE)g_configTitleMatch, &size) != ERROR_SUCCESS
        || type != REG_SZ) {
        strcpy(g_configTitleMatch, "xshell");
    }
    g_configTitleMatch[sizeof(g_configTitleMatch) - 1] = '\0';

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_METHOD, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value <= PASTE_METHOD_HTTP) {
        g_configPasteMethod = (int)value;
    }

    size = sizeof(g_configHttpMessageTemplate);
    if (RegQueryValueExA(hKey, REG_VALUE_HTTP_MESSAGE_TEMPLATE, NULL, &type,
                         (LPBYTE)g_configHttpMessageTemplate, &size) != ERROR_SUCCESS ||
        type != REG_SZ) {
        strcpy(g_configHttpMessageTemplate, DEFAULT_HTTP_MESSAGE_TEMPLATE);
    }
    g_configHttpMessageTemplate[sizeof(g_configHttpMessageTemplate) - 1] = '\0';
    if (!strstr(g_configHttpMessageTemplate, HTTP_URL_PLACEHOLDER)) {
        strcpy(g_configHttpMessageTemplate, DEFAULT_HTTP_MESSAGE_TEMPLATE);
    }

    size = sizeof(g_configBindIp);
    if (RegQueryValueExA(hKey, REG_VALUE_BIND_IP, NULL, &type,
                         (LPBYTE)g_configBindIp, &size) != ERROR_SUCCESS ||
        type != REG_SZ) {
        strcpy(g_configBindIp, "127.0.0.1");
    }
    g_configBindIp[sizeof(g_configBindIp) - 1] = '\0';

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_PORT, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value >= 1 && value <= 65535) {
        g_configHttpPort = (int)value;
    }

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_QUALITY, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value <= 100) {
        g_configJpegQuality = (int)value;
    }

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_HISTORY_LIMIT, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value <= MAX_IMAGE_HISTORY_LIMIT) {
        g_configImageHistoryLimit = (int)value;
    }

    size = sizeof(value);
    result = RegQueryValueExA(hKey, REG_VALUE_COMPATIBILITY_PASTE, NULL, &type,
                              (LPBYTE)&value, &size);
    if (result != ERROR_SUCCESS) {
        size = sizeof(value);
        result = RegQueryValueExA(hKey, REG_VALUE_LEGACY_COMPATIBILITY_PASTE,
                                  NULL, &type, (LPBYTE)&value, &size);
    }
    if (result == ERROR_SUCCESS &&
        type == REG_DWORD && value <= 1) {
        g_configCompatibilityPaste = value != 0;
    }

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_SCREEN_CAPTURE, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value <= 1) {
        g_configScreenCaptureEnabled = value != 0;
    }

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_CAPTURE_GAP_FILL, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value <= CAPTURE_GAP_FILL_BLUR) {
        g_configCaptureGapFill = (int)value;
    }

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_AUTO_UPDATE, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value <= 1) {
        g_configAutoCheckForUpdates = value != 0;
    }

    size = sizeof(g_ignoredUpdateVersion);
    if (RegQueryValueExA(hKey, REG_VALUE_IGNORED_UPDATE_VERSION, NULL, &type,
                         (LPBYTE)g_ignoredUpdateVersion, &size) != ERROR_SUCCESS ||
        type != REG_SZ || size == 0) {
        g_ignoredUpdateVersion[0] = '\0';
    }
    g_ignoredUpdateVersion[sizeof(g_ignoredUpdateVersion) - 1] = '\0';

    size = sizeof(g_configHttpAllowList);
    if (RegQueryValueExA(hKey, REG_VALUE_HTTP_ALLOW_LIST, NULL, &type,
                         (LPBYTE)g_configHttpAllowList, &size) != ERROR_SUCCESS ||
        type != REG_SZ) {
        g_configHttpAllowList[0] = '\0';
    }
    g_configHttpAllowList[sizeof(g_configHttpAllowList) - 1] = '\0';

    size = sizeof(value);
    if (RegQueryValueExA(hKey, REG_VALUE_IMAGE_STORAGE, NULL, &type,
                         (LPBYTE)&value, &size) == ERROR_SUCCESS &&
        type == REG_DWORD && value <= IMAGE_STORAGE_DISK) {
        g_configImageStorage = (int)value;
    }

    RegCloseKey(hKey);
    return TRUE;
}

static void SaveConfigToRegistry(void)
{
    HKEY hKey;
    DWORD disposition;
    char historyText[32];
    LONG result = RegCreateKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                                  REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL,
                                  &hKey, &disposition);
    if (result != ERROR_SUCCESS) return;

    RegSetValueExA(hKey, REG_VALUE_TITLE, 0, REG_SZ,
                   (const BYTE*)g_configTitleMatch,
                   (DWORD)(strlen(g_configTitleMatch) + 1));
    RegSetValueExA(hKey, REG_VALUE_HTTP_MESSAGE_TEMPLATE, 0, REG_SZ,
                   (const BYTE*)g_configHttpMessageTemplate,
                   (DWORD)(strlen(g_configHttpMessageTemplate) + 1));
    RegSetValueExA(hKey, REG_VALUE_BIND_IP, 0, REG_SZ,
                   (const BYTE*)g_configBindIp,
                   (DWORD)(strlen(g_configBindIp) + 1));
    {
        DWORD value = (DWORD)g_configPasteMethod;
        RegSetValueExA(hKey, REG_VALUE_METHOD, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        value = (DWORD)g_configHttpPort;
        RegSetValueExA(hKey, REG_VALUE_PORT, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        value = (DWORD)g_configJpegQuality;
        RegSetValueExA(hKey, REG_VALUE_QUALITY, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        value = (DWORD)g_configImageHistoryLimit;
        RegSetValueExA(hKey, REG_VALUE_HISTORY_LIMIT, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        value = (DWORD)g_configCompatibilityPaste;
        RegSetValueExA(hKey, REG_VALUE_COMPATIBILITY_PASTE, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        RegDeleteValueA(hKey, REG_VALUE_LEGACY_COMPATIBILITY_PASTE);
        value = (DWORD)g_configScreenCaptureEnabled;
        RegSetValueExA(hKey, REG_VALUE_SCREEN_CAPTURE, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        value = (DWORD)g_configCaptureGapFill;
        RegSetValueExA(hKey, REG_VALUE_CAPTURE_GAP_FILL, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        value = (DWORD)g_configAutoCheckForUpdates;
        RegSetValueExA(hKey, REG_VALUE_AUTO_UPDATE, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
        value = (DWORD)g_configImageStorage;
        RegSetValueExA(hKey, REG_VALUE_IMAGE_STORAGE, 0, REG_DWORD,
                       (const BYTE *)&value, sizeof(value));
    }
    RegSetValueExA(hKey, REG_VALUE_IGNORED_UPDATE_VERSION, 0, REG_SZ,
                   (const BYTE *)g_ignoredUpdateVersion,
                   (DWORD)(strlen(g_ignoredUpdateVersion) + 1));
    RegSetValueExA(hKey, REG_VALUE_HTTP_ALLOW_LIST, 0, REG_SZ,
                   (const BYTE *)g_configHttpAllowList,
                   (DWORD)(strlen(g_configHttpAllowList) + 1));

    RegCloseKey(hKey);
    if (g_configImageHistoryLimit == 0) {
        strcpy(historyText, "unlimited");
    } else {
        snprintf(historyText, sizeof(historyText), "%d",
                 g_configImageHistoryLimit);
    }
    LogMessage("Configuration saved: method=%s, shortcut=%s, capture=%s, gap=%s, bind=%s:%d, JPEG=%d%%, history=%s, storage=%s, allow=%s, titles=%s",
               g_configPasteMethod == PASTE_METHOD_HTTP ? "HTTP" : "base64",
               g_configCompatibilityPaste ? "Shift+Insert" : "Ctrl+V",
               g_configScreenCaptureEnabled ? "enabled" : "disabled",
               g_configCaptureGapFill == CAPTURE_GAP_FILL_BLACK ? "black" :
               g_configCaptureGapFill == CAPTURE_GAP_FILL_BLUR ? "blur" :
                                                                  "white",
               g_configBindIp, g_configHttpPort, g_configJpegQuality,
               historyText,
               g_configImageStorage == IMAGE_STORAGE_DISK ? "disk" : "memory",
               g_configHttpAllowList[0] ? g_configHttpAllowList : "all",
               g_configTitleMatch);
}

/* ── Interactive multi-monitor screen capture ─────────────────────────── */

static void EnterScreenCaptureDpiMode(void)
{
    typedef HANDLE (WINAPI *PFN_SetThreadDpiAwarenessContextCompat)(HANDLE);
    HMODULE user32Module = GetModuleHandleW(L"user32.dll");
    PFN_SetThreadDpiAwarenessContextCompat setContext = NULL;

    if (!user32Module || g_capturePreviousDpiContext) return;
    setContext = (PFN_SetThreadDpiAwarenessContextCompat)GetProcAddress(
        user32Module, "SetThreadDpiAwarenessContext");
    if (setContext) {
        g_capturePreviousDpiContext = setContext((HANDLE)(LONG_PTR)-4);
    }
}

static void LeaveScreenCaptureDpiMode(void)
{
    typedef HANDLE (WINAPI *PFN_SetThreadDpiAwarenessContextCompat)(HANDLE);
    HMODULE user32Module;
    PFN_SetThreadDpiAwarenessContextCompat setContext;

    if (!g_capturePreviousDpiContext) return;
    user32Module = GetModuleHandleW(L"user32.dll");
    setContext = user32Module
        ? (PFN_SetThreadDpiAwarenessContextCompat)GetProcAddress(
            user32Module, "SetThreadDpiAwarenessContext")
        : NULL;
    if (setContext) setContext(g_capturePreviousDpiContext);
    g_capturePreviousDpiContext = NULL;
}

static HBITMAP CreateTopDownCaptureBitmap(HDC referenceDc, int width, int height,
                                          DWORD **pixels)
{
    BITMAPINFO bitmapInfo;
    ZeroMemory(&bitmapInfo, sizeof(bitmapInfo));
    bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bitmapInfo.bmiHeader.biWidth = width;
    bitmapInfo.bmiHeader.biHeight = -height;
    bitmapInfo.bmiHeader.biPlanes = 1;
    bitmapInfo.bmiHeader.biBitCount = 32;
    bitmapInfo.bmiHeader.biCompression = BI_RGB;
    return CreateDIBSection(referenceDc, &bitmapInfo, DIB_RGB_COLORS,
                            (void **)pixels, NULL, 0);
}

static void ReleaseScreenCaptureBitmaps(void)
{
    if (g_captureOriginalBitmap) DeleteObject(g_captureOriginalBitmap);
    if (g_captureFrameBitmap) DeleteObject(g_captureFrameBitmap);
    g_captureOriginalBitmap = NULL;
    g_captureFrameBitmap = NULL;
    g_captureOriginalPixels = NULL;
    g_captureWidth = 0;
    g_captureHeight = 0;
}

static BOOL CaptureVirtualDesktopSnapshot(void)
{
    HDC screenDc = NULL;
    HDC memoryDc = NULL;
    HGDIOBJ previousBitmap = NULL;
    DWORD *framePixels = NULL;
    size_t pixelCount;
    BOOL captured = FALSE;

    ReleaseScreenCaptureBitmaps();
    g_captureVirtualX = GetSystemMetrics(SM_XVIRTUALSCREEN);
    g_captureVirtualY = GetSystemMetrics(SM_YVIRTUALSCREEN);
    g_captureWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    g_captureHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    if (g_captureWidth <= 0 || g_captureHeight <= 0 ||
        (size_t)g_captureWidth > (size_t)-1 / (size_t)g_captureHeight) {
        goto cleanup;
    }
    pixelCount = (size_t)g_captureWidth * (size_t)g_captureHeight;
    if (pixelCount > (size_t)-1 / sizeof(DWORD)) goto cleanup;

    screenDc = GetDC(NULL);
    if (!screenDc) goto cleanup;
    memoryDc = CreateCompatibleDC(screenDc);
    if (!memoryDc) goto cleanup;

    g_captureOriginalBitmap = CreateTopDownCaptureBitmap(
        screenDc, g_captureWidth, g_captureHeight, &g_captureOriginalPixels);
    if (!g_captureOriginalBitmap || !g_captureOriginalPixels) {
        goto cleanup;
    }
    g_captureFrameBitmap = CreateTopDownCaptureBitmap(
        screenDc, g_captureWidth, g_captureHeight, &framePixels);
    if (!g_captureFrameBitmap || !framePixels) goto cleanup;

    previousBitmap = SelectObject(memoryDc, g_captureOriginalBitmap);
    if (!BitBlt(memoryDc, 0, 0, g_captureWidth, g_captureHeight,
                screenDc, g_captureVirtualX, g_captureVirtualY,
                SRCCOPY | CAPTUREBLT)) {
        goto cleanup;
    }
    SelectObject(memoryDc, previousBitmap);
    previousBitmap = NULL;

    captured = TRUE;

cleanup:
    if (previousBitmap && memoryDc) SelectObject(memoryDc, previousBitmap);
    if (memoryDc) DeleteDC(memoryDc);
    if (screenDc) ReleaseDC(NULL, screenDc);
    if (!captured) ReleaseScreenCaptureBitmaps();
    return captured;
}

static UINT GetCaptureMonitorDpi(HMONITOR monitor)
{
    typedef HRESULT (WINAPI *PFN_GetDpiForMonitorCompat)(HMONITOR, int,
                                                         UINT *, UINT *);
    HMODULE shcoreModule = LoadLibraryW(L"shcore.dll");
    UINT dpiX = 96;
    UINT dpiY = 96;
    if (shcoreModule) {
        PFN_GetDpiForMonitorCompat getDpiForMonitor =
            (PFN_GetDpiForMonitorCompat)GetProcAddress(
                shcoreModule, "GetDpiForMonitor");
        if (getDpiForMonitor &&
            FAILED(getDpiForMonitor(monitor, 0, &dpiX, &dpiY))) {
            dpiX = 96;
        }
        FreeLibrary(shcoreModule);
    }
    if (dpiX < 72) dpiX = 72;
    if (dpiX > 288) dpiX = 288;
    return dpiX;
}

static int ScaleCaptureUiValue(UINT dpi, int value)
{
    return MulDiv(value, (int)dpi, 96);
}

static BOOL CALLBACK CaptureMonitorEnumProc(HMONITOR monitor, HDC monitorDc,
                                             LPRECT monitorRect, LPARAM context)
{
    CapturePanel *panel;
    UINT dpi;
    int panelWidth;
    int panelHeight;
    int buttonWidth;
    int copyButtonWidth;
    int buttonHeight;
    int buttonGap;
    int separatorGap;
    int panelBottomMargin;
    int panelTopPadding;
    int panelLeft;
    int panelTop;
    int buttonLeft;
    int horizontalPadding;

    (void)monitorDc;
    (void)context;
    if (g_capturePanelCount >= MAX_CAPTURE_MONITORS) return FALSE;

    dpi = GetCaptureMonitorDpi(monitor);
    {
        int monitorWidth = monitorRect->right - monitorRect->left;
        UINT maximumDpi = monitorWidth > 32
            ? (UINT)MulDiv(monitorWidth - 16, 96, CAPTURE_PANEL_WIDTH)
            : 72;
        if (maximumDpi < 72) maximumDpi = 72;
        if (dpi > maximumDpi) dpi = maximumDpi;
    }
    panelWidth = ScaleCaptureUiValue(dpi, CAPTURE_PANEL_WIDTH);
    panelHeight = ScaleCaptureUiValue(dpi, CAPTURE_PANEL_HEIGHT);
    buttonWidth = ScaleCaptureUiValue(dpi, CAPTURE_BUTTON_WIDTH);
    copyButtonWidth = ScaleCaptureUiValue(dpi, CAPTURE_COPY_BUTTON_WIDTH);
    buttonHeight = ScaleCaptureUiValue(dpi, CAPTURE_BUTTON_HEIGHT);
    buttonGap = ScaleCaptureUiValue(dpi, CAPTURE_BUTTON_GAP);
    separatorGap = ScaleCaptureUiValue(dpi, CAPTURE_SEPARATOR_GAP);
    panelBottomMargin = ScaleCaptureUiValue(
        dpi, CAPTURE_PANEL_BOTTOM_MARGIN);
    panelTopPadding = ScaleCaptureUiValue(dpi, 10);
    horizontalPadding =
        (panelWidth - (buttonWidth * (CAPTURE_TOOL_COUNT - 1)) -
         copyButtonWidth -
         (buttonGap * (CAPTURE_TOOL_COUNT - 1)) - separatorGap) / 2;

    panelLeft = monitorRect->left - g_captureVirtualX +
                ((monitorRect->right - monitorRect->left) -
                 panelWidth) / 2;
    panelTop = monitorRect->bottom - g_captureVirtualY -
               panelBottomMargin - panelHeight;
    if (panelTop < monitorRect->top - g_captureVirtualY + 8) {
        panelTop = monitorRect->top - g_captureVirtualY + 8;
    }

    panel = &g_capturePanels[g_capturePanelCount++];
    panel->dpi = dpi;
    SetRect(&panel->monitorRect,
            monitorRect->left - g_captureVirtualX,
            monitorRect->top - g_captureVirtualY,
            monitorRect->right - g_captureVirtualX,
            monitorRect->bottom - g_captureVirtualY);
    SetRect(&panel->panelRect, panelLeft, panelTop,
            panelLeft + panelWidth,
            panelTop + panelHeight);
    buttonLeft = panelLeft + horizontalPadding;
    for (int tool = 0; tool < CAPTURE_TOOL_COUNT; tool++) {
        int width = tool == CAPTURE_TOOL_COPY ? copyButtonWidth : buttonWidth;
        /* Cancel sits apart from the action tools, past a visual divider. */
        if (tool == CAPTURE_TOOL_CANCEL) buttonLeft += separatorGap;
        SetRect(&panel->buttonRects[tool],
                buttonLeft, panelTop + panelTopPadding,
                buttonLeft + width,
                panelTop + panelTopPadding + buttonHeight);
        buttonLeft += width + buttonGap;
    }
    return TRUE;
}

static void BuildCapturePanels(void)
{
    g_capturePanelCount = 0;
    EnumDisplayMonitors(NULL, NULL, CaptureMonitorEnumProc, 0);
    if (g_capturePanelCount == 0) {
        RECT fallback = {
            g_captureVirtualX,
            g_captureVirtualY,
            g_captureVirtualX + g_captureWidth,
            g_captureVirtualY + g_captureHeight
        };
        CaptureMonitorEnumProc(NULL, NULL, &fallback, 0);
    }
}

static POINT GetCaptureCursorPoint(HWND hwnd)
{
    POINT point = {0};
    GetCursorPos(&point);
    ScreenToClient(hwnd, &point);
    if (point.x < 0) point.x = 0;
    if (point.y < 0) point.y = 0;
    if (point.x > g_captureWidth) point.x = g_captureWidth;
    if (point.y > g_captureHeight) point.y = g_captureHeight;
    return point;
}

static RECT NormalizeCaptureRect(POINT start, POINT end)
{
    RECT rect;
    rect.left = start.x < end.x ? start.x : end.x;
    rect.top = start.y < end.y ? start.y : end.y;
    rect.right = start.x > end.x ? start.x : end.x;
    rect.bottom = start.y > end.y ? start.y : end.y;
    if (rect.left < 0) rect.left = 0;
    if (rect.top < 0) rect.top = 0;
    if (rect.right > g_captureWidth) rect.right = g_captureWidth;
    if (rect.bottom > g_captureHeight) rect.bottom = g_captureHeight;
    return rect;
}

static BOOL IsCaptureSelectionValid(const RECT *selection)
{
    return selection &&
           selection->right - selection->left >= 2 &&
           selection->bottom - selection->top >= 2;
}

static BOOL AppendCaptureSelection(const RECT *selection)
{
    RECT *resized;
    size_t newCapacity;

    if (!IsCaptureSelectionValid(selection)) return FALSE;
    if (g_captureSelectionCount == g_captureSelectionCapacity) {
        newCapacity = g_captureSelectionCapacity
            ? g_captureSelectionCapacity * 2 : 8;
        if (newCapacity < g_captureSelectionCapacity ||
            newCapacity > (size_t)-1 / sizeof(RECT)) {
            return FALSE;
        }
        resized = (RECT *)realloc(
            g_captureSelections, newCapacity * sizeof(RECT));
        if (!resized) return FALSE;
        g_captureSelections = resized;
        g_captureSelectionCapacity = newCapacity;
    }
    g_captureSelections[g_captureSelectionCount++] = *selection;
    return TRUE;
}

static void ClearCaptureSelections(void)
{
    g_captureSelectionCount = 0;
}

static void ReleaseCaptureSelections(void)
{
    free(g_captureSelections);
    g_captureSelections = NULL;
    g_captureSelectionCount = 0;
    g_captureSelectionCapacity = 0;
}

static size_t GetDisplayedCaptureSelectionCount(void)
{
    return g_captureSelectionCount + (g_captureDragging ? 1u : 0u);
}

static BOOL GetDisplayedCaptureSelection(size_t index, RECT *selection)
{
    if (!selection) return FALSE;
    if (index < g_captureSelectionCount) {
        *selection = g_captureSelections[index];
        return TRUE;
    }
    if (g_captureDragging && index == g_captureSelectionCount) {
        /* Shown even when degenerate (0 × 0) so the dimension pill and the
           Copy Selection label stay truthful for the whole drag. */
        *selection = NormalizeCaptureRect(
            g_captureDragStart, g_captureDragCurrent);
        return TRUE;
    }
    return FALSE;
}

/* TRUE while any stored selection or an in-progress drag (of any size)
   exists; the Copy tool then acts on the selection instead of the full
   desktop. A drag released at an invalid size adds nothing, after which
   this reverts to FALSE. */
static BOOL CaptureHasSelection(void)
{
    return g_captureSelectionCount > 0 || g_captureDragging;
}

static BOOL GetCaptureSelectionBounds(RECT *bounds)
{
    if (!bounds || g_captureSelectionCount == 0) return FALSE;
    *bounds = g_captureSelections[0];
    for (size_t index = 1; index < g_captureSelectionCount; index++) {
        const RECT *selection = &g_captureSelections[index];
        if (selection->left < bounds->left) bounds->left = selection->left;
        if (selection->top < bounds->top) bounds->top = selection->top;
        if (selection->right > bounds->right) bounds->right = selection->right;
        if (selection->bottom > bounds->bottom) bounds->bottom = selection->bottom;
    }
    return TRUE;
}

static void AlphaFillRect(HDC destinationDc, const RECT *rect,
                          COLORREF color, BYTE opacity)
{
    int width = rect->right - rect->left;
    int height = rect->bottom - rect->top;
    HDC sourceDc;
    HBITMAP sourceBitmap;
    HGDIOBJ oldBitmap;
    BITMAPINFO bitmapInfo;
    DWORD *pixel = NULL;
    BLENDFUNCTION blend = {AC_SRC_OVER, 0, opacity, 0};

    if (width <= 0 || height <= 0) return;
    sourceDc = CreateCompatibleDC(destinationDc);
    if (!sourceDc) return;
    ZeroMemory(&bitmapInfo, sizeof(bitmapInfo));
    bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bitmapInfo.bmiHeader.biWidth = 1;
    bitmapInfo.bmiHeader.biHeight = -1;
    bitmapInfo.bmiHeader.biPlanes = 1;
    bitmapInfo.bmiHeader.biBitCount = 32;
    bitmapInfo.bmiHeader.biCompression = BI_RGB;
    sourceBitmap = CreateDIBSection(destinationDc, &bitmapInfo, DIB_RGB_COLORS,
                                    (void **)&pixel, NULL, 0);
    if (!sourceBitmap || !pixel) {
        if (sourceBitmap) DeleteObject(sourceBitmap);
        DeleteDC(sourceDc);
        return;
    }
    *pixel = (DWORD)GetBValue(color) |
             ((DWORD)GetGValue(color) << 8) |
             ((DWORD)GetRValue(color) << 16);
    oldBitmap = SelectObject(sourceDc, sourceBitmap);
    AlphaBlend(destinationDc, rect->left, rect->top, width, height,
               sourceDc, 0, 0, 1, 1, blend);
    SelectObject(sourceDc, oldBitmap);
    DeleteObject(sourceBitmap);
    DeleteDC(sourceDc);
}

enum {
    CAPTURE_GDIP_COMPOSITING_HIGH_QUALITY = 2,
    CAPTURE_GDIP_SMOOTHING_ANTIALIAS_8X8 = 5,
    CAPTURE_GDIP_PIXEL_OFFSET_HIGH_QUALITY = 2,
    CAPTURE_GDIP_TEXT_ANTIALIAS_GRID_FIT = 3,
    CAPTURE_GDIP_UNIT_PIXEL = 2,
    CAPTURE_GDIP_FONT_REGULAR = 0,
    CAPTURE_GDIP_FONT_BOLD = 1,
    CAPTURE_GDIP_FILL_ALTERNATE = 0,
    CAPTURE_GDIP_LINE_CAP_ROUND = 2,
    CAPTURE_GDIP_LINE_JOIN_ROUND = 2,
    CAPTURE_GDIP_DASH_CAP_ROUND = 2,
    CAPTURE_GDIP_DASH_STYLE_DASH = 1,
    CAPTURE_GDIP_STRING_ALIGN_CENTER = 1,
    CAPTURE_GDIP_STRING_NO_WRAP = 0x1000,
    CAPTURE_GDIP_GRADIENT_VERTICAL = 1,
    CAPTURE_GDIP_WRAP_TILE_FLIP_XY = 3
};

static float ScaleCaptureUiFloat(UINT dpi, float value)
{
    return value * (float)dpi / 96.0f;
}

static DWORD CaptureArgb(BYTE opacity, COLORREF color)
{
    return ((DWORD)opacity << 24) |
           ((DWORD)GetRValue(color) << 16) |
           ((DWORD)GetGValue(color) << 8) |
           (DWORD)GetBValue(color);
}

static GpPath *CreateCaptureRoundedRectPath(const RECT *rect, float radius)
{
    GpPath *path = NULL;
    float left = (float)rect->left;
    float top = (float)rect->top;
    float width = (float)(rect->right - rect->left);
    float height = (float)(rect->bottom - rect->top);
    float diameter;

    if (width <= 0.0f || height <= 0.0f ||
        GdipCreatePath(CAPTURE_GDIP_FILL_ALTERNATE, &path) != 0 || !path) {
        return NULL;
    }
    if (radius < 0.0f) radius = 0.0f;
    if (radius > width / 2.0f) radius = width / 2.0f;
    if (radius > height / 2.0f) radius = height / 2.0f;
    diameter = radius * 2.0f;

    if (diameter < 1.0f) {
        GdipAddPathLine(path, left, top, left + width, top);
        GdipAddPathLine(path, left + width, top,
                        left + width, top + height);
        GdipAddPathLine(path, left + width, top + height,
                        left, top + height);
        GdipAddPathLine(path, left, top + height, left, top);
    } else {
        GdipAddPathArc(path, left, top, diameter, diameter, 180.0f, 90.0f);
        GdipAddPathArc(path, left + width - diameter, top,
                       diameter, diameter, 270.0f, 90.0f);
        GdipAddPathArc(path, left + width - diameter,
                       top + height - diameter,
                       diameter, diameter, 0.0f, 90.0f);
        GdipAddPathArc(path, left, top + height - diameter,
                       diameter, diameter, 90.0f, 90.0f);
    }
    GdipClosePathFigure(path);
    return path;
}

static void FillCaptureRoundedRect(GpGraphics *graphics, const RECT *rect,
                                   COLORREF color, BYTE opacity, float radius)
{
    GpPath *path;
    GpSolidFill *brush = NULL;

    if (!graphics) return;
    path = CreateCaptureRoundedRectPath(rect, radius);
    if (!path) return;
    if (GdipCreateSolidFill(CaptureArgb(opacity, color), &brush) == 0 && brush) {
        GdipFillPath(graphics, (GpBrush *)brush, path);
        GdipDeleteBrush((GpBrush *)brush);
    }
    GdipDeletePath(path);
}

static void StrokeCaptureRoundedRect(GpGraphics *graphics, const RECT *rect,
                                     COLORREF color, BYTE opacity,
                                     float width, float radius)
{
    GpPath *path;
    GpPen *pen = NULL;

    if (!graphics) return;
    path = CreateCaptureRoundedRectPath(rect, radius);
    if (!path) return;
    if (GdipCreatePen1(CaptureArgb(opacity, color), width,
                       CAPTURE_GDIP_UNIT_PIXEL, &pen) == 0 && pen) {
        GdipSetPenLineJoin(pen, CAPTURE_GDIP_LINE_JOIN_ROUND);
        GdipDrawPath(graphics, pen, path);
        GdipDeletePen(pen);
    }
    GdipDeletePath(path);
}

static void FillCaptureRoundedRectGradient(GpGraphics *graphics,
                                           const RECT *rect,
                                           DWORD topColor, DWORD bottomColor,
                                           float radius)
{
    GpPath *path;
    GpLineGradient *brush = NULL;
    CaptureGpRectF brushRect;

    if (!graphics) return;
    path = CreateCaptureRoundedRectPath(rect, radius);
    if (!path) return;
    /* Oversize the brush by a pixel so the gradient cannot wrap at the
       antialiased path edge. */
    brushRect.X = (float)rect->left - 1.0f;
    brushRect.Y = (float)rect->top - 1.0f;
    brushRect.Width = (float)(rect->right - rect->left) + 2.0f;
    brushRect.Height = (float)(rect->bottom - rect->top) + 2.0f;
    if (GdipCreateLineBrushFromRect(&brushRect, topColor, bottomColor,
                                    CAPTURE_GDIP_GRADIENT_VERTICAL,
                                    CAPTURE_GDIP_WRAP_TILE_FLIP_XY,
                                    &brush) == 0 && brush) {
        GdipFillPath(graphics, (GpBrush *)brush, path);
        GdipDeleteBrush((GpBrush *)brush);
    } else {
        GpSolidFill *fallback = NULL;
        if (GdipCreateSolidFill(bottomColor, &fallback) == 0 && fallback) {
            GdipFillPath(graphics, (GpBrush *)fallback, path);
            GdipDeleteBrush((GpBrush *)fallback);
        }
    }
    GdipDeletePath(path);
}

/* Layered low-alpha fills approximate a blurred drop shadow; the overlapping
   rings accumulate toward the shape so the falloff reads as soft. */
static void DrawCaptureSoftShadow(GpGraphics *graphics, const RECT *rect,
                                  float radius, UINT dpi, int layerCount,
                                  float spreadStep, int dropDistance,
                                  BYTE layerAlpha)
{
    float step = ScaleCaptureUiFloat(dpi, spreadStep);
    int offsetY = ScaleCaptureUiValue(dpi, dropDistance);

    if (step < 1.0f) step = 1.0f;
    for (int layer = layerCount; layer >= 1; layer--) {
        RECT layerRect = *rect;
        int grow = (int)(step * (float)layer + 0.5f);
        InflateRect(&layerRect, grow, grow);
        OffsetRect(&layerRect, 0, offsetY);
        FillCaptureRoundedRect(graphics, &layerRect, RGB(0, 0, 0), layerAlpha,
                               radius + (float)grow);
    }
}

static GpFont *CreateCaptureFont(float pixelSize)
{
    GpFontFamily *family = NULL;
    GpFont *font = NULL;
    int style = CAPTURE_GDIP_FONT_REGULAR;

    if (GdipCreateFontFamilyFromName(L"Segoe UI Semibold", NULL, &family) != 0 ||
        !family) {
        family = NULL;
        style = CAPTURE_GDIP_FONT_BOLD;
        if (GdipCreateFontFamilyFromName(L"Segoe UI", NULL, &family) != 0 ||
            !family) {
            family = NULL;
            GdipGetGenericFontFamilySansSerif(&family);
        }
    }
    if (family) {
        if (GdipCreateFont(family, pixelSize, style,
                           CAPTURE_GDIP_UNIT_PIXEL, &font) != 0) {
            font = NULL;
        }
        GdipDeleteFontFamily(family);
    }
    return font;
}

static GpStringFormat *CreateCaptureCenteredTextFormat(void)
{
    GpStringFormat *format = NULL;
    if (GdipCreateStringFormat(0, 0, &format) != 0 || !format) return NULL;
    GdipSetStringFormatFlags(format, CAPTURE_GDIP_STRING_NO_WRAP);
    GdipSetStringFormatAlign(format, CAPTURE_GDIP_STRING_ALIGN_CENTER);
    GdipSetStringFormatLineAlign(format, CAPTURE_GDIP_STRING_ALIGN_CENTER);
    return format;
}

/* Font/format cache for the overlay. The GDI+ font-family lookup is
   expensive and the toolbars repaint on every mouse move while a box is
   drawn near them, so fonts are created once per size (one per monitor
   DPI) and reused until the overlay closes. */
typedef struct {
    float pixelSize;
    GpFont *font;
} CaptureFontCacheEntry;

static CaptureFontCacheEntry g_captureFontCache[MAX_CAPTURE_MONITORS];
static int g_captureFontCacheCount = 0;
static GpStringFormat *g_captureTextFormatCache = NULL;

static GpFont *GetCaptureFont(float pixelSize)
{
    GpFont *font;
    for (int i = 0; i < g_captureFontCacheCount; i++) {
        if (g_captureFontCache[i].pixelSize == pixelSize) {
            return g_captureFontCache[i].font;
        }
    }
    font = CreateCaptureFont(pixelSize);
    if (font) {
        if (g_captureFontCacheCount == MAX_CAPTURE_MONITORS) {
            GdipDeleteFont(g_captureFontCache[0].font);
            memmove(&g_captureFontCache[0], &g_captureFontCache[1],
                    (MAX_CAPTURE_MONITORS - 1) *
                        sizeof(g_captureFontCache[0]));
            g_captureFontCacheCount--;
        }
        g_captureFontCache[g_captureFontCacheCount].pixelSize = pixelSize;
        g_captureFontCache[g_captureFontCacheCount].font = font;
        g_captureFontCacheCount++;
    }
    return font;
}

static GpStringFormat *GetCaptureTextFormat(void)
{
    if (!g_captureTextFormatCache) {
        g_captureTextFormatCache = CreateCaptureCenteredTextFormat();
    }
    return g_captureTextFormatCache;
}

/* Screen-DC graphics used to measure label text outside WM_PAINT, so
   layout decisions (and hit-testing) can depend on the rendered width. */
static GpGraphics *g_captureMeasureGraphics = NULL;
static HDC g_captureMeasureDc = NULL;

static GpGraphics *GetCaptureMeasureGraphics(void)
{
    if (!g_captureMeasureGraphics) {
        g_captureMeasureDc = GetDC(NULL);
        if (g_captureMeasureDc &&
            GdipCreateFromHDC(g_captureMeasureDc,
                              &g_captureMeasureGraphics) != 0) {
            g_captureMeasureGraphics = NULL;
        }
    }
    return g_captureMeasureGraphics;
}

static void ReleaseCaptureDrawingCache(void)
{
    for (int i = 0; i < g_captureFontCacheCount; i++) {
        GdipDeleteFont(g_captureFontCache[i].font);
    }
    g_captureFontCacheCount = 0;
    if (g_captureTextFormatCache) {
        GdipDeleteStringFormat(g_captureTextFormatCache);
        g_captureTextFormatCache = NULL;
    }
    if (g_captureMeasureGraphics) {
        GdipDeleteGraphics(g_captureMeasureGraphics);
        g_captureMeasureGraphics = NULL;
    }
    if (g_captureMeasureDc) {
        ReleaseDC(NULL, g_captureMeasureDc);
        g_captureMeasureDc = NULL;
    }
}

static void DrawCaptureCenteredText(GpGraphics *graphics, const WCHAR *text,
                                    const RECT *rect, COLORREF color,
                                    BYTE alpha, GpFont *font,
                                    GpStringFormat *format)
{
    GpSolidFill *brush = NULL;
    CaptureGpRectF layout;

    if (!graphics || !text || !font || !format) return;
    if (GdipCreateSolidFill(CaptureArgb(alpha, color), &brush) != 0 || !brush) {
        return;
    }
    layout.X = (float)rect->left;
    layout.Y = (float)rect->top;
    layout.Width = (float)(rect->right - rect->left);
    layout.Height = (float)(rect->bottom - rect->top);
    GdipDrawString(graphics, text, -1, font, &layout, format,
                   (GpBrush *)brush);
    GdipDeleteBrush((GpBrush *)brush);
}

static void DrawCaptureToolIcon(GpGraphics *graphics, int tool,
                                const RECT *buttonRect, COLORREF color,
                                BYTE alpha, UINT dpi)
{
    float centerX = ((float)buttonRect->left + (float)buttonRect->right) / 2.0f;
    float top = (float)buttonRect->top + ScaleCaptureUiFloat(dpi, 7.0f);
    float inner = ScaleCaptureUiFloat(dpi, 10.0f);
    float middle = ScaleCaptureUiFloat(dpi, 4.0f);
    float shortStep = ScaleCaptureUiFloat(dpi, 6.0f);
    float iconHeight = ScaleCaptureUiFloat(dpi, 18.0f);
    float penWidth = ScaleCaptureUiFloat(dpi, 1.6f);
    GpPen *pen = NULL;

    if (penWidth < 1.0f) penWidth = 1.0f;
    if (GdipCreatePen1(CaptureArgb(alpha, color), penWidth,
                       CAPTURE_GDIP_UNIT_PIXEL, &pen) != 0 || !pen) {
        return;
    }
    GdipSetPenStartCap(pen, CAPTURE_GDIP_LINE_CAP_ROUND);
    GdipSetPenEndCap(pen, CAPTURE_GDIP_LINE_CAP_ROUND);
    GdipSetPenLineJoin(pen, CAPTURE_GDIP_LINE_JOIN_ROUND);

    if (tool == CAPTURE_TOOL_CLIP) {
        GdipDrawLine(graphics, pen, centerX - inner, top + shortStep,
                     centerX - inner, top);
        GdipDrawLine(graphics, pen, centerX - inner, top,
                     centerX - middle, top);
        GdipDrawLine(graphics, pen, centerX + middle, top,
                     centerX + inner, top);
        GdipDrawLine(graphics, pen, centerX + inner, top,
                     centerX + inner, top + shortStep);
        GdipDrawLine(graphics, pen, centerX + inner,
                     top + iconHeight - shortStep,
                     centerX + inner, top + iconHeight);
        GdipDrawLine(graphics, pen, centerX + inner, top + iconHeight,
                     centerX + middle, top + iconHeight);
        GdipDrawLine(graphics, pen, centerX - middle, top + iconHeight,
                     centerX - inner, top + iconHeight);
        GdipDrawLine(graphics, pen, centerX - inner, top + iconHeight,
                     centerX - inner, top + iconHeight - shortStep);
    } else if (tool == CAPTURE_TOOL_COPY) {
        int overlap = ScaleCaptureUiValue(dpi, 5);
        RECT backRect = {
            (LONG)(centerX - inner), (LONG)top,
            (LONG)(centerX + (float)overlap),
            (LONG)(top + ScaleCaptureUiFloat(dpi, 13.0f))
        };
        RECT frontRect = {
            (LONG)(centerX - (float)overlap), (LONG)(top + (float)overlap),
            (LONG)(centerX + inner), (LONG)(top + iconHeight)
        };
        GpPath *backPath = CreateCaptureRoundedRectPath(
            &backRect, ScaleCaptureUiFloat(dpi, 2.5f));
        GpPath *frontPath = CreateCaptureRoundedRectPath(
            &frontRect, ScaleCaptureUiFloat(dpi, 2.5f));
        if (backPath) {
            GdipDrawPath(graphics, pen, backPath);
            GdipDeletePath(backPath);
        }
        if (frontPath) {
            GdipDrawPath(graphics, pen, frontPath);
            GdipDeletePath(frontPath);
        }
    } else {
        float cross = ScaleCaptureUiFloat(dpi, 7.0f);
        float centerY = top + iconHeight / 2.0f;
        GdipDrawLine(graphics, pen, centerX - cross, centerY - cross,
                     centerX + cross, centerY + cross);
        GdipDrawLine(graphics, pen, centerX + cross, centerY - cross,
                     centerX - cross, centerY + cross);
    }
    GdipDeletePen(pen);
}

static BYTE ScaleCaptureAlpha(BYTE alpha, int percent)
{
    return (BYTE)(((int)alpha * percent) / 100);
}

static void DrawCapturePanel(GpGraphics *graphics, int panelIndex)
{
    const WCHAR *labels[CAPTURE_TOOL_COUNT];
    CapturePanel *panel = &g_capturePanels[panelIndex];
    float panelRadius = ScaleCaptureUiFloat(panel->dpi, 14.0f);
    float buttonRadius = ScaleCaptureUiFloat(panel->dpi, 8.0f);
    float hairline = ScaleCaptureUiFloat(panel->dpi, 1.0f);
    float borderWidth = ScaleCaptureUiFloat(panel->dpi, 0.65f);
    /* Fade the whole toolbar out of the way while a box is being drawn. */
    int opacity = g_captureDragging ? 50 : 100;
    GpFont *font;
    GpStringFormat *format;

    /* The panel under the drag cursor disappears completely so it never
       obstructs the box being drawn. */
    if (g_captureDragging && panelIndex == g_captureVanishedPanel) return;

    labels[CAPTURE_TOOL_CLIP] = L"Clip";
    labels[CAPTURE_TOOL_COPY] = CaptureHasSelection()
        ? L"Copy Selection" : L"Copy All";
    labels[CAPTURE_TOOL_CANCEL] = L"Cancel";

    if (hairline < 1.0f) hairline = 1.0f;
    if (borderWidth < 0.6f) borderWidth = 0.6f;
    DrawCaptureSoftShadow(graphics, &panel->panelRect, panelRadius,
                          panel->dpi, 7, 1.9f, 3,
                          ScaleCaptureAlpha(9, opacity));
    FillCaptureRoundedRectGradient(
        graphics, &panel->panelRect,
        CaptureArgb(ScaleCaptureAlpha(246, opacity), RGB(43, 49, 60)),
        CaptureArgb(ScaleCaptureAlpha(246, opacity), RGB(24, 28, 35)),
        panelRadius);
    StrokeCaptureRoundedRect(graphics, &panel->panelRect,
                             RGB(255, 255, 255),
                             ScaleCaptureAlpha(38, opacity),
                             borderWidth, panelRadius);

    /* Divider that sets the destructive Cancel apart from the action tools. */
    {
        const RECT *copyRect = &panel->buttonRects[CAPTURE_TOOL_COPY];
        const RECT *cancelRect = &panel->buttonRects[CAPTURE_TOOL_CANCEL];
        float separatorX =
            ((float)copyRect->right + (float)cancelRect->left) / 2.0f;
        float inset = ScaleCaptureUiFloat(panel->dpi, 9.0f);
        GpPen *separatorPen = NULL;
        if (GdipCreatePen1(CaptureArgb(ScaleCaptureAlpha(36, opacity),
                                       RGB(255, 255, 255)), hairline,
                           CAPTURE_GDIP_UNIT_PIXEL, &separatorPen) == 0 &&
            separatorPen) {
            GdipSetPenStartCap(separatorPen, CAPTURE_GDIP_LINE_CAP_ROUND);
            GdipSetPenEndCap(separatorPen, CAPTURE_GDIP_LINE_CAP_ROUND);
            GdipDrawLine(graphics, separatorPen,
                         separatorX, (float)copyRect->top + inset,
                         separatorX, (float)copyRect->bottom - inset);
            GdipDeletePen(separatorPen);
        }
    }

    font = GetCaptureFont(ScaleCaptureUiFloat(panel->dpi, 12.0f));
    format = GetCaptureTextFormat();
    for (int tool = 0; tool < CAPTURE_TOOL_COUNT; tool++) {
        RECT buttonRect = panel->buttonRects[tool];
        RECT labelRect = buttonRect;
        BOOL isCancel = tool == CAPTURE_TOOL_CANCEL;
        BOOL selected = tool == g_captureSelectedTool;
        /* While drawing a box the buttons ignore hover and press hints. */
        BOOL hovered = !g_captureDragging &&
                       panelIndex == g_captureHoveredPanel &&
                       tool == g_captureHoveredTool;
        BOOL pressed = !g_captureDragging &&
                       panelIndex == g_capturePressedPanel &&
                       tool == g_capturePressedTool;
        BOOL active = selected || hovered || pressed;
        COLORREF accent = isCancel
            ? RGB(255, 138, 138) : RGB(126, 208, 255);
        COLORREF labelColor = selected
            ? RGB(255, 255, 255)
            : (active ? RGB(238, 244, 250) : RGB(187, 197, 210));

        if (selected) {
            /* Accent-tinted pill with a soft ring instead of a heavy fill. */
            FillCaptureRoundedRect(graphics, &buttonRect,
                                   isCancel ? RGB(232, 82, 82)
                                            : RGB(64, 156, 210),
                                   ScaleCaptureAlpha(pressed ? 84 : 58,
                                                     opacity),
                                   buttonRadius);
            StrokeCaptureRoundedRect(graphics, &buttonRect, accent,
                                     ScaleCaptureAlpha(pressed ? 165 : 130,
                                                       opacity),
                                     hairline, buttonRadius);
            if (hovered && !pressed) {
                FillCaptureRoundedRect(graphics, &buttonRect,
                                       RGB(255, 255, 255),
                                       ScaleCaptureAlpha(14, opacity),
                                       buttonRadius);
            }
        } else if (pressed) {
            FillCaptureRoundedRect(graphics, &buttonRect,
                                   isCancel ? RGB(226, 74, 74)
                                            : RGB(255, 255, 255),
                                   ScaleCaptureAlpha(isCancel ? 64 : 15,
                                                     opacity),
                                   buttonRadius);
        } else if (hovered) {
            FillCaptureRoundedRect(graphics, &buttonRect,
                                   isCancel ? RGB(238, 90, 90)
                                            : RGB(255, 255, 255),
                                   ScaleCaptureAlpha(isCancel ? 46 : 24,
                                                     opacity),
                                   buttonRadius);
        }
        DrawCaptureToolIcon(graphics, tool, &buttonRect,
                            active ? accent : RGB(206, 215, 227),
                            ScaleCaptureAlpha(255, opacity), panel->dpi);
        labelRect.top = buttonRect.top + ScaleCaptureUiValue(panel->dpi, 31);
        DrawCaptureCenteredText(graphics, labels[tool], &labelRect,
                                labelColor, ScaleCaptureAlpha(255, opacity),
                                font, format);
    }
}

static RECT GetCapturePanelPaintRect(int panelIndex)
{
    /* Wide enough to include the layered drop shadow drawn around the
       panel (7 layers x 1.9px spread, dropped 3px). */
    RECT area = g_capturePanels[panelIndex].panelRect;
    InflateRect(&area,
                ScaleCaptureUiValue(g_capturePanels[panelIndex].dpi, 16),
                ScaleCaptureUiValue(g_capturePanels[panelIndex].dpi, 20));
    return area;
}

static UINT GetCaptureDpiAtPoint(POINT point)
{
    for (int panel = 0; panel < g_capturePanelCount; panel++) {
        if (PtInRect(&g_capturePanels[panel].monitorRect, point)) {
            return g_capturePanels[panel].dpi;
        }
    }
    return g_capturePanelCount > 0 ? g_capturePanels[0].dpi : 96;
}

static void FormatCaptureDimensions(const RECT *selection, WCHAR *buffer,
                                    size_t bufferChars)
{
    swprintf(buffer, bufferChars, L"%d × %d",
             selection->right - selection->left,
             selection->bottom - selection->top);
}

static UINT GetCaptureSelectionLabelRect(const RECT *selection,
                                         RECT *labelRect)
{
    POINT anchor = {selection->left, selection->top};
    UINT dpi = GetCaptureDpiAtPoint(anchor);
    int labelWidth = ScaleCaptureUiValue(dpi, 118);
    int labelHeight = ScaleCaptureUiValue(dpi, 26);
    int gap = ScaleCaptureUiValue(dpi, 8);
    BOOL inside = selection->top < labelHeight + gap * 2;

    /* The pill hugs its text; measured here (not at paint time) so the
       remove pill's floor position and hit-testing stay in sync with the
       rendered width. Falls back to the maximum width if measuring fails. */
    {
        GpGraphics *measure = GetCaptureMeasureGraphics();
        GpFont *font = GetCaptureFont(ScaleCaptureUiFloat(dpi, 12.0f));
        GpStringFormat *format = GetCaptureTextFormat();
        WCHAR dimensions[64];
        CaptureGpRectF layout = {0.0f, 0.0f, 512.0f, 64.0f};
        CaptureGpRectF bounds = {0.0f, 0.0f, 0.0f, 0.0f};
        int fitted = 0;
        int lines = 0;
        FormatCaptureDimensions(selection, dimensions, 64);
        if (measure && font && format &&
            GdipMeasureString(measure, dimensions, -1, font, &layout,
                              format, &bounds, &fitted, &lines) == 0 &&
            bounds.Width > 0.0f) {
            int snugWidth = (int)(bounds.Width + 0.5f) +
                            ScaleCaptureUiValue(dpi, 22);
            int minimumWidth = ScaleCaptureUiValue(dpi, 54);
            if (snugWidth < minimumWidth) snugWidth = minimumWidth;
            if (snugWidth < labelWidth) labelWidth = snugWidth;
        }
    }

    /* Inside the box (near the top of the desktop) the pill keeps the same
       margin on the left as on the top; above the box it aligns flush with
       the left edge. */
    labelRect->left = inside ? selection->left + gap : selection->left;
    labelRect->top = inside ? selection->top + gap
                            : selection->top - labelHeight - gap;
    labelRect->right = labelRect->left + labelWidth;
    labelRect->bottom = labelRect->top + labelHeight;
    if (labelRect->right > g_captureWidth) {
        OffsetRect(labelRect, g_captureWidth - labelRect->right, 0);
    }
    if (labelRect->left < 0) OffsetRect(labelRect, -labelRect->left, 0);
    if (labelRect->bottom > g_captureHeight) {
        OffsetRect(labelRect, 0, g_captureHeight - labelRect->bottom);
    }
    if (labelRect->top < 0) OffsetRect(labelRect, 0, -labelRect->top);
    return dpi;
}

/* Round remove pill at a selection's top-right corner, mirroring the
   dimension pill: above the box flush with the right edge, or inside the
   box with equal top/right margins when there is no room above. */
static UINT GetCaptureSelectionRemoveRect(const RECT *selection,
                                          RECT *pillRect)
{
    POINT anchor = {selection->right, selection->top};
    UINT dpi = GetCaptureDpiAtPoint(anchor);
    int diameter = ScaleCaptureUiValue(dpi, 26);
    int gap = ScaleCaptureUiValue(dpi, 8);
    BOOL inside = selection->top < diameter + gap * 2;
    RECT labelRect;

    pillRect->right = inside ? selection->right - gap : selection->right;
    pillRect->top = inside ? selection->top + gap
                           : selection->top - diameter - gap;
    pillRect->left = pillRect->right - diameter;
    pillRect->bottom = pillRect->top + diameter;
    /* A narrow box must not push the pill over the dimension label; the
       worst case is [w × h][gap][✕] using the standard pill gap. */
    GetCaptureSelectionLabelRect(selection, &labelRect);
    if (pillRect->left < labelRect.right + gap) {
        OffsetRect(pillRect, labelRect.right + gap - pillRect->left, 0);
    }
    if (pillRect->right > g_captureWidth) {
        OffsetRect(pillRect, g_captureWidth - pillRect->right, 0);
    }
    if (pillRect->left < 0) OffsetRect(pillRect, -pillRect->left, 0);
    if (pillRect->bottom > g_captureHeight) {
        OffsetRect(pillRect, 0, g_captureHeight - pillRect->bottom);
    }
    if (pillRect->top < 0) OffsetRect(pillRect, 0, -pillRect->top);
    return dpi;
}

static void DrawCaptureSelectionEdges(GpGraphics *graphics, GpPen *pen,
                                      const RECT *selection, float inset)
{
    float left = (float)selection->left + inset;
    float top = (float)selection->top + inset;
    float right = (float)selection->right - inset;
    float bottom = (float)selection->bottom - inset;
    if (right <= left || bottom <= top) return;
    /* Draw the sides independently so the dash phase stays stable as the
       selection grows instead of crawling around the perimeter. */
    GdipDrawLine(graphics, pen, left, top, right, top);
    GdipDrawLine(graphics, pen, right, top, right, bottom);
    GdipDrawLine(graphics, pen, right, bottom, left, bottom);
    GdipDrawLine(graphics, pen, left, bottom, left, top);
}

/* removeState: -1 = no remove pill (in-progress drag), 0 = pill shown,
   1 = pill shown hovered. */
static void DrawCaptureSelection(GpGraphics *graphics, const RECT *selection,
                                 int removeState)
{
    WCHAR dimensions[64];
    RECT labelRect;
    UINT dpi = GetCaptureSelectionLabelRect(selection, &labelRect);
    float outlineWidth = ScaleCaptureUiFloat(dpi, 1.0f);
    float shadowWidth;
    GpPen *shadowPen = NULL;
    GpPen *outlinePen = NULL;
    GpFont *font;
    GpStringFormat *format;

    if (outlineWidth < 0.85f) outlineWidth = 0.85f;
    shadowWidth = outlineWidth + ScaleCaptureUiFloat(dpi, 1.0f);
    if (GdipCreatePen1(CaptureArgb(125, RGB(0, 0, 0)), shadowWidth,
                       CAPTURE_GDIP_UNIT_PIXEL, &shadowPen) == 0 && shadowPen) {
        GdipSetPenStartCap(shadowPen, CAPTURE_GDIP_LINE_CAP_ROUND);
        GdipSetPenEndCap(shadowPen, CAPTURE_GDIP_LINE_CAP_ROUND);
        DrawCaptureSelectionEdges(graphics, shadowPen, selection,
                                  shadowWidth / 2.0f);
        GdipDeletePen(shadowPen);
    }
    if (GdipCreatePen1(CaptureArgb(245, RGB(226, 247, 255)), outlineWidth,
                       CAPTURE_GDIP_UNIT_PIXEL, &outlinePen) == 0 && outlinePen) {
        GdipSetPenStartCap(outlinePen, CAPTURE_GDIP_LINE_CAP_ROUND);
        GdipSetPenEndCap(outlinePen, CAPTURE_GDIP_LINE_CAP_ROUND);
        GdipSetPenDashCap197819(outlinePen, CAPTURE_GDIP_DASH_CAP_ROUND);
        GdipSetPenDashStyle(outlinePen, CAPTURE_GDIP_DASH_STYLE_DASH);
        DrawCaptureSelectionEdges(graphics, outlinePen, selection,
                                  outlineWidth / 2.0f);
        GdipDeletePen(outlinePen);
    }

    /* labelRect is already snug: GetCaptureSelectionLabelRect measures the
       text so geometry and hit-testing agree with what is rendered. */
    FormatCaptureDimensions(selection, dimensions,
                            sizeof(dimensions) / sizeof(dimensions[0]));
    font = GetCaptureFont(ScaleCaptureUiFloat(dpi, 12.0f));
    format = GetCaptureTextFormat();

    {
        float pillRadius = (float)(labelRect.bottom - labelRect.top) / 2.0f;
        float hairline = ScaleCaptureUiFloat(dpi, 0.65f);
        if (hairline < 0.6f) hairline = 0.6f;
        DrawCaptureSoftShadow(graphics, &labelRect, pillRadius, dpi,
                              3, 1.3f, 2, 10);
        FillCaptureRoundedRectGradient(graphics, &labelRect,
                                       CaptureArgb(242, RGB(40, 46, 57)),
                                       CaptureArgb(242, RGB(21, 25, 32)),
                                       pillRadius);
        StrokeCaptureRoundedRect(graphics, &labelRect, RGB(255, 255, 255), 44,
                                 hairline, pillRadius);
        DrawCaptureCenteredText(graphics, dimensions, &labelRect,
                                RGB(245, 249, 255), 255, font, format);
    }

    if (removeState >= 0) {
        RECT pillRect;
        UINT pillDpi = GetCaptureSelectionRemoveRect(selection, &pillRect);
        float pillRadius = (float)(pillRect.bottom - pillRect.top) / 2.0f;
        float hairline = ScaleCaptureUiFloat(pillDpi, 0.65f);
        BOOL hoveredPill = removeState == 1;
        GpPen *crossPen = NULL;
        float crossHalf = ScaleCaptureUiFloat(pillDpi, 4.5f);
        float crossWidth = ScaleCaptureUiFloat(pillDpi, 1.6f);
        float centerX =
            ((float)pillRect.left + (float)pillRect.right) / 2.0f;
        float centerY =
            ((float)pillRect.top + (float)pillRect.bottom) / 2.0f;

        if (hairline < 0.6f) hairline = 0.6f;
        if (crossWidth < 1.2f) crossWidth = 1.2f;
        DrawCaptureSoftShadow(graphics, &pillRect, pillRadius, pillDpi,
                              3, 1.3f, 2, 10);
        if (hoveredPill) {
            FillCaptureRoundedRect(graphics, &pillRect, RGB(224, 66, 66),
                                   235, pillRadius);
            StrokeCaptureRoundedRect(graphics, &pillRect, RGB(255, 158, 158),
                                     150, hairline, pillRadius);
        } else {
            FillCaptureRoundedRectGradient(graphics, &pillRect,
                                           CaptureArgb(242, RGB(40, 46, 57)),
                                           CaptureArgb(242, RGB(21, 25, 32)),
                                           pillRadius);
            StrokeCaptureRoundedRect(graphics, &pillRect, RGB(255, 255, 255),
                                     44, hairline, pillRadius);
        }
        if (GdipCreatePen1(CaptureArgb(255, hoveredPill
                                                ? RGB(255, 255, 255)
                                                : RGB(226, 233, 241)),
                           crossWidth, CAPTURE_GDIP_UNIT_PIXEL,
                           &crossPen) == 0 && crossPen) {
            GdipSetPenStartCap(crossPen, CAPTURE_GDIP_LINE_CAP_ROUND);
            GdipSetPenEndCap(crossPen, CAPTURE_GDIP_LINE_CAP_ROUND);
            GdipDrawLine(graphics, crossPen,
                         centerX - crossHalf, centerY - crossHalf,
                         centerX + crossHalf, centerY + crossHalf);
            GdipDrawLine(graphics, crossPen,
                         centerX + crossHalf, centerY - crossHalf,
                         centerX - crossHalf, centerY + crossHalf);
            GdipDeletePen(crossPen);
        }
    }
}

static void PaintScreenCaptureOverlay(HWND hwnd)
{
    PAINTSTRUCT paint;
    HDC dc = BeginPaint(hwnd, &paint);
    HDC sourceDc = NULL;
    HDC frameDc = NULL;
    HGDIOBJ oldSourceBitmap = NULL;
    HGDIOBJ oldFrameBitmap = NULL;
    GpGraphics *graphics = NULL;
    RECT selection;
    size_t displayedSelectionCount = GetDisplayedCaptureSelectionCount();
    int paintWidth = paint.rcPaint.right - paint.rcPaint.left;
    int paintHeight = paint.rcPaint.bottom - paint.rcPaint.top;

    if (!dc) return;
    if (paintWidth <= 0 || paintHeight <= 0 ||
        !g_captureOriginalBitmap || !g_captureFrameBitmap) {
        EndPaint(hwnd, &paint);
        return;
    }

    sourceDc = CreateCompatibleDC(dc);
    frameDc = CreateCompatibleDC(dc);
    if (!sourceDc || !frameDc) goto cleanup;
    oldSourceBitmap = SelectObject(sourceDc, g_captureOriginalBitmap);
    oldFrameBitmap = SelectObject(frameDc, g_captureFrameBitmap);
    IntersectClipRect(frameDc, paint.rcPaint.left, paint.rcPaint.top,
                      paint.rcPaint.right, paint.rcPaint.bottom);

    /* Assemble the complete dirty region off-screen, then publish it with one
       BitBlt. This prevents the dim layer and bright selection from appearing
       as separate frames while the pointer is moving. */
    BitBlt(frameDc, paint.rcPaint.left, paint.rcPaint.top,
           paintWidth, paintHeight, sourceDc,
           paint.rcPaint.left, paint.rcPaint.top, SRCCOPY);
    AlphaFillRect(frameDc, &paint.rcPaint, RGB(0, 0, 0), 102);
    for (size_t index = 0; index < displayedSelectionCount; index++) {
        if (GetDisplayedCaptureSelection(index, &selection)) {
            BitBlt(frameDc, selection.left, selection.top,
                   selection.right - selection.left,
                   selection.bottom - selection.top,
                   sourceDc, selection.left, selection.top, SRCCOPY);
        }
    }

    if (GdipCreateFromHDC(frameDc, &graphics) == 0 && graphics) {
        GdipSetCompositingQuality(
            graphics, CAPTURE_GDIP_COMPOSITING_HIGH_QUALITY);
        GdipSetSmoothingMode(graphics, CAPTURE_GDIP_SMOOTHING_ANTIALIAS_8X8);
        GdipSetPixelOffsetMode(
            graphics, CAPTURE_GDIP_PIXEL_OFFSET_HIGH_QUALITY);
        GdipSetTextRenderingHint(
            graphics, CAPTURE_GDIP_TEXT_ANTIALIAS_GRID_FIT);
        for (size_t index = 0; index < displayedSelectionCount; index++) {
            if (GetDisplayedCaptureSelection(index, &selection)) {
                int removeState = -1;
                if (index < g_captureSelectionCount) {
                    removeState =
                        (int)index == g_captureHoveredRemoval ? 1 : 0;
                }
                DrawCaptureSelection(graphics, &selection, removeState);
            }
        }
        for (int panel = 0; panel < g_capturePanelCount; panel++) {
            RECT intersection;
            RECT panelPaintRect = GetCapturePanelPaintRect(panel);
            if (IntersectRect(&intersection, &paint.rcPaint, &panelPaintRect)) {
                DrawCapturePanel(graphics, panel);
            }
        }
        GdipDeleteGraphics(graphics);
        graphics = NULL;
    }
    BitBlt(dc, paint.rcPaint.left, paint.rcPaint.top,
           paintWidth, paintHeight, frameDc,
           paint.rcPaint.left, paint.rcPaint.top, SRCCOPY);

cleanup:
    if (graphics) GdipDeleteGraphics(graphics);
    if (oldFrameBitmap && frameDc) SelectObject(frameDc, oldFrameBitmap);
    if (oldSourceBitmap && sourceDc) SelectObject(sourceDc, oldSourceBitmap);
    if (frameDc) DeleteDC(frameDc);
    if (sourceDc) DeleteDC(sourceDc);
    EndPaint(hwnd, &paint);
}

static int HitTestCaptureToolbar(POINT point, int *panelIndex)
{
    for (int panel = 0; panel < g_capturePanelCount; panel++) {
        for (int tool = 0; tool < CAPTURE_TOOL_COUNT; tool++) {
            if (PtInRect(&g_capturePanels[panel].buttonRects[tool], point)) {
                if (panelIndex) *panelIndex = panel;
                return tool;
            }
        }
    }
    if (panelIndex) *panelIndex = -1;
    return -1;
}

static int FindCapturePanelAtPoint(POINT point)
{
    for (int panel = 0; panel < g_capturePanelCount; panel++) {
        if (PtInRect(&g_capturePanels[panel].panelRect, point)) return panel;
    }
    return -1;
}

static BOOL IsPointInCapturePanel(POINT point)
{
    return FindCapturePanelAtPoint(point) >= 0;
}

static void InvalidateCapturePanel(HWND hwnd, int panelIndex)
{
    RECT area;
    if (panelIndex < 0 || panelIndex >= g_capturePanelCount) return;
    area = GetCapturePanelPaintRect(panelIndex);
    InvalidateRect(hwnd, &area, FALSE);
}

static void InvalidateAllCapturePanels(HWND hwnd)
{
    for (int panel = 0; panel < g_capturePanelCount; panel++) {
        InvalidateCapturePanel(hwnd, panel);
    }
}

/* Repaints every toolbar when the Copy label needs to flip between
   "Copy All" and "Copy Selection". */
static void RefreshCaptureCopyLabels(HWND hwnd)
{
    BOOL showsSelection = CaptureHasSelection();
    if (showsSelection == g_captureCopyShowsSelection) return;
    g_captureCopyShowsSelection = showsSelection;
    InvalidateAllCapturePanels(hwnd);
}

static void UnionCaptureSelectionDirtyRect(RECT *dirtyRect, BOOL *hasDirty,
                                           const RECT *selection);

static int HitTestCaptureRemovePills(POINT point)
{
    for (size_t i = 0; i < g_captureSelectionCount; i++) {
        RECT pillRect;
        GetCaptureSelectionRemoveRect(&g_captureSelections[i], &pillRect);
        if (PtInRect(&pillRect, point)) return (int)i;
    }
    return -1;
}

/* Newest box wins when selections overlap. */
static int HitTestCaptureSelections(POINT point)
{
    for (size_t i = g_captureSelectionCount; i > 0; i--) {
        if (PtInRect(&g_captureSelections[i - 1], point)) {
            return (int)(i - 1);
        }
    }
    return -1;
}

/* Finds a box wall or corner under the cursor for resizing. The grab band
   straddles each edge; corners take precedence over walls and the newest
   box wins. Returns the index or -1; *edge receives CAPTURE_EDGE_*. */
static int HitTestCaptureSelectionEdges(POINT point, int *edge)
{
    int margin = ScaleCaptureUiValue(GetCaptureDpiAtPoint(point), 5);

    *edge = CAPTURE_EDGE_NONE;
    for (size_t i = g_captureSelectionCount; i > 0; i--) {
        const RECT *box = &g_captureSelections[i - 1];
        BOOL nearLeft = abs(point.x - box->left) <= margin;
        BOOL nearRight = abs(point.x - box->right) <= margin;
        BOOL nearTop = abs(point.y - box->top) <= margin;
        BOOL nearBottom = abs(point.y - box->bottom) <= margin;
        BOOL inVerticalSpan = point.y >= box->top - margin &&
                              point.y <= box->bottom + margin;
        BOOL inHorizontalSpan = point.x >= box->left - margin &&
                                point.x <= box->right + margin;

        if (nearLeft && nearTop) *edge = CAPTURE_EDGE_TOPLEFT;
        else if (nearRight && nearTop) *edge = CAPTURE_EDGE_TOPRIGHT;
        else if (nearLeft && nearBottom) *edge = CAPTURE_EDGE_BOTTOMLEFT;
        else if (nearRight && nearBottom) *edge = CAPTURE_EDGE_BOTTOMRIGHT;
        else if (inVerticalSpan && nearLeft) *edge = CAPTURE_EDGE_LEFT;
        else if (inVerticalSpan && nearRight) *edge = CAPTURE_EDGE_RIGHT;
        else if (inHorizontalSpan && nearTop) *edge = CAPTURE_EDGE_TOP;
        else if (inHorizontalSpan && nearBottom) *edge = CAPTURE_EDGE_BOTTOM;
        if (*edge != CAPTURE_EDGE_NONE) return (int)(i - 1);
    }
    return -1;
}

static LPCWSTR GetCaptureEdgeCursor(int edge)
{
    switch (edge) {
    case CAPTURE_EDGE_LEFT:
    case CAPTURE_EDGE_RIGHT:
        return IDC_SIZEWE;
    case CAPTURE_EDGE_TOP:
    case CAPTURE_EDGE_BOTTOM:
        return IDC_SIZENS;
    case CAPTURE_EDGE_TOPLEFT:
    case CAPTURE_EDGE_BOTTOMRIGHT:
        return IDC_SIZENWSE;
    default:
        return IDC_SIZENESW;
    }
}

static void InvalidateCaptureRemovePill(HWND hwnd, int index)
{
    RECT pillRect;
    UINT dpi;
    if (index < 0 || (size_t)index >= g_captureSelectionCount) return;
    dpi = GetCaptureSelectionRemoveRect(&g_captureSelections[index],
                                        &pillRect);
    InflateRect(&pillRect, ScaleCaptureUiValue(dpi, 8),
                ScaleCaptureUiValue(dpi, 8));
    InvalidateRect(hwnd, &pillRect, FALSE);
}

static void RemoveCaptureSelectionAt(HWND hwnd, size_t index)
{
    RECT dirtyRect = {0};
    BOOL hasDirty = FALSE;

    if (index >= g_captureSelectionCount) return;
    UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty,
                                   &g_captureSelections[index]);
    g_captureSelectionCount--;
    if (index < g_captureSelectionCount) {
        memmove(&g_captureSelections[index], &g_captureSelections[index + 1],
                (g_captureSelectionCount - index) *
                    sizeof(*g_captureSelections));
    }
    g_captureHoveredRemoval = -1;
    if (hasDirty) InvalidateRect(hwnd, &dirtyRect, FALSE);
    RefreshCaptureCopyLabels(hwnd);
}

static void UnionCaptureSelectionDirtyRect(RECT *dirtyRect, BOOL *hasDirty,
                                           const RECT *selection)
{
    RECT area;
    RECT labelRect;
    RECT pillRect;
    UINT dpi;

    /* Degenerate rects still contribute: a 0 × 0 in-progress drag shows a
       dimension pill whose area must repaint. */
    area = *selection;
    dpi = GetCaptureSelectionLabelRect(selection, &labelRect);
    GetCaptureSelectionRemoveRect(selection, &pillRect);
    InflateRect(&area, ScaleCaptureUiValue(dpi, 4),
                 ScaleCaptureUiValue(dpi, 4));
    UnionRect(&area, &area, &labelRect);
    UnionRect(&area, &area, &pillRect);
    /* Margin covers the pills' layered drop shadows. */
    InflateRect(&area, ScaleCaptureUiValue(dpi, 8),
                 ScaleCaptureUiValue(dpi, 8));
    if (*hasDirty) {
        UnionRect(dirtyRect, dirtyRect, &area);
    } else {
        *dirtyRect = area;
        *hasDirty = TRUE;
    }
}

static void InvalidateStoredCaptureSelections(HWND hwnd)
{
    for (size_t index = 0; index < g_captureSelectionCount; index++) {
        RECT dirtyRect = {0};
        BOOL hasDirty = FALSE;
        UnionCaptureSelectionDirtyRect(
            &dirtyRect, &hasDirty, &g_captureSelections[index]);
        if (hasDirty) InvalidateRect(hwnd, &dirtyRect, FALSE);
    }
}

static void FinishCaptureSelectionDrag(HWND hwnd, POINT endPoint)
{
    RECT oldRect = NormalizeCaptureRect(
        g_captureDragStart, g_captureDragCurrent);
    RECT newRect;
    RECT dirtyRect = {0};
    BOOL hasDirty = FALSE;

    g_captureDragCurrent = endPoint;
    newRect = NormalizeCaptureRect(g_captureDragStart, g_captureDragCurrent);
    g_captureDragging = FALSE;
    if (IsCaptureSelectionValid(&newRect) &&
        !AppendCaptureSelection(&newRect)) {
        LogMessage("ERROR: Could not retain another screen capture selection");
        MessageBeep(MB_ICONWARNING);
    }

    /* Always repaint the drag area: a committed selection gains its remove
       pill, and an invalid release must erase the 0 × 0 dimension pill. */
    UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty, &oldRect);
    UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty, &newRect);
    if (hasDirty) InvalidateRect(hwnd, &dirtyRect, FALSE);
    RefreshCaptureCopyLabels(hwnd);
    /* Toolbars return to full opacity (and any vanished one reappears). */
    g_captureVanishedPanel = -1;
    InvalidateAllCapturePanels(hwnd);
}

static CaptureBlurSample GetCaptureBlurSample(int coordinate, int blockSize,
                                              int sampleCount)
{
    CaptureBlurSample sample = {0, 0, 0};
    LONGLONG fixedPosition;

    if (sampleCount <= 1) return sample;
    fixedPosition = ((2LL * coordinate + 1) * 256) /
                    (2LL * blockSize) - 128;
    if (fixedPosition <= 0) return sample;

    sample.first = (int)(fixedPosition / 256);
    if (sample.first >= sampleCount - 1) {
        sample.first = sampleCount - 1;
        sample.second = sample.first;
        return sample;
    }
    sample.second = sample.first + 1;
    sample.fraction = (unsigned int)(fixedPosition % 256);
    return sample;
}

static BOOL FillHeavilyBlurredCaptureBackground(
    BYTE *destinationPixels, size_t rowBytes, const RECT *sourceRect)
{
    int width = sourceRect->right - sourceRect->left;
    int height = sourceRect->bottom - sourceRect->top;
    int minimumDimension = width < height ? width : height;
    int blockSize = minimumDimension / 24;
    int reducedWidth;
    int reducedHeight;
    size_t reducedCount;
    DWORD *reduced = NULL;
    DWORD *smoothed = NULL;
    CaptureBlurSample *horizontalSamples = NULL;
    BOOL success = FALSE;

    if (!destinationPixels || !g_captureOriginalPixels ||
        width <= 0 || height <= 0) {
        return FALSE;
    }
    if (blockSize < 32) blockSize = 32;
    if (blockSize > 96) blockSize = 96;
    reducedWidth = (width - 1) / blockSize + 1;
    reducedHeight = (height - 1) / blockSize + 1;
    if ((size_t)reducedWidth > (size_t)-1 / (size_t)reducedHeight) {
        return FALSE;
    }
    reducedCount = (size_t)reducedWidth * (size_t)reducedHeight;
    if (reducedCount > (size_t)-1 / sizeof(DWORD) ||
        (size_t)width > (size_t)-1 / sizeof(CaptureBlurSample)) {
        return FALSE;
    }

    reduced = (DWORD *)malloc(reducedCount * sizeof(DWORD));
    smoothed = (DWORD *)malloc(reducedCount * sizeof(DWORD));
    horizontalSamples = (CaptureBlurSample *)malloc(
        (size_t)width * sizeof(CaptureBlurSample));
    if (!reduced || !smoothed || !horizontalSamples) goto cleanup;

    for (int reducedY = 0; reducedY < reducedHeight; reducedY++) {
        int localTop = reducedY * blockSize;
        int localBottom = localTop + blockSize;
        if (localBottom > height) localBottom = height;
        for (int reducedX = 0; reducedX < reducedWidth; reducedX++) {
            int localLeft = reducedX * blockSize;
            int localRight = localLeft + blockSize;
            ULONGLONG blue = 0;
            ULONGLONG green = 0;
            ULONGLONG red = 0;
            ULONGLONG samples = 0;
            if (localRight > width) localRight = width;

            for (int localY = localTop; localY < localBottom; localY++) {
                const DWORD *source = g_captureOriginalPixels +
                    ((size_t)(sourceRect->top + localY) *
                     (size_t)g_captureWidth) +
                    (size_t)(sourceRect->left + localLeft);
                for (int localX = localLeft; localX < localRight; localX++) {
                    DWORD pixel = *source++;
                    blue += pixel & 0xffu;
                    green += (pixel >> 8) & 0xffu;
                    red += (pixel >> 16) & 0xffu;
                    samples++;
                }
            }
            reduced[(size_t)reducedY * (size_t)reducedWidth +
                    (size_t)reducedX] =
                (DWORD)((blue + samples / 2) / samples) |
                ((DWORD)((green + samples / 2) / samples) << 8) |
                ((DWORD)((red + samples / 2) / samples) << 16);
        }
    }

    /* Smooth the reduced image before interpolation. Combined with aggressive
       downsampling this removes readable detail without creating blocky gaps. */
    for (int reducedY = 0; reducedY < reducedHeight; reducedY++) {
        for (int reducedX = 0; reducedX < reducedWidth; reducedX++) {
            unsigned int blue = 0;
            unsigned int green = 0;
            unsigned int red = 0;
            unsigned int samples = 0;
            int top = reducedY > 0 ? reducedY - 1 : 0;
            int bottom = reducedY + 1 < reducedHeight
                ? reducedY + 1 : reducedHeight - 1;
            int left = reducedX > 0 ? reducedX - 1 : 0;
            int right = reducedX + 1 < reducedWidth
                ? reducedX + 1 : reducedWidth - 1;
            for (int y = top; y <= bottom; y++) {
                for (int x = left; x <= right; x++) {
                    DWORD pixel = reduced[
                        (size_t)y * (size_t)reducedWidth + (size_t)x];
                    blue += pixel & 0xffu;
                    green += (pixel >> 8) & 0xffu;
                    red += (pixel >> 16) & 0xffu;
                    samples++;
                }
            }
            smoothed[(size_t)reducedY * (size_t)reducedWidth +
                     (size_t)reducedX] =
                (DWORD)((blue + samples / 2) / samples) |
                ((DWORD)((green + samples / 2) / samples) << 8) |
                ((DWORD)((red + samples / 2) / samples) << 16);
        }
    }

    for (int x = 0; x < width; x++) {
        horizontalSamples[x] = GetCaptureBlurSample(
            x, blockSize, reducedWidth);
    }
    for (int y = 0; y < height; y++) {
        CaptureBlurSample vertical = GetCaptureBlurSample(
            y, blockSize, reducedHeight);
        DWORD *destination = (DWORD *)(destinationPixels +
            (size_t)(height - 1 - y) * rowBytes);
        for (int x = 0; x < width; x++) {
            CaptureBlurSample horizontal = horizontalSamples[x];
            DWORD topLeft = smoothed[
                (size_t)vertical.first * (size_t)reducedWidth +
                (size_t)horizontal.first];
            DWORD topRight = smoothed[
                (size_t)vertical.first * (size_t)reducedWidth +
                (size_t)horizontal.second];
            DWORD bottomLeft = smoothed[
                (size_t)vertical.second * (size_t)reducedWidth +
                (size_t)horizontal.first];
            DWORD bottomRight = smoothed[
                (size_t)vertical.second * (size_t)reducedWidth +
                (size_t)horizontal.second];
            DWORD result = 0;
            unsigned int inverseX = 256 - horizontal.fraction;
            unsigned int inverseY = 256 - vertical.fraction;

            for (unsigned int shift = 0; shift <= 16; shift += 8) {
                unsigned int top =
                    ((topLeft >> shift) & 0xffu) * inverseX +
                    ((topRight >> shift) & 0xffu) * horizontal.fraction;
                unsigned int bottom =
                    ((bottomLeft >> shift) & 0xffu) * inverseX +
                    ((bottomRight >> shift) & 0xffu) * horizontal.fraction;
                unsigned int channel =
                    (top * inverseY + bottom * vertical.fraction + 32768u) >> 16;
                result |= (DWORD)channel << shift;
            }
            destination[x] = result;
        }
    }
    success = TRUE;

cleanup:
    free(horizontalSamples);
    free(smoothed);
    free(reduced);
    return success;
}

static HGLOBAL CreateCaptureClipboardDib(const RECT *sourceRect,
                                         BOOL selectedRegionsOnly)
{
    int width = sourceRect->right - sourceRect->left;
    int height = sourceRect->bottom - sourceRect->top;
    size_t rowBytes;
    size_t pixelBytes;
    size_t totalBytes;
    HGLOBAL dibMemory;
    BYTE *dibData;
    BITMAPINFOHEADER *header;
    BYTE *destinationPixels;
    BOOL combineSelections =
        selectedRegionsOnly && g_captureSelectionCount > 1;

    if (!g_captureOriginalPixels || width <= 0 || height <= 0 ||
        (size_t)width > (size_t)-1 / 4u) return NULL;
    rowBytes = (size_t)width * 4u;
    if ((size_t)height > (size_t)-1 / rowBytes) return NULL;
    pixelBytes = rowBytes * (size_t)height;
    if (pixelBytes > 0xffffffffu ||
        pixelBytes > (size_t)-1 - sizeof(BITMAPINFOHEADER)) return NULL;
    totalBytes = sizeof(BITMAPINFOHEADER) + pixelBytes;

    dibMemory = GlobalAlloc(GMEM_MOVEABLE, totalBytes);
    if (!dibMemory) return NULL;
    dibData = (BYTE *)GlobalLock(dibMemory);
    if (!dibData) {
        GlobalFree(dibMemory);
        return NULL;
    }

    header = (BITMAPINFOHEADER *)dibData;
    ZeroMemory(header, sizeof(*header));
    header->biSize = sizeof(*header);
    header->biWidth = width;
    header->biHeight = height;
    header->biPlanes = 1;
    header->biBitCount = 32;
    header->biCompression = BI_RGB;
    header->biSizeImage = (DWORD)pixelBytes;
    destinationPixels = dibData + sizeof(*header);

    if (combineSelections) {
        if (g_configCaptureGapFill == CAPTURE_GAP_FILL_BLUR) {
            if (!FillHeavilyBlurredCaptureBackground(
                    destinationPixels, rowBytes, sourceRect)) {
                GlobalUnlock(dibMemory);
                GlobalFree(dibMemory);
                return NULL;
            }
        } else if (g_configCaptureGapFill == CAPTURE_GAP_FILL_BLACK) {
            ZeroMemory(destinationPixels, pixelBytes);
        } else {
            DWORD *pixels = (DWORD *)destinationPixels;
            size_t pixelCount = pixelBytes / sizeof(DWORD);
            for (size_t index = 0; index < pixelCount; index++) {
                pixels[index] = 0x00ffffffu;
            }
        }
    }

    {
        size_t copyCount = combineSelections ? g_captureSelectionCount : 1;
        for (size_t index = 0; index < copyCount; index++) {
            RECT copyRect;
            if (combineSelections) {
                if (!IntersectRect(&copyRect, &g_captureSelections[index],
                                   sourceRect)) {
                    continue;
                }
            } else {
                copyRect = *sourceRect;
            }

            size_t copyBytes =
                (size_t)(copyRect.right - copyRect.left) * sizeof(DWORD);
            for (int sourceY = copyRect.top;
                 sourceY < copyRect.bottom; sourceY++) {
                const DWORD *source = g_captureOriginalPixels +
                    ((size_t)sourceY * (size_t)g_captureWidth) +
                    (size_t)copyRect.left;
                size_t destinationY =
                    (size_t)(sourceY - sourceRect->top);
                BYTE *destination = destinationPixels +
                    ((size_t)(height - 1) - destinationY) * rowBytes +
                    (size_t)(copyRect.left - sourceRect->left) * sizeof(DWORD);
                memcpy(destination, source, copyBytes);
            }
        }
    }
    GlobalUnlock(dibMemory);
    return dibMemory;
}

static void CloseScreenCaptureOverlay(void)
{
    HWND overlay = g_captureOverlayHwnd;
    if (overlay) {
        ShowWindow(overlay, SW_HIDE);
        DestroyWindow(overlay);
    } else {
        ReleaseCaptureSelections();
        ReleaseScreenCaptureBitmaps();
    }
}

static LRESULT CALLBACK ScreenCaptureWndProc(HWND hwnd, UINT message,
                                             WPARAM wParam, LPARAM lParam)
{
    switch (message) {
    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT:
        PaintScreenCaptureOverlay(hwnd);
        return 0;

    case WM_LBUTTONDOWN:
        {
            POINT point = GetCaptureCursorPoint(hwnd);
            int panelIndex = -1;
            int removeIndex = -1;
            int moveIndex = -1;
            int resizeIndex = -1;
            int resizeEdge = CAPTURE_EDGE_NONE;
            int tool = HitTestCaptureToolbar(point, &panelIndex);
            if (tool >= 0) {
                g_capturePressedPanel = panelIndex;
                g_capturePressedTool = tool;
                SetCapture(hwnd);
                InvalidateCapturePanel(hwnd, panelIndex);
            } else if ((removeIndex = HitTestCaptureRemovePills(point)) >= 0) {
                RemoveCaptureSelectionAt(hwnd, (size_t)removeIndex);
            } else if ((wParam & MK_SHIFT) == 0 &&
                       (resizeIndex =
                            HitTestCaptureSelectionEdges(point,
                                                         &resizeEdge)) >= 0) {
                g_captureResizingIndex = resizeIndex;
                g_captureResizingEdge = resizeEdge;
                SetCapture(hwnd);
            } else if ((wParam & MK_SHIFT) == 0 &&
                       (moveIndex = HitTestCaptureSelections(point)) >= 0) {
                /* Plain click inside a box picks it up; Shift+drag still
                   draws a new additive box across existing ones. */
                g_captureMovingIndex = moveIndex;
                g_captureMoveGrabOffset.x =
                    point.x - g_captureSelections[moveIndex].left;
                g_captureMoveGrabOffset.y =
                    point.y - g_captureSelections[moveIndex].top;
                SetCapture(hwnd);
            } else if (!IsPointInCapturePanel(point) &&
                       g_captureSelectedTool == CAPTURE_TOOL_CLIP) {
                BOOL additive = (wParam & MK_SHIFT) != 0;
                if (!additive) {
                    InvalidateStoredCaptureSelections(hwnd);
                    ClearCaptureSelections();
                }
                g_captureDragging = TRUE;
                g_captureDragStart = point;
                g_captureDragCurrent = point;
                SetCapture(hwnd);
                RefreshCaptureCopyLabels(hwnd);
                /* Toolbars fade to half opacity while drawing. */
                InvalidateAllCapturePanels(hwnd);
            }
        }
        return 0;

    case WM_MOUSEMOVE:
        {
            POINT point = GetCaptureCursorPoint(hwnd);
            int panelIndex = -1;
            int tool = HitTestCaptureToolbar(point, &panelIndex);
            /* Hover feedback is suppressed while drawing, so skip the state
               churn (and the panel repaints it triggers) during a drag. */
            if (!g_captureDragging &&
                (panelIndex != g_captureHoveredPanel ||
                 tool != g_captureHoveredTool)) {
                int oldPanel = g_captureHoveredPanel;
                g_captureHoveredPanel = panelIndex;
                g_captureHoveredTool = tool;
                InvalidateCapturePanel(hwnd, oldPanel);
                if (panelIndex != oldPanel) {
                    InvalidateCapturePanel(hwnd, panelIndex);
                }
            }
            if (!g_captureDragging && g_captureMovingIndex < 0 &&
                g_captureResizingIndex < 0) {
                int hoveredRemoval = HitTestCaptureRemovePills(point);
                if (hoveredRemoval != g_captureHoveredRemoval) {
                    int oldRemoval = g_captureHoveredRemoval;
                    g_captureHoveredRemoval = hoveredRemoval;
                    InvalidateCaptureRemovePill(hwnd, oldRemoval);
                    InvalidateCaptureRemovePill(hwnd, hoveredRemoval);
                }
            }
            if (g_captureResizingIndex >= 0 &&
                (size_t)g_captureResizingIndex < g_captureSelectionCount) {
                RECT *box = &g_captureSelections[g_captureResizingIndex];
                RECT oldBox = *box;
                int edge = g_captureResizingEdge;
                if (edge == CAPTURE_EDGE_LEFT ||
                    edge == CAPTURE_EDGE_TOPLEFT ||
                    edge == CAPTURE_EDGE_BOTTOMLEFT) {
                    box->left = point.x > box->right - 2
                        ? box->right - 2 : point.x;
                }
                if (edge == CAPTURE_EDGE_RIGHT ||
                    edge == CAPTURE_EDGE_TOPRIGHT ||
                    edge == CAPTURE_EDGE_BOTTOMRIGHT) {
                    box->right = point.x < box->left + 2
                        ? box->left + 2 : point.x;
                }
                if (edge == CAPTURE_EDGE_TOP ||
                    edge == CAPTURE_EDGE_TOPLEFT ||
                    edge == CAPTURE_EDGE_TOPRIGHT) {
                    box->top = point.y > box->bottom - 2
                        ? box->bottom - 2 : point.y;
                }
                if (edge == CAPTURE_EDGE_BOTTOM ||
                    edge == CAPTURE_EDGE_BOTTOMLEFT ||
                    edge == CAPTURE_EDGE_BOTTOMRIGHT) {
                    box->bottom = point.y < box->top + 2
                        ? box->top + 2 : point.y;
                }
                if (!EqualRect(&oldBox, box)) {
                    RECT dirtyRect = {0};
                    BOOL hasDirty = FALSE;
                    UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty,
                                                   &oldBox);
                    UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty,
                                                   box);
                    if (hasDirty) InvalidateRect(hwnd, &dirtyRect, FALSE);
                }
            }
            if (g_captureMovingIndex >= 0 &&
                (size_t)g_captureMovingIndex < g_captureSelectionCount) {
                RECT *box = &g_captureSelections[g_captureMovingIndex];
                int boxWidth = box->right - box->left;
                int boxHeight = box->bottom - box->top;
                int newLeft = point.x - g_captureMoveGrabOffset.x;
                int newTop = point.y - g_captureMoveGrabOffset.y;
                if (newLeft < 0) newLeft = 0;
                if (newTop < 0) newTop = 0;
                if (newLeft + boxWidth > g_captureWidth) {
                    newLeft = g_captureWidth - boxWidth;
                }
                if (newTop + boxHeight > g_captureHeight) {
                    newTop = g_captureHeight - boxHeight;
                }
                if (newLeft != box->left || newTop != box->top) {
                    RECT oldBox = *box;
                    RECT dirtyRect = {0};
                    BOOL hasDirty = FALSE;
                    SetRect(box, newLeft, newTop,
                            newLeft + boxWidth, newTop + boxHeight);
                    /* Both dirty rects include the corner pills at their
                       inside or outside placement for that box position. */
                    UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty,
                                                   &oldBox);
                    UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty,
                                                   box);
                    if (hasDirty) InvalidateRect(hwnd, &dirtyRect, FALSE);
                }
            }
            if (g_captureDragging) {
                RECT oldRect = NormalizeCaptureRect(g_captureDragStart,
                                                    g_captureDragCurrent);
                RECT newRect;
                RECT dirtyRect = {0};
                BOOL hasDirty = FALSE;
                int vanishedPanel = FindCapturePanelAtPoint(point);
                g_captureDragCurrent = point;
                newRect = NormalizeCaptureRect(g_captureDragStart,
                                               g_captureDragCurrent);
                UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty, &oldRect);
                UnionCaptureSelectionDirtyRect(&dirtyRect, &hasDirty, &newRect);
                if (hasDirty) InvalidateRect(hwnd, &dirtyRect, FALSE);
                if (vanishedPanel != g_captureVanishedPanel) {
                    int oldVanished = g_captureVanishedPanel;
                    g_captureVanishedPanel = vanishedPanel;
                    InvalidateCapturePanel(hwnd, oldVanished);
                    InvalidateCapturePanel(hwnd, vanishedPanel);
                }
                RefreshCaptureCopyLabels(hwnd);
            }
        }
        return 0;

    case WM_LBUTTONUP:
        if (g_capturePressedTool >= 0) {
            POINT point = GetCaptureCursorPoint(hwnd);
            int releasePanel = -1;
            int releaseTool = HitTestCaptureToolbar(point, &releasePanel);
            int pressedPanel = g_capturePressedPanel;
            int pressedTool = g_capturePressedTool;
            g_capturePressedPanel = -1;
            g_capturePressedTool = -1;
            if (GetCapture() == hwnd) ReleaseCapture();
            InvalidateCapturePanel(hwnd, pressedPanel);
            if (releasePanel == pressedPanel && releaseTool == pressedTool) {
                if (pressedTool == CAPTURE_TOOL_CLIP) {
                    g_captureSelectedTool = CAPTURE_TOOL_CLIP;
                    InvalidateStoredCaptureSelections(hwnd);
                    ClearCaptureSelections();
                    RefreshCaptureCopyLabels(hwnd);
                } else if (pressedTool == CAPTURE_TOOL_COPY) {
                    PostMessage(g_hWndMain, WM_SCREEN_CAPTURE_COPY, FALSE, 0);
                } else if (pressedTool == CAPTURE_TOOL_CANCEL) {
                    PostMessage(g_hWndMain, WM_SCREEN_CAPTURE_CANCEL, 0, 0);
                }
            }
            return 0;
        }
        if (g_captureMovingIndex >= 0 || g_captureResizingIndex >= 0) {
            g_captureMovingIndex = -1;
            g_captureResizingIndex = -1;
            g_captureResizingEdge = CAPTURE_EDGE_NONE;
            if (GetCapture() == hwnd) ReleaseCapture();
            return 0;
        }
        if (g_captureDragging) {
            POINT point = GetCaptureCursorPoint(hwnd);
            FinishCaptureSelectionDrag(hwnd, point);
            if (GetCapture() == hwnd) ReleaseCapture();
        }
        return 0;

    case WM_CAPTURECHANGED:
        if (g_capturePressedTool >= 0) {
            int oldPanel = g_capturePressedPanel;
            g_capturePressedPanel = -1;
            g_capturePressedTool = -1;
            InvalidateCapturePanel(hwnd, oldPanel);
        }
        g_captureMovingIndex = -1;
        g_captureResizingIndex = -1;
        g_captureResizingEdge = CAPTURE_EDGE_NONE;
        if (g_captureDragging) {
            POINT point = GetCaptureCursorPoint(hwnd);
            FinishCaptureSelectionDrag(hwnd, point);
        }
        return 0;

    case WM_SETCURSOR:
        if (LOWORD(lParam) == HTCLIENT) {
            POINT point = GetCaptureCursorPoint(hwnd);
            int hoverEdge = CAPTURE_EDGE_NONE;
            LPCWSTR cursorName;
            if (g_captureDragging) {
                cursorName = IDC_CROSS;
            } else if (g_captureResizingIndex >= 0) {
                cursorName = GetCaptureEdgeCursor(g_captureResizingEdge);
            } else if (g_captureMovingIndex >= 0 ||
                       HitTestCaptureToolbar(point, NULL) >= 0 ||
                       HitTestCaptureRemovePills(point) >= 0) {
                cursorName = IDC_HAND;
            } else if (IsPointInCapturePanel(point)) {
                cursorName = IDC_ARROW;
            } else if (HitTestCaptureSelectionEdges(point, &hoverEdge) >= 0) {
                cursorName = GetCaptureEdgeCursor(hoverEdge);
            } else if (HitTestCaptureSelections(point) >= 0) {
                cursorName = IDC_HAND;
            } else {
                cursorName = IDC_CROSS;
            }
            SetCursor(LoadCursor(NULL, cursorName));
            return TRUE;
        }
        break;

    case WM_KEYDOWN:
    case WM_SYSKEYDOWN:
        if (wParam == VK_ESCAPE) {
            PostMessage(g_hWndMain, WM_SCREEN_CAPTURE_CANCEL, 0, 0);
            return 0;
        }
        if (wParam == VK_SNAPSHOT) {
            PostMessage(g_hWndMain, WM_SCREEN_CAPTURE_COPY, TRUE, 0);
            return 0;
        }
        if (wParam == VK_RETURN && g_captureSelectionCount > 0) {
            PostMessage(g_hWndMain, WM_SCREEN_CAPTURE_COPY, FALSE, 0);
            return 0;
        }
        break;

    case WM_DISPLAYCHANGE:
        PostMessage(g_hWndMain, WM_SCREEN_CAPTURE_CANCEL, 1, 0);
        return 0;

    case WM_DESTROY:
        g_captureDragging = FALSE;
        g_capturePressedPanel = -1;
        g_capturePressedTool = -1;
        if (GetCapture() == hwnd) ReleaseCapture();
        if (g_captureOverlayHwnd == hwnd) {
            g_captureOverlayHwnd = NULL;
            SetHookCaptureFlag(HOOK_CAPTURE_ACTIVE, FALSE);
        }
        g_captureHoveredPanel = -1;
        g_captureHoveredTool = -1;
        g_captureHoveredRemoval = -1;
        g_captureMovingIndex = -1;
        g_captureResizingIndex = -1;
        g_captureResizingEdge = CAPTURE_EDGE_NONE;
        g_captureVanishedPanel = -1;
        ReleaseCaptureDrawingCache();
        ReleaseCaptureSelections();
        ReleaseScreenCaptureBitmaps();
        return 0;
    }
    return DefWindowProcW(hwnd, message, wParam, lParam);
}

static BOOL BeginScreenCapture(void)
{
    static BOOL classRegistered = FALSE;
    WNDCLASSEXW windowClass;

    /* The Print Screen hook checks g_configScreenCaptureEnabled before
       posting; the tray menu's Capture entry works regardless. */
    if (g_captureOverlayHwnd) return TRUE;
    EnterScreenCaptureDpiMode();
    if (!CaptureVirtualDesktopSnapshot()) {
        LogMessage("ERROR: Could not capture the virtual desktop (%lu)",
                   GetLastError());
        MessageBeep(MB_ICONWARNING);
        LeaveScreenCaptureDpiMode();
        return FALSE;
    }
    BuildCapturePanels();

    if (!classRegistered) {
        ZeroMemory(&windowClass, sizeof(windowClass));
        windowClass.cbSize = sizeof(windowClass);
        windowClass.style = CS_HREDRAW | CS_VREDRAW;
        windowClass.lpfnWndProc = ScreenCaptureWndProc;
        windowClass.hInstance = g_hInstance;
        windowClass.hCursor = LoadCursor(NULL, IDC_CROSS);
        windowClass.lpszClassName = L"ImagePasterCaptureOverlay";
        if (!RegisterClassExW(&windowClass)) {
            LogMessage("ERROR: Could not register screen capture overlay (%lu)",
                       GetLastError());
            ReleaseScreenCaptureBitmaps();
            LeaveScreenCaptureDpiMode();
            return FALSE;
        }
        classRegistered = TRUE;
    }

    g_captureSelectedTool = CAPTURE_TOOL_CLIP;
    g_captureHoveredPanel = -1;
    g_captureHoveredTool = -1;
    g_capturePressedPanel = -1;
    g_capturePressedTool = -1;
    g_captureDragging = FALSE;
    g_captureCopyShowsSelection = FALSE;
    g_captureHoveredRemoval = -1;
    g_captureMovingIndex = -1;
    g_captureResizingIndex = -1;
    g_captureResizingEdge = CAPTURE_EDGE_NONE;
    g_captureVanishedPanel = -1;
    ReleaseCaptureSelections();
    g_captureOverlayHwnd = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        L"ImagePasterCaptureOverlay", L"ImagePaster Screen Capture",
        WS_POPUP, g_captureVirtualX, g_captureVirtualY,
        g_captureWidth, g_captureHeight,
        NULL, NULL, g_hInstance, NULL);
    if (!g_captureOverlayHwnd) {
        LogMessage("ERROR: Could not create screen capture overlay (%lu)",
                   GetLastError());
        ReleaseScreenCaptureBitmaps();
        LeaveScreenCaptureDpiMode();
        return FALSE;
    }

    SetHookCaptureFlag(HOOK_CAPTURE_ACTIVE, TRUE);
    ShowWindow(g_captureOverlayHwnd, SW_SHOW);
    SetWindowPos(g_captureOverlayHwnd, HWND_TOPMOST,
                 g_captureVirtualX, g_captureVirtualY,
                 g_captureWidth, g_captureHeight,
                 SWP_SHOWWINDOW);
    SetForegroundWindow(g_captureOverlayHwnd);
    SetFocus(g_captureOverlayHwnd);
    UpdateWindow(g_captureOverlayHwnd);
    LeaveScreenCaptureDpiMode();
    LogMessage("Print Screen capture opened across %d monitor(s)",
               g_capturePanelCount);
    return TRUE;
}

static BOOL CompleteScreenCapture(BOOL forceFullDesktop)
{
    RECT sourceRect;
    HGLOBAL dibMemory;
    int width;
    int height;
    BOOL usedSelection = FALSE;
    size_t selectionCount = 0;

    if (!g_captureOverlayHwnd || !g_captureOriginalPixels) return FALSE;
    if (!forceFullDesktop && GetCaptureSelectionBounds(&sourceRect)) {
        usedSelection = TRUE;
        selectionCount = g_captureSelectionCount;
    } else {
        SetRect(&sourceRect, 0, 0, g_captureWidth, g_captureHeight);
    }
    width = sourceRect.right - sourceRect.left;
    height = sourceRect.bottom - sourceRect.top;
    dibMemory = CreateCaptureClipboardDib(&sourceRect, usedSelection);
    if (!dibMemory) {
        LogMessage("ERROR: Could not allocate the screen capture clipboard image");
        MessageBeep(MB_ICONWARNING);
        return FALSE;
    }

    if (!OpenClipboard(g_hWndMain)) {
        LogMessage("ERROR: Could not open the clipboard for screen capture (%lu)",
                   GetLastError());
        GlobalFree(dibMemory);
        MessageBeep(MB_ICONWARNING);
        return FALSE;
    }
    if (!EmptyClipboard() || !SetClipboardData(CF_DIB, dibMemory)) {
        DWORD errorCode = GetLastError();
        CloseClipboard();
        GlobalFree(dibMemory);
        LogMessage("ERROR: Could not place screen capture on the clipboard (%lu)",
                   errorCode);
        MessageBeep(MB_ICONWARNING);
        return FALSE;
    }
    CloseClipboard();

    CloseScreenCaptureOverlay();
    if (usedSelection) {
        LogMessage("Screen capture copied: %dx%d (%llu selection%s)",
                   width, height, (unsigned long long)selectionCount,
                   selectionCount == 1 ? "" : "s");
    } else {
        LogMessage("Screen capture copied: %dx%d (full desktop)",
                   width, height);
    }
    SetCachedClipboardSequence(0);
    RefreshClipboardImageCache();
    return TRUE;
}

static void CancelScreenCapture(const char *reason)
{
    if (!g_captureOverlayHwnd) return;
    CloseScreenCaptureOverlay();
    LogMessage("Screen capture cancelled: %s", reason);
}

/* ── Low-level keyboard hook ────────────────────────────────────────────── */

/* GetWindowText does not send messages across processes, but it can wait on
   our own UI thread. Our settings/log/history windows are never paste targets. */
static BOOL HookForegroundTitleMatches(HWND foreground)
{
    DWORD processId = 0;
    WCHAR title[512];
    BOOL matchFound = FALSE;

    if (!foreground ||
        !GetWindowThreadProcessId(foreground, &processId) ||
        processId == GetCurrentProcessId()) {
        return FALSE;
    }
    if (GetWindowTextW(foreground, title, 512) <= 0) return FALSE;
    for (WCHAR *p = title; *p; p++) {
        if (*p >= L'A' && *p <= L'Z') *p = *p - L'A' + L'a';
    }
    if (!TryAcquireSRWLockShared(&g_keywordLock)) return FALSE;
    for (int i = 0; i < g_keywordCount; i++) {
        if (wcsstr(title, g_keywords[i]) != NULL) {
            matchFound = TRUE;
            break;
        }
    }
    ReleaseSRWLockShared(&g_keywordLock);
    return matchFound;
}

/* At most one Print Screen and one cancel action may be outstanding. Keep the
   slot until UI work finishes, including capture/encoding: repeated taps must
   not flood a busy queue or unexpectedly copy the whole screen on recovery. */
static BOOL QueueHookCapture(UINT message, WPARAM argument)
{
    volatile LONG *pending = message == WM_SCREEN_CAPTURE_CANCEL
        ? &g_hookCancelPending : &g_hookPrintPending;
    if (InterlockedCompareExchange(pending, TRUE, FALSE)) return TRUE;
    if (PostMessageW(g_hWndMain, message, argument, HOOK_CAPTURE_MESSAGE)) return TRUE;
    InterlockedExchange(pending, FALSE);
    return FALSE;
}

static BOOL QueueHookPrintScreen(LONG captureState)
{
    if (!(captureState & (HOOK_CAPTURE_ENABLED | HOOK_CAPTURE_ACTIVE))) return FALSE;
    return QueueHookCapture((captureState & HOOK_CAPTURE_ACTIVE)
                            ? WM_SCREEN_CAPTURE_COPY : WM_SCREEN_CAPTURE_BEGIN,
                            (captureState & HOOK_CAPTURE_ACTIVE) ? TRUE : FALSE);
}

/* Bound paste traffic as well: even a held Ctrl+V must not crowd capture out of
   a stalled UI queue. Coalesce only the same target/clipboard request; a new
   target or clipboard passes through normally instead of using stale data. */
static BOOL QueueHookPaste(HWND target, DWORD sequence)
{
    if (InterlockedCompareExchange(&g_hookPastePending, TRUE, FALSE)) {
        return target == g_hookPasteTarget && sequence == g_hookPasteSequence;
    }
    g_hookPasteTarget = target;
    g_hookPasteSequence = sequence;
    if (PostMessageW(g_hWndMain, WM_DO_PASTE, sequence, (LPARAM)target)) return TRUE;
    InterlockedExchange(&g_hookPastePending, FALSE);
    return FALSE;
}

/* Runs only on the hook thread. Never log, encode, touch files/WebView, allocate,
   or wait on the UI/a lock. Print Screen does one atomic state read and at most
   one asynchronous post; unrelated keys do not even read shared state. */
static LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode != HC_ACTION) return CallNextHookEx(NULL, nCode, wParam, lParam);
    const KBDLLHOOKSTRUCT *pKb = (const KBDLLHOOKSTRUCT *)lParam;
    if (pKb->vkCode != VK_SNAPSHOT && pKb->vkCode != VK_ESCAPE && pKb->vkCode != 'V') {
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }
    BOOL keyDown = wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN;
    BOOL keyUp = wParam == WM_KEYUP || wParam == WM_SYSKEYUP;
    if (!keyDown && !keyUp) return CallNextHookEx(NULL, nCode, wParam, lParam);

    if (pKb->vkCode == VK_SNAPSHOT) {
        LONG state = InterlockedCompareExchange(&g_hookCaptureState, 0, 0);
        if (!(state & (HOOK_CAPTURE_ENABLED | HOOK_CAPTURE_ACTIVE)) &&
            !g_printScreenKeyDown) return CallNextHookEx(NULL, nCode, wParam, lParam);
        /* Some keyboards expose only the key-up transition. */
        if (!g_printScreenKeyDown && !QueueHookPrintScreen(state)) {
            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }
        g_printScreenKeyDown = keyDown;
        return 1;
    }

    if (pKb->vkCode == VK_ESCAPE) {
        LONG state = InterlockedCompareExchange(&g_hookCaptureState, 0, 0);
        if (!(state & HOOK_CAPTURE_ACTIVE) && !g_escapeKeyDown) {
            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }
        if (!g_escapeKeyDown && !QueueHookCapture(WM_SCREEN_CAPTURE_CANCEL, 0)) {
            return CallNextHookEx(NULL, nCode, wParam, lParam);
        }
        g_escapeKeyDown = keyDown;
        return 1;
    }

    /* Paste is the only path needing routing queries. Keep normal typing,
       injected pastes and non-image pastes out of title matching entirely. */
    if (!keyDown || ((pKb->flags & LLKHF_INJECTED) &&
                    pKb->dwExtraInfo == PASTE_INPUT_TAG) ||
        !InterlockedCompareExchange(&g_hookHasKeywords, 0, 0) ||
        !(GetAsyncKeyState(VK_CONTROL) & 0x8000) ||
        (GetAsyncKeyState(VK_MENU) & 0x8000)) {
        return CallNextHookEx(NULL, nCode, wParam, lParam);
    }
    DWORD sequence = GetClipboardSequenceNumber();
    BOOL imagePasteAvailable = IsClipboardFormatAvailable(CF_DIB) ||
        (sequence == GetCachedClipboardSequence() && TryHasCachedImage());
    if (imagePasteAvailable) {
        HWND foreground = GetForegroundWindow();
        if (HookForegroundTitleMatches(foreground) &&
            QueueHookPaste(foreground, sequence)) return 1;
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}

/* A registered hotkey is not subject to LowLevelHooksTimeout. The hook normally
   swallows Print Screen before WM_HOTKEY is generated; this backup receives it
   if the hook is removed. Registration may be denied by another app/the OS, so
   keep the hook in either case and retry a failed registration at renewal. */
static void ReconcilePrintScreenHotkey(void)
{
    BOOL wanted = (InterlockedCompareExchange(&g_hookCaptureState, 0, 0) &
                   (HOOK_CAPTURE_ENABLED | HOOK_CAPTURE_ACTIVE)) != 0;
    if (!wanted) {
        if (g_printHotkeyRegistered) UnregisterHotKey(NULL, ID_HOTKEY_PRINT_SCREEN);
        g_printHotkeyRegistered = FALSE;
        g_printHotkeyError = 0;
    } else if (!g_printHotkeyRegistered) {
        if (RegisterHotKey(NULL, ID_HOTKEY_PRINT_SCREEN, MOD_NOREPEAT, VK_SNAPSHOT)) {
            g_printHotkeyRegistered = TRUE;
            g_printHotkeyError = 0;
            PostMessageW(g_hWndMain, WM_KEYBOARD_HOTKEY_STATUS, TRUE, 0);
        } else {
            DWORD error = GetLastError();
            if (error != g_printHotkeyError) {
                PostMessageW(g_hWndMain, WM_KEYBOARD_HOTKEY_STATUS, FALSE, error);
                g_printHotkeyError = error;
            }
        }
    }
}

static void HandlePrintScreenHotkey(void)
{
    /* A hotkey message may outlive its registration/configuration. Recheck it. */
    if (g_printHotkeyRegistered) {
        if (QueueHookPrintScreen(InterlockedCompareExchange(&g_hookCaptureState, 0, 0))) {
            /* This fallback key was not swallowed by our hook, so async state
               is usable here. Its release must not become a second action. */
            g_printScreenKeyDown = (GetAsyncKeyState(VK_SNAPSHOT) & 0x8000) != 0;
        }
        /* Reinstall immediately, instead of waiting for the next deadline. */
        InterlockedExchange(&g_hookRenewRequested, TRUE);
    }
}

/* Windows provides no notification when LowLevelHooksTimeout silently removes
   a hook. Renew periodically on its owner thread, even while the UI is busy.
   Install first so a failed replacement never discards a working hook. */
static BOOL RenewKeyboardHook(BOOL reportSuccess)
{
    HHOOK replacement = SetWindowsHookExW(
        WH_KEYBOARD_LL, LowLevelKeyboardProc, g_hInstance, 0);
    if (!replacement) {
        DWORD error = GetLastError();
        if (error != g_hookInstallError) {
            PostMessageW(g_hWndMain, WM_KEYBOARD_HOOK_STATUS, FALSE, error);
            g_hookInstallError = error;
        }
        return FALSE;
    }

    g_hookInstallError = 0;
    HHOOK previous = g_hHook;
    g_hHook = replacement;
    if (previous) UnhookWindowsHookEx(previous);
    /* Preserve key-down guards across renewal. Swallowed keys do not update
       async key state: querying it here would mistake a held key for a release
       and turn auto-repeat/key-up into an unwanted second capture. */
    if (reportSuccess) {
        PostMessageW(g_hWndMain, WM_KEYBOARD_HOOK_STATUS, TRUE, 0);
    }
    return TRUE;
}

static DWORD WINAPI KeyboardHookThreadProc(LPVOID parameter)
{
    MSG message;
    HANDLE events[] = {g_keyboardHookStopEvent, g_keyboardHookWakeEvent};
    ULONGLONG renewAt = 0;
    BOOL reportSuccess = TRUE;
    BOOL reconcileHotkey = TRUE;
    (void)parameter;

    /* Prefer this tiny input pump to our CPU-heavy image work, without raising
       the process to high/realtime priority or disabling normal OS boosts. */
    if (!SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_ABOVE_NORMAL)) {
        PostMessageW(g_hWndMain, WM_KEYBOARD_HOOK_STATUS, 2, GetLastError());
    }
    /* Create this thread's message queue before installing the hook. */
    PeekMessageW(&message, NULL, WM_USER, WM_USER, PM_NOREMOVE);
    for (;;) {
        if (WaitForSingleObject(g_keyboardHookStopEvent, 0) == WAIT_OBJECT_0) break;
        ULONGLONG now = GetTickCount64();
        if (reconcileHotkey || now >= renewAt) {
            ReconcilePrintScreenHotkey();
            reconcileHotkey = FALSE;
        }
        if (InterlockedExchange(&g_hookRenewRequested, FALSE) || now >= renewAt) {
            reportSuccess = !RenewKeyboardHook(reportSuccess);
            renewAt = GetTickCount64() + KEYBOARD_HOOK_RENEW_MS;
        }

        now = GetTickCount64();
        DWORD waitMs = now < renewAt ? (DWORD)(renewAt - now) : 0;
        DWORD result = MsgWaitForMultipleObjectsEx(
            2, events, waitMs, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
        if (result == WAIT_OBJECT_0) break;
        if (result == WAIT_OBJECT_0 + 1) reconcileHotkey = TRUE;
        if (result == WAIT_FAILED) {
            PostMessageW(g_hWndMain, WM_KEYBOARD_HOOK_STATUS, FALSE, GetLastError());
            break;
        }
        /* Bound each drain so a stream of input cannot starve renewal or exit. */
        for (int i = 0; i < 64 &&
             PeekMessageW(&message, NULL, 0, 0, PM_REMOVE); i++) {
            if (message.message == WM_QUIT) goto cleanup;
            if (message.message == WM_HOTKEY &&
                message.wParam == ID_HOTKEY_PRINT_SCREEN) HandlePrintScreenHotkey();
            /* This thread owns no windows and never runs application/UI
               handlers. PeekMessage itself dispatches low-level hook calls. */
        }
    }

cleanup:
    if (g_printHotkeyRegistered) UnregisterHotKey(NULL, ID_HOTKEY_PRINT_SCREEN);
    g_printHotkeyRegistered = FALSE;
    g_printHotkeyError = 0;
    g_hookInstallError = 0;
    if (g_hHook) UnhookWindowsHookEx(g_hHook);
    g_hHook = NULL;
    g_printScreenKeyDown = FALSE;
    g_escapeKeyDown = FALSE;
    return 0;
}

static void StartKeyboardHook(void)
{
    g_keyboardHookStopEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
    if (!g_keyboardHookStopEvent) {
        LogMessage("ERROR: Could not create keyboard hook stop event (%lu)", GetLastError());
        return;
    }
    g_keyboardHookWakeEvent = CreateEventW(NULL, FALSE, FALSE, NULL);
    if (!g_keyboardHookWakeEvent) {
        LogMessage("ERROR: Could not create keyboard hook wake event (%lu)", GetLastError());
        CloseHandle(g_keyboardHookStopEvent);
        g_keyboardHookStopEvent = NULL;
        return;
    }
    g_keyboardHookThread = CreateThread(
        NULL, 0, KeyboardHookThreadProc, NULL, 0, NULL);
    if (!g_keyboardHookThread) {
        LogMessage("ERROR: Could not start keyboard hook thread (%lu)", GetLastError());
        CloseHandle(g_keyboardHookWakeEvent);
        CloseHandle(g_keyboardHookStopEvent);
        g_keyboardHookWakeEvent = NULL;
        g_keyboardHookStopEvent = NULL;
    }
}

static void StopKeyboardHook(void)
{
    if (!g_keyboardHookThread) return;
    SetEvent(g_keyboardHookStopEvent);
    /* Join before destroying windows/cache that the callback can inspect.
       The hook never synchronously calls the UI, so there is no wait cycle. */
    WaitForSingleObject(g_keyboardHookThread, INFINITE);
    CloseHandle(g_keyboardHookThread);
    CloseHandle(g_keyboardHookWakeEvent);
    CloseHandle(g_keyboardHookStopEvent);
    g_keyboardHookThread = NULL;
    g_keyboardHookWakeEvent = NULL;
    g_keyboardHookStopEvent = NULL;
}

/* A hook can queue work while the UI is busy. Never paste that request into
   a subsequently focused window or replace a newer clipboard item. */
static BOOL IsPasteRequestCurrent(HWND target, DWORD clipboardSequence)
{
    return target && GetForegroundWindow() == target &&
           GetClipboardSequenceNumber() == clipboardSequence;
}

/* ── System tray icon ──────────────────────────────────────────────────── */

static void InitTrayIcon(HWND hwnd)
{
    ZeroMemory(&g_nid, sizeof(g_nid));
    g_nid.cbSize = sizeof(NOTIFYICONDATAW);
    g_nid.hWnd = hwnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = g_hAppIcon;
    wcscpy(g_nid.szTip, L"ImagePaster");
    Shell_NotifyIconW(NIM_ADD, &g_nid);
}

static void UpdateTooltip(void)
{
    WCHAR titleMatch[sizeof(g_configTitleMatch)] = L"None";
    WCHAR httpServer[32] = L"Disabled";
    WCHAR prefix[64];
    WCHAR suffix[96];
    size_t tipCapacity = sizeof(g_nid.szTip) / sizeof(g_nid.szTip[0]);
    size_t fixedLength;
    size_t titleCapacity;
    size_t titleLength;

    if (!g_nid.hWnd) return;

    if (g_configTitleMatch[0] != '\0' && g_keywordCount > 0) {
        if (MultiByteToWideChar(
                CP_UTF8, 0, g_configTitleMatch, -1, titleMatch,
                sizeof(titleMatch) / sizeof(titleMatch[0])) <= 0) {
            wcscpy(titleMatch, L"None");
        }
    }
    for (WCHAR *character = titleMatch; *character; character++) {
        if (*character == L'\r' || *character == L'\n') {
            *character = L' ';
        }
    }
    if (g_httpState == 1) {
        swprintf(httpServer,
                 sizeof(httpServer) / sizeof(httpServer[0]),
                 L"%hs:%d", g_configBindIp, g_configHttpPort);
    }

    swprintf(prefix, sizeof(prefix) / sizeof(prefix[0]),
             L"Paste Method: %s\nTitle Match: ",
             g_configPasteMethod == PASTE_METHOD_HTTP ? L"HTTP" : L"Base64");
    swprintf(suffix, sizeof(suffix) / sizeof(suffix[0]),
             L"\nScreen Capture: %s\nHTTP Server: %s",
             g_configScreenCaptureEnabled ? L"Enabled" : L"Disabled",
             httpServer);

    fixedLength = wcslen(prefix) + wcslen(suffix);
    titleCapacity = fixedLength < tipCapacity - 1
        ? tipCapacity - 1 - fixedLength : 0;
    titleLength = wcslen(titleMatch);
    if (titleLength > titleCapacity) {
        if (titleCapacity >= 3) {
            size_t ellipsisAt = titleCapacity - 3;
            if (ellipsisAt > 0 &&
                titleMatch[ellipsisAt - 1] >= 0xd800 &&
                titleMatch[ellipsisAt - 1] <= 0xdbff) {
                ellipsisAt--;
            }
            titleMatch[ellipsisAt] = L'.';
            titleMatch[ellipsisAt + 1] = L'.';
            titleMatch[ellipsisAt + 2] = L'.';
            titleMatch[ellipsisAt + 3] = L'\0';
        } else {
            titleMatch[titleCapacity] = L'\0';
        }
    }

    swprintf(g_nid.szTip, tipCapacity, L"%s%s%s",
             prefix, titleMatch, suffix);
    g_nid.szTip[tipCapacity - 1] = L'\0';
    g_nid.uFlags = NIF_TIP;
    Shell_NotifyIconW(NIM_MODIFY, &g_nid);
}

static void CreateContextMenu(void)
{
    g_hMenu = CreatePopupMenu();
    AppendMenuW(g_hMenu, MF_STRING, ID_TRAY_CAPTURE, L"Capture");
    AppendMenuW(g_hMenu, MF_STRING, ID_TRAY_HISTORY, L"History");
    AppendMenuW(g_hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(g_hMenu, MF_STRING, ID_TRAY_LOG, L"Activity Log");
    AppendMenuW(g_hMenu, MF_STRING, ID_TRAY_CONFIGURE, L"Configure");
    AppendMenuW(g_hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenuW(g_hMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");
}

/* ── WebView2 helper functions ─────────────────────────────────────────── */

static BOOL load_webview2_loader(void)
{
    HRSRC hRes = FindResource(NULL, MAKEINTRESOURCE(IDR_WEBVIEW2_DLL), RT_RCDATA);
    if (!hRes) {
        MessageBoxW(NULL, L"Failed to find WebView2Loader.dll in embedded resources.",
            APP_NAME, MB_ICONERROR);
        return FALSE;
    }
    HGLOBAL hData = LoadResource(NULL, hRes);
    DWORD dllSize = SizeofResource(NULL, hRes);
    const void *dllBytes = LockResource(hData);
    if (!dllBytes || dllSize == 0) {
        MessageBoxW(NULL, L"Failed to load WebView2Loader.dll from embedded resources.",
            APP_NAME, MB_ICONERROR);
        return FALSE;
    }
    WCHAR tempDir[MAX_PATH];
    DWORD tempLen = GetTempPathW(MAX_PATH, tempDir);
    if (tempLen == 0 || tempLen >= MAX_PATH - 50) {
        MessageBoxW(NULL, L"Failed to get temp directory path.", APP_NAME, MB_ICONERROR);
        return FALSE;
    }
    swprintf(g_extractedDllPath, MAX_PATH, L"%sImagePaster", tempDir);
    CreateDirectoryW(g_extractedDllPath, NULL);
    swprintf(g_extractedDllPath, MAX_PATH, L"%sImagePaster\\WebView2Loader.dll", tempDir);

    HMODULE hMod = LoadLibraryW(g_extractedDllPath);
    if (!hMod) {
        HANDLE hFile = CreateFileW(g_extractedDllPath, GENERIC_WRITE, 0, NULL,
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile == INVALID_HANDLE_VALUE) {
            MessageBoxW(NULL, L"Failed to write WebView2Loader.dll to temp directory.",
                APP_NAME, MB_ICONERROR);
            return FALSE;
        }
        DWORD written = 0;
        WriteFile(hFile, dllBytes, dllSize, &written, NULL);
        CloseHandle(hFile);
        if (written != dllSize) {
            MessageBoxW(NULL, L"Failed to write complete WebView2Loader.dll.",
                APP_NAME, MB_ICONERROR);
            return FALSE;
        }
        hMod = LoadLibraryW(g_extractedDllPath);
    }
    if (!hMod) {
        MessageBoxW(NULL, L"Failed to load WebView2Loader.dll.", APP_NAME, MB_ICONERROR);
        return FALSE;
    }
    fnCreateEnvironment = (PFN_CreateCoreWebView2EnvironmentWithOptions)
        GetProcAddress(hMod, "CreateCoreWebView2EnvironmentWithOptions");
    if (!fnCreateEnvironment) {
        MessageBoxW(NULL, L"WebView2Loader.dll loaded but entry point not found.",
            APP_NAME, MB_ICONERROR);
        return FALSE;
    }
    return TRUE;
}

static void webview_execute_script(const wchar_t* script)
{
    if (g_webviewView) {
        g_webviewView->lpVtbl->ExecuteScript(g_webviewView, script, NULL);
    }
}

static void webview_sync_controller_bounds(void)
{
    if (!g_webviewController || !g_webviewHwnd) return;
    RECT bounds;
    GetClientRect(g_webviewHwnd, &bounds);
    g_webviewController->lpVtbl->put_Bounds(g_webviewController, bounds);
    g_webviewController->lpVtbl->put_IsVisible(g_webviewController, TRUE);
}

/* The manifest makes the process per-monitor DPI aware, so WebView-reported
 * CSS dimensions must be converted to the physical pixels used by Win32. */
typedef UINT (WINAPI *PFN_GetDpiForWindow)(HWND);

static UINT GetWebViewWindowDpi(HWND hwnd)
{
    static PFN_GetDpiForWindow fnGetDpiForWindow = NULL;
    static BOOL resolved = FALSE;
    HDC hdc;
    UINT dpi;

    if (!resolved) {
        fnGetDpiForWindow = (PFN_GetDpiForWindow)GetProcAddress(
            GetModuleHandleW(L"user32.dll"), "GetDpiForWindow");
        resolved = TRUE;
    }
    if (fnGetDpiForWindow && hwnd) {
        dpi = fnGetDpiForWindow(hwnd);
        if (dpi) return dpi;
    }
    hdc = GetDC(hwnd);
    dpi = hdc ? (UINT)GetDeviceCaps(hdc, LOGPIXELSX) : 96;
    if (hdc) ReleaseDC(hwnd, hdc);
    return dpi ? dpi : 96;
}

static void UpdateDebugPrint(const wchar_t *format, ...)
{
    wchar_t buffer[1024];
    va_list args;
    va_start(args, format);
    vswprintf_s(buffer, sizeof(buffer) / sizeof(wchar_t), format, args);
    va_end(args);
    OutputDebugStringW(buffer);
}

/* ── Self update ───────────────────────────────────────────────────────── */

typedef struct {
    HINTERNET session;
    HINTERNET connection;
    HINTERNET request;
} UpdateHttpRequest;

static void CloseUpdateHttpRequest(UpdateHttpRequest* http) {
    if (!http) return;
    if (http->request) WinHttpCloseHandle(http->request);
    if (http->connection) WinHttpCloseHandle(http->connection);
    if (http->session) WinHttpCloseHandle(http->session);
    ZeroMemory(http, sizeof(*http));
}

static void SetUpdateTaskError(UpdateCheckTask* task, LPCWSTR message,
                               DWORD errorCode) {
    if (!task) return;
    task->kind = UPDATE_CHECK_ERROR;
    if (errorCode) {
        swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                   L"%s (Windows error %lu).", message, (unsigned long)errorCode);
    } else {
        wcscpy_s(task->message, sizeof(task->message) / sizeof(wchar_t), message);
    }
}

static BOOL CancelUpdateTaskIfRequested(UpdateCheckTask* task) {
    if (!task || !g_updateCancelEvent ||
        WaitForSingleObject(g_updateCancelEvent, 0) != WAIT_OBJECT_0) {
        return FALSE;
    }
    task->kind = UPDATE_CHECK_CANCELLED;
    task->message[0] = L'\0';
    return TRUE;
}

static void PublishUpdateProgress(UpdateCheckTask* task, DWORD percent) {
    if (!task || CancelUpdateTaskIfRequested(task) ||
        !IsWindow(task->targetWindow)) {
        return;
    }

    InterlockedExchange(&g_updateProgressPercent, (LONG)percent);
    if (InterlockedCompareExchange(&g_updateProgressPosted, TRUE, FALSE) == FALSE &&
        !PostMessageW(task->targetWindow, WM_APP_UPDATE_PROGRESS, 0, 0)) {
        InterlockedExchange(&g_updateProgressPosted, FALSE);
    }
}

static DWORD CalculateUpdateProgressPercent(ULONGLONG receivedBytes,
                                            ULONGLONG totalBytes) {
    if (!totalBytes) return 0;
    if (receivedBytes >= totalBytes) return 100;
    // Whole percent completed; 100 is reserved for the final byte.
    DWORD percent = (DWORD)((double)receivedBytes * 100.0 / (double)totalBytes);
    return percent > 99 ? 99 : percent;
}

static BOOL OpenUpdateHttpRequest(LPCWSTR verb, ULONGLONG cacheBuster,
                                  UpdateHttpRequest* http, DWORD* statusCode) {
    if (!verb || !http) return FALSE;
    ZeroMemory(http, sizeof(*http));
    if (statusCode) *statusCode = 0;

    wchar_t hostName[256] = L"";
    wchar_t urlPath[2048] = L"";
    wchar_t extraInfo[512] = L"";
    URL_COMPONENTS components = {0};
    components.dwStructSize = sizeof(components);
    components.lpszHostName = hostName;
    components.dwHostNameLength = sizeof(hostName) / sizeof(wchar_t);
    components.lpszUrlPath = urlPath;
    components.dwUrlPathLength = sizeof(urlPath) / sizeof(wchar_t);
    components.lpszExtraInfo = extraInfo;
    components.dwExtraInfoLength = sizeof(extraInfo) / sizeof(wchar_t);
    if (!WinHttpCrackUrl(UPDATE_URL, 0, 0, &components)) return FALSE;

    wchar_t objectName[2560];
    if (wcscpy_s(objectName, sizeof(objectName) / sizeof(wchar_t), urlPath) != 0 ||
        wcscat_s(objectName, sizeof(objectName) / sizeof(wchar_t), extraInfo) != 0) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    // A unique query value makes each button click reach the current branch
    // artifact even when an HTTP proxy or GitHub edge cache retains the
    // previous response. HEAD and GET share the same value within one check.
    wchar_t cacheSuffix[64];
    wchar_t separator = wcschr(objectName, L'?') ? L'&' : L'?';
    int cacheSuffixLength = swprintf_s(cacheSuffix,
        sizeof(cacheSuffix) / sizeof(wchar_t), L"%lcipUpdate=%016llx",
        separator, (unsigned long long)cacheBuster);
    if (cacheSuffixLength <= 0 ||
        wcscat_s(objectName, sizeof(objectName) / sizeof(wchar_t), cacheSuffix) != 0) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    http->session = WinHttpOpen(L"ImagePaster Update",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!http->session) goto fail;
    WinHttpSetTimeouts(http->session, 10000, 10000, 15000, 30000);

    http->connection = WinHttpConnect(http->session, hostName,
                                      components.nPort, 0);
    if (!http->connection) goto fail;

    DWORD flags = WINHTTP_FLAG_REFRESH;
    if (components.nScheme == INTERNET_SCHEME_HTTPS) flags |= WINHTTP_FLAG_SECURE;
    http->request = WinHttpOpenRequest(http->connection, verb, objectName,
        NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!http->request) goto fail;

    static const wchar_t noCacheHeaders[] =
        L"Cache-Control: no-cache, no-store, max-age=0\r\nPragma: no-cache\r\n";
    WinHttpAddRequestHeaders(http->request, noCacheHeaders, (DWORD)-1L,
                            WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE);
    if (!WinHttpSendRequest(http->request, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) ||
        !WinHttpReceiveResponse(http->request, NULL)) {
        goto fail;
    }

    DWORD status = 0;
    DWORD statusSize = sizeof(status);
    if (!WinHttpQueryHeaders(http->request,
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize,
            WINHTTP_NO_HEADER_INDEX)) {
        goto fail;
    }
    if (statusCode) *statusCode = status;
    if (status != 200) {
        CloseUpdateHttpRequest(http);
        SetLastError(ERROR_WINHTTP_INVALID_SERVER_RESPONSE);
        return FALSE;
    }
    return TRUE;

fail: {
        DWORD errorCode = GetLastError();
        CloseUpdateHttpRequest(http);
        SetLastError(errorCode);
        return FALSE;
    }
}

static BOOL QueryUpdateContentLength(HINTERNET request, ULONGLONG* size) {
    if (!request || !size) return FALSE;
    wchar_t lengthText[64] = L"";
    DWORD lengthBytes = sizeof(lengthText);
    if (!WinHttpQueryHeaders(request, WINHTTP_QUERY_CONTENT_LENGTH,
            WINHTTP_HEADER_NAME_BY_INDEX, lengthText, &lengthBytes,
            WINHTTP_NO_HEADER_INDEX)) {
        return FALSE;
    }
    lengthText[(sizeof(lengthText) / sizeof(wchar_t)) - 1] = L'\0';

    wchar_t* end = NULL;
    unsigned long long parsed = _wcstoui64(lengthText, &end, 10);
    if (end == lengthText || !end || *end != L'\0' || parsed == 0) {
        SetLastError(ERROR_INVALID_DATA);
        return FALSE;
    }
    *size = (ULONGLONG)parsed;
    return TRUE;
}

static BOOL GetExecutableVersion(LPCWSTR path, ExecutableVersion* version) {
    if (!path || !*path || !version) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    DWORD ignored = 0;
    DWORD infoSize = GetFileVersionInfoSizeW(path, &ignored);
    if (infoSize == 0) {
        DWORD errorCode = GetLastError();
        SetLastError(errorCode ? errorCode : ERROR_RESOURCE_DATA_NOT_FOUND);
        return FALSE;
    }

    BYTE* infoData = (BYTE*)malloc(infoSize);
    if (!infoData) {
        SetLastError(ERROR_NOT_ENOUGH_MEMORY);
        return FALSE;
    }

    if (!GetFileVersionInfoW(path, 0, infoSize, infoData)) {
        DWORD errorCode = GetLastError();
        free(infoData);
        SetLastError(errorCode ? errorCode : ERROR_INVALID_DATA);
        return FALSE;
    }

    VS_FIXEDFILEINFO* fixedInfo = NULL;
    UINT fixedInfoSize = 0;
    if (!VerQueryValueW(infoData, L"\\", (LPVOID*)&fixedInfo, &fixedInfoSize) ||
        !fixedInfo || fixedInfoSize < sizeof(*fixedInfo) ||
        fixedInfo->dwSignature != VS_FFI_SIGNATURE) {
        free(infoData);
        SetLastError(ERROR_INVALID_DATA);
        return FALSE;
    }

    version->major = HIWORD(fixedInfo->dwFileVersionMS);
    version->minor = LOWORD(fixedInfo->dwFileVersionMS);
    version->patch = HIWORD(fixedInfo->dwFileVersionLS);
    version->build = LOWORD(fixedInfo->dwFileVersionLS);
    free(infoData);
    return TRUE;
}

static int CompareExecutableVersions(const ExecutableVersion* left,
                                     const ExecutableVersion* right) {
    const WORD leftParts[] = {
        left->major, left->minor, left->patch, left->build
    };
    const WORD rightParts[] = {
        right->major, right->minor, right->patch, right->build
    };
    for (size_t index = 0; index < sizeof(leftParts) / sizeof(leftParts[0]); index++) {
        if (leftParts[index] < rightParts[index]) return -1;
        if (leftParts[index] > rightParts[index]) return 1;
    }
    return 0;
}

static void FormatExecutableVersion(const ExecutableVersion* version,
                                    wchar_t* text, size_t textCch) {
    if (!version || !text || textCch == 0) return;
    if (swprintf_s(text, textCch, L"%u.%u.%u.%u",
                   (unsigned int)version->major,
                   (unsigned int)version->minor,
                   (unsigned int)version->patch,
                   (unsigned int)version->build) <= 0) {
        text[0] = L'\0';
    }
}

static void FormatExecutableVersionForDisplay(const ExecutableVersion* version,
                                              wchar_t* text, size_t textCch) {
    if (!version || !text || textCch == 0) return;
    if (version->build == 0) {
        if (swprintf_s(text, textCch, L"%u.%u.%u",
                       (unsigned int)version->major,
                       (unsigned int)version->minor,
                       (unsigned int)version->patch) <= 0) {
            text[0] = L'\0';
        }
        return;
    }
    FormatExecutableVersion(version, text, textCch);
}

static BOOL BuildUpdateTempPath(wchar_t path[MAX_PATH], LPCWSTR role,
                                DWORD processId) {
    if (!path || !role || !*role || processId == 0) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    wchar_t tempDirectory[MAX_PATH];
    DWORD tempLength = GetTempPathW(MAX_PATH, tempDirectory);
    if (tempLength == 0) return FALSE;
    if (tempLength >= MAX_PATH) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    int length = swprintf_s(path, MAX_PATH, L"%s%s-%s-%lu.exe",
                            tempDirectory, APP_NAME, role,
                            (unsigned long)processId);
    if (length <= 0 || length >= MAX_PATH) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }
    return TRUE;
}

static BOOL DeleteUpdateTempFile(LPCWSTR path) {
    if (!path || !*path) return FALSE;
    SetFileAttributesW(path, FILE_ATTRIBUTE_NORMAL);
    if (DeleteFileW(path)) return TRUE;
    DWORD errorCode = GetLastError();
    return errorCode == ERROR_FILE_NOT_FOUND || errorCode == ERROR_PATH_NOT_FOUND;
}

static BOOL QueryRemoteUpdateSize(UpdateCheckTask* task, ULONGLONG* size) {
    if (CancelUpdateTaskIfRequested(task)) return FALSE;

    UpdateHttpRequest http;
    DWORD status = 0;
    if (!OpenUpdateHttpRequest(L"HEAD", task->cacheBuster, &http, &status)) {
        DWORD errorCode = GetLastError();
        if (CancelUpdateTaskIfRequested(task)) return FALSE;
        if (status) {
            swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                       L"The update server returned HTTP status %lu.",
                       (unsigned long)status);
            task->kind = UPDATE_CHECK_ERROR;
        } else {
            SetUpdateTaskError(task, L"Could not contact the update server", errorCode);
        }
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    BOOL ok = QueryUpdateContentLength(http.request, size);
    DWORD errorCode = ok ? ERROR_SUCCESS : GetLastError();
    CloseUpdateHttpRequest(&http);
    if (CancelUpdateTaskIfRequested(task)) return FALSE;
    if (!ok) {
        SetUpdateTaskError(task, L"The update server did not report a valid file size",
                           errorCode);
        return FALSE;
    }
    if (*size > UPDATE_MAX_BYTES) {
        SetUpdateTaskError(task, L"The available update is unexpectedly large", 0);
        return FALSE;
    }
    return TRUE;
}

static BOOL DownloadUpdateFile(UpdateCheckTask* task, ULONGLONG expectedSize) {
    if (CancelUpdateTaskIfRequested(task)) return FALSE;

    UpdateHttpRequest http;
    DWORD status = 0;
    if (!OpenUpdateHttpRequest(L"GET", task->cacheBuster, &http, &status)) {
        DWORD errorCode = GetLastError();
        if (CancelUpdateTaskIfRequested(task)) return FALSE;
        if (status) {
            swprintf_s(task->message, sizeof(task->message) / sizeof(wchar_t),
                       L"The update download returned HTTP status %lu.",
                       (unsigned long)status);
            task->kind = UPDATE_CHECK_ERROR;
        } else {
            SetUpdateTaskError(task, L"Could not download the update", errorCode);
        }
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    ULONGLONG downloadSize = 0;
    if (QueryUpdateContentLength(http.request, &downloadSize) &&
        downloadSize != expectedSize) {
        CloseUpdateHttpRequest(&http);
        SetUpdateTaskError(task,
            L"The available update changed while it was being downloaded. Try again", 0);
        return FALSE;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        CloseUpdateHttpRequest(&http);
        return FALSE;
    }

    DeleteUpdateTempFile(task->stagedPath);
    HANDLE file = CreateFileW(task->stagedPath, GENERIC_WRITE, 0, NULL,
                              CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) {
        DWORD errorCode = GetLastError();
        CloseUpdateHttpRequest(&http);
        SetUpdateTaskError(task, L"Could not create the staged update file", errorCode);
        return FALSE;
    }

    BOOL ok = TRUE;
    ULONGLONG totalWritten = 0;
    DWORD reportedPercent = 0;
    PublishUpdateProgress(task, reportedPercent);
    BYTE buffer[64 * 1024];
    while (ok) {
        if (CancelUpdateTaskIfRequested(task)) {
            ok = FALSE;
            break;
        }

        DWORD bytesRead = 0;
        if (!WinHttpReadData(http.request, buffer, sizeof(buffer), &bytesRead)) {
            DWORD errorCode = GetLastError();
            if (!CancelUpdateTaskIfRequested(task)) {
                SetUpdateTaskError(task, L"The update download was interrupted", errorCode);
            }
            ok = FALSE;
            break;
        }
        if (CancelUpdateTaskIfRequested(task)) {
            ok = FALSE;
            break;
        }
        if (bytesRead == 0) break;
        if (totalWritten + bytesRead > expectedSize) {
            SetUpdateTaskError(task, L"The downloaded update has an invalid size", 0);
            ok = FALSE;
            break;
        }

        DWORD bytesWritten = 0;
        if (!WriteFile(file, buffer, bytesRead, &bytesWritten, NULL)) {
            SetUpdateTaskError(task, L"Could not write the staged update", GetLastError());
            ok = FALSE;
            break;
        }
        if (bytesWritten != bytesRead) {
            SetUpdateTaskError(task, L"Could not write the staged update",
                               ERROR_WRITE_FAULT);
            ok = FALSE;
            break;
        }
        totalWritten += bytesWritten;

        // The size was validated above, so progress is a share of the whole
        // file. Only whole-percent changes are posted to the UI thread.
        DWORD percent = CalculateUpdateProgressPercent(totalWritten, expectedSize);
        if (percent != reportedPercent) {
            PublishUpdateProgress(task, percent);
            reportedPercent = percent;
        }
    }

    if (ok && CancelUpdateTaskIfRequested(task)) ok = FALSE;
    if (ok && totalWritten != expectedSize) {
        SetUpdateTaskError(task, L"The downloaded update is incomplete", 0);
        ok = FALSE;
    }
    if (ok && !FlushFileBuffers(file)) {
        SetUpdateTaskError(task, L"Could not finish writing the staged update", GetLastError());
        ok = FALSE;
    }
    CloseHandle(file);
    CloseUpdateHttpRequest(&http);

    if (ok && CancelUpdateTaskIfRequested(task)) ok = FALSE;
    DWORD binaryType = 0;
    if (ok && (!GetBinaryTypeW(task->stagedPath, &binaryType) ||
               binaryType != SCS_64BIT_BINARY)) {
        SetUpdateTaskError(task, L"The downloaded file is not a valid 64-bit application", 0);
        ok = FALSE;
    }
    if (!ok) DeleteUpdateTempFile(task->stagedPath);
    return ok;
}

static void DiscardUpdateTask(UpdateCheckTask* task) {
    if (!task) return;
    if (task->stagedPath[0]) {
        DeleteUpdateTempFile(task->stagedPath);
    }
    free(task);
}

static void PublishUpdateTask(UpdateCheckTask* task) {
    CancelUpdateTaskIfRequested(task);
    HWND targetWindow = task ? task->targetWindow : NULL;
    if (task && IsWindow(targetWindow)) {
        // Hand the result over before clearing the pending flag, so the UI
        // thread never sees an idle updater while this result is in flight.
        UpdateCheckTask* previous = (UpdateCheckTask*)InterlockedExchangePointer(
            (PVOID volatile*)&g_updatePostedResult, task);
        DiscardUpdateTask(previous);
    } else {
        DiscardUpdateTask(task);
        targetWindow = NULL;
    }
    InterlockedExchange(&g_updateCheckPending, FALSE);
    InterlockedExchange(&g_updateCheckAutomatic, FALSE);

    if (targetWindow &&
        !PostMessageW(targetWindow, WM_APP_UPDATE_RESULT, 0, 0)) {
        UpdateCheckTask* unclaimed = (UpdateCheckTask*)InterlockedExchangePointer(
            (PVOID volatile*)&g_updatePostedResult, NULL);
        DiscardUpdateTask(unclaimed);
    }
}

static DWORD WINAPI UpdateCheckThread(LPVOID parameter) {
    UpdateCheckTask* task = (UpdateCheckTask*)parameter;
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    DWORD pathLength = GetModuleFileNameW(NULL, task->targetPath,
                                         sizeof(task->targetPath) / sizeof(wchar_t));
    if (pathLength == 0 || pathLength >= sizeof(task->targetPath) / sizeof(wchar_t)) {
        SetUpdateTaskError(task, L"Could not determine the running executable path",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }

    if (!GetExecutableVersion(task->targetPath, &task->runningVersion)) {
        SetUpdateTaskError(task, L"Could not read the running application version",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    ULONGLONG remoteSize = 0;
    if (!QueryRemoteUpdateSize(task, &remoteSize)) {
        PublishUpdateTask(task);
        return 0;
    }
    if (!BuildUpdateTempPath(task->stagedPath, L"download",
                             GetCurrentProcessId())) {
        SetUpdateTaskError(task, L"Could not create the temporary update path",
                           GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    if (!DownloadUpdateFile(task, remoteSize)) {
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    if (!GetExecutableVersion(task->stagedPath, &task->availableVersion)) {
        SetUpdateTaskError(task,
            L"The downloaded application does not contain valid version information",
            GetLastError());
        PublishUpdateTask(task);
        return 0;
    }
    if (CancelUpdateTaskIfRequested(task)) {
        PublishUpdateTask(task);
        return 0;
    }

    UpdateDebugPrint(L"[INFO] Update versions: running %u.%u.%u.%u, available %u.%u.%u.%u\n",
               (unsigned int)task->runningVersion.major,
               (unsigned int)task->runningVersion.minor,
               (unsigned int)task->runningVersion.patch,
               (unsigned int)task->runningVersion.build,
               (unsigned int)task->availableVersion.major,
               (unsigned int)task->availableVersion.minor,
               (unsigned int)task->availableVersion.patch,
               (unsigned int)task->availableVersion.build);
    int comparison = CompareExecutableVersions(&task->availableVersion,
                                               &task->runningVersion);
    task->kind = comparison > 0 ? UPDATE_CHECK_NEWER
               : comparison < 0 ? UPDATE_CHECK_OLDER
                                : UPDATE_CHECK_SAME;
    PublishUpdateTask(task);
    return 0;
}

static BOOL ParseUpdateProcessId(LPCWSTR text, DWORD* processId) {
    if (!text || !processId || !*text) return FALSE;
    wchar_t* end = NULL;
    unsigned long value = wcstoul(text, &end, 10);
    if (!end || *end != L'\0' || value == 0) return FALSE;
    *processId = (DWORD)value;
    return TRUE;
}

static BOOL ValidateUpdateTempFilePair(LPCWSTR helperPath,
                                       LPCWSTR stagedPath,
                                       DWORD processId) {
    if (!helperPath || !stagedPath || !*helperPath || !*stagedPath) return FALSE;

    wchar_t expectedHelperName[96], expectedStagedName[96];
    int helperNameLength = swprintf_s(expectedHelperName,
        sizeof(expectedHelperName) / sizeof(wchar_t),
        APP_NAME L"-updater-%lu.exe", (unsigned long)processId);
    int stagedNameLength = swprintf_s(expectedStagedName,
        sizeof(expectedStagedName) / sizeof(wchar_t),
        APP_NAME L"-download-%lu.exe", (unsigned long)processId);
    if (helperNameLength <= 0 || stagedNameLength <= 0 ||
        _wcsicmp(PathFindFileNameW(helperPath), expectedHelperName) != 0 ||
        _wcsicmp(PathFindFileNameW(stagedPath), expectedStagedName) != 0) {
        return FALSE;
    }

    wchar_t helperDirectory[MAX_PATH], stagedDirectory[MAX_PATH];
    if (wcscpy_s(helperDirectory, MAX_PATH, helperPath) != 0 ||
        wcscpy_s(stagedDirectory, MAX_PATH, stagedPath) != 0 ||
        !PathRemoveFileSpecW(helperDirectory) ||
        !PathRemoveFileSpecW(stagedDirectory)) {
        return FALSE;
    }
    return _wcsicmp(helperDirectory, stagedDirectory) == 0;
}

static HANDLE DuplicateUpdateLaunchToken(HANDLE process) {
    HANDLE processToken = NULL;
    HANDLE launchToken = NULL;
    if (!process ||
        !OpenProcessToken(process, TOKEN_QUERY | TOKEN_DUPLICATE,
                          &processToken)) {
        return NULL;
    }
    DuplicateTokenEx(processToken, MAXIMUM_ALLOWED, NULL,
                     SecurityImpersonation, TokenPrimary, &launchToken);
    CloseHandle(processToken);
    return launchToken;
}

static BOOL LaunchUpdateTarget(LPCWSTR targetPath, LPCWSTR stagedPath,
                               LPCWSTR helperPath, DWORD helperProcessId,
                               DWORD oldProcessId, HANDLE launchToken,
                               BOOL successfulUpdate, BOOL reopenSettings) {
    wchar_t commandLine[MAX_PATH * 3 + 256];
    LPCWSTR finishAction = successfulUpdate
        ? L"--finish-update"
        : L"--finish-update-cleanup";
    int commandLength = swprintf_s(commandLine,
        sizeof(commandLine) / sizeof(wchar_t),
        L"\"%s\" %s %lu %lu \"%s\" \"%s\"%s", targetPath, finishAction,
        (unsigned long)helperProcessId, (unsigned long)oldProcessId,
        stagedPath, helperPath,
        successfulUpdate && reopenSettings ? L" --reopen-settings" : L"");
    if (commandLength <= 0 ||
        commandLength >= (int)(sizeof(commandLine) / sizeof(wchar_t))) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    STARTUPINFOW startupInfo = {0};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo = {0};
    BOOL launched = FALSE;
    if (launchToken) {
        wchar_t tokenCommandLine[MAX_PATH * 3 + 256];
        wcscpy_s(tokenCommandLine,
                 sizeof(tokenCommandLine) / sizeof(wchar_t), commandLine);
        LPVOID environment = NULL;
        BOOL hasEnvironment = CreateEnvironmentBlock(&environment, launchToken,
                                                     FALSE);
        launched = CreateProcessWithTokenW(launchToken, 0, targetPath,
                                           tokenCommandLine,
                                           hasEnvironment
                                               ? CREATE_UNICODE_ENVIRONMENT : 0,
                                           environment, NULL,
                                           &startupInfo, &processInfo);
        if (environment) DestroyEnvironmentBlock(environment);
    }
    if (!launched) {
        ZeroMemory(&processInfo, sizeof(processInfo));
        launched = CreateProcessW(targetPath, commandLine, NULL, NULL, FALSE,
                                  0, NULL, NULL, &startupInfo, &processInfo);
    }
    if (launched) {
        CloseHandle(processInfo.hProcess);
        CloseHandle(processInfo.hThread);
    }
    return launched;
}

static int RestartAfterUpdateFailure(LPCWSTR targetPath, LPCWSTR stagedPath,
                                     LPCWSTR helperPath, DWORD oldProcessId,
                                     HANDLE launchToken, LPCWSTR message) {
    MessageBoxW(NULL, message, APP_NAME L" Update", MB_OK | MB_ICONERROR);
    LaunchUpdateTarget(targetPath, stagedPath, helperPath,
                       GetCurrentProcessId(), oldProcessId, launchToken, FALSE, FALSE);
    if (launchToken) CloseHandle(launchToken);
    DeleteUpdateTempFile(stagedPath);
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
    MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    return 1;
}

static int RunUpdateApplyHelper(DWORD oldProcessId, LPCWSTR readyEventName,
                                LPCWSTR targetPath, LPCWSTR stagedPath,
                                BOOL reopenSettings) {
    wchar_t expectedEventPrefix[96];
    int prefixLength = swprintf_s(expectedEventPrefix,
        sizeof(expectedEventPrefix) / sizeof(wchar_t),
        L"Local\\ImagePaster_UpdateReady_%lu_",
        (unsigned long)oldProcessId);
    if (prefixLength <= 0 || !readyEventName ||
        _wcsnicmp(readyEventName, expectedEventPrefix,
                  (size_t)prefixLength) != 0) {
        return ERROR_INVALID_DATA;
    }

    HANDLE readyEvent = OpenEventW(EVENT_MODIFY_STATE, FALSE, readyEventName);
    if (!readyEvent) return (int)GetLastError();

    HANDLE oldProcess = OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                    FALSE, oldProcessId);
    if (!oldProcess) {
        DWORD errorCode = GetLastError();
        CloseHandle(readyEvent);
        return (int)errorCode;
    }

    wchar_t oldProcessPath[MAX_PATH];
    DWORD oldProcessPathLength = sizeof(oldProcessPath) / sizeof(wchar_t);
    if (!QueryFullProcessImageNameW(oldProcess, 0, oldProcessPath,
                                    &oldProcessPathLength)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return (int)errorCode;
    }
    if (_wcsicmp(oldProcessPath, targetPath) != 0) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_INVALID_DATA;
    }

    wchar_t helperPath[MAX_PATH];
    DWORD helperPathLength = GetModuleFileNameW(NULL, helperPath,
                                               sizeof(helperPath) / sizeof(wchar_t));
    DWORD binaryType = 0;
    if (helperPathLength == 0 || helperPathLength >= MAX_PATH ||
        !ValidateUpdateTempFilePair(helperPath, stagedPath, oldProcessId)) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_INVALID_DATA;
    }
    if (!GetBinaryTypeW(stagedPath, &binaryType)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return (int)errorCode;
    }
    if (binaryType != SCS_64BIT_BINARY) {
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        return ERROR_BAD_EXE_FORMAT;
    }

    // The helper is elevated only for file replacement. Preserve a primary
    // token from the original process so the restarted launcher normally
    // returns to the user's non-elevated session.
    HANDLE launchToken = DuplicateUpdateLaunchToken(oldProcess);

    // Only let the parent exit once this helper has verified every path and
    // owns the process handle it must wait on.
    if (!SetEvent(readyEvent)) {
        DWORD errorCode = GetLastError();
        CloseHandle(oldProcess);
        CloseHandle(readyEvent);
        if (launchToken) CloseHandle(launchToken);
        return (int)errorCode;
    }
    CloseHandle(readyEvent);

    DWORD waitResult = WaitForSingleObject(oldProcess, UPDATE_HELPER_WAIT_MS);
    CloseHandle(oldProcess);
    if (waitResult != WAIT_OBJECT_0) {
        if (launchToken) CloseHandle(launchToken);
        DeleteUpdateTempFile(stagedPath);
        SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
        MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
        MessageBoxW(NULL, L"The running application did not close in time.",
                    APP_NAME L" Update", MB_OK | MB_ICONERROR);
        return ERROR_TIMEOUT;
    }

    wchar_t replacementPath[MAX_PATH], backupPath[MAX_PATH];
    DWORD helperProcessId = GetCurrentProcessId();
    int replacementLength = swprintf_s(replacementPath,
        sizeof(replacementPath) / sizeof(wchar_t), L"%s.new.%lu.exe",
        targetPath, (unsigned long)helperProcessId);
    int backupLength = swprintf_s(backupPath,
        sizeof(backupPath) / sizeof(wchar_t), L"%s.backup.%lu.exe",
        targetPath, (unsigned long)helperProcessId);
    if (replacementLength <= 0 || replacementLength >= MAX_PATH ||
        backupLength <= 0 || backupLength >= MAX_PATH) {
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The update paths were too long. The previous version will restart.");
    }

    SetFileAttributesW(replacementPath, FILE_ATTRIBUTE_NORMAL);
    DeleteFileW(replacementPath);
    SetFileAttributesW(backupPath, FILE_ATTRIBUTE_NORMAL);
    DeleteFileW(backupPath);
    if (!CopyFileW(stagedPath, replacementPath, FALSE)) {
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The update could not be prepared. The previous version will restart.");
    }

    DWORD targetAttributes = GetFileAttributesW(targetPath);
    BOOL clearedReadOnly = FALSE;
    if (targetAttributes != INVALID_FILE_ATTRIBUTES &&
        (targetAttributes & FILE_ATTRIBUTE_READONLY)) {
        clearedReadOnly = SetFileAttributesW(
            targetPath, targetAttributes & ~FILE_ATTRIBUTE_READONLY);
    }
    if (!ReplaceFileW(targetPath, replacementPath, backupPath,
                      REPLACEFILE_WRITE_THROUGH, NULL, NULL)) {
        if (clearedReadOnly) SetFileAttributesW(targetPath, targetAttributes);
        DeleteFileW(replacementPath);
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The executable could not be replaced. The previous version will restart.");
    }

    if (!LaunchUpdateTarget(targetPath, stagedPath, helperPath,
                            helperProcessId, oldProcessId, launchToken,
                            TRUE, reopenSettings)) {
        DeleteFileW(targetPath);
        if (!MoveFileExW(backupPath, targetPath,
                         MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
            if (launchToken) CloseHandle(launchToken);
            DeleteUpdateTempFile(stagedPath);
            SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
            MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
            MessageBoxW(NULL,
                L"The updated application could not start and the previous executable "
                L"could not be restored. A backup remains beside the application.",
                APP_NAME L" Update", MB_OK | MB_ICONERROR);
            return 1;
        }
        return RestartAfterUpdateFailure(targetPath, stagedPath, helperPath,
            oldProcessId, launchToken,
            L"The updated application could not start. The previous version was restored.");
    }
    if (launchToken) CloseHandle(launchToken);

    if (!DeleteFileW(backupPath)) {
        MoveFileExW(backupPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    }
    DeleteUpdateTempFile(stagedPath);
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);
    MoveFileExW(helperPath, NULL, MOVEFILE_DELAY_UNTIL_REBOOT);
    return 0;
}

static BOOL FinishUpdateCleanup(DWORD helperProcessId, DWORD oldProcessId,
                                LPCWSTR stagedPath, LPCWSTR helperPath) {
    wchar_t targetPath[MAX_PATH], expectedStagedPath[MAX_PATH];
    wchar_t expectedHelperPath[MAX_PATH];
    DWORD targetLength = GetModuleFileNameW(NULL, targetPath,
                                           sizeof(targetPath) / sizeof(wchar_t));
    if (targetLength == 0 || targetLength >= MAX_PATH ||
        !BuildUpdateTempPath(expectedStagedPath, L"download", oldProcessId) ||
        !BuildUpdateTempPath(expectedHelperPath, L"updater", oldProcessId) ||
        _wcsicmp(stagedPath, expectedStagedPath) != 0 ||
        _wcsicmp(helperPath, expectedHelperPath) != 0) {
        return FALSE;
    }

    HANDLE helperProcess = OpenProcess(SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                       FALSE, helperProcessId);
    if (helperProcess) {
        wchar_t runningHelperPath[MAX_PATH];
        DWORD runningHelperPathLength = MAX_PATH;
        if (QueryFullProcessImageNameW(helperProcess, 0, runningHelperPath,
                                      &runningHelperPathLength) &&
            _wcsicmp(runningHelperPath, helperPath) == 0) {
            WaitForSingleObject(helperProcess, UPDATE_HELPER_WAIT_MS);
        }
        CloseHandle(helperProcess);
    }
    for (int attempt = 0;
         attempt < 20 && !DeleteUpdateTempFile(stagedPath);
         ++attempt) {
        Sleep(100);
    }
    for (int attempt = 0;
         attempt < 20 && !DeleteUpdateTempFile(helperPath);
         ++attempt) {
        Sleep(100);
    }
    return TRUE;
}

// Returns an exit code and sets handled for the temporary updater process.
// Both finish modes perform cleanup and continue normal application startup;
// updateCompleted identifies only a successful executable replacement.
static int HandleUpdateCommandLine(BOOL* handled, BOOL* updateCompleted,
                                   BOOL* reopenSettings) {
    if (handled) *handled = FALSE;
    if (updateCompleted) *updateCompleted = FALSE;
    if (reopenSettings) *reopenSettings = FALSE;
    int argumentCount = 0;
    LPWSTR* arguments = CommandLineToArgvW(GetCommandLineW(), &argumentCount);
    if (!arguments) return 0;

    // The optional choice belongs only to this handoff, never the registry.
    BOOL requestedReopen = argumentCount == 7 &&
        wcscmp(arguments[6], L"--reopen-settings") == 0;
    BOOL validArguments = argumentCount == 6 || requestedReopen;
    int result = 0;
    if (validArguments && wcscmp(arguments[1], L"--apply-update") == 0) {
        DWORD oldProcessId = 0;
        if (handled) *handled = TRUE;
        if (!ParseUpdateProcessId(arguments[2], &oldProcessId)) {
            result = ERROR_INVALID_PARAMETER;
        } else {
            result = RunUpdateApplyHelper(oldProcessId, arguments[3],
                                          arguments[4], arguments[5], requestedReopen);
        }
    } else if (validArguments &&
               (wcscmp(arguments[1], L"--finish-update") == 0 ||
                wcscmp(arguments[1], L"--finish-update-cleanup") == 0)) {
        DWORD helperProcessId = 0, oldProcessId = 0;
        if (ParseUpdateProcessId(arguments[2], &helperProcessId) &&
            ParseUpdateProcessId(arguments[3], &oldProcessId)) {
            BOOL recognizedHandoff = FinishUpdateCleanup(
                helperProcessId, oldProcessId, arguments[4], arguments[5]);
            if (recognizedHandoff && updateCompleted &&
                wcscmp(arguments[1], L"--finish-update") == 0) {
                *updateCompleted = TRUE;
                if (reopenSettings) *reopenSettings = requestedReopen;
            }
        }
    }
    LocalFree(arguments);
    return result;
}


/* ── JSON helpers ──────────────────────────────────────────────────────── */

static int json_hex_value(char value)
{
    if (value >= '0' && value <= '9') return value - '0';
    if (value >= 'a' && value <= 'f') return value - 'a' + 10;
    if (value >= 'A' && value <= 'F') return value - 'A' + 10;
    return -1;
}

static BOOL json_append_utf8(unsigned int codePoint, char *out, size_t outLen,
                             size_t *position)
{
    unsigned char encoded[4];
    size_t encodedLen;

    if (codePoint == 0) {
        return FALSE;
    } else if (codePoint <= 0x7f) {
        encoded[0] = (unsigned char)codePoint;
        encodedLen = 1;
    } else if (codePoint <= 0x7ff) {
        encoded[0] = (unsigned char)(0xc0 | (codePoint >> 6));
        encoded[1] = (unsigned char)(0x80 | (codePoint & 0x3f));
        encodedLen = 2;
    } else if (codePoint <= 0xffff) {
        encoded[0] = (unsigned char)(0xe0 | (codePoint >> 12));
        encoded[1] = (unsigned char)(0x80 | ((codePoint >> 6) & 0x3f));
        encoded[2] = (unsigned char)(0x80 | (codePoint & 0x3f));
        encodedLen = 3;
    } else if (codePoint <= 0x10ffff) {
        encoded[0] = (unsigned char)(0xf0 | (codePoint >> 18));
        encoded[1] = (unsigned char)(0x80 | ((codePoint >> 12) & 0x3f));
        encoded[2] = (unsigned char)(0x80 | ((codePoint >> 6) & 0x3f));
        encoded[3] = (unsigned char)(0x80 | (codePoint & 0x3f));
        encodedLen = 4;
    } else {
        return FALSE;
    }

    if (*position + encodedLen >= outLen) return FALSE;
    memcpy(out + *position, encoded, encodedLen);
    *position += encodedLen;
    return TRUE;
}

static const char *json_find_value(const char *json, const char *key)
{
    const char *p = json;
    size_t keyLen = strlen(key);

    while (*p) {
        const char *nameStart;
        const char *nameEnd;
        BOOL escaped = FALSE;

        if (*p++ != '"') continue;
        nameStart = p;
        while (*p) {
            if (*p == '\\') {
                escaped = TRUE;
                p++;
                if (*p) p++;
                continue;
            }
            if (*p == '"') break;
            p++;
        }
        if (*p != '"') return NULL;
        nameEnd = p++;
        while (*p && isspace((unsigned char)*p)) p++;
        if (*p != ':') continue;
        if (!escaped && (size_t)(nameEnd - nameStart) == keyLen &&
            memcmp(nameStart, key, keyLen) == 0) {
            p++;
            while (*p && isspace((unsigned char)*p)) p++;
            return p;
        }
    }
    return NULL;
}

static BOOL json_get_string(const char *json, const char *key, char *out, size_t outLen)
{
    const char *p;
    size_t position = 0;

    if (!json || !key || !out || outLen == 0) return FALSE;
    out[0] = '\0';
    p = json_find_value(json, key);
    if (!p) return FALSE;
    if (*p != '"') return FALSE;
    p++;

    while (*p) {
        unsigned char value = (unsigned char)*p++;
        if (value == '"') {
            out[position] = '\0';
            return TRUE;
        }
        if (value < 0x20) return FALSE;
        if (value != '\\') {
            if (position + 1 >= outLen) return FALSE;
            out[position++] = (char)value;
            continue;
        }

        value = (unsigned char)*p++;
        if (!value) return FALSE;
        if (value == '"' || value == '\\' || value == '/') {
            if (position + 1 >= outLen) return FALSE;
            out[position++] = (char)value;
        } else if (value == 'b' || value == 'f' || value == 'n' ||
                   value == 'r' || value == 't') {
            static const char escapedValues[] = {'\b', '\f', '\n', '\r', '\t'};
            const char escapedNames[] = {'b', 'f', 'n', 'r', 't'};
            size_t escapeIndex = 0;
            while (escapedNames[escapeIndex] != (char)value) escapeIndex++;
            if (position + 1 >= outLen) return FALSE;
            out[position++] = escapedValues[escapeIndex];
        } else if (value == 'u') {
            unsigned int codePoint = 0;
            if (strlen(p) < 4) return FALSE;
            for (int digit = 0; digit < 4; digit++) {
                int hexValue = json_hex_value(p[digit]);
                if (hexValue < 0) return FALSE;
                codePoint = (codePoint << 4) | (unsigned int)hexValue;
            }
            p += 4;

            if (codePoint >= 0xd800 && codePoint <= 0xdbff &&
                p[0] == '\\' && p[1] == 'u') {
                unsigned int lowSurrogate = 0;
                if (strlen(p) < 6) return FALSE;
                for (int digit = 0; digit < 4; digit++) {
                    int hexValue = json_hex_value(p[2 + digit]);
                    if (hexValue < 0) return FALSE;
                    lowSurrogate = (lowSurrogate << 4) | (unsigned int)hexValue;
                }
                if (lowSurrogate >= 0xdc00 && lowSurrogate <= 0xdfff) {
                    codePoint = 0x10000 + ((codePoint - 0xd800) << 10) +
                                (lowSurrogate - 0xdc00);
                    p += 6;
                } else {
                    codePoint = 0xfffd;
                }
            } else if (codePoint >= 0xd800 && codePoint <= 0xdfff) {
                codePoint = 0xfffd;
            }

            if (!json_append_utf8(codePoint, out, outLen, &position)) return FALSE;
        } else {
            return FALSE;
        }
    }
    return FALSE;
}

static BOOL json_get_int(const char *json, const char *key, int *out)
{
    const char *p;
    if (!json || !key || !out) return FALSE;
    p = json_find_value(json, key);
    if (!p) return FALSE;
    *out = atoi(p);
    return TRUE;
}

static void json_escape_wstring(const wchar_t *in, wchar_t *out, size_t outLen);

static void json_escape_string(const char *in, wchar_t *out, size_t outLen)
{
    int wideLen;
    wchar_t *wide;

    if (!out || outLen == 0) return;
    out[0] = L'\0';
    if (!in) return;

    wideLen = MultiByteToWideChar(CP_UTF8, 0, in, -1, NULL, 0);
    if (wideLen <= 0) return;
    wide = (wchar_t *)malloc((size_t)wideLen * sizeof(wchar_t));
    if (!wide) return;
    if (MultiByteToWideChar(CP_UTF8, 0, in, -1, wide, wideLen) <= 0) {
        free(wide);
        return;
    }
    json_escape_wstring(wide, out, outLen);
    free(wide);
}

static void json_escape_wstring(const wchar_t *in, wchar_t *out, size_t outLen)
{
    size_t j = 0;
    for (size_t i = 0; in[i] && j < outLen - 2; i++) {
        wchar_t c = in[i];
        if (c == L'"' || c == L'\\') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = c;
        } else if (c == L'\n') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'n';
        } else if (c == L'\r') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'r';
        } else if (c == L'\t') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L't';
        } else if (c == L'\b') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'b';
        } else if (c == L'\f') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'f';
        } else if (c < 0x20 || c == 0x2028 || c == 0x2029) {
            static const wchar_t hex[] = L"0123456789abcdef";
            if (j + 6 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'u';
            out[j++] = hex[(c >> 12) & 0xf];
            out[j++] = hex[(c >> 8) & 0xf];
            out[j++] = hex[(c >> 4) & 0xf];
            out[j++] = hex[c & 0xf];
        } else {
            out[j++] = c;
        }
    }
    out[j] = L'\0';
}

static void CfgSendUpdateResultWithVersions(LPCWSTR status, LPCWSTR title,
                                            LPCWSTR message,
                                            LPCWSTR currentVersion,
                                            LPCWSTR remoteVersion,
                                            BOOL automatic) {
    if (!g_webviewView || !status || !title || !message ||
        !currentVersion || !remoteVersion) return;
    wchar_t escapedStatus[64], escapedTitle[256], escapedMessage[1024];
    wchar_t escapedCurrentVersion[64], escapedRemoteVersion[64];
    json_escape_wstring(status, escapedStatus,
                        sizeof(escapedStatus) / sizeof(wchar_t));
    json_escape_wstring(title, escapedTitle,
                        sizeof(escapedTitle) / sizeof(wchar_t));
    json_escape_wstring(message, escapedMessage,
                        sizeof(escapedMessage) / sizeof(wchar_t));
    json_escape_wstring(currentVersion, escapedCurrentVersion,
                        sizeof(escapedCurrentVersion) / sizeof(wchar_t));
    json_escape_wstring(remoteVersion, escapedRemoteVersion,
                        sizeof(escapedRemoteVersion) / sizeof(wchar_t));

    wchar_t script[1792];
    int written = swprintf_s(script, sizeof(script) / sizeof(wchar_t),
        L"window.onUpdateResult({\"status\":\"%s\",\"title\":\"%s\","
        L"\"message\":\"%s\",\"currentVersion\":\"%s\","
        L"\"remoteVersion\":\"%s\",\"automatic\":%s})",
        escapedStatus, escapedTitle, escapedMessage,
        escapedCurrentVersion, escapedRemoteVersion,
        automatic ? L"true" : L"false");
    if (written > 0) webview_execute_script(script);
}

static void CfgSendUpdateResult(LPCWSTR status, LPCWSTR title, LPCWSTR message) {
    CfgSendUpdateResultWithVersions(status, title, message, L"", L"", FALSE);
}

static void CfgSendUpdateProgress(DWORD percent) {
    wchar_t script[160];
    int written = swprintf_s(script, sizeof(script) / sizeof(wchar_t),
        L"window.onUpdateProgress({\"percentComplete\":%lu})",
        (unsigned long)percent);
    if (written > 0) webview_execute_script(script);
}

static void DiscardPendingUpdateNotice(void) {
    UpdateCheckTask* task = g_updateNoticeTask;
    g_updateNoticeTask = NULL;
    DiscardUpdateTask(task);
}

static BOOL IsUpdateWorkerActive(void) {
    // Read the pending flag first: a worker publishes its result before
    // clearing that flag, so an idle answer means no result is in flight.
    if (InterlockedCompareExchange(&g_updateCheckPending, FALSE, FALSE) == TRUE) {
        return TRUE;
    }
    return InterlockedCompareExchangePointer(
        (PVOID volatile*)&g_updatePostedResult, NULL, NULL) != NULL;
}

/* UI thread only. Whether settings should show a check in progress. */
static BOOL IsUpdateCheckVisible(void) {
    return (IsUpdateWorkerActive() && !g_updateCheckAbandoned) ||
           g_updateQueuedCheck;
}

static void StartUpdateCheck(BOOL automatic) {
    if (!g_hWndMain) return;
    if (automatic && (g_updateNoticeTask || g_updateReadyTask)) return;
    if (g_updateCheckAbandoned && IsUpdateWorkerActive()) {
        // Run this request once the stopped worker exits instead of
        // reporting a conflict. A manual request replaces an automatic one.
        if (!automatic || !g_updateQueuedCheck) {
            g_updateQueuedCheck = TRUE;
            g_updateQueuedAutomatic = automatic;
        }
        return;
    }
    if (IsUpdateWorkerActive() ||
        InterlockedCompareExchange(&g_updateCheckPending, TRUE, FALSE) != FALSE) {
        if (!automatic) {
            CfgSendUpdateResult(L"error", L"Update check in progress",
                L"Another update check is still finishing. Try again shortly.");
        }
        return;
    }
    // The check starting now also satisfies any queued request.
    if (g_updateQueuedCheck && !g_updateQueuedAutomatic) automatic = FALSE;
    g_updateQueuedCheck = FALSE;
    g_updateQueuedAutomatic = FALSE;
    g_updateCheckAbandoned = FALSE;
    InterlockedExchange(&g_updateCheckAutomatic, automatic ? TRUE : FALSE);

    // Every accepted request starts from scratch. Automatic requests are
    // skipped above while a result is awaiting user action, avoiding an
    // hourly re-download of the same prepared executable.
    DiscardPendingUpdateNotice();
    DiscardPreparedUpdate();

    if (!g_updateCancelEvent) {
        g_updateCancelEvent = CreateEventW(NULL, TRUE, FALSE, NULL);
        if (!g_updateCancelEvent) {
            DWORD errorCode = GetLastError();
            InterlockedExchange(&g_updateCheckPending, FALSE);
            InterlockedExchange(&g_updateCheckAutomatic, FALSE);
            wchar_t message[256];
            swprintf_s(message, sizeof(message) / sizeof(wchar_t),
                L"Could not initialize update cancellation (Windows error %lu).",
                (unsigned long)errorCode);
            CfgSendUpdateResultWithVersions(L"error", L"Update failed", message,
                                            L"", L"", automatic);
            return;
        }
    }
    ResetEvent(g_updateCancelEvent);
    InterlockedExchange(&g_updateProgressPercent, 0);
    InterlockedExchange(&g_updateProgressPosted, FALSE);

    UpdateCheckTask* task = (UpdateCheckTask*)calloc(1, sizeof(UpdateCheckTask));
    if (!task) {
        InterlockedExchange(&g_updateCheckPending, FALSE);
        InterlockedExchange(&g_updateCheckAutomatic, FALSE);
        CfgSendUpdateResultWithVersions(L"error", L"Update failed",
            L"There was not enough memory to check for updates.",
            L"", L"", automatic);
        return;
    }
    task->targetWindow = g_hWndMain;
    task->automatic = automatic;
    LONG sequence = InterlockedIncrement(&g_updateRequestSequence);
    task->cacheBuster =
        ((GetTickCount64() ^ GetCurrentProcessId()) << 32) | (DWORD)sequence;
    if (task->cacheBuster == 0) task->cacheBuster = 1;

    HANDLE thread = CreateThread(NULL, 0, UpdateCheckThread, task, 0, NULL);
    if (!thread) {
        DWORD errorCode = GetLastError();
        free(task);
        InterlockedExchange(&g_updateCheckPending, FALSE);
        InterlockedExchange(&g_updateCheckAutomatic, FALSE);
        wchar_t message[256];
        swprintf_s(message, sizeof(message) / sizeof(wchar_t),
                   L"Could not start the update check (Windows error %lu).",
                   (unsigned long)errorCode);
        CfgSendUpdateResultWithVersions(L"error", L"Update failed", message,
                                        L"", L"", automatic);
        return;
    }
    CloseHandle(thread);
}

static BOOL IsConfigurationViewOpen(void) {
    return g_webviewHwnd && IsWindow(g_webviewHwnd) &&
           strcmp(g_pendingView, "config") == 0;
}

static BOOL IsIgnoredUpdateVersion(const ExecutableVersion* version) {
    char formatted[32];
    if (!version || !g_ignoredUpdateVersion[0]) return FALSE;
    snprintf(formatted, sizeof(formatted), "%u.%u.%u.%u",
             (unsigned int)version->major,
             (unsigned int)version->minor,
             (unsigned int)version->patch,
             (unsigned int)version->build);
    return strcmp(formatted, g_ignoredUpdateVersion) == 0;
}

static void SaveIgnoredUpdateVersion(const char* version) {
    HKEY key;
    DWORD disposition;
    if (!version) return;
    strncpy(g_ignoredUpdateVersion, version,
            sizeof(g_ignoredUpdateVersion) - 1);
    g_ignoredUpdateVersion[sizeof(g_ignoredUpdateVersion) - 1] = '\0';
    if (RegCreateKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &key,
                        &disposition) == ERROR_SUCCESS) {
        RegSetValueExA(key, REG_VALUE_IGNORED_UPDATE_VERSION, 0, REG_SZ,
                       (const BYTE*)g_ignoredUpdateVersion,
                       (DWORD)(strlen(g_ignoredUpdateVersion) + 1));
        RegCloseKey(key);
    }
}

static void PresentPendingUpdateNotice(void) {
    if (!g_configViewReady || !g_webviewView || !g_updateNoticeTask) return;

    UpdateCheckTask* task = g_updateNoticeTask;
    g_updateNoticeTask = NULL;
    LPCWSTR status = NULL;
    LPCWSTR title = NULL;
    LPCWSTR message = NULL;
    wchar_t currentVersion[32] = L"";
    wchar_t remoteVersion[32] = L"";
    BOOL installable = FALSE;

    if (task->kind == UPDATE_CHECK_CANCELLED) {
        status = L"cancelled";
        title = L"";
        message = L"";
    } else if (task->kind == UPDATE_CHECK_ERROR) {
        status = L"error";
        title = L"Update failed";
        message = task->message;
    } else {
        FormatExecutableVersionForDisplay(
            &task->runningVersion, currentVersion,
            sizeof(currentVersion) / sizeof(wchar_t));
        FormatExecutableVersionForDisplay(
            &task->availableVersion, remoteVersion,
            sizeof(remoteVersion) / sizeof(wchar_t));
        if (task->kind == UPDATE_CHECK_NEWER) {
            status = L"newer";
            title = L"Update available";
            message = L"A newer version is ready to install.";
            installable = TRUE;
        } else if (task->kind == UPDATE_CHECK_SAME) {
            status = L"same";
            title = L"You're up to date";
            message = L"The remote build matches your current version. "
                      L"You can force a reinstall if needed.";
            installable = !task->automatic;
        } else if (task->kind == UPDATE_CHECK_OLDER) {
            status = L"older";
            title = L"No update available";
            message = L"The remote build is older than your current version.";
        }
    }

    if (!status) {
        DiscardUpdateTask(task);
        return;
    }
    if (installable) {
        DiscardPreparedUpdate();
        g_updateReadyTask = task;
    }
    CfgSendUpdateResultWithVersions(status, title, message,
                                    currentVersion, remoteVersion,
                                    task->automatic);
    if (!installable) DiscardUpdateTask(task);
}

static void QueueUpdateNotice(UpdateCheckTask* task) {
    DiscardPendingUpdateNotice();
    g_updateNoticeTask = task;
    PresentPendingUpdateNotice();
}

static void StartQueuedUpdateCheck(void) {
    if (!g_updateQueuedCheck) return;
    BOOL automatic = g_updateQueuedAutomatic;
    g_updateQueuedCheck = FALSE;
    g_updateQueuedAutomatic = FALSE;
    StartUpdateCheck(automatic);
}

static void HandleCompletedUpdateCheck(UpdateCheckTask* task) {
    if (!task) return;
    InterlockedExchange(&g_updateProgressPosted, FALSE);
    InterlockedExchange(&g_updateProgressPercent, 0);

    if (g_updateCheckAbandoned) {
        // Settings already showed this check as stopped; drop its result
        // and any staged download instead of presenting it.
        g_updateCheckAbandoned = FALSE;
        UpdateDebugPrint(L"[INFO] Stopped update check has exited\n");
        DiscardUpdateTask(task);
        StartQueuedUpdateCheck();
        return;
    }

    if (task->kind == UPDATE_CHECK_CANCELLED) {
        UpdateDebugPrint(L"[INFO] Update check cancelled\n");
    } else if (task->kind == UPDATE_CHECK_ERROR) {
        UpdateDebugPrint(L"[WARNING] Update check failed: %s\n", task->message);
    }

    if (task->automatic && !g_configAutoCheckForUpdates) {
        DiscardUpdateTask(task);
        return;
    }

    if (task->automatic && task->kind == UPDATE_CHECK_NEWER &&
        IsIgnoredUpdateVersion(&task->availableVersion)) {
        UpdateDebugPrint(L"[INFO] Automatic update prompt suppressed for ignored version\n");
        task->kind = UPDATE_CHECK_CANCELLED;
        if (IsConfigurationViewOpen()) {
            QueueUpdateNotice(task);
        } else {
            DiscardUpdateTask(task);
        }
        return;
    }

    if (task->automatic && task->kind == UPDATE_CHECK_NEWER) {
        if (g_webviewHwnd && !IsConfigurationViewOpen()) {
            SendMessageW(g_webviewHwnd, WM_CLOSE, 0, 0);
        }
        QueueUpdateNotice(task);
        ShowWebViewDialog("config", 560, 520);
        if (!IsConfigurationViewOpen()) DiscardPendingUpdateNotice();
        return;
    }

    if (IsConfigurationViewOpen()) {
        QueueUpdateNotice(task);
    } else {
        DiscardUpdateTask(task);
    }
}

static void IgnorePreparedUpdateVersion(const char* requestedVersion) {
    UpdateCheckTask* task = g_updateReadyTask;
    char preparedVersion[32];
    if (!task || !task->automatic || task->kind != UPDATE_CHECK_NEWER ||
        !requestedVersion) {
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The update version could not be ignored. Check for updates again.");
        return;
    }
    snprintf(preparedVersion, sizeof(preparedVersion), "%u.%u.%u.%u",
             (unsigned int)task->availableVersion.major,
             (unsigned int)task->availableVersion.minor,
             (unsigned int)task->availableVersion.patch,
             (unsigned int)task->availableVersion.build);
    if (strcmp(preparedVersion, requestedVersion) != 0) {
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The update version changed. Check for updates again.");
        return;
    }
    SaveIgnoredUpdateVersion(preparedVersion);
    UpdateDebugPrint(L"[INFO] Automatic update version added to the ignore list\n");
    DiscardPreparedUpdate();
}

static void CancelUpdateCheck(void) {
    BOOL stopped = g_updateQueuedCheck;
    g_updateQueuedCheck = FALSE;
    g_updateQueuedAutomatic = FALSE;
    if (!g_updateCheckAbandoned && IsUpdateWorkerActive()) {
        UpdateDebugPrint(L"[INFO] Update check cancellation requested\n");
        if (g_updateCancelEvent) SetEvent(g_updateCancelEvent);
        g_updateCheckAbandoned = TRUE;
        stopped = TRUE;
    }
    // Report the stop at once. A blocking network call may not return until
    // its timeout; the worker then deletes its partial download and its
    // late result is discarded rather than shown.
    if (stopped) CfgSendUpdateResult(L"cancelled", L"", L"");
}

static HANDLE CreateUpdateReadyEvent(DWORD processId, wchar_t* eventName,
                                     size_t eventNameCch) {
    if (!processId || !eventName || eventNameCch < 96) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return NULL;
    }

    ULONGLONG nonce = GetTickCount64() ^
                      ((ULONGLONG)GetCurrentThreadId() << 32);
    BCryptGenRandom(NULL, (PUCHAR)&nonce, sizeof(nonce),
                    BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    int nameLength = swprintf_s(eventName, eventNameCch,
        L"Local\\ImagePaster_UpdateReady_%lu_%016llx",
        (unsigned long)processId, (unsigned long long)nonce);
    if (nameLength <= 0 || nameLength >= (int)eventNameCch) {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return NULL;
    }

    // The elevated helper can run with a split administrator token (or with
    // alternate administrator credentials). Grant interactive users access
    // to this random, session-local event so either UAC path can acknowledge
    // readiness without exposing any file or process permissions.
    PSECURITY_DESCRIPTOR descriptor = NULL;
    if (!ConvertStringSecurityDescriptorToSecurityDescriptorW(
            L"D:P(A;;GA;;;IU)(A;;GA;;;BA)(A;;GA;;;SY)",
            SDDL_REVISION_1, &descriptor, NULL)) {
        return NULL;
    }
    SECURITY_ATTRIBUTES securityAttributes = {0};
    securityAttributes.nLength = sizeof(securityAttributes);
    securityAttributes.lpSecurityDescriptor = descriptor;

    HANDLE readyEvent = CreateEventW(&securityAttributes, TRUE, FALSE,
                                     eventName);
    DWORD errorCode = readyEvent ? ERROR_SUCCESS : GetLastError();
    LocalFree(descriptor);
    if (!readyEvent) SetLastError(errorCode);
    return readyEvent;
}

static BOOL LaunchStagedUpdate(LPCWSTR stagedPath, LPCWSTR targetPath,
                               BOOL reopenSettings) {
    if (!stagedPath || !targetPath || !*stagedPath || !*targetPath) {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    DWORD oldProcessId = GetCurrentProcessId();
    wchar_t helperPath[MAX_PATH];
    if (!BuildUpdateTempPath(helperPath, L"updater", oldProcessId)) {
        return FALSE;
    }
    DeleteUpdateTempFile(helperPath);
    if (!CopyFileW(targetPath, helperPath, TRUE)) return FALSE;
    SetFileAttributesW(helperPath, FILE_ATTRIBUTE_NORMAL);

    // CopyFile preserves alternate data streams. The source is already the
    // running, user-approved executable, so do not carry its download-zone
    // marker onto the short-lived updater copy and trigger a second warning.
    wchar_t zonePath[MAX_PATH + 32];
    if (swprintf_s(zonePath, sizeof(zonePath) / sizeof(wchar_t),
                   L"%s:Zone.Identifier", helperPath) > 0) {
        DeleteFileW(zonePath);
    }

    wchar_t readyEventName[160];
    HANDLE readyEvent = CreateUpdateReadyEvent(oldProcessId, readyEventName,
        sizeof(readyEventName) / sizeof(wchar_t));
    if (!readyEvent) {
        DWORD errorCode = GetLastError();
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }

    wchar_t parameters[MAX_PATH * 2 + 512];
    int parameterLength = swprintf_s(parameters,
        sizeof(parameters) / sizeof(wchar_t),
        L"--apply-update %lu \"%s\" \"%s\" \"%s\"%s",
        (unsigned long)oldProcessId, readyEventName, targetPath, stagedPath,
        reopenSettings ? L" --reopen-settings" : L"");
    if (parameterLength <= 0 ||
        parameterLength >= (int)(sizeof(parameters) / sizeof(wchar_t))) {
        CloseHandle(readyEvent);
        DeleteUpdateTempFile(helperPath);
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    SHELLEXECUTEINFOW executeInfo = {0};
    executeInfo.cbSize = sizeof(executeInfo);
    executeInfo.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC;
    executeInfo.hwnd = g_webviewHwnd;
    executeInfo.lpVerb = L"runas";
    executeInfo.lpFile = helperPath;
    executeInfo.lpParameters = parameters;
    executeInfo.nShow = SW_HIDE;
    BOOL elevated = ShellExecuteExW(&executeInfo);
    if (!elevated || !executeInfo.hProcess) {
        DWORD errorCode = elevated ? ERROR_INVALID_HANDLE : GetLastError();
        if (!errorCode) errorCode = ERROR_ACCESS_DENIED;
        CloseHandle(readyEvent);
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }

    HANDLE waitHandles[2] = { readyEvent, executeInfo.hProcess };
    DWORD waitResult = WaitForMultipleObjects(2, waitHandles, FALSE,
                                              UPDATE_HELPER_READY_MS);
    DWORD errorCode = ERROR_SUCCESS;
    if (waitResult != WAIT_OBJECT_0) {
        if (waitResult == WAIT_OBJECT_0 + 1) {
            DWORD exitCode = ERROR_INSTALL_FAILURE;
            if (!GetExitCodeProcess(executeInfo.hProcess, &exitCode) ||
                exitCode == ERROR_SUCCESS || exitCode == STILL_ACTIVE) {
                exitCode = ERROR_INSTALL_FAILURE;
            }
            errorCode = exitCode;
        } else {
            errorCode = waitResult == WAIT_TIMEOUT ? ERROR_TIMEOUT
                                                   : GetLastError();
            if (!errorCode) errorCode = ERROR_INSTALL_FAILURE;
        }
    }
    CloseHandle(executeInfo.hProcess);
    CloseHandle(readyEvent);

    if (waitResult != WAIT_OBJECT_0) {
        DeleteUpdateTempFile(helperPath);
        SetLastError(errorCode);
        return FALSE;
    }
    return TRUE;
}

static void DiscardPreparedUpdate(void) {
    UpdateCheckTask* task = g_updateReadyTask;
    g_updateReadyTask = NULL;
    DiscardUpdateTask(task);
}

static void InstallPreparedUpdate(BOOL reopenSettings) {
    UpdateCheckTask* task = g_updateReadyTask;
    g_updateReadyTask = NULL;
    if (!task || (task->kind != UPDATE_CHECK_NEWER &&
                  task->kind != UPDATE_CHECK_SAME)) {
        DiscardUpdateTask(task);
        CfgSendUpdateResult(L"error", L"Update unavailable",
            L"The prepared update is no longer available. Check for updates again.");
        return;
    }

    if (LaunchStagedUpdate(task->stagedPath, task->targetPath, reopenSettings)) {
        UpdateDebugPrint(L"[INFO] Update accepted; exiting for replacement\n");
        g_updateInstallReady = TRUE;
        free(task);  // The updater process now owns the staged file.
        if (g_webviewHwnd) PostMessageW(g_webviewHwnd, WM_CLOSE, 0, 0);
        return;
    }

    DWORD errorCode = GetLastError();
    wchar_t message[384];
    LPCWSTR title = L"Update failed";
    if (errorCode == ERROR_CANCELLED) {
        title = L"Update cancelled";
        wcscpy_s(message, sizeof(message) / sizeof(wchar_t),
            L"Administrator approval was cancelled. Your current version is still running.");
    } else {
        swprintf_s(message, sizeof(message) / sizeof(wchar_t),
            L"The elevated update process could not be started (Windows error %lu).",
            (unsigned long)errorCode);
    }
    UpdateDebugPrint(L"[WARNING] %s\n", message);
    CfgSendUpdateResult(L"error", title, message);
    DiscardUpdateTask(task);
}


/* ── Push functions (C -> JS) ──────────────────────────────────────────── */

static void webview_push_init_config(void)
{
    wchar_t wTitleMatch[4096];
    wchar_t wHttpMessageTemplate[MAX_HTTP_MESSAGE_TEMPLATE_BYTES * 2];
    wchar_t wBindIp[128];
    wchar_t wHttpAllowList[MAX_HTTP_ALLOW_LIST_BYTES * 2];
    wchar_t wHttpStatus[512];
    wchar_t wUpdateCompletedVersion[64];
    WCHAR storageDir[MAX_PATH];
    wchar_t wStorageDir[MAX_PATH * 2];
    wchar_t ipsJson[4096];
    char ips[MAX_DETECTED_IPS][INET_ADDRSTRLEN];
    int ipCount = EnumerateDetectedIpv4Addresses(ips, MAX_DETECTED_IPS);
    const wchar_t *captureGapFill =
        g_configCaptureGapFill == CAPTURE_GAP_FILL_BLACK ? L"black" :
        g_configCaptureGapFill == CAPTURE_GAP_FILL_BLUR ? L"blur" : L"white";
    size_t pos = 0;
    json_escape_string(g_configTitleMatch, wTitleMatch, 4096);
    json_escape_string(g_configHttpMessageTemplate, wHttpMessageTemplate,
                       MAX_HTTP_MESSAGE_TEMPLATE_BYTES * 2);
    json_escape_string(g_configBindIp, wBindIp, 128);
    json_escape_string(g_configHttpAllowList, wHttpAllowList,
                       MAX_HTTP_ALLOW_LIST_BYTES * 2);
    if (!BuildImageStorageDirectoryPath(storageDir)) storageDir[0] = L'\0';
    json_escape_wstring(storageDir, wStorageDir, MAX_PATH * 2);
    json_escape_string(g_httpStatus, wHttpStatus, 512);
    json_escape_wstring(g_updateConfirmationPending ? APP_VERSION_W : L"",
                        wUpdateCompletedVersion, 64);

    pos += swprintf(ipsJson + pos, 4096 - pos, L"[");
    for (int i = 0; i < ipCount && pos < 4000; i++) {
        wchar_t wIp[64];
        MultiByteToWideChar(CP_UTF8, 0, ips[i], -1, wIp, 64);
        pos += swprintf(ipsJson + pos, 4096 - pos,
                        i == 0 ? L"\"%s\"" : L",\"%s\"", wIp);
    }
    swprintf(ipsJson + pos, 4096 - pos, L"]");

    wchar_t script[32768];
    swprintf(script, 32768,
        L"window.onInit({\"view\":\"config\",\"config\":{"
        L"\"titleMatch\":\"%s\","
        L"\"pasteMethod\":\"%s\","
        L"\"httpMessageTemplate\":\"%s\","
        L"\"bindIp\":\"%s\","
        L"\"httpPort\":%d,"
        L"\"httpAllowList\":\"%s\","
        L"\"jpegQuality\":%d,"
        L"\"imageHistoryLimit\":%d,"
        L"\"imageStorage\":\"%s\","
        L"\"imageStorageDir\":\"%s\","
        L"\"compatibilityPaste\":%s,"
        L"\"screenCaptureEnabled\":%s,"
        L"\"captureGapFill\":\"%s\","
        L"\"autoCheckForUpdates\":%s,"
        L"\"updateCheckPending\":%s,"
        L"\"updatePromptPending\":%s,"
        L"\"availableIps\":%s,"
        L"\"bindIpAvailable\":%s,"
        L"\"serverStatus\":\"%s\","
        L"\"version\":\"%s\"},"
        L"\"updateCompletedVersion\":\"%s\"})",
        wTitleMatch,
        g_configPasteMethod == PASTE_METHOD_HTTP ? L"http" : L"base64",
        wHttpMessageTemplate, wBindIp, g_configHttpPort, wHttpAllowList,
        g_configJpegQuality, g_configImageHistoryLimit,
        g_configImageStorage == IMAGE_STORAGE_DISK ? L"disk" : L"memory",
        wStorageDir,
        g_configCompatibilityPaste ? L"true" : L"false",
        g_configScreenCaptureEnabled ? L"true" : L"false",
        captureGapFill,
        g_configAutoCheckForUpdates ? L"true" : L"false",
        IsUpdateCheckVisible() ? L"true" : L"false",
        g_updateNoticeTask ? L"true" : L"false",
        ipsJson,
        IsConfiguredBindAddressPresent() ? L"true" : L"false",
        wHttpStatus, APP_VERSION_W, wUpdateCompletedVersion);
    webview_execute_script(script);
}

static void webview_push_init_log(void)
{
    /* Build a JSON array of all log entries */
    size_t bufLen = (size_t)g_logCount * 600 + 256;
    if (bufLen < 1024) bufLen = 1024;
    wchar_t *logJson = (wchar_t*)malloc(bufLen * sizeof(wchar_t));
    if (!logJson) return;

    size_t pos = 0;
    pos += swprintf(logJson + pos, bufLen - pos, L"[");
    for (int i = 0; i < g_logCount && pos < bufLen - 600; i++) {
        /* Display oldest first: index 0 = oldest */
        int bufIdx;
        if (g_logCount < LOG_RING_CAPACITY) {
            bufIdx = i;
        } else {
            bufIdx = (g_logHead + i) % LOG_RING_CAPACITY;
        }
        LogEntry *entry = &g_logRing[bufIdx];

        if (i > 0) pos += swprintf(logJson + pos, bufLen - pos, L",");

        wchar_t wTime[32], wMsg[1024];
        MultiByteToWideChar(CP_UTF8, 0, entry->time, -1, wTime, 32);
        json_escape_string(entry->message, wMsg, 1024);

        pos += swprintf(logJson + pos, bufLen - pos,
            L"{\"time\":\"%s\",\"message\":\"%s\"}",
            wTime, wMsg);
    }
    if (pos < bufLen - 1) pos += swprintf(logJson + pos, bufLen - pos, L"]");

    size_t scriptLen = bufLen + 256;
    wchar_t *script = (wchar_t*)malloc(scriptLen * sizeof(wchar_t));
    if (!script) { free(logJson); return; }
    swprintf(script, scriptLen,
        L"window.onInit({\"view\":\"log\",\"log\":%s})", logJson);
    webview_execute_script(script);
    free(script);
    free(logJson);
}

/* ── History view ──────────────────────────────────────────────────────── */

static const GUID g_wicImagingFactoryClsid =
    {0xcacaf262, 0x9370, 0x4615, {0xa1, 0x3b, 0x9f, 0x55, 0x39, 0xda, 0x4c, 0x0a}};
static const GUID g_wicImagingFactoryIid =
    {0xec5ec8a9, 0xc395, 0x4314, {0x9c, 0x77, 0x54, 0xd7, 0xa9, 0x35, 0xff, 0x70}};
static const GUID g_wicContainerFormatJpeg =
    {0x19e4a5aa, 0x5662, 0x4fc5, {0xa0, 0xc0, 0x17, 0x58, 0x02, 0x8e, 0x10, 0x57}};
static const GUID g_wicPixelFormat24bppBGR =
    {0x6fddc324, 0x4e03, 0x4bfe, {0xb1, 0x85, 0x3d, 0x77, 0x76, 0x8d, 0xc9, 0x0c}};
/* Catmull-Rom scaling on Windows 10 and later; older systems use Fant. */
#define WIC_INTERPOLATION_HIGH_QUALITY_CUBIC ((WICBitmapInterpolationMode)4)

typedef struct {
    char token[IMAGE_TOKEN_HEX_LEN + 1];
    char *dataUri;      /* malloc'd JPEG data URI */
    ULONGLONG lastUsed; /* the least recently used entry is replaced first */
} HistoryThumbnail;

/* Shared by the UI thread and the thumbnail worker under g_thumbLock. The
   queue only holds rows the History page can currently see, top first. */
static SRWLOCK g_thumbLock = SRWLOCK_INIT;
static HistoryThumbnail g_thumbCache[HISTORY_THUMB_CACHE_ENTRIES];
static size_t g_thumbCacheCount = 0;
static ULONGLONG g_thumbClock = 0;
static char g_thumbQueue[HISTORY_THUMB_QUEUE_MAX][IMAGE_TOKEN_HEX_LEN + 1];
static size_t g_thumbQueueCount = 0;
static char g_thumbReady[HISTORY_THUMB_QUEUE_MAX * 2][IMAGE_TOKEN_HEX_LEN + 1];
static size_t g_thumbReadyCount = 0;
static BOOL g_thumbStop = FALSE;
static volatile LONG g_thumbReadyPosted = FALSE;
static HANDLE g_thumbWorkEvent = NULL;
static HANDLE g_thumbThread = NULL;  /* UI thread only */
static size_t g_historyViewPage = 0; /* UI thread only; zero-based */

static void CalculateHistoryThumbnailSize(UINT width, UINT height,
                                          UINT *thumbWidth, UINT *thumbHeight)
{
    if (width >= height) {
        *thumbWidth = width > HISTORY_THUMB_MAX_DIM ? HISTORY_THUMB_MAX_DIM : width;
        *thumbHeight = (UINT)(((ULONGLONG)height * *thumbWidth + width / 2) / width);
    } else {
        *thumbHeight = height > HISTORY_THUMB_MAX_DIM ? HISTORY_THUMB_MAX_DIM : height;
        *thumbWidth = (UINT)(((ULONGLONG)width * *thumbHeight + height / 2) / height);
    }
    if (*thumbWidth < 1) *thumbWidth = 1;
    if (*thumbHeight < 1) *thumbHeight = 1;
}

static BOOL CopyStreamBytes(IStream *stream, BYTE **data, DWORD *size)
{
    STATSTG stat;
    LARGE_INTEGER zero;
    ULONG bytesRead = 0;

    *data = NULL;
    *size = 0;
    ZeroMemory(&stat, sizeof(stat));
    if (IStream_Stat(stream, &stat, STATFLAG_NONAME) != S_OK ||
        stat.cbSize.QuadPart == 0 || stat.cbSize.QuadPart > 0xffffffffULL) {
        return FALSE;
    }
    *data = (BYTE *)malloc((size_t)stat.cbSize.QuadPart);
    if (!*data) return FALSE;
    zero.QuadPart = 0;
    if (IStream_Seek(stream, zero, STREAM_SEEK_SET, NULL) != S_OK ||
        IStream_Read(stream, *data, (ULONG)stat.cbSize.QuadPart,
                     &bytesRead) != S_OK ||
        bytesRead != (ULONG)stat.cbSize.QuadPart) {
        free(*data);
        *data = NULL;
        return FALSE;
    }
    *size = (DWORD)bytesRead;
    return TRUE;
}

/* Scaling the decoder's frame directly lets the WIC JPEG decoder drop
   resolution in the DCT domain before the cubic filter runs, so a large
   screenshot is never decoded at full size. */
static BOOL CreateHistoryThumbnailJpegWic(IWICImagingFactory *factory,
                                          const BYTE *jpegData, DWORD jpegSize,
                                          BYTE **thumbJpeg, DWORD *thumbSize)
{
    IWICStream *input = NULL;
    IWICBitmapDecoder *decoder = NULL;
    IWICBitmapFrameDecode *frame = NULL;
    IWICBitmapScaler *scaler = NULL;
    IWICFormatConverter *converter = NULL;
    IStream *output = NULL;
    IWICBitmapEncoder *encoder = NULL;
    IWICBitmapFrameEncode *target = NULL;
    IPropertyBag2 *options = NULL;
    WICPixelFormatGUID format = g_wicPixelFormat24bppBGR;
    UINT width = 0;
    UINT height = 0;
    UINT thumbWidth = 0;
    UINT thumbHeight = 0;
    BOOL ok = FALSE;
    HRESULT hr;

    *thumbJpeg = NULL;
    *thumbSize = 0;
    hr = IWICImagingFactory_CreateStream(factory, &input);
    if (SUCCEEDED(hr)) {
        hr = IWICStream_InitializeFromMemory(input, (BYTE *)jpegData, jpegSize);
    }
    if (SUCCEEDED(hr)) {
        hr = IWICImagingFactory_CreateDecoderFromStream(factory, (IStream *)input,
            NULL, WICDecodeMetadataCacheOnDemand, &decoder);
    }
    if (SUCCEEDED(hr)) hr = IWICBitmapDecoder_GetFrame(decoder, 0, &frame);
    if (SUCCEEDED(hr)) hr = IWICBitmapFrameDecode_GetSize(frame, &width, &height);
    if (SUCCEEDED(hr) && (width == 0 || height == 0)) hr = E_FAIL;
    if (SUCCEEDED(hr)) {
        CalculateHistoryThumbnailSize(width, height, &thumbWidth, &thumbHeight);
        hr = IWICImagingFactory_CreateBitmapScaler(factory, &scaler);
    }
    if (SUCCEEDED(hr)) {
        hr = IWICBitmapScaler_Initialize(scaler, (IWICBitmapSource *)frame,
            thumbWidth, thumbHeight, WIC_INTERPOLATION_HIGH_QUALITY_CUBIC);
        if (FAILED(hr)) {
            IWICBitmapScaler_Release(scaler);
            scaler = NULL;
            hr = IWICImagingFactory_CreateBitmapScaler(factory, &scaler);
            if (SUCCEEDED(hr)) {
                hr = IWICBitmapScaler_Initialize(scaler, (IWICBitmapSource *)frame,
                    thumbWidth, thumbHeight, WICBitmapInterpolationModeFant);
            }
        }
    }
    if (SUCCEEDED(hr)) hr = IWICImagingFactory_CreateFormatConverter(factory, &converter);
    if (SUCCEEDED(hr)) {
        hr = IWICFormatConverter_Initialize(converter, (IWICBitmapSource *)scaler,
            &g_wicPixelFormat24bppBGR, WICBitmapDitherTypeNone, NULL, 0.0,
            WICBitmapPaletteTypeCustom);
    }
    if (SUCCEEDED(hr)) hr = CreateStreamOnHGlobal(NULL, TRUE, &output);
    if (SUCCEEDED(hr)) {
        hr = IWICImagingFactory_CreateEncoder(factory, &g_wicContainerFormatJpeg,
                                              NULL, &encoder);
    }
    if (SUCCEEDED(hr)) {
        hr = IWICBitmapEncoder_Initialize(encoder, output, WICBitmapEncoderNoCache);
    }
    if (SUCCEEDED(hr)) hr = IWICBitmapEncoder_CreateNewFrame(encoder, &target, &options);
    if (SUCCEEDED(hr)) {
        PROPBAG2 option;
        VARIANT quality;
        ZeroMemory(&option, sizeof(option));
        ZeroMemory(&quality, sizeof(quality));
        option.pstrName = (LPOLESTR)L"ImageQuality";
        quality.vt = VT_R4;
        quality.fltVal = HISTORY_THUMB_JPEG_QUALITY / 100.0f;
        IPropertyBag2_Write(options, 1, &option, &quality); /* else default */
        hr = IWICBitmapFrameEncode_Initialize(target, options);
    }
    if (SUCCEEDED(hr)) {
        hr = IWICBitmapFrameEncode_SetSize(target, thumbWidth, thumbHeight);
    }
    if (SUCCEEDED(hr)) hr = IWICBitmapFrameEncode_SetPixelFormat(target, &format);
    if (SUCCEEDED(hr) && !IsEqualGUID(&format, &g_wicPixelFormat24bppBGR)) {
        hr = E_FAIL;
    }
    if (SUCCEEDED(hr)) {
        hr = IWICBitmapFrameEncode_WriteSource(target,
            (IWICBitmapSource *)converter, NULL);
    }
    if (SUCCEEDED(hr)) hr = IWICBitmapFrameEncode_Commit(target);
    if (SUCCEEDED(hr)) hr = IWICBitmapEncoder_Commit(encoder);
    if (SUCCEEDED(hr)) ok = CopyStreamBytes(output, thumbJpeg, thumbSize);

    if (options) IPropertyBag2_Release(options);
    if (target) IWICBitmapFrameEncode_Release(target);
    if (encoder) IWICBitmapEncoder_Release(encoder);
    if (output) IStream_Release(output);
    if (converter) IWICFormatConverter_Release(converter);
    if (scaler) IWICBitmapScaler_Release(scaler);
    if (frame) IWICBitmapFrameDecode_Release(frame);
    if (decoder) IWICBitmapDecoder_Release(decoder);
    if (input) IWICStream_Release(input);
    return ok;
}

/* Fallback if WIC is unavailable: full GDI+ decode and bicubic resample. */
static BOOL CreateHistoryThumbnailJpegGdiplus(const BYTE *jpegData,
                                              DWORD jpegSize,
                                              BYTE **thumbJpeg, DWORD *thumbSize)
{
    IStream *stream = NULL;
    GpBitmap *source = NULL;
    GpBitmap *thumbnail = NULL;
    GpGraphics *graphics = NULL;
    UINT width = 0;
    UINT height = 0;
    UINT thumbWidth;
    UINT thumbHeight;
    ULONG written = 0;
    LARGE_INTEGER zero;
    BOOL ok = FALSE;

    *thumbJpeg = NULL;
    *thumbSize = 0;
    if (CreateStreamOnHGlobal(NULL, TRUE, &stream) != S_OK) return FALSE;
    zero.QuadPart = 0;
    if (IStream_Write(stream, jpegData, jpegSize, &written) != S_OK ||
        written != jpegSize ||
        IStream_Seek(stream, zero, STREAM_SEEK_SET, NULL) != S_OK) {
        goto cleanup;
    }
    if (GdipCreateBitmapFromStream(stream, &source) != 0 || !source) {
        goto cleanup;
    }
    GdipGetImageWidth((GpImage *)source, &width);
    GdipGetImageHeight((GpImage *)source, &height);
    if (width == 0 || height == 0) goto cleanup;
    CalculateHistoryThumbnailSize(width, height, &thumbWidth, &thumbHeight);

    if (GdipCreateBitmapFromScan0((INT)thumbWidth, (INT)thumbHeight, 0,
                                  HISTORY_THUMB_PIXEL_FORMAT, NULL,
                                  &thumbnail) != 0 || !thumbnail) {
        goto cleanup;
    }
    if (GdipGetImageGraphicsContext((GpImage *)thumbnail, &graphics) != 0 ||
        !graphics) {
        goto cleanup;
    }
    GdipSetInterpolationMode(graphics, GDIP_INTERPOLATION_HIGH_BICUBIC);
    GdipSetPixelOffsetMode(graphics, CAPTURE_GDIP_PIXEL_OFFSET_HIGH_QUALITY);
    if (GdipDrawImageRectI(graphics, (GpImage *)source, 0, 0,
                           (INT)thumbWidth, (INT)thumbHeight) != 0) {
        goto cleanup;
    }
    GdipDeleteGraphics(graphics);
    graphics = NULL;
    ok = EncodeImageToMemory((GpImage *)thumbnail, L"image/jpeg",
                             HISTORY_THUMB_JPEG_QUALITY, thumbJpeg, thumbSize);

cleanup:
    if (graphics) GdipDeleteGraphics(graphics);
    if (thumbnail) GdipDisposeImage((GpImage *)thumbnail);
    if (source) GdipDisposeImage((GpImage *)source);
    if (stream) IStream_Release(stream);
    return ok;
}

static char *CreateJpegDataUri(const BYTE *jpegData, DWORD jpegSize)
{
    static const char prefix[] = "data:image/jpeg;base64,";
    DWORD base64Len = 0;
    char *base64 = Base64Encode(jpegData, jpegSize, &base64Len);
    char *dataUri;

    if (!base64) return NULL;
    dataUri = (char *)malloc(sizeof(prefix) + base64Len);
    if (dataUri) {
        memcpy(dataUri, prefix, sizeof(prefix) - 1);
        memcpy(dataUri + sizeof(prefix) - 1, base64, base64Len + 1);
    }
    free(base64);
    return dataUri;
}

/* Worker thread. Returns a malloc'd data URI, or NULL if the image is gone
   or unreadable. Disk files are read and decoded outside g_imageLock, so a
   slow disk never holds up clipboard updates. */
static char *CreateHistoryThumbnail(IWICImagingFactory *factory,
                                    const char *token)
{
    WCHAR *path = GetRetainedImageDiskPathCopy(token);
    BYTE *jpegData = NULL;
    DWORD jpegSize = 0;
    BYTE *thumbJpeg = NULL;
    DWORD thumbSize = 0;
    char *dataUri = NULL;

    if (path) {
        jpegData = ReadRetainedImageFile(path, &jpegSize);
        free(path);
    } else {
        CopyRetainedImageByToken(token, &jpegData, &jpegSize, NULL);
    }
    if (!jpegData) return NULL;
    if ((factory && CreateHistoryThumbnailJpegWic(factory, jpegData, jpegSize,
                                                  &thumbJpeg, &thumbSize)) ||
        CreateHistoryThumbnailJpegGdiplus(jpegData, jpegSize,
                                          &thumbJpeg, &thumbSize)) {
        dataUri = CreateJpegDataUri(thumbJpeg, thumbSize);
        free(thumbJpeg);
    }
    free(jpegData);
    return dataUri;
}

/* Caller holds g_thumbLock exclusively. */
static HistoryThumbnail *FindHistoryThumbnailLocked(const char *token)
{
    for (size_t i = 0; i < g_thumbCacheCount; i++) {
        if (strcmp(g_thumbCache[i].token, token) == 0) {
            g_thumbCache[i].lastUsed = ++g_thumbClock;
            return &g_thumbCache[i];
        }
    }
    return NULL;
}

/* Caller holds g_thumbLock exclusively; takes ownership of dataUri. */
static void StoreHistoryThumbnailLocked(const char *token, char *dataUri)
{
    HistoryThumbnail *slot = FindHistoryThumbnailLocked(token);

    if (!slot && g_thumbCacheCount < HISTORY_THUMB_CACHE_ENTRIES) {
        slot = &g_thumbCache[g_thumbCacheCount++];
    } else if (!slot) {
        slot = &g_thumbCache[0];
        for (size_t i = 1; i < g_thumbCacheCount; i++) {
            if (g_thumbCache[i].lastUsed < slot->lastUsed) slot = &g_thumbCache[i];
        }
    }
    free(slot->dataUri);
    memcpy(slot->token, token, sizeof(slot->token));
    slot->dataUri = dataUri;
    slot->lastUsed = ++g_thumbClock;
}

/* Caller holds g_thumbLock exclusively. */
static void MarkHistoryThumbnailReadyLocked(const char *token)
{
    size_t capacity = sizeof(g_thumbReady) / sizeof(g_thumbReady[0]);
    for (size_t i = 0; i < g_thumbReadyCount; i++) {
        if (strcmp(g_thumbReady[i], token) == 0) return;
    }
    if (g_thumbReadyCount < capacity) {
        memcpy(g_thumbReady[g_thumbReadyCount++], token, sizeof(g_thumbReady[0]));
    }
}

static DWORD WINAPI HistoryThumbnailThread(LPVOID parameter)
{
    HRESULT comResult = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    IWICImagingFactory *factory = NULL;

    (void)parameter;
    if (SUCCEEDED(comResult) &&
        FAILED(CoCreateInstance(&g_wicImagingFactoryClsid, NULL,
                                CLSCTX_INPROC_SERVER, &g_wicImagingFactoryIid,
                                (void **)&factory))) {
        factory = NULL;
    }
    for (;;) {
        char token[IMAGE_TOKEN_HEX_LEN + 1];
        BOOL haveWork = FALSE;
        BOOL cached = FALSE;
        BOOL stop;

        AcquireSRWLockExclusive(&g_thumbLock);
        stop = g_thumbStop;
        if (!stop && g_thumbQueueCount > 0) {
            memcpy(token, g_thumbQueue[0], sizeof(token));
            g_thumbQueueCount--;
            memmove(g_thumbQueue[0], g_thumbQueue[1],
                    g_thumbQueueCount * sizeof(g_thumbQueue[0]));
            haveWork = TRUE;
            if (FindHistoryThumbnailLocked(token)) {
                MarkHistoryThumbnailReadyLocked(token);
                cached = TRUE;
            }
        }
        ReleaseSRWLockExclusive(&g_thumbLock);
        if (stop) break;
        if (!haveWork) {
            WaitForSingleObject(g_thumbWorkEvent, INFINITE);
            continue;
        }

        if (!cached) {
            // A failure is reported but not cached, so reopening retries.
            char *dataUri = CreateHistoryThumbnail(factory, token);
            AcquireSRWLockExclusive(&g_thumbLock);
            if (dataUri) StoreHistoryThumbnailLocked(token, dataUri);
            MarkHistoryThumbnailReadyLocked(token);
            ReleaseSRWLockExclusive(&g_thumbLock);
        }
        if (InterlockedCompareExchange(&g_thumbReadyPosted, TRUE, FALSE) == FALSE &&
            !PostMessageW(g_hWndMain, WM_HISTORY_THUMBS_READY, 0, 0)) {
            InterlockedExchange(&g_thumbReadyPosted, FALSE);
        }
    }
    if (factory) IWICImagingFactory_Release(factory);
    if (SUCCEEDED(comResult)) CoUninitialize();
    return 0;
}

static BOOL EnsureHistoryThumbnailWorker(void)
{
    if (g_thumbThread) return TRUE;
    if (!g_thumbWorkEvent) {
        g_thumbWorkEvent = CreateEventW(NULL, FALSE, FALSE, NULL);
        if (!g_thumbWorkEvent) return FALSE;
    }
    g_thumbThread = CreateThread(NULL, 0, HistoryThumbnailThread, NULL, 0, NULL);
    if (!g_thumbThread) return FALSE;
    // Previews should never compete with input handling or the dialog.
    SetThreadPriority(g_thumbThread, THREAD_PRIORITY_BELOW_NORMAL);
    return TRUE;
}

static void StopHistoryThumbnailWorker(void)
{
    if (!g_thumbThread) return;
    AcquireSRWLockExclusive(&g_thumbLock);
    g_thumbStop = TRUE;
    g_thumbQueueCount = 0;
    ReleaseSRWLockExclusive(&g_thumbLock);
    SetEvent(g_thumbWorkEvent);
    // One preview takes far less than this; exit continues regardless.
    WaitForSingleObject(g_thumbThread, 2000);
    CloseHandle(g_thumbThread);
    g_thumbThread = NULL;
}

/* UI thread. Sends finished thumbnails to the History page in small batches.
   A token without a cached thumbnail goes out empty: no preview available. */
static void PushReadyHistoryThumbnails(void)
{
    BOOL historyOpen = g_webviewView && strcmp(g_pendingView, "history") == 0;

    for (;;) {
        char tokens[HISTORY_THUMB_PUSH_BATCH][IMAGE_TOKEN_HEX_LEN + 1];
        size_t count;
        size_t scriptCap = 128;
        size_t pos = 0;
        wchar_t *script = NULL;
        int written = -1;

        AcquireSRWLockExclusive(&g_thumbLock);
        if (!historyOpen) g_thumbReadyCount = 0;
        count = g_thumbReadyCount < HISTORY_THUMB_PUSH_BATCH
            ? g_thumbReadyCount : HISTORY_THUMB_PUSH_BATCH;
        memcpy(tokens, g_thumbReady, count * sizeof(g_thumbReady[0]));
        g_thumbReadyCount -= count;
        memmove(g_thumbReady[0], g_thumbReady[count],
                g_thumbReadyCount * sizeof(g_thumbReady[0]));
        for (size_t i = 0; i < count; i++) {
            HistoryThumbnail *thumb = FindHistoryThumbnailLocked(tokens[i]);
            scriptCap += (thumb ? strlen(thumb->dataUri) : 0) +
                         IMAGE_TOKEN_HEX_LEN + 32;
        }
        if (count > 0) script = (wchar_t *)malloc(scriptCap * sizeof(wchar_t));
        if (script) {
            written = swprintf(script, scriptCap,
                L"window.onHistoryThumbs && window.onHistoryThumbs([");
            for (size_t i = 0; i < count && written >= 0; i++) {
                HistoryThumbnail *thumb = FindHistoryThumbnailLocked(tokens[i]);
                pos += (size_t)written;
                written = swprintf(script + pos, scriptCap - pos,
                    L"%s{\"token\":\"%hs\",\"thumb\":\"%hs\"}", i ? L"," : L"",
                    tokens[i], thumb ? thumb->dataUri : "");
            }
            if (written >= 0) {
                pos += (size_t)written;
                written = swprintf(script + pos, scriptCap - pos, L"])");
            }
        }
        ReleaseSRWLockExclusive(&g_thumbLock);

        if (count == 0) return;
        if (script && written >= 0) webview_execute_script(script);
        free(script);
    }
}

/* UI thread. tokens lists the rows the History page can see that still lack
   a thumbnail, top first. It replaces the previous request so rows already
   scrolled past are never generated; cached thumbnails are sent at once. */
static void RequestHistoryThumbnails(const char *tokens)
{
    char queue[HISTORY_THUMB_QUEUE_MAX][IMAGE_TOKEN_HEX_LEN + 1];
    size_t queued = 0;
    BOOL anyReady = FALSE;
    const char *cursor = tokens;

    AcquireSRWLockExclusive(&g_thumbLock);
    while (cursor && *cursor && queued < HISTORY_THUMB_QUEUE_MAX) {
        const char *start = cursor;
        const char *end = strchr(start, ',');
        size_t length = end ? (size_t)(end - start) : strlen(start);
        char token[IMAGE_TOKEN_HEX_LEN + 1];
        BOOL duplicate = FALSE;

        cursor = end ? end + 1 : NULL;
        if (length != IMAGE_TOKEN_HEX_LEN) continue;
        memcpy(token, start, length);
        token[length] = '\0';
        if (!IsValidImageToken(token)) continue;
        if (FindHistoryThumbnailLocked(token)) {
            MarkHistoryThumbnailReadyLocked(token);
            anyReady = TRUE;
            continue;
        }
        for (size_t i = 0; i < queued && !duplicate; i++) {
            duplicate = strcmp(queue[i], token) == 0;
        }
        if (!duplicate) memcpy(queue[queued++], token, sizeof(token));
    }
    memcpy(g_thumbQueue, queue, queued * sizeof(queue[0]));
    g_thumbQueueCount = queued;
    ReleaseSRWLockExclusive(&g_thumbLock);

    if (queued > 0) {
        if (EnsureHistoryThumbnailWorker()) {
            SetEvent(g_thumbWorkEvent);
        } else {
            AcquireSRWLockExclusive(&g_thumbLock);
            for (size_t i = 0; i < queued; i++) {
                MarkHistoryThumbnailReadyLocked(queue[i]);
            }
            g_thumbQueueCount = 0;
            ReleaseSRWLockExclusive(&g_thumbLock);
            anyReady = TRUE;
        }
    }
    if (anyReady) PushReadyHistoryThumbnails();
}

/* UI thread. Pending previews are pointless once the History dialog closes;
   finished ones stay cached for the next time it opens. */
static void DropHistoryThumbnailRequests(void)
{
    AcquireSRWLockExclusive(&g_thumbLock);
    g_thumbQueueCount = 0;
    g_thumbReadyCount = 0;
    ReleaseSRWLockExclusive(&g_thumbLock);
}

static void ForgetHistoryThumbnail(const char *token)
{
    AcquireSRWLockExclusive(&g_thumbLock);
    for (size_t i = 0; i < g_thumbCacheCount; i++) {
        if (strcmp(g_thumbCache[i].token, token) != 0) continue;
        free(g_thumbCache[i].dataUri);
        g_thumbCache[i] = g_thumbCache[--g_thumbCacheCount];
        ZeroMemory(&g_thumbCache[g_thumbCacheCount], sizeof(g_thumbCache[0]));
        break;
    }
    ReleaseSRWLockExclusive(&g_thumbLock);
}

static void ClearHistoryThumbnails(void)
{
    AcquireSRWLockExclusive(&g_thumbLock);
    for (size_t i = 0; i < g_thumbCacheCount; i++) free(g_thumbCache[i].dataUri);
    ZeroMemory(g_thumbCache, sizeof(g_thumbCache));
    g_thumbCacheCount = 0;
    g_thumbQueueCount = 0;
    g_thumbReadyCount = 0;
    ReleaseSRWLockExclusive(&g_thumbLock);
}

/* Sends one page of History metadata, newest first with the current image
   leading page 0. Thumbnails follow separately as rows become visible. The
   first push answers getInit; later ones refresh the open page in place. */
static void webview_push_history(BOOL initial)
{
    BOOL hasCurrent;
    size_t total;
    size_t pageCount;
    size_t first;
    size_t count;
    ULONGLONG totalBytes = 0;
    size_t scriptCap;
    wchar_t *script = NULL;
    size_t pos = 0;
    BOOL formatted = TRUE;
    int written;

    if (!g_webviewView) return;

    /* Mutations only happen on this thread, so the shared lock is held for
       the whole snapshot without risking a self-deadlock. */
    AcquireSRWLockShared(&g_imageLock);

    hasCurrent = g_cachedImage.token[0] &&
                 (g_cachedImage.jpegData || g_cachedImage.diskPath);
    total = (hasCurrent ? 1 : 0) + g_imageHistoryCount;
    if (hasCurrent) totalBytes += g_cachedImage.jpegSize;
    for (size_t i = 0; i < g_imageHistoryCount; i++) {
        totalBytes += g_imageHistory[i].jpegSize;
    }
    pageCount = total ? (total + HISTORY_VIEW_PAGE_SIZE - 1) / HISTORY_VIEW_PAGE_SIZE
                      : 1;
    if (g_historyViewPage >= pageCount) g_historyViewPage = pageCount - 1;
    first = g_historyViewPage * HISTORY_VIEW_PAGE_SIZE;
    count = total - first < HISTORY_VIEW_PAGE_SIZE ? total - first
                                                   : HISTORY_VIEW_PAGE_SIZE;

    scriptCap = 512 + count * (MAX_PATH * 2 + 640);
    script = (wchar_t *)malloc(scriptCap * sizeof(wchar_t));
    if (script) {
        written = swprintf(script + pos, scriptCap - pos,
            L"%s{\"pasteMethod\":\"%s\",\"historyLimit\":%d,"
            L"\"total\":%I64u,\"totalBytes\":%I64u,"
            L"\"page\":%I64u,\"pageSize\":%d,\"entries\":[",
            initial ? L"window.onInit({\"view\":\"history\",\"history\":"
                    : L"window.onHistoryData && window.onHistoryData(",
            g_configPasteMethod == PASTE_METHOD_HTTP ? L"http" : L"base64",
            g_configImageHistoryLimit, (ULONGLONG)total, totalBytes,
            (ULONGLONG)g_historyViewPage, HISTORY_VIEW_PAGE_SIZE);
        if (written < 0) formatted = FALSE; else pos += (size_t)written;
        for (size_t i = 0; i < count && formatted; i++) {
            size_t index = first + i;
            BOOL isCurrent = hasCurrent && index == 0;
            const CachedImage *image = isCurrent
                ? &g_cachedImage
                : &g_imageHistory[g_imageHistoryCount - 1 -
                                  (index - (hasCurrent ? 1 : 0))];
            wchar_t wDiskPath[MAX_PATH * 2];
            if (image->diskPath) {
                json_escape_wstring(image->diskPath, wDiskPath, MAX_PATH * 2);
            } else {
                wDiskPath[0] = L'\0';
            }
            written = swprintf(script + pos, scriptCap - pos,
                L"%s{\"token\":\"%hs\",\"current\":%s,"
                L"\"width\":%u,\"height\":%u,\"bytes\":%lu,"
                L"\"capturedAt\":%I64u,"
                L"\"storage\":\"%s\",\"path\":\"%s\","
                L"\"url\":\"http://%hs:%d/%hs.jpg\"}",
                i == 0 ? L"" : L",",
                image->token,
                isCurrent ? L"true" : L"false",
                image->width, image->height,
                (unsigned long)image->jpegSize,
                image->capturedAtMs,
                image->diskPath ? L"disk" : L"memory", wDiskPath,
                g_configBindIp, g_configHttpPort, image->token);
            if (written < 0) formatted = FALSE; else pos += (size_t)written;
        }
        if (formatted) {
            written = swprintf(script + pos, scriptCap - pos, L"]}%s",
                               initial ? L"})" : L")");
            if (written < 0) formatted = FALSE;
        }
    }

    ReleaseSRWLockShared(&g_imageLock);

    if (script) {
        if (formatted) webview_execute_script(script);
        free(script);
    }
}

static void NotifyHistoryViewChanged(void)
{
    if (g_webviewView && strcmp(g_pendingView, "history") == 0) {
        webview_push_history(FALSE);
    }
}

static void SendHistoryActionResult(BOOL ok, const wchar_t *message)
{
    wchar_t escaped[640];
    wchar_t script[768];

    if (!g_webviewView) return;
    json_escape_wstring(message, escaped, 640);
    swprintf(script, 768,
             L"window.onHistoryActionResult && window.onHistoryActionResult("
             L"{\"ok\":%s,\"message\":\"%s\"})",
             ok ? L"true" : L"false", escaped);
    webview_execute_script(script);
}

static void SaveHistoryImageToDisk(const char *token)
{
    BYTE *jpegData = NULL;
    DWORD jpegSize = 0;
    ULONGLONG capturedAtMs = 0;
    WCHAR filePath[MAX_PATH];
    OPENFILENAMEW dialog;
    SYSTEMTIME localTime;
    BOOL haveTime = FALSE;

    if (!CopyRetainedImageByToken(token, &jpegData, &jpegSize,
                                  &capturedAtMs)) {
        SendHistoryActionResult(FALSE,
            L"This image is no longer retained.");
        NotifyHistoryViewChanged();
        return;
    }

    if (capturedAtMs > 0) {
        FILETIME utcFile;
        FILETIME localFile;
        ULARGE_INTEGER value;
        value.QuadPart = capturedAtMs * 10000ULL + 116444736000000000ULL;
        utcFile.dwLowDateTime = value.LowPart;
        utcFile.dwHighDateTime = value.HighPart;
        haveTime = FileTimeToLocalFileTime(&utcFile, &localFile) &&
                   FileTimeToSystemTime(&localFile, &localTime);
    }
    if (haveTime) {
        swprintf(filePath, MAX_PATH,
                 L"ImagePaster-%04u%02u%02u-%02u%02u%02u.jpg",
                 localTime.wYear, localTime.wMonth, localTime.wDay,
                 localTime.wHour, localTime.wMinute, localTime.wSecond);
    } else {
        wcscpy(filePath, L"ImagePaster-image.jpg");
    }

    ZeroMemory(&dialog, sizeof(dialog));
    dialog.lStructSize = sizeof(dialog);
    dialog.hwndOwner = g_webviewHwnd;
    dialog.lpstrFilter = L"JPEG image (*.jpg)\0*.jpg\0All files (*.*)\0*.*\0";
    dialog.lpstrFile = filePath;
    dialog.nMaxFile = MAX_PATH;
    dialog.lpstrDefExt = L"jpg";
    dialog.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;

    if (GetSaveFileNameW(&dialog)) {
        HANDLE file = CreateFileW(filePath, GENERIC_WRITE, 0, NULL,
                                  CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (file != INVALID_HANDLE_VALUE) {
            DWORD bytesWritten = 0;
            BOOL wrote = WriteFile(file, jpegData, jpegSize,
                                   &bytesWritten, NULL) &&
                         bytesWritten == jpegSize;
            CloseHandle(file);
            if (wrote) {
                wchar_t message[MAX_PATH + 32];
                swprintf(message, MAX_PATH + 32, L"Saved to %s", filePath);
                LogMessage("History image %.12s... saved to disk", token);
                SendHistoryActionResult(TRUE, message);
            } else {
                DeleteFileW(filePath);
                SendHistoryActionResult(FALSE,
                    L"The file could not be written.");
            }
        } else {
            SendHistoryActionResult(FALSE, L"The file could not be created.");
        }
    }
    free(jpegData);
}

/* ── COM callback handler implementations ────────────────────────────── */

static HRESULT STDMETHODCALLTYPE EnvCompleted_Invoke(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*, HRESULT, ICoreWebView2Environment*);
static HRESULT STDMETHODCALLTYPE CtrlCompleted_Invoke(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, HRESULT, ICoreWebView2Controller*);
static HRESULT STDMETHODCALLTYPE MsgReceived_Invoke(ICoreWebView2WebMessageReceivedEventHandler*, ICoreWebView2*, ICoreWebView2WebMessageReceivedEventArgs*);

static HRESULT STDMETHODCALLTYPE EnvCompleted_QueryInterface(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This, REFIID riid, void **ppv) {
    (void)riid;
    *ppv = This;
    This->lpVtbl->AddRef(This);
    return S_OK;
}
static ULONG STDMETHODCALLTYPE EnvCompleted_AddRef(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This) {
    return ++This->refCount;
}
static ULONG STDMETHODCALLTYPE EnvCompleted_Release(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This) {
    ULONG rc = --This->refCount;
    if (rc == 0) free(This);
    return rc;
}

static HRESULT STDMETHODCALLTYPE EnvCompleted_Invoke(ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *This, HRESULT result, ICoreWebView2Environment *env) {
    (void)This;
    if (FAILED(result) || !env) return result;
    g_webviewEnv = env;
    env->lpVtbl->AddRef(env);

    static ControllerCompletedHandlerVtbl ctrlVtbl = {0};
    static BOOL ctrlVtblInit = FALSE;
    if (!ctrlVtblInit) {
        ctrlVtbl.QueryInterface = (HRESULT (STDMETHODCALLTYPE *)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*, REFIID, void**))EnvCompleted_QueryInterface;
        ctrlVtbl.AddRef = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*))EnvCompleted_AddRef;
        ctrlVtbl.Release = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*))EnvCompleted_Release;
        ctrlVtbl.Invoke = CtrlCompleted_Invoke;
        ctrlVtblInit = TRUE;
    }

    ICoreWebView2CreateCoreWebView2ControllerCompletedHandler *handler = malloc(sizeof(*handler));
    handler->lpVtbl = &ctrlVtbl;
    handler->refCount = 1;

    env->lpVtbl->CreateCoreWebView2Controller(env, g_webviewHwnd, handler);
    handler->lpVtbl->Release(handler);
    return S_OK;
}

static EnvironmentCompletedHandlerVtbl g_envCompletedVtbl = {
    EnvCompleted_QueryInterface,
    EnvCompleted_AddRef,
    EnvCompleted_Release,
    EnvCompleted_Invoke
};

static HRESULT STDMETHODCALLTYPE CtrlCompleted_Invoke(ICoreWebView2CreateCoreWebView2ControllerCompletedHandler *This, HRESULT result, ICoreWebView2Controller *controller) {
    (void)This;
    if (FAILED(result) || !controller) return result;

    g_webviewController = controller;
    controller->lpVtbl->AddRef(controller);

    RECT bounds;
    GetClientRect(g_webviewHwnd, &bounds);
    controller->lpVtbl->put_Bounds(controller, bounds);
    controller->lpVtbl->put_IsVisible(controller, TRUE);

    ICoreWebView2 *webview = NULL;
    controller->lpVtbl->get_CoreWebView2(controller, &webview);
    if (!webview) return E_FAIL;
    g_webviewView = webview;

    ICoreWebView2Settings *settings = NULL;
    webview->lpVtbl->get_Settings(webview, &settings);
    if (settings) {
        settings->lpVtbl->put_AreDefaultContextMenusEnabled(settings, FALSE);
        settings->lpVtbl->put_AreDevToolsEnabled(settings, FALSE);
        settings->lpVtbl->put_IsStatusBarEnabled(settings, FALSE);
        settings->lpVtbl->put_IsZoomControlEnabled(settings, FALSE);
        settings->lpVtbl->Release(settings);
    }

    static WebMessageReceivedHandlerVtbl msgVtbl = {0};
    static BOOL msgVtblInit = FALSE;
    if (!msgVtblInit) {
        msgVtbl.QueryInterface = (HRESULT (STDMETHODCALLTYPE *)(ICoreWebView2WebMessageReceivedEventHandler*, REFIID, void**))EnvCompleted_QueryInterface;
        msgVtbl.AddRef = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2WebMessageReceivedEventHandler*))EnvCompleted_AddRef;
        msgVtbl.Release = (ULONG (STDMETHODCALLTYPE *)(ICoreWebView2WebMessageReceivedEventHandler*))EnvCompleted_Release;
        msgVtbl.Invoke = MsgReceived_Invoke;
        msgVtblInit = TRUE;
    }

    ICoreWebView2WebMessageReceivedEventHandler *msgHandler = malloc(sizeof(*msgHandler));
    msgHandler->lpVtbl = &msgVtbl;
    msgHandler->refCount = 1;

    EventRegistrationToken token;
    webview->lpVtbl->add_WebMessageReceived(webview, msgHandler, &token);
    msgHandler->lpVtbl->Release(msgHandler);

    /* Load embedded HTML from resources */
    HRSRC hRes = FindResource(NULL, MAKEINTRESOURCE(IDR_HTML_UI), RT_RCDATA);
    if (hRes) {
        HGLOBAL hData = LoadResource(NULL, hRes);
        if (hData) {
            DWORD htmlSize = SizeofResource(NULL, hRes);
            const char *htmlUtf8 = (const char *)LockResource(hData);
            if (htmlUtf8 && htmlSize > 0) {
                int wLen = MultiByteToWideChar(CP_UTF8, 0, htmlUtf8, (int)htmlSize, NULL, 0);
                wchar_t *wHtml = malloc((wLen + 1) * sizeof(wchar_t));
                MultiByteToWideChar(CP_UTF8, 0, htmlUtf8, (int)htmlSize, wHtml, wLen);
                wHtml[wLen] = L'\0';
                webview->lpVtbl->NavigateToString(webview, wHtml);
                free(wHtml);
            }
        }
    }

    return S_OK;
}

/* ── WebMessageReceivedHandler ─────────────────────────────────────────── */

static HRESULT STDMETHODCALLTYPE MsgReceived_Invoke(ICoreWebView2WebMessageReceivedEventHandler *This, ICoreWebView2 *sender, ICoreWebView2WebMessageReceivedEventArgs *args) {
    (void)This; (void)sender;

    LPWSTR wMsg = NULL;
    args->lpVtbl->TryGetWebMessageAsString(args, &wMsg);
    if (!wMsg) return S_OK;

    int len = WideCharToMultiByte(CP_UTF8, 0, wMsg, -1, NULL, 0, NULL, NULL);
    char *msg = malloc(len);
    WideCharToMultiByte(CP_UTF8, 0, wMsg, -1, msg, len, NULL, NULL);
    CoTaskMemFree(wMsg);

    char action[64] = {0};
    json_get_string(msg, "action", action, sizeof(action));

    if (strcmp(action, "getInit") == 0) {
        if (strcmp(g_pendingView, "config") == 0) {
            webview_push_init_config();
        } else if (strcmp(g_pendingView, "log") == 0) {
            webview_push_init_log();
        } else if (strcmp(g_pendingView, "history") == 0) {
            g_historyViewPage = 0;
            webview_push_history(TRUE);
        }
    } else if (strcmp(action, "configReady") == 0) {
        if (IsConfigurationViewOpen()) {
            int checkAutomatically = 0;
            // A stopped worker that is still exiting does not count; the
            // automatic check is queued behind it by StartUpdateCheck.
            BOOL updateWorkAlreadyActive = IsUpdateCheckVisible() ||
                g_updateNoticeTask || g_updateReadyTask;
            json_get_int(msg, "checkAutomatically", &checkAutomatically);
            g_configViewReady = TRUE;
            PresentPendingUpdateNotice();
            if (checkAutomatically != 0 && g_configAutoCheckForUpdates &&
                !g_updateConfirmationPending && !updateWorkAlreadyActive) {
                StartUpdateCheck(TRUE);
            }
        }
    } else if (strcmp(action, "checkUpdate") == 0) {
        int automatic = 0;
        json_get_int(msg, "automatic", &automatic);
        StartUpdateCheck(automatic != 0);
    } else if (strcmp(action, "cancelUpdateCheck") == 0) {
        CancelUpdateCheck();
    } else if (strcmp(action, "installUpdate") == 0) {
        int reopenSettings = 0;
        json_get_int(msg, "reopenSettingsAfterUpdate", &reopenSettings);
        InstallPreparedUpdate(reopenSettings == 1);
    } else if (strcmp(action, "dismissUpdate") == 0) {
        DiscardPreparedUpdate();
    } else if (strcmp(action, "ignoreUpdateVersion") == 0) {
        char version[32] = {0};
        json_get_string(msg, "version", version, sizeof(version));
        IgnorePreparedUpdateVersion(version);
    } else if (strcmp(action, "dismissUpdateConfirmation") == 0) {
        g_updateConfirmationPending = FALSE;
    } else if (strcmp(action, "saveSettings") == 0) {
        char titleMatch[2048] = {0};
        char pasteMethod[32] = {0};
        char httpMessageTemplate[MAX_HTTP_MESSAGE_TEMPLATE_BYTES] = {0};
        char captureGapFill[16] = {0};
        char bindIp[INET_ADDRSTRLEN] = {0};
        char httpAllowList[MAX_HTTP_ALLOW_LIST_BYTES] = {0};
        char imageStorage[16] = {0};
        HttpAllowRule allowScratch[MAX_HTTP_ALLOW_RULES];
        int allowScratchCount = 0;
        int httpPort = 0;
        int jpegQuality = -1;
        int imageHistoryLimit = -1;
        int compatibilityPaste = -1;
        int screenCaptureEnabled = -1;
        int autoCheckForUpdates = -1;
        BOOL httpMessageParsed;
        IN_ADDR parsedAddress;

        json_get_string(msg, "titleMatch", titleMatch, sizeof(titleMatch));
        json_get_string(msg, "pasteMethod", pasteMethod, sizeof(pasteMethod));
        httpMessageParsed = json_get_string(
            msg, "httpMessageTemplate", httpMessageTemplate,
            sizeof(httpMessageTemplate));
        json_get_string(msg, "bindIp", bindIp, sizeof(bindIp));
        json_get_string(msg, "httpAllowList", httpAllowList,
                        sizeof(httpAllowList));
        json_get_string(msg, "imageStorage", imageStorage,
                        sizeof(imageStorage));
        json_get_int(msg, "httpPort", &httpPort);
        json_get_int(msg, "jpegQuality", &jpegQuality);
        json_get_int(msg, "imageHistoryLimit", &imageHistoryLimit);
        json_get_int(msg, "compatibilityPaste", &compatibilityPaste);
        json_get_int(msg, "screenCaptureEnabled", &screenCaptureEnabled);
        json_get_string(msg, "captureGapFill", captureGapFill,
                        sizeof(captureGapFill));
        json_get_int(msg, "autoCheckForUpdates", &autoCheckForUpdates);

        if (strcmp(pasteMethod, "base64") != 0 && strcmp(pasteMethod, "http") != 0) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Select a valid paste method.'})");
            free(msg);
            return S_OK;
        }
        if (!httpMessageParsed) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'The HTTP paste message is invalid or too long.'})");
            free(msg);
            return S_OK;
        }
        if (!strstr(httpMessageTemplate, HTTP_URL_PLACEHOLDER)) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'The HTTP paste message must include {URL}.'})");
            free(msg);
            return S_OK;
        }
        if (InetPtonA(AF_INET, bindIp, &parsedAddress) != 1) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Enter a valid IPv4 bind address.'})");
            free(msg);
            return S_OK;
        }
        if (httpPort < 1 || httpPort > 65535) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'HTTP port must be between 1 and 65535.'})");
            free(msg);
            return S_OK;
        }
        if (!ParseHttpAllowList(httpAllowList, allowScratch,
                                MAX_HTTP_ALLOW_RULES, &allowScratchCount)) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Allowed clients must be IPv4 addresses or CIDR subnets separated by commas (at most 64 entries).'})");
            free(msg);
            return S_OK;
        }
        if (strcmp(imageStorage, "memory") != 0 &&
            strcmp(imageStorage, "disk") != 0) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Select a valid image storage location.'})");
            free(msg);
            return S_OK;
        }
        if (jpegQuality < 0 || jpegQuality > 100) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'JPEG quality must be between 0 and 100.'})");
            free(msg);
            return S_OK;
        }
        if (imageHistoryLimit < 0 || imageHistoryLimit > MAX_IMAGE_HISTORY_LIMIT) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Image history must be Unlimited or between 1 and 1000.'})");
            free(msg);
            return S_OK;
        }
        if (compatibilityPaste < 0 || compatibilityPaste > 1) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Select a valid text-paste shortcut.'})");
            free(msg);
            return S_OK;
        }
        if (screenCaptureEnabled < 0 || screenCaptureEnabled > 1) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Select a valid Print Screen capture setting.'})");
            free(msg);
            return S_OK;
        }
        if (strcmp(captureGapFill, "white") != 0 &&
            strcmp(captureGapFill, "black") != 0 &&
            strcmp(captureGapFill, "blur") != 0) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Select a valid multi-region gap fill.'})");
            free(msg);
            return S_OK;
        }
        if (autoCheckForUpdates < 0 || autoCheckForUpdates > 1) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Select a valid automatic update setting.'})");
            free(msg);
            return S_OK;
        }

        BOOL networkChanged = strcmp(g_configBindIp, bindIp) != 0 ||
                              g_configHttpPort != httpPort;
        BOOL qualityChanged = g_configJpegQuality != jpegQuality;
        if (networkChanged) StopHttpServer();

        strncpy(g_configTitleMatch, titleMatch, sizeof(g_configTitleMatch) - 1);
        g_configTitleMatch[sizeof(g_configTitleMatch) - 1] = '\0';
        g_configPasteMethod = strcmp(pasteMethod, "http") == 0
            ? PASTE_METHOD_HTTP : PASTE_METHOD_BASE64;
        strncpy(g_configHttpMessageTemplate, httpMessageTemplate,
                sizeof(g_configHttpMessageTemplate) - 1);
        g_configHttpMessageTemplate[
            sizeof(g_configHttpMessageTemplate) - 1] = '\0';
        strncpy(g_configBindIp, bindIp, sizeof(g_configBindIp) - 1);
        g_configBindIp[sizeof(g_configBindIp) - 1] = '\0';
        g_configHttpPort = httpPort;
        strncpy(g_configHttpAllowList, httpAllowList,
                sizeof(g_configHttpAllowList) - 1);
        g_configHttpAllowList[sizeof(g_configHttpAllowList) - 1] = '\0';
        ApplyHttpAllowList(); /* takes effect immediately, no restart needed */
        /* Applies to newly copied images; existing entries keep their
           current location until they are evicted. */
        g_configImageStorage = strcmp(imageStorage, "disk") == 0
            ? IMAGE_STORAGE_DISK : IMAGE_STORAGE_MEMORY;
        g_configJpegQuality = jpegQuality;
        {
            size_t evicted = SetImageHistoryLimit(imageHistoryLimit);
            if (evicted > 0) {
                LogMessage("Image history limit evicted %llu image(s)",
                           (unsigned long long)evicted);
            }
        }
        if (g_configImageStorage == IMAGE_STORAGE_DISK) {
            LoadPersistedDiskHistory();
        }
        g_configCompatibilityPaste = compatibilityPaste != 0;
        g_configScreenCaptureEnabled = screenCaptureEnabled != 0;
        g_configCaptureGapFill = strcmp(captureGapFill, "black") == 0
            ? CAPTURE_GAP_FILL_BLACK
            : strcmp(captureGapFill, "blur") == 0
                ? CAPTURE_GAP_FILL_BLUR
                : CAPTURE_GAP_FILL_WHITE;
        if (!g_configScreenCaptureEnabled && g_captureOverlayHwnd) {
            CancelScreenCapture("feature was disabled");
        }
        g_configAutoCheckForUpdates = autoCheckForUpdates != 0;
        SaveConfigToRegistry();
        ParseKeywords();
        UpdateTooltip();
        if (qualityChanged && IsClipboardFormatAvailable(CF_DIB)) {
            SetCachedClipboardSequence(0);
            RefreshClipboardImageCache();
        }
        ReconcileHttpServer();
        LogMessage("Configuration updated");
        PostMessage(g_webviewHwnd, WM_CLOSE, 0, 0);
    } else if (strcmp(action, "close") == 0) {
        PostMessage(g_webviewHwnd, WM_CLOSE, 0, 0);
    } else if (strcmp(action, "clearLog") == 0) {
        g_logCount = 0;
        g_logHead = 0;
        /* Push empty log array back to JS */
        webview_execute_script(L"window.onInit && window.onInit({\"view\":\"log\",\"log\":[]})");
    } else if (strcmp(action, "copyLog") == 0) {
        CopyActivityLogToClipboard();
    } else if (strcmp(action, "historyPage") == 0) {
        int page = 0;
        json_get_int(msg, "page", &page);
        if (strcmp(g_pendingView, "history") == 0) {
            g_historyViewPage = page > 0 ? (size_t)page : 0;
            webview_push_history(FALSE);
        }
    } else if (strcmp(action, "historyThumbs") == 0) {
        char tokens[HISTORY_THUMB_QUEUE_MAX * (IMAGE_TOKEN_HEX_LEN + 1) + 1];
        if (!json_get_string(msg, "tokens", tokens, sizeof(tokens))) {
            tokens[0] = '\0';
        }
        if (strcmp(g_pendingView, "history") == 0) {
            RequestHistoryThumbnails(tokens);
        }
    } else if (strcmp(action, "historyCopyUrl") == 0) {
        char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
        char url[192];
        int urlWritten;
        json_get_string(msg, "token", token, sizeof(token));
        urlWritten = snprintf(url, sizeof(url), "http://%s:%d/%s.jpg",
                              g_configBindIp, g_configHttpPort, token);
        if (token[0] && urlWritten > 0 && urlWritten < (int)sizeof(url) &&
            PlaceUtf8TextOnClipboard(url, TRUE)) {
            LogMessage("History image URL copied to clipboard (id %.12s...)",
                       token);
            SendHistoryActionResult(TRUE, L"Image URL copied to the clipboard.");
        } else {
            SendHistoryActionResult(FALSE, L"The URL could not be copied.");
        }
    } else if (strcmp(action, "historyOpenUrl") == 0) {
        char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
        wchar_t url[192];
        int urlWritten;
        json_get_string(msg, "token", token, sizeof(token));
        urlWritten = swprintf(url, 192, L"http://%hs:%d/%hs.jpg",
                              g_configBindIp, g_configHttpPort, token);
        if (token[0] && urlWritten > 0 &&
            (INT_PTR)ShellExecuteW(NULL, L"open", url, NULL, NULL,
                                   SW_SHOWNORMAL) > 32) {
            LogMessage("History image URL opened in browser (id %.12s...)",
                       token);
        } else {
            SendHistoryActionResult(FALSE, L"The link could not be opened.");
        }
    } else if (strcmp(action, "historySave") == 0) {
        char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
        json_get_string(msg, "token", token, sizeof(token));
        SaveHistoryImageToDisk(token);
    } else if (strcmp(action, "historyRevealFile") == 0) {
        char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
        WCHAR *path;
        json_get_string(msg, "token", token, sizeof(token));
        path = GetRetainedImageDiskPathCopy(token);
        if (!path) {
            SendHistoryActionResult(FALSE,
                L"This image is not stored on disk.");
        } else {
            LPITEMIDLIST item = ILCreateFromPathW(path);
            BOOL revealed = FALSE;
            if (item) {
                revealed = SUCCEEDED(
                    SHOpenFolderAndSelectItems(item, 0, NULL, 0));
                ILFree(item);
            }
            if (revealed) {
                LogMessage("History image file shown in Explorer (id %.12s...)",
                           token);
            } else {
                SendHistoryActionResult(FALSE,
                    L"The file could not be shown in Explorer.");
            }
            free(path);
        }
    } else if (strcmp(action, "historyDelete") == 0) {
        char token[IMAGE_TOKEN_HEX_LEN + 1] = {0};
        json_get_string(msg, "token", token, sizeof(token));
        if (DeleteRetainedImageByToken(token)) {
            ForgetHistoryThumbnail(token);
            LogMessage("History image %.12s... removed by user", token);
            SendHistoryActionResult(TRUE, L"Image removed.");
        } else {
            SendHistoryActionResult(FALSE,
                L"This image is no longer retained.");
        }
        NotifyHistoryViewChanged();
    } else if (strcmp(action, "historyClearAll") == 0) {
        size_t removed = ClearAllRetainedImages();
        wchar_t message[80];
        ClearHistoryThumbnails();
        LogMessage("History cleared by user (%llu image(s) discarded)",
                   (unsigned long long)removed);
        swprintf(message, 80, L"Removed %I64u retained image%s.",
                 (ULONGLONG)removed, removed == 1 ? L"" : L"s");
        SendHistoryActionResult(TRUE, message);
        NotifyHistoryViewChanged();
    } else if (strcmp(action, "resize") == 0) {
        int contentHeight = 0;
        int contentWidth = 0;
        json_get_int(msg, "height", &contentHeight);
        json_get_int(msg, "width", &contentWidth);
        if (contentHeight > 0 && g_webviewHwnd &&
            !IsZoomed(g_webviewHwnd) && !IsIconic(g_webviewHwnd)) {
            int physicalHeight = MulDiv(
                contentHeight, (int)GetWebViewWindowDpi(g_webviewHwnd), 96);
            RECT clientRect = {0}, windowRect = {0};
            GetClientRect(g_webviewHwnd, &clientRect);
            GetWindowRect(g_webviewHwnd, &windowRect);
            int chromeH = (windowRect.bottom - windowRect.top) - (clientRect.bottom - clientRect.top);
            int newWindowH = physicalHeight + chromeH;
            int windowW = windowRect.right - windowRect.left;
            if (contentWidth > 0) {
                int chromeW = (windowRect.right - windowRect.left) -
                              (clientRect.right - clientRect.left);
                windowW = MulDiv(contentWidth,
                                 (int)GetWebViewWindowDpi(g_webviewHwnd),
                                 96) + chromeW;
            }
            int newX = windowRect.left;
            int newY = windowRect.top;
            HMONITOR monitor = MonitorFromWindow(g_webviewHwnd,
                                                  MONITOR_DEFAULTTONEAREST);
            MONITORINFO monitorInfo = {0};
            UINT flags = SWP_NOZORDER;
            monitorInfo.cbSize = sizeof(monitorInfo);
            if (monitor && GetMonitorInfoW(monitor, &monitorInfo)) {
                int availableHeight = monitorInfo.rcWork.bottom -
                                      monitorInfo.rcWork.top - 24;
                int availableWidth = monitorInfo.rcWork.right -
                                     monitorInfo.rcWork.left - 24;
                if (newWindowH > availableHeight) newWindowH = availableHeight;
                if (windowW > availableWidth) windowW = availableWidth;
                if (newY + newWindowH > monitorInfo.rcWork.bottom - 12) {
                    newY = monitorInfo.rcWork.bottom - newWindowH - 12;
                }
                if (newY < monitorInfo.rcWork.top + 12) {
                    newY = monitorInfo.rcWork.top + 12;
                }
                if (newX + windowW > monitorInfo.rcWork.right - 12) {
                    newX = monitorInfo.rcWork.right - windowW - 12;
                }
                if (newX < monitorInfo.rcWork.left + 12) {
                    newX = monitorInfo.rcWork.left + 12;
                }
            }
            if (g_webviewWindowShown) {
                flags |= SWP_NOACTIVATE;
            } else {
                flags |= SWP_SHOWWINDOW;
                KillTimer(g_webviewHwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
            }
            SetWindowPos(g_webviewHwnd, NULL, newX, newY,
                         windowW, newWindowH, flags);
            g_webviewWindowShown = TRUE;
            webview_sync_controller_bounds();
        }
    }

    free(msg);
    return S_OK;
}

/* ── WebView2 window ───────────────────────────────────────────────────── */

static LRESULT CALLBACK WebViewWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_APP_UPDATE_PROGRESS:
            InterlockedExchange(&g_updateProgressPosted, FALSE);
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE &&
                !g_updateCheckAbandoned) {
                DWORD percent = (DWORD)InterlockedCompareExchange(
                    &g_updateProgressPercent, 0, 0);
                CfgSendUpdateProgress(percent);
            }
            return 0;

        case WM_APP_UPDATE_RESULT: {
            UpdateCheckTask* task = (UpdateCheckTask*)InterlockedExchangePointer(
                (PVOID volatile*)&g_updatePostedResult, NULL);
            HandleCompletedUpdateCheck(task);
            return 0;
        }

        case WM_SIZE:
            webview_sync_controller_bounds();
            return 0;

        case WM_DPICHANGED: {
            const RECT *suggested = (const RECT *)lParam;
            SetWindowPos(hwnd, NULL, suggested->left, suggested->top,
                         suggested->right - suggested->left,
                         suggested->bottom - suggested->top,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            return 0;
        }

        case WM_TIMER:
            if (wParam == ID_TIMER_WEBVIEW_SHOW_FALLBACK) {
                KillTimer(hwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
                if (!g_webviewWindowShown) {
                    ShowWindow(hwnd, SW_SHOWNOACTIVATE);
                    UpdateWindow(hwnd);
                    g_webviewWindowShown = TRUE;
                    webview_sync_controller_bounds();
                }
                return 0;
            }
            break;

        case WM_CLOSE:
            g_webviewWindowShown = FALSE;
            KillTimer(hwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
            if (g_webviewController) {
                g_webviewController->lpVtbl->Close(g_webviewController);
                g_webviewController->lpVtbl->Release(g_webviewController);
                g_webviewController = NULL;
            }
            if (g_webviewView) {
                g_webviewView->lpVtbl->Release(g_webviewView);
                g_webviewView = NULL;
            }
            if (g_webviewEnv) {
                g_webviewEnv->lpVtbl->Release(g_webviewEnv);
                g_webviewEnv = NULL;
            }
            DestroyWindow(hwnd);
            return 0;

        case WM_DESTROY:
            // Closing settings stops a manual check, including a queued one.
            if (g_updateQueuedCheck && !g_updateQueuedAutomatic) {
                g_updateQueuedCheck = FALSE;
            }
            if (InterlockedCompareExchange(&g_updateCheckPending,
                                           FALSE, FALSE) == TRUE &&
                InterlockedCompareExchange(&g_updateCheckAutomatic,
                                           FALSE, FALSE) == FALSE) {
                if (g_updateCancelEvent) SetEvent(g_updateCancelEvent);
                g_updateCheckAbandoned = TRUE;
            }
            DropHistoryThumbnailRequests();
            DiscardPendingUpdateNotice();
            DiscardPreparedUpdate();
            g_webviewHwnd = NULL;
            g_webviewWindowShown = FALSE;
            g_configViewReady = FALSE;
            KillTimer(hwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
            if (g_updateInstallReady && g_hWndMain) {
                PostMessageW(g_hWndMain, WM_COMMAND, ID_TRAY_EXIT, 0);
            }
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowWebViewDialog(const char* view, int width, int height) {
    if (g_webviewHwnd != NULL) {
        if (IsIconic(g_webviewHwnd)) ShowWindow(g_webviewHwnd, SW_RESTORE);
        else ShowWindow(g_webviewHwnd, SW_SHOW);
        SetWindowPos(g_webviewHwnd, HWND_TOP, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
        SetForegroundWindow(g_webviewHwnd);
        return;
    }

    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    if (!fnCreateEnvironment && !load_webview2_loader()) {
        return;
    }

    strncpy(g_pendingView, view, sizeof(g_pendingView) - 1);
    g_pendingView[sizeof(g_pendingView) - 1] = '\0';

    /* Register window class (once) */
    static BOOL classRegistered = FALSE;
    if (!classRegistered) {
        WNDCLASSEXW wc = {0};
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = WebViewWndProc;
        wc.hInstance = g_hInstance;
        wc.hIcon = (HICON)LoadImageW(g_hInstance, MAKEINTRESOURCEW(IDI_APPICON),
            IMAGE_ICON, GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON), LR_DEFAULTCOLOR);
        wc.hCursor = LoadCursor(NULL, IDC_ARROW);
        wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = L"ImagePasterWebViewWnd";
        wc.hIconSm = (HICON)LoadImageW(g_hInstance, MAKEINTRESOURCEW(IDI_APPICON),
            IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);
        RegisterClassExW(&wc);
        classRegistered = TRUE;
    }

    const wchar_t *title = L"Configuration";
    if (strcmp(view, "log") == 0) title = L"Activity Log";
    else if (strcmp(view, "history") == 0) title = L"Image History";

    RECT workArea;
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &workArea, 0);
    int dpi = (int)GetWebViewWindowDpi(NULL);
    width = MulDiv(width, dpi, 96);
    height = MulDiv(height, dpi, 96);
    if (height > workArea.bottom - workArea.top) {
        height = workArea.bottom - workArea.top;
    }
    int posX = workArea.left + ((workArea.right - workArea.left) - width) / 2;
    int posY = workArea.top + ((workArea.bottom - workArea.top) - height) / 2;

    g_webviewHwnd = CreateWindowExW(0, L"ImagePasterWebViewWnd", title,
        WS_OVERLAPPEDWINDOW,
        posX, posY, width, height,
        NULL, NULL, g_hInstance, NULL);

    if (!g_webviewHwnd) {
        LogMessage("ERROR: Failed to create WebView2 window.");
        return;
    }
    g_webviewWindowShown = FALSE;
    g_configViewReady = FALSE;
    SetTimer(g_webviewHwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK, WEBVIEW_SHOW_FALLBACK_DELAY_MS, NULL);

    WCHAR userDataFolder[MAX_PATH];
    DWORD tempLen = GetTempPathW(MAX_PATH, userDataFolder);
    if (tempLen > 0 && tempLen < MAX_PATH - 30) {
        wcscat(userDataFolder, L"ImagePaster.WebView2");
    } else {
        wcscpy(userDataFolder, L"");
    }

    ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler *envHandler = malloc(sizeof(*envHandler));
    envHandler->lpVtbl = &g_envCompletedVtbl;
    envHandler->refCount = 1;

    HRESULT hr = fnCreateEnvironment(NULL, userDataFolder[0] ? userDataFolder : NULL, NULL, envHandler);
    envHandler->lpVtbl->Release(envHandler);

    if (FAILED(hr)) {
        LogMessage("ERROR: Failed to initialize WebView2 environment.");
        MessageBoxW(NULL,
            L"Failed to initialize WebView2.\n\n"
            L"Please ensure the Microsoft Edge WebView2 Runtime is installed.\n"
            L"Download from: https://developer.microsoft.com/en-us/microsoft-edge/webview2/",
            APP_NAME, MB_ICONERROR | MB_OK);
        DestroyWindow(g_webviewHwnd);
        g_webviewHwnd = NULL;
    }
}

/* ── Window procedure (hidden message window) ──────────────────────────── */

static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_TRAYICON:
        if (lParam == WM_LBUTTONDBLCLK) {
            if (g_configScreenCaptureEnabled) {
                LogMessage("Screen capture requested by tray icon double-click");
                PostMessage(hWnd, WM_SCREEN_CAPTURE_BEGIN, 0, 0);
            } else {
                LogMessage("Opening Configuration dialog from tray icon double-click");
                if (g_webviewHwnd && !IsConfigurationViewOpen()) {
                    SendMessageW(g_webviewHwnd, WM_CLOSE, 0, 0);
                }
                ShowWebViewDialog("config", 560, 520);
            }
        } else if (lParam == WM_RBUTTONUP) {
            POINT pt;
            GetCursorPos(&pt);
            SetForegroundWindow(hWnd);
            EnableMenuItem(g_hMenu, ID_TRAY_CAPTURE,
                           g_captureOverlayHwnd ? MF_GRAYED : MF_ENABLED);
            EnableMenuItem(g_hMenu, ID_TRAY_HISTORY,
                           (g_webviewHwnd || !AnyRetainedImages())
                               ? MF_GRAYED : MF_ENABLED);
            EnableMenuItem(g_hMenu, ID_TRAY_LOG, g_webviewHwnd ? MF_GRAYED : MF_ENABLED);
            EnableMenuItem(g_hMenu, ID_TRAY_CONFIGURE,
                           (!g_webviewHwnd || IsConfigurationViewOpen())
                               ? MF_ENABLED : MF_GRAYED);
            TrackPopupMenu(g_hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
        }
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_TRAY_CAPTURE:
            LogMessage("Screen capture requested from the tray menu");
            PostMessage(hWnd, WM_SCREEN_CAPTURE_BEGIN, 0, 0);
            break;
        case ID_TRAY_HISTORY:
            LogMessage("Opening History dialog");
            ShowWebViewDialog("history", 640, 560);
            break;
        case ID_TRAY_LOG:
            LogMessage("Opening Activity Log dialog");
            ShowWebViewDialog("log", 700, 500);
            break;
        case ID_TRAY_CONFIGURE:
            LogMessage("Opening Configuration dialog");
            ShowWebViewDialog("config", 560, 520);
            break;
        case ID_TRAY_EXIT:
            LogMessage("User selected Exit");
            StopKeyboardHook();
            CancelScreenCapture("application is exiting");
            /* Close WebView if open */
            if (g_webviewHwnd) SendMessage(g_webviewHwnd, WM_CLOSE, 0, 0);
            KillTimer(hWnd, ID_TIMER_HTTP_RECONCILE);
            KillTimer(hWnd, ID_TIMER_CLIPBOARD_RETRY);
            KillTimer(hWnd, ID_TIMER_DEFERRED_PASTE);
            KillTimer(hWnd, ID_TIMER_AUTO_UPDATE);
            if (g_updateCancelEvent) SetEvent(g_updateCancelEvent);
            DiscardUpdateTask((UpdateCheckTask*)InterlockedExchangePointer(
                (PVOID volatile*)&g_updatePostedResult, NULL));
            DiscardPendingUpdateNotice();
            DiscardPreparedUpdate();
            StopHistoryThumbnailWorker();
            if (fnRemoveClipboardFormatListener) fnRemoveClipboardFormatListener(hWnd);
            StopHttpServer();
            Shell_NotifyIconW(NIM_DELETE, &g_nid);
            if (g_hAppIcon) DestroyIcon(g_hAppIcon);
            if (g_hMenu) DestroyMenu(g_hMenu);
            DestroyImageCache();
            if (g_httpSocketLockReady && !g_httpThread) {
                DeleteCriticalSection(&g_httpSocketLock);
                g_httpSocketLockReady = FALSE;
            }
            if (g_winsockReady && !g_httpThread) {
                WSACleanup();
                g_winsockReady = FALSE;
            }
            GdiplusShutdown(g_gdipToken);
            CoUninitialize();
            if (g_hMutex) {
                ReleaseMutex(g_hMutex);
                CloseHandle(g_hMutex);
            }
            DestroyWindow(hWnd);
            break;
        }
        return 0;

    case WM_SCREEN_CAPTURE_BEGIN:
        if (lParam != HOOK_CAPTURE_MESSAGE || g_configScreenCaptureEnabled ||
            g_captureOverlayHwnd) BeginScreenCapture();
        if (lParam == HOOK_CAPTURE_MESSAGE) InterlockedExchange(&g_hookPrintPending, FALSE);
        return 0;

    case WM_SCREEN_CAPTURE_COPY:
        CompleteScreenCapture((BOOL)wParam);
        if (lParam == HOOK_CAPTURE_MESSAGE) InterlockedExchange(&g_hookPrintPending, FALSE);
        return 0;

    case WM_SCREEN_CAPTURE_CANCEL:
        CancelScreenCapture(wParam
            ? "display configuration changed" : "requested by user");
        if (lParam == HOOK_CAPTURE_MESSAGE) InterlockedExchange(&g_hookCancelPending, FALSE);
        return 0;

    case WM_KEYBOARD_HOTKEY_STATUS:
        if (wParam) {
            LogMessage("Print Screen backup hotkey registered (independent of keyboard hook)");
        } else {
            LogMessage("Print Screen backup hotkey unavailable (%lu); keyboard hook remains enabled",
                       (DWORD)lParam);
        }
        return 0;

    case WM_APP_UPDATE_PROGRESS:
        InterlockedExchange(&g_updateProgressPosted, FALSE);
        if (InterlockedCompareExchange(&g_updateCheckPending,
                                       FALSE, FALSE) == TRUE &&
            !g_updateCheckAbandoned && g_configViewReady) {
            DWORD percent = (DWORD)InterlockedCompareExchange(
                &g_updateProgressPercent, 0, 0);
            CfgSendUpdateProgress(percent);
        }
        return 0;

    case WM_APP_UPDATE_RESULT:
        {
            UpdateCheckTask* task = (UpdateCheckTask*)InterlockedExchangePointer(
                (PVOID volatile*)&g_updatePostedResult, NULL);
            HandleCompletedUpdateCheck(task);
        }
        return 0;

    case WM_HISTORY_THUMBS_READY:
        InterlockedExchange(&g_thumbReadyPosted, FALSE);
        PushReadyHistoryThumbnails();
        return 0;

    case WM_KEYBOARD_HOOK_STATUS:
        if (wParam == 2) {
            LogMessage("Could not raise input thread priority (%lu); continuing at normal priority",
                       (DWORD)lParam);
        } else if (wParam) {
            LogMessage("Keyboard hook active on dedicated input thread (automatic renewal every %u seconds)",
                       KEYBOARD_HOOK_RENEW_MS / 1000u);
        } else {
            LogMessage("ERROR: Keyboard hook installation/message loop failed (%lu)",
                       (DWORD)lParam);
        }
        return 0;

    case WM_DO_PASTE:
        {
            HWND target = (HWND)lParam;
            DWORD requestedSequence = (DWORD)wParam;
            if (!IsPasteRequestCurrent(target, requestedSequence)) {
                KillTimer(hWnd, ID_TIMER_DEFERRED_PASTE);
                g_pasteDeferred = FALSE;
                InterlockedExchange(&g_hookPastePending, FALSE);
                LogMessage("Paste cancelled: focused window or clipboard changed while queued");
                return 0;
            }
            if (g_configCompatibilityPaste &&
                (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0) {
                if (!g_pasteDeferred) {
                    LogMessage("Waiting for Ctrl release before Shift+Insert paste");
                    g_pasteDeferred = TRUE;
                }
                g_deferredPasteTarget = target;
                g_deferredPasteSequence = requestedSequence;
                if (!SetTimer(hWnd, ID_TIMER_DEFERRED_PASTE,
                              DEFERRED_PASTE_DELAY_MS, NULL)) {
                    g_pasteDeferred = FALSE;
                    InterlockedExchange(&g_hookPastePending, FALSE);
                    LogMessage("Paste cancelled: could not schedule Ctrl-release check");
                }
                return 0;
            }
            KillTimer(hWnd, ID_TIMER_DEFERRED_PASTE);
            g_pasteDeferred = FALSE;
            DWORD sequence = GetClipboardSequenceNumber();
            if (sequence != GetCachedClipboardSequence()) {
                if (!RefreshClipboardImageCache()) {
                    InterlockedExchange(&g_hookPastePending, FALSE);
                    LogMessage("Paste cancelled while waiting for the current clipboard image");
                    return 0;
                }
            }
            if (IsPasteRequestCurrent(target, requestedSequence)) {
                PasteCachedImage();
            } else {
                LogMessage("Paste cancelled: focused window or clipboard changed during preparation");
            }
            InterlockedExchange(&g_hookPastePending, FALSE);
        }
        return 0;

    case WM_CLIPBOARDUPDATE:
        {
            DWORD sequence = GetClipboardSequenceNumber();
            if (g_writingClipboardText) return 0;
            if (sequence == GetCachedClipboardSequence()) return 0;
            if (sequence == g_ownClipboardSequence) {
                SetCachedClipboardSequence(sequence);
                return 0;
            }
            RefreshClipboardImageCache();
        }
        return 0;

    case WM_TIMER:
        if (wParam == ID_TIMER_HTTP_RECONCILE) {
            ReconcileHttpServer();
            return 0;
        }
        if (wParam == ID_TIMER_CLIPBOARD_RETRY) {
            RefreshClipboardImageCache();
            return 0;
        }
        if (wParam == ID_TIMER_DEFERRED_PASTE) {
            KillTimer(hWnd, ID_TIMER_DEFERRED_PASTE);
            if (g_pasteDeferred && !PostMessage(hWnd, WM_DO_PASTE,
                    g_deferredPasteSequence, (LPARAM)g_deferredPasteTarget)) {
                g_pasteDeferred = FALSE;
                InterlockedExchange(&g_hookPastePending, FALSE);
                LogMessage("Paste cancelled: could not queue Ctrl-release check");
            }
            return 0;
        }
        if (wParam == ID_TIMER_AUTO_UPDATE) {
            if (g_configAutoCheckForUpdates) StartUpdateCheck(TRUE);
            return 0;
        }
        break;

    case WM_HTTP_EVENT:
        if ((int)wParam == HTTP_EVENT_SERVED) {
            LogMessage("HTTP image request served (200 OK)");
        } else if ((int)wParam == HTTP_EVENT_GONE) {
            LogMessage("HTTP request for evicted image returned 410 Gone");
        } else if ((int)wParam == HTTP_EVENT_NOT_FOUND) {
            LogMessage("HTTP request for unknown image returned 404 Not Found");
        } else if ((int)wParam == 500) {
            LogMessage("HTTP request failed: stored image could not be read (500)");
        } else if ((int)wParam == HTTP_EVENT_DENIED) {
            struct in_addr deniedAddress;
            char addressText[INET_ADDRSTRLEN];
            deniedAddress.s_addr = (ULONG)lParam;
            if (!InetNtopA(AF_INET, &deniedAddress, addressText,
                           sizeof(addressText))) {
                strcpy(addressText, "unknown");
            }
            LogMessage("HTTP connection from %s rejected by the client allowlist",
                       addressText);
        }
        return 0;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

/* ── Entry point ────────────────────────────────────────────────────────── */

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                   LPSTR lpCmdLine, int nCmdShow)
{
    BOOL updateHelperHandled = FALSE;
    BOOL updateCompleted = FALSE;
    BOOL reopenSettings = FALSE;
    int updateHelperResult = HandleUpdateCommandLine(
        &updateHelperHandled, &updateCompleted, &reopenSettings);
    if (updateHelperHandled) return updateHelperResult;

    MSG msg;
    GdiplusStartupInput gdipInput;

    (void)hPrevInstance;
    (void)lpCmdLine;
    (void)nCmdShow;

    g_hInstance = hInstance;

    /* Single-instance check */
    g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
    for (int attempt = 0;
         g_hMutex && GetLastError() == ERROR_ALREADY_EXISTS && attempt < 10;
         attempt++) {
        CloseHandle(g_hMutex);
        Sleep(250);
        g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
    }
    if (g_hMutex && GetLastError() == ERROR_ALREADY_EXISTS) {
        CloseHandle(g_hMutex);
        g_hMutex = NULL;
        MessageBoxW(NULL, L"ImagePaster is already running.", APP_NAME,
                    MB_OK | MB_ICONINFORMATION);
        return 0;
    }

    /* Initialize COM (needed for IStream and WebView2) */
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    /* Initialize GDI+ */
    ZeroMemory(&gdipInput, sizeof(gdipInput));
    gdipInput.GdiplusVersion = 1;
    if (GdiplusStartup(&g_gdipToken, &gdipInput, NULL) != 0) {
        MessageBoxW(NULL, L"Failed to initialize GDI+.", APP_NAME, MB_OK | MB_ICONERROR);
        return 1;
    }

    /* Load configuration */
    LoadConfigFromRegistry();
    ApplyHttpAllowList();
    ParseKeywords();

    /* Load application icon */
    g_hAppIcon = (HICON)LoadImageW(hInstance, MAKEINTRESOURCEW(IDI_APPICON),
                                    IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);

    /* Register hidden message window class */
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"ImagePasterMsgClass";
    if (!RegisterClassExW(&wc)) {
        MessageBoxW(NULL, L"Failed to register window class.", APP_NAME, MB_OK | MB_ICONERROR);
        return 1;
    }

    /* Create hidden message window */
    g_hWndMain = CreateWindowExW(0, L"ImagePasterMsgClass", L"ImagePaster", 0,
                                  0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL);
    if (!g_hWndMain) {
        MessageBoxW(NULL, L"Failed to create message window.", APP_NAME, MB_OK | MB_ICONERROR);
        GdiplusShutdown(g_gdipToken);
        return 1;
    }

    /* System tray */
    InitTrayIcon(g_hWndMain);
    CreateContextMenu();
    UpdateTooltip();

    LogMessage("ImagePaster %s started", APP_VERSION_A);
    LogMessage("GDI+ initialized");
    LogMessage("Title match keywords: %s", g_configTitleMatch);
    LogMessage("Paste method: %s", g_configPasteMethod == PASTE_METHOD_HTTP ? "HTTP" : "base64");
    LogMessage("Text paste shortcut: %s",
               g_configCompatibilityPaste ? "Shift+Insert" : "Ctrl+V");
    LogMessage("Interactive Print Screen capture: %s",
               g_configScreenCaptureEnabled ? "enabled" : "disabled");
    if (g_configImageHistoryLimit == 0) {
        LogMessage("Image history limit: unlimited");
    } else {
        LogMessage("Image history limit: %d image(s)",
                   g_configImageHistoryLimit);
    }
    LogMessage("Image storage: %s",
               g_configImageStorage == IMAGE_STORAGE_DISK ? "disk" : "memory");
    if (g_configImageStorage == IMAGE_STORAGE_DISK) {
        LoadPersistedDiskHistory();
    }

    /* Clipboard monitoring keeps both paste modes synchronized with the user's
       latest clipboard image, independent of which mode is currently selected. */
    {
        HMODULE user32Module = GetModuleHandleW(L"user32.dll");
        if (user32Module) {
            fnAddClipboardFormatListener = (PFN_AddClipboardFormatListener)
                GetProcAddress(user32Module, "AddClipboardFormatListener");
            fnRemoveClipboardFormatListener = (PFN_RemoveClipboardFormatListener)
                GetProcAddress(user32Module, "RemoveClipboardFormatListener");
        }
    }
    if (!fnAddClipboardFormatListener || !fnAddClipboardFormatListener(g_hWndMain)) {
        LogMessage("ERROR: Failed to register clipboard listener (%lu)", GetLastError());
    }
    RefreshClipboardImageCache();

    /* The image server remains active in both paste modes. */
    {
        WSADATA winsockData;
        if (WSAStartup(MAKEWORD(2, 2), &winsockData) == 0) {
            g_winsockReady = TRUE;
            InitializeCriticalSection(&g_httpSocketLock);
            g_httpSocketLockReady = TRUE;
        } else {
            SetHttpStatus(3, "Winsock initialization failed");
        }
    }
    ReconcileHttpServer();
    SetTimer(g_hWndMain, ID_TIMER_HTTP_RECONCILE,
             HTTP_RECONCILE_INTERVAL_MS, NULL);

    /* Keep low-level input delivery independent of encoding, disk I/O and UI. */
    StartKeyboardHook();

    if (updateCompleted) {
        LogMessage("Application update completed: %s", APP_VERSION_A);
        if (reopenSettings) {
            g_updateConfirmationPending = TRUE;
            ShowWebViewDialog("config", 560, 520);
        }
    }

    SetTimer(g_hWndMain, ID_TIMER_AUTO_UPDATE,
             AUTO_UPDATE_INTERVAL_MS, NULL);
    if (g_configAutoCheckForUpdates) StartUpdateCheck(TRUE);

    /* Message loop */
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    StopKeyboardHook();
    return (int)msg.wParam;
}
