/*
 * ImagePaster - main.c
 *
 * System tray utility that intercepts Ctrl+V when a matching window is focused
 * and the clipboard contains an image. It can paste either a raw base64-encoded
 * PNG string or a short URL served by the built-in HTTP image server.
 *
 * Features:
 *   - Shared in-memory clipboard image cache (PNG/base64 and configurable JPEG)
 *   - Configurable base64 or HTTP URL paste mode
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
#include <objbase.h>
#include <shellapi.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>
#include "resource.h"

/* ── GDI+ flat API declarations ─────────────────────────────────────────── */

typedef int GpStatus;
typedef void GpBitmap;
typedef void GpImage;

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
GpStatus __stdcall GdipGetImageEncodersSize(UINT *numEncoders, UINT *size);
GpStatus __stdcall GdipGetImageEncoders(UINT numEncoders, UINT size, ImageCodecInfo *encoders);
GpStatus __stdcall GdipSaveImageToStream(GpImage *image, IStream *stream, const CLSID *clsidEncoder, const void *encoderParams);
GpStatus __stdcall GdipDisposeImage(GpImage *image);
GpStatus __stdcall GdipGetImageWidth(GpImage *image, UINT *width);
GpStatus __stdcall GdipGetImageHeight(GpImage *image, UINT *height);

/* ── Constants ──────────────────────────────────────────────────────────── */

#define APP_NAME          L"ImagePaster"
#define APP_VERSION_A     "1.0.0"
#define APP_VERSION_W     L"1.0.0"
#define MUTEX_NAME        L"ImagePaster_SingleInstance"
#define WM_TRAYICON       (WM_USER + 1)
#define WM_DO_PASTE       (WM_APP + 1)
#define WM_HTTP_EVENT      (WM_APP + 2)
#define ID_TRAY_LOG       1001
#define ID_TRAY_CONFIGURE 1002
#define ID_TRAY_EXIT      1003
#define ID_TIMER_WEBVIEW_SHOW_FALLBACK 1006
#define ID_TIMER_HTTP_RECONCILE         1007
#define ID_TIMER_CLIPBOARD_RETRY        1008
#define WEBVIEW_SHOW_FALLBACK_DELAY_MS 350
#define HTTP_RECONCILE_INTERVAL_MS      3000
#define CLIPBOARD_RETRY_DELAY_MS        150

#define REG_KEY_PATH       "SOFTWARE\\JPIT\\ImagePaster"
#define REG_VALUE_TITLE    "TitleMatch"
#define REG_VALUE_METHOD   "PasteMethod"
#define REG_VALUE_BIND_IP  "BindIp"
#define REG_VALUE_PORT     "HttpPort"
#define REG_VALUE_QUALITY  "JpegQuality"

#define LOG_RING_CAPACITY  500
#define MAX_KEYWORDS       64
#define MAX_DETECTED_IPS    64
#define IMAGE_TOKEN_BYTES   32
#define IMAGE_TOKEN_HEX_LEN (IMAGE_TOKEN_BYTES * 2)
#define DEFAULT_HTTP_PORT   10444
#define DEFAULT_JPEG_QUALITY 80
#define PASTE_METHOD_BASE64 0
#define PASTE_METHOD_HTTP   1
#define HTTP_EVENT_SERVED   200
#define HTTP_EVENT_GONE     410
#define HTTP_EVENT_NOT_FOUND 404

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
static HHOOK     g_hHook;
static ULONG_PTR g_gdipToken;
static HANDLE    g_hMutex;
static HICON     g_hAppIcon;
static NOTIFYICONDATAW g_nid;
static HMENU     g_hMenu;

static volatile BOOL g_bSkipNextPaste = FALSE;
static BOOL g_writingClipboardText = FALSE;
static DWORD g_ownClipboardSequence = 0;
static DWORD g_lastClipboardSequence = 0;
typedef BOOL (WINAPI *PFN_AddClipboardFormatListener)(HWND);
typedef BOOL (WINAPI *PFN_RemoveClipboardFormatListener)(HWND);
static PFN_AddClipboardFormatListener fnAddClipboardFormatListener = NULL;
static PFN_RemoveClipboardFormatListener fnRemoveClipboardFormatListener = NULL;

/* Persisted configuration */
static char g_configTitleMatch[2048] = "xshell";
static int  g_configPasteMethod = PASTE_METHOD_BASE64;
static char g_configBindIp[INET_ADDRSTRLEN] = "127.0.0.1";
static int  g_configHttpPort = DEFAULT_HTTP_PORT;
static int  g_configJpegQuality = DEFAULT_JPEG_QUALITY;
static WCHAR g_keywords[MAX_KEYWORDS][128];
static int   g_keywordCount = 0;

/* Clipboard image cache shared with the HTTP worker thread. */
typedef struct {
    BYTE *jpegData;
    DWORD jpegSize;
    char *base64Data;
    DWORD base64Len;
    char token[IMAGE_TOKEN_HEX_LEN + 1];
    UINT width;
    UINT height;
} CachedImage;

static CachedImage g_cachedImage = {0};
static SRWLOCK g_imageLock = SRWLOCK_INIT;
static char (*g_goneTokens)[IMAGE_TOKEN_HEX_LEN + 1] = NULL;
static size_t g_goneTokenCount = 0;
static size_t g_goneTokenCapacity = 0;

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
static void ReconcileHttpServer(void);
static void StopHttpServer(void);
static BOOL RefreshClipboardImageCache(void);
static void ClearCachedImage(const char *reason);

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

static void ReplaceCachedImage(BYTE *jpegData, DWORD jpegSize,
                               char *base64Data, DWORD base64Len,
                               const char *token, UINT width, UINT height)
{
    AcquireSRWLockExclusive(&g_imageLock);
    RememberGoneTokenLocked(g_cachedImage.token);
    free(g_cachedImage.jpegData);
    free(g_cachedImage.base64Data);
    ZeroMemory(&g_cachedImage, sizeof(g_cachedImage));
    g_cachedImage.jpegData = jpegData;
    g_cachedImage.jpegSize = jpegSize;
    g_cachedImage.base64Data = base64Data;
    g_cachedImage.base64Len = base64Len;
    strncpy(g_cachedImage.token, token, sizeof(g_cachedImage.token) - 1);
    g_cachedImage.width = width;
    g_cachedImage.height = height;
    ReleaseSRWLockExclusive(&g_imageLock);
}

static void ClearCachedImage(const char *reason)
{
    BOOL hadImage;

    AcquireSRWLockExclusive(&g_imageLock);
    hadImage = g_cachedImage.token[0] != '\0';
    RememberGoneTokenLocked(g_cachedImage.token);
    free(g_cachedImage.jpegData);
    free(g_cachedImage.base64Data);
    ZeroMemory(&g_cachedImage, sizeof(g_cachedImage));
    ReleaseSRWLockExclusive(&g_imageLock);

    if (hadImage) LogMessage("Clipboard image cache cleared: %s", reason);
}

static BOOL HasCachedImage(void)
{
    BOOL available;
    AcquireSRWLockShared(&g_imageLock);
    available = g_cachedImage.token[0] != '\0' &&
                g_cachedImage.jpegData != NULL && g_cachedImage.base64Data != NULL;
    ReleaseSRWLockShared(&g_imageLock);
    return available;
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
    char token[IMAGE_TOKEN_HEX_LEN + 1];
    UINT width = 0, height = 0;
    BOOL success = FALSE;

    if (!IsClipboardFormatAvailable(CF_DIB)) {
        g_lastClipboardSequence = sequence;
        KillTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY);
        ClearCachedImage("clipboard now contains non-image data");
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

    /* Once the replacement image is readable, the previous URL is no longer
       current even while the new encodings are being prepared. */
    ClearCachedImage("a newer clipboard image is being prepared");

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
    if (!base64Data || !GenerateImageToken(token)) {
        LogMessage("ERROR: Failed to prepare clipboard image cache");
        goto cleanup;
    }

    /* Do not publish stale data if the clipboard changed during compression. */
    if (GetClipboardSequenceNumber() != sequence) {
        LogMessage("Clipboard changed during image encoding; retrying latest contents");
        SetTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY, CLIPBOARD_RETRY_DELAY_MS, NULL);
        goto cleanup;
    }

    ReplaceCachedImage(jpegData, jpegSize, base64Data, base64Len,
                       token, width, height);
    jpegData = NULL;
    base64Data = NULL;
    g_lastClipboardSequence = sequence;
    KillTimer(g_hWndMain, ID_TIMER_CLIPBOARD_RETRY);
    LogMessage("Clipboard image cached: %ux%u, JPEG=%lu bytes, quality=%d%%, id=%.12s...",
               width, height, jpegSize, g_configJpegQuality, token);
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
        BYTE *copy = (BYTE *)malloc(g_cachedImage.jpegSize);
        if (copy) {
            memcpy(copy, g_cachedImage.jpegData, g_cachedImage.jpegSize);
            *jpegCopy = copy;
            *jpegSize = g_cachedImage.jpegSize;
            status = HTTP_EVENT_SERVED;
        } else {
            status = 500;
        }
    } else {
        for (size_t i = 0; i < g_goneTokenCount; i++) {
            if (strcmp(g_goneTokens[i], token) == 0) {
                status = HTTP_EVENT_GONE;
                break;
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
        messageHeader = "Current clipboard image";
        body = NULL;
        bodySize = imageSize;
        break;
    case HTTP_EVENT_GONE:
        reason = "Gone";
        contentType = "text/plain; charset=utf-8";
        messageHeader = "This image is no longer available because the clipboard image changed";
        body = "This image is no longer available. A newer clipboard image replaced it, or the clipboard no longer contains an image.\n";
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
        status == HTTP_EVENT_NOT_FOUND) {
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

        SOCKET client = accept(listenSocket, NULL, NULL);
        if (client == INVALID_SOCKET) {
            if (InterlockedCompareExchange(&g_httpStopRequested, 0, 0) != 0) break;
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

static void SimulateCtrlV(void)
{
    INPUT inputs[4];
    ZeroMemory(inputs, sizeof(inputs));

    g_bSkipNextPaste = TRUE;

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

    SendInput(4, inputs, sizeof(INPUT));
    LogMessage("Simulated Ctrl+V (re-injection)");
}

static BOOL PlaceUtf8TextOnClipboard(const char *text)
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

    /* The cache represents the user's last copied image. Ignore our own text update. */
    g_ownClipboardSequence = GetClipboardSequenceNumber();
    g_lastClipboardSequence = g_ownClipboardSequence;
    g_writingClipboardText = FALSE;
    return TRUE;
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
        const char *format =
            "[ image available at http://%s:%d/%s.jpg - if you feel this image "
            "will be useful later on be sure to save it to /tmp or a temp location "
            "for later use ]";
        char formatted[512];
        int formattedLen = snprintf(formatted, sizeof(formatted), format,
                                    g_configBindIp, g_configHttpPort, token);
        if (formattedLen > 0 && formattedLen < (int)sizeof(formatted)) {
            pasteText = (char *)malloc((size_t)formattedLen + 1);
            if (pasteText) {
                memcpy(pasteText, formatted, (size_t)formattedLen + 1);
                textLen = (DWORD)formattedLen;
            }
        }
    }

    if (!pasteText) {
        LogMessage("ERROR: Could not allocate paste text");
        return FALSE;
    }

    if (PlaceUtf8TextOnClipboard(pasteText)) {
        LogMessage("Prepared %s paste (%lu characters, image id %.12s...)",
                   g_configPasteMethod == PASTE_METHOD_HTTP ? "HTTP URL" : "base64",
                   textLen, token);
        SimulateCtrlV();
        success = TRUE;
    }
    free(pasteText);
    return success;
}

/* ── Keyword parsing ───────────────────────────────────────────────────── */

static void ParseKeywords(void)
{
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

    RegCloseKey(hKey);
    return TRUE;
}

static void SaveConfigToRegistry(void)
{
    HKEY hKey;
    DWORD disposition;
    LONG result = RegCreateKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                                  REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL,
                                  &hKey, &disposition);
    if (result != ERROR_SUCCESS) return;

    RegSetValueExA(hKey, REG_VALUE_TITLE, 0, REG_SZ,
                   (const BYTE*)g_configTitleMatch,
                   (DWORD)(strlen(g_configTitleMatch) + 1));
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
    }

    RegCloseKey(hKey);
    LogMessage("Configuration saved: method=%s, bind=%s:%d, JPEG=%d%%, titles=%s",
               g_configPasteMethod == PASTE_METHOD_HTTP ? "HTTP" : "base64",
               g_configBindIp, g_configHttpPort, g_configJpegQuality,
               g_configTitleMatch);
}

/* ── Low-level keyboard hook ────────────────────────────────────────────── */

static LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
    if (nCode == HC_ACTION && wParam == WM_KEYDOWN) {
        KBDLLHOOKSTRUCT *pKb = (KBDLLHOOKSTRUCT *)lParam;

        if (pKb->vkCode == 'V') {
            BOOL ctrlDown = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
            BOOL altDown  = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;

            if (ctrlDown && !altDown) {
                /* Recursion guard: skip if this is our re-injected paste */
                if (g_bSkipNextPaste) {
                    g_bSkipNextPaste = FALSE;
                    LogMessage("Re-injected Ctrl+V detected, passing through");
                    return CallNextHookEx(g_hHook, nCode, wParam, lParam);
                }

                LogMessage("--- Ctrl+V detected ---");

                /* Check if a matching window is focused */
                BOOL matchFound = FALSE;
                {
                    HWND hFg = GetForegroundWindow();
                    if (hFg) {
                        WCHAR title[512];
                        if (GetWindowTextW(hFg, title, 512) > 0) {
                            /* Lowercase the title */
                            for (WCHAR *p = title; *p; p++) {
                                if (*p >= L'A' && *p <= L'Z')
                                    *p = *p - L'A' + L'a';
                            }
                            /* Check each keyword */
                            for (int i = 0; i < g_keywordCount; i++) {
                                if (wcsstr(title, g_keywords[i]) != NULL) {
                                    matchFound = TRUE;
                                    break;
                                }
                            }
                        }
                    }
                }
                LogMessage("Title match: %s", matchFound ? "YES" : "NO");

                /* A user image may be awaiting WM_CLIPBOARDUPDATE processing. Our
                   own text paste keeps the cached image authoritative. */
                DWORD sequence = GetClipboardSequenceNumber();
                BOOL clipboardHasImage = IsClipboardFormatAvailable(CF_DIB);
                BOOL cacheIsCurrent = sequence == g_lastClipboardSequence && HasCachedImage();
                BOOL imagePasteAvailable = clipboardHasImage || cacheIsCurrent;
                LogMessage("Current clipboard image available: %s",
                           imagePasteAvailable ? "YES" : "NO");

                if (matchFound && imagePasteAvailable) {
                    LogMessage("Intercepting image paste using %s mode",
                               g_configPasteMethod == PASTE_METHOD_HTTP ? "HTTP" : "base64");
                    PostMessage(g_hWndMain, WM_DO_PASTE, 0, 0);

                    /* Block the original Ctrl+V. Encoding and clipboard writes are
                       deliberately deferred out of this low-level hook. */
                    return 1;
                }
            }
        }
    }

    return CallNextHookEx(g_hHook, nCode, wParam, lParam);
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
    if (g_configTitleMatch[0] == '\0' || g_keywordCount == 0) {
        wcscpy(g_nid.szTip, L"Image pasting is inactive");
    } else {
        WCHAR wMatch[128];
        MultiByteToWideChar(CP_UTF8, 0, g_configTitleMatch, -1, wMatch, 128);
        /* szTip is 128 wchars max; prefix is ~27 chars, leave room */
        WCHAR tip[128];
        swprintf(tip, 128, L"Image pasting active for %s", wMatch);
        tip[127] = L'\0';
        wcscpy(g_nid.szTip, tip);
    }
    g_nid.uFlags = NIF_TIP;
    Shell_NotifyIconW(NIM_MODIFY, &g_nid);
}

static void CreateContextMenu(void)
{
    g_hMenu = CreatePopupMenu();
    AppendMenuW(g_hMenu, MF_STRING, ID_TRAY_LOG, L"Activity Log");
    AppendMenuW(g_hMenu, MF_STRING, ID_TRAY_CONFIGURE, L"Configuration");
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

/* ── JSON helpers ──────────────────────────────────────────────────────── */

static BOOL json_get_string(const char *json, const char *key, char *out, size_t outLen)
{
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return FALSE;
    p += strlen(search);
    while (*p == ' ' || *p == ':') p++;
    if (*p != '"') return FALSE;
    p++;
    size_t i = 0;
    while (*p && *p != '"' && i < outLen - 1) {
        out[i++] = *p++;
    }
    out[i] = '\0';
    return TRUE;
}

static BOOL json_get_int(const char *json, const char *key, int *out)
{
    char search[128];
    snprintf(search, sizeof(search), "\"%s\"", key);
    const char *p = strstr(json, search);
    if (!p) return FALSE;
    p += strlen(search);
    while (*p == ' ' || *p == ':') p++;
    *out = atoi(p);
    return TRUE;
}

static void json_escape_string(const char *in, wchar_t *out, size_t outLen)
{
    size_t j = 0;
    for (size_t i = 0; in[i] && j < outLen - 2; i++) {
        char c = in[i];
        if (c == '"' || c == '\\') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = (wchar_t)c;
        } else if (c == '\n') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'n';
        } else if (c == '\r') {
            if (j + 2 >= outLen) break;
            out[j++] = L'\\';
            out[j++] = L'r';
        } else {
            out[j++] = (wchar_t)(unsigned char)c;
        }
    }
    out[j] = L'\0';
}

/* ── Push functions (C -> JS) ──────────────────────────────────────────── */

static void webview_push_init_config(void)
{
    wchar_t wTitleMatch[4096];
    wchar_t wBindIp[128];
    wchar_t wHttpStatus[512];
    wchar_t ipsJson[4096];
    char ips[MAX_DETECTED_IPS][INET_ADDRSTRLEN];
    int ipCount = EnumerateDetectedIpv4Addresses(ips, MAX_DETECTED_IPS);
    size_t pos = 0;
    json_escape_string(g_configTitleMatch, wTitleMatch, 4096);
    json_escape_string(g_configBindIp, wBindIp, 128);
    json_escape_string(g_httpStatus, wHttpStatus, 512);

    pos += swprintf(ipsJson + pos, 4096 - pos, L"[");
    for (int i = 0; i < ipCount && pos < 4000; i++) {
        wchar_t wIp[64];
        MultiByteToWideChar(CP_UTF8, 0, ips[i], -1, wIp, 64);
        pos += swprintf(ipsJson + pos, 4096 - pos,
                        i == 0 ? L"\"%s\"" : L",\"%s\"", wIp);
    }
    swprintf(ipsJson + pos, 4096 - pos, L"]");

    wchar_t script[16384];
    swprintf(script, 16384,
        L"window.onInit({\"view\":\"config\",\"config\":{"
        L"\"titleMatch\":\"%s\","
        L"\"pasteMethod\":\"%s\","
        L"\"bindIp\":\"%s\","
        L"\"httpPort\":%d,"
        L"\"jpegQuality\":%d,"
        L"\"availableIps\":%s,"
        L"\"bindIpAvailable\":%s,"
        L"\"serverStatus\":\"%s\","
        L"\"version\":\"%s\"}})",
        wTitleMatch,
        g_configPasteMethod == PASTE_METHOD_HTTP ? L"http" : L"base64",
        wBindIp, g_configHttpPort, g_configJpegQuality, ipsJson,
        IsConfiguredBindAddressPresent() ? L"true" : L"false",
        wHttpStatus, APP_VERSION_W);
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
        }
    } else if (strcmp(action, "saveSettings") == 0) {
        char titleMatch[2048] = {0};
        char pasteMethod[32] = {0};
        char bindIp[INET_ADDRSTRLEN] = {0};
        int httpPort = 0;
        int jpegQuality = -1;
        IN_ADDR parsedAddress;

        json_get_string(msg, "titleMatch", titleMatch, sizeof(titleMatch));
        json_get_string(msg, "pasteMethod", pasteMethod, sizeof(pasteMethod));
        json_get_string(msg, "bindIp", bindIp, sizeof(bindIp));
        json_get_int(msg, "httpPort", &httpPort);
        json_get_int(msg, "jpegQuality", &jpegQuality);

        if (strcmp(pasteMethod, "base64") != 0 && strcmp(pasteMethod, "http") != 0) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'Select a valid paste method.'})");
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
        if (jpegQuality < 0 || jpegQuality > 100) {
            webview_execute_script(L"window.onSaveResult && window.onSaveResult({ok:false,message:'JPEG quality must be between 0 and 100.'})");
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
        strncpy(g_configBindIp, bindIp, sizeof(g_configBindIp) - 1);
        g_configBindIp[sizeof(g_configBindIp) - 1] = '\0';
        g_configHttpPort = httpPort;
        g_configJpegQuality = jpegQuality;
        SaveConfigToRegistry();
        ParseKeywords();
        UpdateTooltip();
        if (qualityChanged && IsClipboardFormatAvailable(CF_DIB)) {
            g_lastClipboardSequence = 0;
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
    } else if (strcmp(action, "resize") == 0) {
        int contentHeight = 0;
        json_get_int(msg, "height", &contentHeight);
        if (contentHeight > 0 && g_webviewHwnd) {
            RECT clientRect = {0}, windowRect = {0};
            GetClientRect(g_webviewHwnd, &clientRect);
            GetWindowRect(g_webviewHwnd, &windowRect);
            int chromeH = (windowRect.bottom - windowRect.top) - (clientRect.bottom - clientRect.top);
            int newWindowH = contentHeight + chromeH;
            int windowW = windowRect.right - windowRect.left;
            UINT flags = SWP_NOMOVE | SWP_NOZORDER;
            if (g_webviewWindowShown) {
                flags |= SWP_NOACTIVATE;
            } else {
                flags |= SWP_SHOWWINDOW;
                KillTimer(g_webviewHwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
            }
            SetWindowPos(g_webviewHwnd, NULL, 0, 0, windowW, newWindowH, flags);
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
        case WM_SIZE:
            webview_sync_controller_bounds();
            return 0;

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
            g_webviewHwnd = NULL;
            g_webviewWindowShown = FALSE;
            KillTimer(hwnd, ID_TIMER_WEBVIEW_SHOW_FALLBACK);
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowWebViewDialog(const char* view, int width, int height) {
    if (g_webviewHwnd != NULL) {
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

    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int posX = (screenW - width) / 2;
    int posY = (screenH - height) / 2;

    g_webviewHwnd = CreateWindowExW(0, L"ImagePasterWebViewWnd", title,
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        posX, posY, width, height,
        NULL, NULL, g_hInstance, NULL);

    if (!g_webviewHwnd) {
        LogMessage("ERROR: Failed to create WebView2 window.");
        return;
    }
    g_webviewWindowShown = FALSE;
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
        if (lParam == WM_RBUTTONUP) {
            POINT pt;
            GetCursorPos(&pt);
            SetForegroundWindow(hWnd);
            EnableMenuItem(g_hMenu, ID_TRAY_LOG, g_webviewHwnd ? MF_GRAYED : MF_ENABLED);
            EnableMenuItem(g_hMenu, ID_TRAY_CONFIGURE, g_webviewHwnd ? MF_GRAYED : MF_ENABLED);
            TrackPopupMenu(g_hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);
        }
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
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
            /* Close WebView if open */
            if (g_webviewHwnd) SendMessage(g_webviewHwnd, WM_CLOSE, 0, 0);
            KillTimer(hWnd, ID_TIMER_HTTP_RECONCILE);
            KillTimer(hWnd, ID_TIMER_CLIPBOARD_RETRY);
            if (fnRemoveClipboardFormatListener) fnRemoveClipboardFormatListener(hWnd);
            StopHttpServer();
            Shell_NotifyIconW(NIM_DELETE, &g_nid);
            if (g_hAppIcon) DestroyIcon(g_hAppIcon);
            if (g_hMenu) DestroyMenu(g_hMenu);
            if (g_hHook) UnhookWindowsHookEx(g_hHook);
            ClearCachedImage("application is exiting");
            AcquireSRWLockExclusive(&g_imageLock);
            free(g_goneTokens);
            g_goneTokens = NULL;
            g_goneTokenCount = 0;
            g_goneTokenCapacity = 0;
            ReleaseSRWLockExclusive(&g_imageLock);
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

    case WM_DO_PASTE:
        {
            DWORD sequence = GetClipboardSequenceNumber();
            if (sequence != g_lastClipboardSequence) {
                if (!RefreshClipboardImageCache()) {
                    LogMessage("Paste cancelled while waiting for the current clipboard image");
                    return 0;
                }
            }
            PasteCachedImage();
        }
        return 0;

    case WM_CLIPBOARDUPDATE:
        {
            DWORD sequence = GetClipboardSequenceNumber();
            if (g_writingClipboardText) return 0;
            if (sequence == g_lastClipboardSequence) return 0;
            if (sequence == g_ownClipboardSequence) {
                g_lastClipboardSequence = sequence;
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
        break;

    case WM_HTTP_EVENT:
        if ((int)wParam == HTTP_EVENT_SERVED) {
            LogMessage("HTTP image request served (200 OK)");
        } else if ((int)wParam == HTTP_EVENT_GONE) {
            LogMessage("HTTP request for superseded image returned 410 Gone");
        } else if ((int)wParam == HTTP_EVENT_NOT_FOUND) {
            LogMessage("HTTP request for unknown image returned 404 Not Found");
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
    MSG msg;
    GdiplusStartupInput gdipInput;

    (void)hPrevInstance;
    (void)lpCmdLine;
    (void)nCmdShow;

    g_hInstance = hInstance;

    /* Single-instance check */
    g_hMutex = CreateMutexW(NULL, TRUE, MUTEX_NAME);
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
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

    /* Install keyboard hook */
    g_hHook = SetWindowsHookExW(WH_KEYBOARD_LL, LowLevelKeyboardProc, hInstance, 0);
    if (!g_hHook) {
        LogMessage("ERROR: Failed to install keyboard hook (%lu)", GetLastError());
    } else {
        LogMessage("Keyboard hook installed (WH_KEYBOARD_LL)");
        LogMessage("Monitoring for Ctrl+V with image clipboard...");
    }

    /* Message loop */
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return (int)msg.wParam;
}
