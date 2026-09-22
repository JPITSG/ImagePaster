"""Run the History paging and lazy-thumbnail pipeline from main.c.

The production functions are compiled with inert Windows boundaries and an
MSVCRT-compatible swprintf, so every script pushed to the page is parsed as
JSON here. The worker loop runs synchronously; decoding is replaced by a stub.
"""
import json
from pathlib import Path
import re
import subprocess
import tempfile
import unittest

from test_updater import SOURCE, function


def between(start, end):
    first = SOURCE.index(start)
    return SOURCE[first:SOURCE.index(end, first) + len(end)]


class HistoryTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory(prefix='imagepaster-history-')
        cls.addClassCleanup(cls.directory.cleanup)
        path = Path(cls.directory.name)
        definitions = '\n'.join(re.findall(
            r'^#define (?:HISTORY_\w+|IMAGE_TOKEN_\w+|WM_HISTORY_THUMBS_READY)\s+.*$',
            SOURCE, re.MULTILINE))
        cached_image = between('typedef struct {\n    BYTE *jpegData;',
                               '} CachedImage;')
        thumbnail_state = between('typedef struct {\n    char token[IMAGE_TOKEN_HEX_LEN + 1];\n    char *dataUri;',
                                  'static size_t g_historyViewPage = 0;')
        production = '\n'.join(function(name) for name in (
            'IsValidImageToken', 'json_escape_wstring',
            'FindHistoryThumbnailLocked', 'StoreHistoryThumbnailLocked',
            'MarkHistoryThumbnailReadyLocked', 'HistoryThumbnailThread',
            'EnsureHistoryThumbnailWorker', 'StopHistoryThumbnailWorker',
            'PushReadyHistoryThumbnails', 'RequestHistoryThumbnails',
            'DropHistoryThumbnailRequests', 'ForgetHistoryThumbnail',
            'ClearHistoryThumbnails', 'webview_push_history'))
        window = function('WndProc')
        ready = window[window.index('    case WM_HISTORY_THUMBS_READY:'):
                       window.index('    case WM_KEYBOARD_HOOK_STATUS:')]
        production += ('\nstatic LRESULT uiThumbsReady(void) {\n'
                       'switch (WM_HISTORY_THUMBS_READY) {\n' + ready +
                       '}\nreturn -1;\n}\n')
        source = (PREFIX + definitions + '\n' + cached_image + '\n' +
                  thumbnail_state + '\n' + STUBS + production + SCENARIOS)
        (path / 'history.c').write_text(source)
        cls.executable = path / 'history'
        subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                        str(path / 'history.c'), '-o', str(cls.executable)],
                       check=True)

    def scripts(self, scenario):
        output = subprocess.run([str(self.executable), scenario], check=True,
                                capture_output=True, text=True, timeout=10).stdout
        pushed = []
        for line in output.splitlines():
            label, script = line.split('\t', 1)
            for prefix, suffix, kind in (
                    ('window.onInit(', ')', 'init'),
                    ('window.onHistoryData && window.onHistoryData(', ')', 'data'),
                    ('window.onHistoryThumbs && window.onHistoryThumbs(', ')', 'thumbs')):
                if script.startswith(prefix) and script.endswith(suffix):
                    pushed.append((label, kind,
                                   json.loads(script[len(prefix):-len(suffix)])))
                    break
            else:
                self.fail('Unexpected script: ' + script[:120])
        return pushed

    def test_pages_carry_metadata_only_newest_first(self):
        pushed = self.scripts('pages')
        labels = [label for label, _, _ in pushed]
        self.assertEqual(labels, ['first', 'second', 'clamped', 'empty'])

        _, kind, init = pushed[0]
        self.assertEqual(kind, 'init')
        self.assertEqual(init['view'], 'history')
        first = init['history']
        self.assertEqual((first['total'], first['page'], first['pageSize']), (120, 0, 50))
        self.assertEqual(first['totalBytes'], sum(1000 + i for i in range(120)))
        entries = first['entries']
        self.assertEqual(len(entries), 50)
        self.assertTrue(entries[0]['current'])
        self.assertEqual(entries[0]['token'], 'c' * 64)
        # Then history newest first: image 118 (the newest) down to 70.
        self.assertEqual([e['bytes'] for e in entries[1:]],
                         [1000 + i for i in range(118, 69, -1)])
        self.assertFalse(any(e['current'] for e in entries[1:]))
        disk = [e for e in entries if e['storage'] == 'disk']
        self.assertTrue(disk)
        self.assertTrue(all(e['path'].startswith('C:\\Users\\Zoë "Q"\\') for e in disk))
        self.assertTrue(all(e['url'] == f"http://192.168.1.5:10444/{e['token']}.jpg"
                            for e in entries))
        self.assertTrue(all('thumb' not in e for e in entries))

        second = pushed[1][2]
        self.assertEqual(pushed[1][1], 'data')
        self.assertEqual(second['page'], 1)
        self.assertEqual([e['bytes'] for e in second['entries']],
                         [1000 + i for i in range(69, 19, -1)])

        clamped = pushed[2][2]
        self.assertEqual(clamped['page'], 2)  # page 99 of 3 shows the last page
        self.assertEqual([e['bytes'] for e in clamped['entries']],
                         [1000 + i for i in range(19, -1, -1)])

        empty = pushed[3][2]
        self.assertEqual((empty['total'], empty['page'], empty['entries']), (0, 0, []))

    def test_thumbnails_follow_the_latest_request_and_are_cached(self):
        pushed = self.scripts('thumbs')
        thumbs = {}
        for label, kind, payload in pushed:
            self.assertEqual(kind, 'thumbs')
            self.assertLessEqual(len(payload), 8)  # small scripts, never one blob
            thumbs.setdefault(label, []).append(payload)

        def tokens(label):
            return [item['token'][0] for batch in thumbs[label] for item in batch]

        # A newer request replaced the queue; only its rows were generated.
        self.assertEqual(tokens('replaced'), ['c', 'd'])
        self.assertTrue(all(item['thumb'].startswith('data:image/jpeg;base64,')
                            for batch in thumbs['replaced'] for item in batch))
        self.assertEqual(tokens('cached'), ['c', 'd'])  # sent before any work
        self.assertEqual(tokens('generated'), ['e'])
        failed = [item for batch in thumbs['failed'] for item in batch]
        self.assertEqual([(item['token'][0], item['thumb']) for item in failed],
                         [('f', '')])
        self.assertEqual([len(batch) for batch in thumbs['batched']], [8, 8, 4])
        self.assertEqual(tokens('rerequested'), ['5'])  # sent once
        self.assertEqual(tokens('no-worker'), ['a'])
        self.assertEqual(thumbs['no-worker'][0][0]['thumb'], '')


PREFIX = r'''
#include <assert.h>
#include <locale.h>
#include <stdarg.h>
#include <stdint.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
typedef int BOOL;
typedef int32_t LONG;
typedef uint8_t BYTE;
typedef uint32_t DWORD, UINT;
typedef uint64_t ULONGLONG;
typedef uintptr_t WPARAM;
typedef intptr_t LPARAM, LRESULT;
typedef long HRESULT;
typedef void *HWND, *HANDLE, *LPVOID;
typedef wchar_t WCHAR;
typedef struct { int exclusive, shared; } SRWLOCK;
typedef struct { int unused; } GUID;
typedef struct IWICImagingFactory IWICImagingFactory;
typedef DWORD (*LPTHREAD_START_ROUTINE)(LPVOID);
#define SRWLOCK_INIT {0, 0}
#define TRUE 1
#define FALSE 0
#define WINAPI
#define MAX_PATH 260
#define WM_APP 0x8000
#define INFINITE 0xffffffffu
#define S_OK ((HRESULT)0)
#define E_FAIL ((HRESULT)(int32_t)0x80004005u)
#define SUCCEEDED(hr) ((HRESULT)(hr) >= 0)
#define FAILED(hr) ((HRESULT)(hr) < 0)
#define COINIT_MULTITHREADED 0
#define CLSCTX_INPROC_SERVER 1
#define THREAD_PRIORITY_BELOW_NORMAL (-1)
#define PASTE_METHOD_HTTP 1
#define ZeroMemory(pointer, size) memset((pointer), 0, (size))
#define IWICImagingFactory_Release(factory) ((void)(factory))
'''

STUBS = r'''
static HWND g_hWndMain = (HWND)1;
static void *g_webviewView = (void *)2;
static char g_pendingView[16] = "history";
static int g_configPasteMethod = PASTE_METHOD_HTTP;
static int g_configImageHistoryLimit = 0;
static char g_configBindIp[16] = "192.168.1.5";
static int g_configHttpPort = 10444;
static SRWLOCK g_imageLock = SRWLOCK_INIT;
static CachedImage g_cachedImage;
static CachedImage *g_imageHistory;
static size_t g_imageHistoryCount;
static const GUID g_wicImagingFactoryClsid, g_wicImagingFactoryIid;
static const char *label = "";
static int threadsCreated, signals, readyPosts, generated, failThread;
static int rerequestWhileDecoding;
static void RequestHistoryThumbnails(const char *tokens);
static char generatedOrder[512];

/* MSVCRT's swprintf reads %s as wide, %hs as narrow and knows %I64u. */
static int msvc_swprintf(wchar_t *buffer, size_t count, const wchar_t *format, ...)
{
    wchar_t translated[512];
    size_t out = 0;
    va_list args;
    int written;
    for (const wchar_t *p = format; *p; p++) {
        translated[out++] = *p;
        if (*p != L'%') continue;
        p++;
        if (*p == L'%') { translated[out++] = *p; continue; }
        while (*p && wcschr(L"-+ #0123456789.", *p)) translated[out++] = *p++;
        if (p[0] == L'I' && p[1] == L'6' && p[2] == L'4') {
            translated[out++] = L'l';
            translated[out++] = L'l';
            p += 3;
        } else if (p[0] == L'h' && p[1] == L's') {
            p++;
        } else if (*p == L's') {
            translated[out++] = L'l';
        } else if (*p == L'l') {
            translated[out++] = *p++;
        }
        translated[out++] = *p;
        assert(out < 500);
    }
    translated[out] = L'\0';
    va_start(args, format);
    written = vswprintf(buffer, count, translated, args);
    va_end(args);
    return written;
}
#define swprintf msvc_swprintf

/* SRW locks are not reentrant; nothing may re-acquire one it holds. */
static void AcquireSRWLockExclusive(SRWLOCK *lock) {
    assert(!lock->exclusive && !lock->shared);
    lock->exclusive = 1;
}
static void ReleaseSRWLockExclusive(SRWLOCK *lock) {
    assert(lock->exclusive);
    lock->exclusive = 0;
}
static void AcquireSRWLockShared(SRWLOCK *lock) {
    assert(!lock->exclusive && !lock->shared);
    lock->shared = 1;
}
static void ReleaseSRWLockShared(SRWLOCK *lock) {
    assert(lock->shared);
    lock->shared = 0;
}
static LONG InterlockedExchange(volatile LONG *target, LONG value) {
    LONG previous = *target;
    *target = value;
    return previous;
}
static LONG InterlockedCompareExchange(volatile LONG *target, LONG value,
                                       LONG comparand) {
    LONG previous = *target;
    if (previous == comparand) *target = value;
    return previous;
}
static HRESULT CoInitializeEx(void *reserved, DWORD mode) {
    (void)reserved; (void)mode;
    return S_OK;
}
static void CoUninitialize(void) {}
static HRESULT CoCreateInstance(const GUID *clsid, void *outer, DWORD context,
                                const GUID *iid, void **object) {
    (void)clsid; (void)outer; (void)context; (void)iid;
    *object = NULL;
    return E_FAIL; /* the stub below stands in for WIC and GDI+ */
}
static HANDLE CreateEventW(void *attributes, BOOL manualReset, BOOL initial,
                           const wchar_t *name) {
    (void)attributes; (void)initial; (void)name;
    assert(!manualReset);
    return (HANDLE)0x51;
}
static BOOL SetEvent(HANDLE event) {
    assert(event == (HANDLE)0x51);
    signals++;
    return TRUE;
}
static DWORD HistoryThumbnailThread(LPVOID parameter);
static HANDLE CreateThread(void *attributes, size_t stack,
                           LPTHREAD_START_ROUTINE start, LPVOID parameter,
                           DWORD flags, DWORD *id) {
    (void)attributes; (void)stack; (void)parameter; (void)flags; (void)id;
    assert(start == HistoryThumbnailThread);
    if (failThread) return NULL;
    threadsCreated++;
    return (HANDLE)0x52;
}
static BOOL SetThreadPriority(HANDLE thread, int priority) {
    assert(thread == (HANDLE)0x52 && priority == THREAD_PRIORITY_BELOW_NORMAL);
    return TRUE;
}
static BOOL CloseHandle(HANDLE handle) { return handle == (HANDLE)0x52; }
static BOOL PostMessageW(HWND window, UINT message, WPARAM wParam, LPARAM lParam) {
    (void)wParam; (void)lParam;
    assert(window == g_hWndMain && message == WM_HISTORY_THUMBS_READY);
    readyPosts++;
    return TRUE;
}
static void webview_execute_script(const wchar_t *script) {
    // Nothing calls into the page while holding either lock.
    assert(!g_thumbLock.exclusive && !g_imageLock.shared);
    printf("%s\t%ls\n", label, script);
}
/* Stands in for WIC/GDI+ decoding: tokens starting with f cannot decode. */
static char *CreateHistoryThumbnail(IWICImagingFactory *factory, const char *token) {
    char *dataUri;
    (void)factory;
    assert(!g_thumbLock.exclusive && !g_imageLock.shared); /* decode unlocked */
    generatedOrder[generated++] = token[0];
    if (rerequestWhileDecoding) { /* the page asks again meanwhile */
        rerequestWhileDecoding = 0;
        RequestHistoryThumbnails(token);
    }
    if (token[0] == 'f') return NULL;
    dataUri = malloc(64);
    snprintf(dataUri, 64, "data:image/jpeg;base64,%.12s", token);
    return dataUri;
}
static DWORD WaitForSingleObject(HANDLE handle, DWORD milliseconds);
'''

SCENARIOS = r'''
/* The worker idles on its event; ending the synchronous run there. */
static DWORD WaitForSingleObject(HANDLE handle, DWORD milliseconds) {
    (void)milliseconds;
    if (handle == (HANDLE)0x51) g_thumbStop = TRUE;
    else assert(handle == (HANDLE)0x52);
    return 0;
}

static void run_worker(void) {
    HistoryThumbnailThread(NULL);
    g_thumbStop = FALSE;
}

static void deliver(const char *name) {
    label = name;
    if (g_thumbReadyPosted) uiThumbsReady();
    assert(!g_thumbReadyPosted);
}

static void make_token(char out[IMAGE_TOKEN_HEX_LEN + 1], char first, int index) {
    snprintf(out, IMAGE_TOKEN_HEX_LEN + 1, "%c%063x", first, index);
}

/* For immediate use only; the buffer is reused by the next call. */
static const char *token(char first, int index) {
    static char value[IMAGE_TOKEN_HEX_LEN + 1];
    make_token(value, first, index);
    return value;
}

static void request(const char *list) {
    RequestHistoryThumbnails(list);
}

static void pages(void) {
    static CachedImage history[119];
    static WCHAR paths[119][MAX_PATH];
    memset(g_cachedImage.token, 'c', IMAGE_TOKEN_HEX_LEN);
    g_cachedImage.jpegData = (BYTE *)"x";
    g_cachedImage.jpegSize = 1000 + 119;
    g_cachedImage.width = 1920;
    g_cachedImage.height = 1080;
    for (int i = 0; i < 119; i++) {
        snprintf(history[i].token, sizeof(history[i].token), "%064x", i + 1);
        history[i].jpegSize = 1000 + i;
        history[i].width = 800;
        history[i].height = 600;
        history[i].capturedAtMs = 1700000000000ULL + (ULONGLONG)i;
        if (i % 2) {
            swprintf(paths[i], MAX_PATH, L"C:\\Users\\Zo\u00eb \"Q\"\\%hs.jpg",
                     history[i].token);
            history[i].diskPath = paths[i];
        } else {
            history[i].jpegData = (BYTE *)"x";
        }
    }
    g_imageHistory = history;
    g_imageHistoryCount = 119;

    label = "first";
    g_historyViewPage = 0;
    webview_push_history(TRUE);
    label = "second";
    g_historyViewPage = 1;
    webview_push_history(FALSE);
    label = "clamped";
    g_historyViewPage = 99;
    webview_push_history(FALSE);
    assert(g_historyViewPage == 2);

    label = "empty";
    memset(&g_cachedImage, 0, sizeof(g_cachedImage));
    g_imageHistoryCount = 0;
    webview_push_history(FALSE);
    assert(g_historyViewPage == 0);
}

static void thumbs(void) {
    char list[4096];
    char a[IMAGE_TOKEN_HEX_LEN + 1], b[IMAGE_TOKEN_HEX_LEN + 1];
    char c[IMAGE_TOKEN_HEX_LEN + 1], d[IMAGE_TOKEN_HEX_LEN + 1];
    char e[IMAGE_TOKEN_HEX_LEN + 1], f[IMAGE_TOKEN_HEX_LEN + 1];
    make_token(a, 'a', 1);
    make_token(b, 'b', 2);
    make_token(c, 'c', 3);
    make_token(d, 'd', 4);
    make_token(e, 'e', 5);
    make_token(f, 'f', 6);

    snprintf(list, sizeof(list), "%s,xyz,%s,%s,%s", a, b, a, c);
    request(list); /* invalid and duplicate tokens are skipped */
    assert(g_thumbQueueCount == 3 && threadsCreated == 1 && signals == 1);
    snprintf(list, sizeof(list), "%s,%s", c, d);
    request(list); /* the user scrolled: a and b are no longer wanted */
    assert(g_thumbQueueCount == 2);
    run_worker();
    assert(generated == 2 && memcmp(generatedOrder, "cd", 2) == 0);
    assert(readyPosts == 1); /* one wake-up for both previews */
    deliver("replaced");

    snprintf(list, sizeof(list), "%s,%s,%s", c, d, e);
    label = "cached";
    request(list); /* cached rows go out at once, only e is queued */
    assert(g_thumbQueueCount == 1 && generated == 2);
    run_worker();
    deliver("generated");
    assert(generated == 3);

    request(f);
    run_worker();
    deliver("failed");
    request(f); /* failures are not cached: a later view retries */
    assert(g_thumbQueueCount == 1);
    DropHistoryThumbnailRequests(); /* the dialog closed first */
    run_worker();
    assert(generated == 4 && readyPosts == 3);

    list[0] = '\0';
    for (int i = 0; i < 20; i++) {
        size_t used = strlen(list);
        snprintf(list + used, sizeof(list) - used, "%s%s", i ? "," : "",
                 token('9', 100 + i));
    }
    request(list);
    run_worker();
    deliver("batched"); /* 20 previews leave in batches of 8, 8 and 4 */

    rerequestWhileDecoding = 1; /* asked again while being made: made once */
    request(token('5', 1));
    run_worker();
    assert(generated == 25);
    deliver("rerequested");

    ForgetHistoryThumbnail(c); /* deleted images drop their preview */
    request(c);
    assert(g_thumbQueueCount == 1);
    ClearHistoryThumbnails();
    assert(g_thumbCacheCount == 0 && g_thumbQueueCount == 0);

    for (int i = 0; i < HISTORY_THUMB_CACHE_ENTRIES + 10; i++) {
        label = "lru";
        request(token('7', 1000 + i));
        run_worker();
        deliver("lru");
        if (i == HISTORY_THUMB_CACHE_ENTRIES - 20) {
            request(token('7', 1000)); /* touching it keeps it cached */
        }
    }
    assert(g_thumbCacheCount == HISTORY_THUMB_CACHE_ENTRIES);
    AcquireSRWLockExclusive(&g_thumbLock);
    assert(FindHistoryThumbnailLocked(token('7', 1000)) != NULL);
    assert(FindHistoryThumbnailLocked(token('7', 1001)) == NULL); /* LRU */
    assert(g_thumbReadyCount == 0);
    ReleaseSRWLockExclusive(&g_thumbLock);

    strcpy(g_pendingView, "config"); /* previews for a closed view are dropped */
    request(token('7', 1000));
    assert(g_thumbReadyCount == 0);
    strcpy(g_pendingView, "history");

    StopHistoryThumbnailWorker();
    assert(g_thumbStop && !g_thumbThread && g_thumbQueueCount == 0);
    g_thumbStop = FALSE;
    failThread = 1;
    label = "no-worker";
    request(a); /* without a worker the row shows that no preview exists */
    assert(g_thumbQueueCount == 0);
}

int main(int argc, char **argv) {
    setlocale(LC_ALL, "C.UTF-8");
    if (argc == 2 && strcmp(argv[1], "pages") == 0) pages();
    else if (argc == 2 && strcmp(argv[1], "thumbs") == 0) thumbs();
    else return 2;
    return 0;
}
'''


if __name__ == '__main__':
    unittest.main()
