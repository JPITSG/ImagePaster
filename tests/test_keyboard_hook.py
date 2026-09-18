"""Run production hook functions with deterministic Windows API fault injection.

No desktop input is generated. The fake UI queue is deliberately never drained.
These tests cover decisions, renewal, and the owner-thread loop, not Windows'
actual hook delivery (which still requires a Windows desktop smoke test).
"""
from pathlib import Path
import re
import subprocess
import tempfile
import unittest

from test_updater import SOURCE


def function(name):
    match = re.search(r'^static [^\n]*\b' + re.escape(name) +
                      r'\([^;]*?\)\s*\{', SOURCE, re.MULTILINE)
    if not match:
        raise AssertionError('Missing production function: ' + name)
    depth = 1
    end = match.end()
    while depth:
        depth += (SOURCE[end] == '{') - (SOURCE[end] == '}')
        end += 1
    return SOURCE[match.start():end]


class KeyboardHookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory(prefix='imagepaster-hook-test-')
        cls.addClassCleanup(cls.directory.cleanup)
        path = Path(cls.directory.name)
        definitions = '\n'.join(re.findall(
            r'^#define (?:WM_DO_PASTE|WM_SCREEN_CAPTURE_\w+|WM_KEYBOARD_HOOK_\w+|'
            r'WM_KEYBOARD_HOTKEY_STATUS|KEYBOARD_HOOK_RENEW_MS|PASTE_INPUT_TAG|'
            r'HOOK_CAPTURE_\w+|ID_HOTKEY_PRINT_SCREEN|MOD_NOREPEAT|'
            r'ID_TIMER_DEFERRED_PASTE|DEFERRED_PASTE_DELAY_MS)\s+.*$',
            SOURCE, re.MULTILINE))
        # Keep the production constants as well as the production functions.
        production = '\n'.join(function(name) for name in (
            'GetCachedClipboardSequence', 'SetCachedClipboardSequence',
            'TryHasCachedImage', 'HookForegroundTitleMatches',
            'SetHookCaptureFlag', 'QueueHookCapture', 'QueueHookPrintScreen', 'QueueHookPaste',
            'LowLevelKeyboardProc', 'ReconcilePrintScreenHotkey',
            'HandlePrintScreenHotkey', 'RenewKeyboardHook',
            'KeyboardHookThreadProc', 'StartKeyboardHook', 'StopKeyboardHook',
            'SimulateStandardTextPaste',
            'IsPasteRequestCurrent'))
        # Also exercise actual UI case bodies that release the bounded slots.
        window = function('WndProc')
        for name, start, end in (
                ('uiCapture', '    case WM_SCREEN_CAPTURE_BEGIN:',
                 '    case WM_KEYBOARD_HOTKEY_STATUS:'),
                ('uiPaste', '    case WM_DO_PASTE:', '    case WM_CLIPBOARDUPDATE:')):
            cases = window[window.index(start):window.index(end)]
            production += ('\nstatic LRESULT ' + name +
                           '(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {\n'
                           '(void)hWnd; switch (msg) {\n' + cases + '} return -1; }\n')
        timer = window[window.index('        if (wParam == ID_TIMER_DEFERRED_PASTE)'):
                       window.index('        if (wParam == ID_TIMER_AUTO_UPDATE)')]
        production += ('\nstatic LRESULT uiTimer(HWND hWnd, WPARAM wParam) {\n' +
                       timer + 'return -1; }\n')
        source = PREFIX + definitions + '\n' + STUBS + production + TESTS
        (path / 'hook.c').write_text(source)
        cls.executable = path / 'hook'
        subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                        str(path / 'hook.c'), '-o', str(cls.executable)], check=True)

    def check(self, scenario):
        subprocess.run([str(self.executable), scenario], check=True, timeout=5)

    def test_capture_and_paste_with_busy_ui_and_contended_locks(self):
        self.check('routing')

    def test_renewal_recovers_removed_hook_and_retains_hook_on_install_failure(self):
        self.check('renewal')

    def test_owner_loop_recovers_without_ui_pump_and_cleans_up(self):
        self.check('loop')

    def test_injected_paste_tag_does_not_skip_next_physical_paste(self):
        self.check('injection')

    def test_queued_paste_rejects_changed_window_or_clipboard(self):
        self.check('stale')

    def test_unrelated_input_has_no_shared_state_or_routing_work(self):
        self.check('fastpath')

    def test_capture_flood_has_bounded_ui_queue_and_no_routing_queries(self):
        self.check('flood')

    def test_hotkey_fallback_conflicts_config_changes_and_key_release(self):
        self.check('hotkey')

    def test_hotkey_recovers_before_periodic_deadline(self):
        self.check('hotkey_loop')

    def test_start_failure_and_shutdown_release_all_handles(self):
        self.check('lifecycle')

    def test_repeated_pastes_cannot_flood_capture_queue(self):
        self.check('paste_flood')

    def test_ui_finishes_or_cancels_pending_slots_including_timer_failures(self):
        self.check('ui_slots')

    def test_configuration_wakes_input_thread_without_waiting_for_renewal(self):
        self.check('wake_loop')


PREFIX = r'''
#include <assert.h>
#include <stdint.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
typedef int BOOL;
typedef uint32_t DWORD, UINT;
typedef int32_t LONG;
typedef uint64_t ULONGLONG;
typedef uintptr_t WPARAM, ULONG_PTR;
typedef intptr_t LPARAM, LRESULT;
typedef wchar_t WCHAR;
typedef void *HWND, *HHOOK, *HANDLE, *LPVOID, *HINSTANCE;
typedef int SRWLOCK;
typedef struct { UINT message; WPARAM wParam; } MSG;
typedef struct { DWORD vkCode, flags; ULONG_PTR dwExtraInfo; } KBDLLHOOKSTRUCT;
typedef struct { UINT type; struct { DWORD wVk, dwFlags; ULONG_PTR dwExtraInfo; } ki; } INPUT;
typedef LRESULT (*HOOKPROC)(int, WPARAM, LPARAM);
#define CALLBACK
#define WINAPI
#define TRUE 1
#define FALSE 0
#define WM_APP 0x8000
#define WM_USER 0x0400
#define WM_QUIT 0x0012
#define WM_HOTKEY 0x0312
#define WM_KEYDOWN 0x0100
#define WM_KEYUP 0x0101
#define WM_SYSKEYDOWN 0x0104
#define WM_SYSKEYUP 0x0105
#define HC_ACTION 0
#define VK_SNAPSHOT 0x2c
#define VK_ESCAPE 0x1b
#define VK_CONTROL 0x11
#define VK_MENU 0x12
#define CF_DIB 8
#define LLKHF_INJECTED 0x10
#define WH_KEYBOARD_LL 13
#define INPUT_KEYBOARD 1
#define KEYEVENTF_KEYUP 2
#define WAIT_OBJECT_0 0
#define WAIT_TIMEOUT 258
#define WAIT_FAILED UINT32_MAX
#define PM_NOREMOVE 0
#define PM_REMOVE 1
#define QS_ALLINPUT 0x04ff
#define MWMO_INPUTAVAILABLE 4
#define INFINITE UINT32_MAX
#define THREAD_PRIORITY_ABOVE_NORMAL 1
#define ZeroMemory(ptr, size) memset(ptr, 0, size)
'''

STUBS = r'''
static HINSTANCE g_hInstance = (void *)1;
static HWND g_hWndMain = (void *)2;
static HHOOK g_hHook;
static HANDLE g_keyboardHookStopEvent = (void *)3;
static HANDLE g_keyboardHookWakeEvent = (void *)6, g_keyboardHookThread;
static volatile LONG g_hookCaptureState, g_lastClipboardSequence;
static volatile LONG g_hookHasKeywords = TRUE;
static volatile LONG g_hookPrintPending, g_hookCancelPending, g_hookRenewRequested;
static volatile LONG g_hookPastePending;
static HWND g_hookPasteTarget;
static DWORD g_hookPasteSequence;
static BOOL g_printHotkeyRegistered;
static DWORD g_printHotkeyError, g_hookInstallError;
static BOOL g_printScreenKeyDown, g_escapeKeyDown;
static WCHAR g_keywords[64][128] = { L"xshell" };
static int g_keywordCount = 1;
static SRWLOCK g_keywordLock, g_imageLock;
static struct { char token[65]; void *jpegData, *diskPath, *base64Data; } g_cachedImage;
static DWORD clipboardSequence = 12, foregroundPid = 99;
static HWND foregroundWindow = (void *)4;
static BOOL clipboardImage, ctrlDown, altDown, printDown, escapeDown;
static BOOL failPost, failInstall, installed, inHook, stop, flood;
static BOOL failHotkey, failPriority, loopHotkey, pendingHotkey;
static BOOL loopWake;
static BOOL failThread, joined;
static BOOL removedOnFirstWait;
static UINT postCount[0x8010], installCount, unhookCount, waits, logCount;
static UINT atomicOps, routingQueries, lockAttempts, postAttempts;
static UINT registerCount, unregisterCount, priorityCount, wakeCount;
static UINT eventCount, failEventAt, closeCount, liveHandles;
static BOOL g_configScreenCaptureEnabled, g_configCompatibilityPaste, g_pasteDeferred;
static HWND g_captureOverlayHwnd, g_deferredPasteTarget;
static DWORD g_deferredPasteSequence;
static BOOL failTimer, failRefresh, changeOnRefresh;
static UINT pasteCount, beginCount, copyCount, cancelCount;
static ULONGLONG now;
static HHOOK nextHook = (void *)10;
static DWORD GetLastError(void) { return 123; }
static LONG InterlockedCompareExchange(volatile LONG *p, LONG value, LONG expected) {
    atomicOps++;
    LONG old = *p;
    if (old == expected) *p = value;
    return old;
}
static LONG InterlockedExchange(volatile LONG *p, LONG value) {
    atomicOps++;
    LONG old = *p; *p = value; return old;
}
static LONG InterlockedOr(volatile LONG *p, LONG value) {
    atomicOps++; LONG old = *p; *p |= value; return old;
}
static LONG InterlockedAnd(volatile LONG *p, LONG value) {
    atomicOps++; LONG old = *p; *p &= value; return old;
}
static BOOL TryAcquireSRWLockShared(SRWLOCK *lock) { lockAttempts++; return !*lock; }
static void ReleaseSRWLockShared(SRWLOCK *lock) { assert(!*lock); }
static HWND GetForegroundWindow(void) { routingQueries++; return foregroundWindow; }
static DWORD GetWindowThreadProcessId(HWND h, DWORD *pid) {
    routingQueries++;
    (void)h; *pid = foregroundPid; return 10;
}
static DWORD GetCurrentProcessId(void) { return 42; }
static int GetWindowTextW(HWND h, WCHAR *text, int count) {
    routingQueries++;
    (void)h; assert(count >= 12);
    // Calling this for our own process would synchronously wait for the UI.
    assert(foregroundPid != GetCurrentProcessId());
    wcscpy(text, L"XShell test"); return (int)wcslen(text);
}
static short GetAsyncKeyState(int key) {
    routingQueries++;
    BOOL down = key == VK_CONTROL ? ctrlDown : key == VK_MENU ? altDown :
                key == VK_SNAPSHOT ? printDown : key == VK_ESCAPE ? escapeDown : FALSE;
    return down ? (short)0x8000 : 0;
}
static DWORD GetClipboardSequenceNumber(void) { routingQueries++; return clipboardSequence; }
static BOOL IsClipboardFormatAvailable(UINT format) {
    routingQueries++; assert(format == CF_DIB); return clipboardImage;
}
static BOOL PostMessageW(HWND h, UINT message, WPARAM wp, LPARAM lp) {
    postAttempts++;
    assert(h == g_hWndMain); (void)wp; (void)lp;
    if (message == WM_DO_PASTE) {
        assert(wp == clipboardSequence && (HWND)lp == foregroundWindow);
    }
    if (message == WM_SCREEN_CAPTURE_BEGIN || message == WM_SCREEN_CAPTURE_COPY ||
        message == WM_SCREEN_CAPTURE_CANCEL) assert(lp == HOOK_CAPTURE_MESSAGE);
    // Never dispatch UI work inline, including logging.
    if (failPost) return FALSE;
    assert(message < sizeof(postCount) / sizeof(postCount[0]));
    postCount[message]++;
    return TRUE;
}
static LRESULT CallNextHookEx(HHOOK h, int code, WPARAM wp, LPARAM lp) {
    (void)h; (void)code; (void)wp; (void)lp; return 77;
}
static HHOOK SetWindowsHookExW(int kind, HOOKPROC callback, HINSTANCE instance, DWORD thread) {
    assert(kind == WH_KEYBOARD_LL && callback && instance == g_hInstance && thread == 0);
    installCount++;
    if (failInstall) return NULL;
    installed = TRUE;
    nextHook = (void *)((uintptr_t)nextHook + 1);
    return nextHook;
}
static BOOL UnhookWindowsHookEx(HHOOK hook) {
    assert(hook); unhookCount++;
    if (hook == g_hHook) installed = FALSE;
    return TRUE;
}
static BOOL RegisterHotKey(HWND window, int id, UINT modifiers, UINT vk) {
    assert(!inHook && !window && id == ID_HOTKEY_PRINT_SCREEN &&
           modifiers == MOD_NOREPEAT && vk == VK_SNAPSHOT);
    registerCount++;
    return !failHotkey;
}
static BOOL UnregisterHotKey(HWND window, int id) {
    assert(!inHook && !window && id == ID_HOTKEY_PRINT_SCREEN);
    unregisterCount++; return TRUE;
}
static HANDLE GetCurrentThread(void) { return (void *)7; }
static BOOL SetThreadPriority(HANDLE thread, int priority) {
    assert(!inHook && thread == GetCurrentThread() && priority == THREAD_PRIORITY_ABOVE_NORMAL);
    priorityCount++;
    return !failPriority;
}
static BOOL SetEvent(HANDLE event) {
    assert(!inHook);
    if (event == g_keyboardHookStopEvent) stop = TRUE;
    else { assert(event == g_keyboardHookWakeEvent); wakeCount++; }
    return TRUE;
}
static HANDLE CreateEventW(void *attributes, BOOL manual, BOOL initial, const WCHAR *name) {
    assert(!attributes && !initial && !name && manual == !(eventCount % 2));
    eventCount++;
    if (eventCount == failEventAt) return NULL;
    liveHandles++;
    return (void *)(uintptr_t)(100 + eventCount);
}
static HANDLE CreateThread(void *attributes, size_t stack, DWORD (*proc)(LPVOID),
                           LPVOID parameter, DWORD flags, DWORD *id) {
    assert(!attributes && !stack && proc && !parameter && !flags && !id);
    if (failThread) return NULL;
    liveHandles++;
    return (void *)8;
}
static BOOL CloseHandle(HANDLE handle) {
    assert(handle && liveHandles); liveHandles--; closeCount++;
    if (handle == g_keyboardHookThread) assert(joined);
    return TRUE;
}
static ULONGLONG GetTickCount64(void) { return now; }
static DWORD WaitForSingleObject(HANDLE event, DWORD timeout) {
    if (event == g_keyboardHookThread) {
        assert(stop && timeout == INFINITE); joined = TRUE; return WAIT_OBJECT_0;
    }
    assert(event == g_keyboardHookStopEvent && timeout == 0);
    return stop ? WAIT_OBJECT_0 : WAIT_TIMEOUT;
}
static BOOL PeekMessageW(MSG *msg, HWND window, UINT low, UINT high, UINT mode) {
    (void)window; (void)low; (void)high;
    if (mode == PM_NOREMOVE) return FALSE;
    if (pendingHotkey) {
        pendingHotkey = FALSE;
        msg->message = WM_HOTKEY;
        msg->wParam = ID_HOTKEY_PRINT_SCREEN;
        return TRUE;
    }
    if (!flood) return FALSE;
    msg->message = WM_USER;
    now += 1000;
    return TRUE;
}
static DWORD MsgWaitForMultipleObjectsEx(DWORD count, const HANDLE *events, DWORD timeout,
                                         DWORD mask, DWORD flags) {
    assert(count == 2 && events[0] == g_keyboardHookStopEvent &&
           events[1] == g_keyboardHookWakeEvent);
    assert(mask == QS_ALLINPUT && flags == MWMO_INPUTAVAILABLE);
    assert(timeout <= KEYBOARD_HOOK_RENEW_MS);
    waits++;
    if (loopWake) {
        now++;
        assert(installCount == 1 && now < KEYBOARD_HOOK_RENEW_MS);
        if (waits == 1) {
            assert(!g_printHotkeyRegistered);
            g_hookCaptureState = HOOK_CAPTURE_ENABLED; // UI publishes and wakes.
            return WAIT_OBJECT_0 + 1;
        }
        if (waits == 2) {
            assert(g_printHotkeyRegistered);
            g_hookCaptureState = 0;
            return WAIT_OBJECT_0 + 1;
        }
        assert(!g_printHotkeyRegistered);
        stop = TRUE;
        return WAIT_OBJECT_0;
    }
    if (loopHotkey) {
        if (waits == 1) {
            installed = FALSE; // Removal just after install, before deadline.
            pendingHotkey = TRUE;
            printDown = TRUE;
            now++;
            return WAIT_OBJECT_0 + 2;
        }
        assert(installed && installCount == 2 && now < KEYBOARD_HOOK_RENEW_MS);
        stop = TRUE;
        return WAIT_OBJECT_0;
    }
    if (waits == 1) { // Simulate silent OS removal, with the UI still stuck.
        installed = FALSE;
        removedOnFirstWait = TRUE;
        failInstall = TRUE;
    } else if (waits == 2) {
        assert(!installed);
        failInstall = FALSE;
    } else if (waits == 3) {
        assert(installed); // Renewed independently of the undrained UI queue.
    } else {
        stop = TRUE;
        return WAIT_OBJECT_0;
    }
    if (!flood) now += timeout;
    return flood ? WAIT_OBJECT_0 + 2 : WAIT_TIMEOUT;
}
static void LogMessage(const char *format, ...) {
    (void)format; assert(!inHook); logCount++;
}
#define PostMessage PostMessageW
static BOOL KillTimer(HWND hwnd, uintptr_t id) {
    assert(hwnd == g_hWndMain && id == ID_TIMER_DEFERRED_PASTE); return TRUE;
}
static uintptr_t SetTimer(HWND hwnd, uintptr_t id, UINT delay, void *proc) {
    assert(hwnd == g_hWndMain && id == ID_TIMER_DEFERRED_PASTE &&
           delay == DEFERRED_PASTE_DELAY_MS && !proc);
    return failTimer ? 0 : id;
}
static BOOL RefreshClipboardImageCache(void) {
    assert(g_hookPastePending);
    if (changeOnRefresh) clipboardSequence++;
    return !failRefresh;
}
static BOOL PasteCachedImage(void) { assert(g_hookPastePending); pasteCount++; return TRUE; }
static BOOL BeginScreenCapture(void) { assert(g_hookPrintPending); beginCount++; return TRUE; }
static BOOL CompleteScreenCapture(BOOL full) {
    assert(g_hookPrintPending && full); copyCount++; return FALSE;
}
static void CancelScreenCapture(const char *reason) {
    assert(g_hookCancelPending && reason); cancelCount++;
}
static UINT SendInput(UINT count, INPUT *inputs, int size);
'''

TESTS = r'''
static LRESULT key(DWORD vk, BOOL down, DWORD flags, ULONG_PTR tag) {
    KBDLLHOOKSTRUCT event = {vk, flags, tag};
    inHook = TRUE;
    LRESULT result = LowLevelKeyboardProc(HC_ACTION, down ? WM_KEYDOWN : WM_KEYUP,
                                          (LPARAM)&event);
    inHook = FALSE;
    return result;
}
static UINT SendInput(UINT count, INPUT *inputs, int size) {
    assert(count == 4 && size == sizeof(INPUT));
    for (UINT i = 0; i < count; i++) {
        assert(inputs[i].ki.dwExtraInfo == PASTE_INPUT_TAG);
        BOOL down = !(inputs[i].ki.dwFlags & KEYEVENTF_KEYUP);
        if (inputs[i].ki.wVk == VK_CONTROL) ctrlDown = down;
        assert(key(inputs[i].ki.wVk, down, LLKHF_INJECTED,
                   inputs[i].ki.dwExtraInfo) == 77);
    }
    return 4;
}
static void routing(void) {
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 77);
    SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, TRUE);
    g_keywordLock = g_imageLock = 1; // UI holds locks, never pumps its queue.
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1);
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1); // Auto-repeat is ignored.
    assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1);
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == 1);
    g_hookPrintPending = FALSE; // UI finished the first action.
    assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1); // Key-up-only keyboard.
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == 2);
    g_hookPrintPending = FALSE;
    SetHookCaptureFlag(HOOK_CAPTURE_ACTIVE, TRUE);
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1);
    assert(postCount[WM_SCREEN_CAPTURE_COPY] == 1);
    assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1);
    assert(key(VK_ESCAPE, TRUE, 0, 0) == 1);
    SetHookCaptureFlag(HOOK_CAPTURE_ACTIVE, FALSE);
    assert(key(VK_ESCAPE, FALSE, 0, 0) == 1);
    assert(postCount[WM_SCREEN_CAPTURE_CANCEL] == 1);
    ctrlDown = clipboardImage = TRUE;
    assert(key('V', TRUE, 0, 0) == 77); // Busy keyword lock: no waiting.
    g_keywordLock = 0;
    assert(key('V', TRUE, 0, 0) == 1); // New image needs no cache lock.
    clipboardImage = FALSE;
    SetCachedClipboardSequence(clipboardSequence);
    strcpy(g_cachedImage.token, "image");
    g_cachedImage.diskPath = g_cachedImage.base64Data = (void *)1;
    assert(key('V', TRUE, 0, 0) == 77); // Busy image lock: no waiting.
    g_imageLock = 0;
    assert(key('V', TRUE, 0, 0) == 1); // Own text clipboard keeps cached image.
    clipboardSequence++;
    assert(key('V', TRUE, 0, 0) == 77); // Stale cache must not intercept text.
    clipboardImage = TRUE;
    foregroundPid = GetCurrentProcessId();
    assert(key('V', TRUE, 0, 0) == 77); // Never send WM_GETTEXT to busy own UI.
    foregroundPid = 99;
    failPost = TRUE;
    g_hookPrintPending = FALSE;
    g_hookPastePending = FALSE;
    assert(key('V', TRUE, 0, 0) == 77);
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 77); // Don't swallow failed dispatch.
    assert(!g_printScreenKeyDown && logCount == 0);
}
static void renewal(void) {
    assert(RenewKeyboardHook(TRUE));
    HHOOK previous = g_hHook;
    failInstall = TRUE;
    assert(!RenewKeyboardHook(FALSE));
    assert(installed && g_hHook == previous && unhookCount == 0);
    assert(!RenewKeyboardHook(TRUE)); // Identical errors must not flood the UI.
    installed = FALSE; // Windows silently removed the hook; handle stays stale.
    failInstall = FALSE;
    g_printScreenKeyDown = g_escapeKeyDown = TRUE;
    printDown = escapeDown = FALSE; // Swallowed down events never reach async state.
    assert(RenewKeyboardHook(TRUE));
    assert(installed && g_hHook != previous && unhookCount == 1);
    assert(g_printScreenKeyDown && g_escapeKeyDown);
    assert(RenewKeyboardHook(FALSE));
    assert(g_printScreenKeyDown && g_escapeKeyDown);
    assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1);
    assert(key(VK_ESCAPE, FALSE, 0, 0) == 1);
    assert(!g_printScreenKeyDown && !g_escapeKeyDown);
    assert(!postCount[WM_SCREEN_CAPTURE_BEGIN] && !postCount[WM_SCREEN_CAPTURE_CANCEL]);
    assert(postCount[WM_KEYBOARD_HOOK_STATUS] == 3);
}
static void loop(void) {
    for (int busy = 0; busy < 2; busy++) {
        stop = failInstall = installed = FALSE;
        waits = installCount = unhookCount = 0;
        now = 0;
        flood = busy;
        failPriority = busy; // Scheduling enhancement failure is non-fatal.
        SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, TRUE);
        assert(KeyboardHookThreadProc(NULL) == 0);
        assert(removedOnFirstWait && waits == 4 && installCount == 4);
        assert(unhookCount == 3 && !installed && !g_hHook);
        assert(!g_printScreenKeyDown && !g_escapeKeyDown && !logCount);
        assert(!g_printHotkeyRegistered && priorityCount == (UINT)busy + 1);
        assert(unregisterCount == (UINT)busy + 1);
    }
}
static void injection(void) {
    ctrlDown = clipboardImage = TRUE;
    assert(key('V', TRUE, 0, 0) == 1);
    assert(postCount[WM_DO_PASTE] == 1);
    SimulateStandardTextPaste();
    assert(postCount[WM_DO_PASTE] == 1 && logCount == 1);
    g_hookPastePending = FALSE; // UI completed the first paste.
    ctrlDown = TRUE;
    assert(key('V', TRUE, 0, 0) == 1); // Physical paste is not skipped.
    assert(postCount[WM_DO_PASTE] == 2);
    assert(key('V', TRUE, LLKHF_INJECTED, 123) == 1); // Other tools still work.
    SetCachedClipboardSequence(UINT32_MAX);
    assert(GetCachedClipboardSequence() == UINT32_MAX);
}
static void stale(void) {
    HWND target = GetForegroundWindow();
    assert(IsPasteRequestCurrent(target, clipboardSequence));
    assert(!IsPasteRequestCurrent(NULL, clipboardSequence));
    assert(!IsPasteRequestCurrent(target, clipboardSequence - 1));
    foregroundWindow = (void *)5;
    assert(!IsPasteRequestCurrent(target, clipboardSequence));
}
static void fastpath(void) {
    for (int i = 0; i < 1000000; i++) {
        assert(key('A', TRUE, 0, 0) == 77);
        assert(key('A', FALSE, 0, 0) == 77);
    }
    // Negative nCode must not even dereference lParam.
    assert(LowLevelKeyboardProc(-1, WM_KEYDOWN, 0) == 77);
    assert(key('V', FALSE, 0, 0) == 77);
    assert(key('V', TRUE, LLKHF_INJECTED, PASTE_INPUT_TAG) == 77);
    assert(atomicOps == 0 && routingQueries == 0 && lockAttempts == 0 && postAttempts == 0);
    g_hookHasKeywords = FALSE;
    assert(key('V', TRUE, 0, 0) == 77);
    assert(routingQueries == 0 && lockAttempts == 0 && postAttempts == 0);
    g_hookHasKeywords = TRUE;
    ctrlDown = TRUE; // Text-only clipboard: do not inspect foreground titles.
    assert(key('V', TRUE, 0, 0) == 77);
    assert(routingQueries == 4 && lockAttempts == 0 && postAttempts == 0);
}
static void capture_flood(void) {
    SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, TRUE);
    routingQueries = lockAttempts = postAttempts = atomicOps = 0;
    g_keywordLock = g_imageLock = 1;
    for (int i = 0; i < 100000; i++) {
        assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1);
        assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1); // Autorepeat.
        assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1);
    }
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == 1 && postAttempts == 1);
    assert(!routingQueries && !lockAttempts && !logCount);
    assert(atomicOps == 400000); // No hidden key-state/title/clipboard calls.
    g_hookPrintPending = FALSE;
    SetHookCaptureFlag(HOOK_CAPTURE_ACTIVE, TRUE);
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1);
    assert(postCount[WM_SCREEN_CAPTURE_COPY] == 1);
    // Renewal during a swallowed hold must not cause a second action on release.
    assert(RenewKeyboardHook(FALSE));
    g_hookPrintPending = FALSE;
    assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1);
    assert(postCount[WM_SCREEN_CAPTURE_COPY] == 1);
    for (int i = 0; i < 100000; i++) {
        assert(key(VK_ESCAPE, TRUE, 0, 0) == 1);
        assert(key(VK_ESCAPE, FALSE, 0, 0) == 1);
    }
    assert(postCount[WM_SCREEN_CAPTURE_CANCEL] == 1);
    g_hookCancelPending = FALSE;
    failPost = TRUE;
    assert(key(VK_ESCAPE, TRUE, 0, 0) == 77);
    assert(!g_hookCancelPending && !g_escapeKeyDown);
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 77);
    assert(!g_hookPrintPending && !g_printScreenKeyDown);
    failPost = FALSE;
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1); // Failed post does not poison the slot.
}
static void hotkey(void) {
    ReconcilePrintScreenHotkey();
    assert(!registerCount);
    SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, TRUE);
    failHotkey = TRUE;
    ReconcilePrintScreenHotkey();
    ReconcilePrintScreenHotkey();
    assert(!g_printHotkeyRegistered && registerCount == 2);
    assert(postCount[WM_KEYBOARD_HOTKEY_STATUS] == 1); // One warning per error.
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1); // Conflict does not disable hook.
    assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1);
    g_hookPrintPending = FALSE;
    failHotkey = FALSE;
    ReconcilePrintScreenHotkey();
    assert(g_printHotkeyRegistered && registerCount == 3);
    ReconcilePrintScreenHotkey();
    assert(registerCount == 3); // Never accumulate duplicate registrations.
    installed = FALSE;
    printDown = TRUE;
    HandlePrintScreenHotkey();
    assert(g_hookPrintPending && g_hookRenewRequested && g_printScreenKeyDown);
    UINT actions = postCount[WM_SCREEN_CAPTURE_BEGIN];
    HandlePrintScreenHotkey();
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == actions);
    assert(RenewKeyboardHook(FALSE));
    g_hookPrintPending = FALSE;
    printDown = FALSE;
    assert(key(VK_SNAPSHOT, FALSE, 0, 0) == 1);
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == actions); // No duplicate on key-up.
    SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, FALSE);
    HandlePrintScreenHotkey(); // Already queued hotkey after disabling capture.
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == actions);
    ReconcilePrintScreenHotkey();
    assert(!g_printHotkeyRegistered && unregisterCount == 1);
    HandlePrintScreenHotkey();
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == actions);
    SetHookCaptureFlag(HOOK_CAPTURE_ACTIVE, TRUE); // Manual tray capture still supports Print.
    ReconcilePrintScreenHotkey();
    HandlePrintScreenHotkey();
    assert(postCount[WM_SCREEN_CAPTURE_COPY] == 1);
    SetHookCaptureFlag(HOOK_CAPTURE_ACTIVE, FALSE);
    ReconcilePrintScreenHotkey();
    assert(!g_printHotkeyRegistered && unregisterCount == 2 && wakeCount == 4);
}
static void hotkey_loop(void) {
    loopHotkey = TRUE;
    SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, TRUE);
    assert(KeyboardHookThreadProc(NULL) == 0);
    assert(waits == 2 && installCount == 2 && now < KEYBOARD_HOOK_RENEW_MS);
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == 1 && !installed);
    assert(!g_printHotkeyRegistered && unregisterCount == 1);
}
static void lifecycle(void) {
    g_keyboardHookStopEvent = g_keyboardHookWakeEvent = NULL;
    for (UINT fail = 1; fail <= 3; fail++) {
        eventCount = closeCount = 0;
        failEventAt = fail;
        failThread = fail == 3;
        StartKeyboardHook();
        assert(!g_keyboardHookThread && !g_keyboardHookStopEvent && !g_keyboardHookWakeEvent);
        assert(!liveHandles && closeCount == fail - 1);
        StopKeyboardHook();
    }
    failEventAt = eventCount = closeCount = 0;
    failThread = FALSE;
    StartKeyboardHook();
    assert(g_keyboardHookThread && g_keyboardHookStopEvent && g_keyboardHookWakeEvent);
    assert(liveHandles == 3);
    StopKeyboardHook();
    assert(joined && !liveHandles && closeCount == 3);
    assert(!g_keyboardHookThread && !g_keyboardHookStopEvent && !g_keyboardHookWakeEvent);
    StopKeyboardHook(); // Cleanup is idempotent.
    assert(closeCount == 3);
}
static void paste_flood(void) {
    ctrlDown = clipboardImage = TRUE;
    SetHookCaptureFlag(HOOK_CAPTURE_ENABLED, TRUE);
    for (int i = 0; i < 100000; i++) assert(key('V', TRUE, 0, 0) == 1);
    assert(postCount[WM_DO_PASTE] == 1 && postAttempts == 1 && g_hookPastePending);
    assert(key(VK_SNAPSHOT, TRUE, 0, 0) == 1);
    assert(postCount[WM_SCREEN_CAPTURE_BEGIN] == 1); // Paste did not fill the queue.
    foregroundWindow = (void *)5;
    assert(key('V', TRUE, 0, 0) == 77); // Busy slot for a different target: fail open.
    foregroundWindow = (void *)4;
    clipboardSequence++;
    assert(key('V', TRUE, 0, 0) == 77); // Never coalesce a new clipboard with the old one.
    g_hookPastePending = FALSE;
    failPost = TRUE;
    assert(key('V', TRUE, 0, 0) == 77 && !g_hookPastePending);
    failPost = FALSE;
    assert(key('V', TRUE, 0, 0) == 1 && g_hookPastePending);
    assert(postCount[WM_DO_PASTE] == 2);
}
static void ui_slots(void) {
    g_hookPrintPending = TRUE;
    // Configuration disabled while capture was queued: release without opening.
    uiCapture(g_hWndMain, WM_SCREEN_CAPTURE_BEGIN, 0, HOOK_CAPTURE_MESSAGE);
    assert(!g_hookPrintPending && !beginCount);
    g_hookPrintPending = g_configScreenCaptureEnabled = TRUE;
    uiCapture(g_hWndMain, WM_SCREEN_CAPTURE_BEGIN, 0, HOOK_CAPTURE_MESSAGE);
    assert(!g_hookPrintPending && beginCount == 1);
    g_hookPrintPending = TRUE;
    uiCapture(g_hWndMain, WM_SCREEN_CAPTURE_COPY, TRUE, HOOK_CAPTURE_MESSAGE);
    assert(!g_hookPrintPending && copyCount == 1); // Release even if copy failed.
    g_hookCancelPending = TRUE;
    uiCapture(g_hWndMain, WM_SCREEN_CAPTURE_CANCEL, 0, HOOK_CAPTURE_MESSAGE);
    assert(!g_hookCancelPending && cancelCount == 1);
    g_hookPastePending = TRUE;
    uiPaste(g_hWndMain, WM_DO_PASTE, clipboardSequence - 1, (LPARAM)foregroundWindow);
    assert(!g_hookPastePending && !pasteCount);
    g_hookPastePending = failRefresh = TRUE;
    uiPaste(g_hWndMain, WM_DO_PASTE, clipboardSequence, (LPARAM)foregroundWindow);
    assert(!g_hookPastePending && !pasteCount);
    failRefresh = FALSE;
    g_hookPastePending = changeOnRefresh = TRUE;
    uiPaste(g_hWndMain, WM_DO_PASTE, clipboardSequence, (LPARAM)foregroundWindow);
    assert(!g_hookPastePending && !pasteCount);
    changeOnRefresh = FALSE;
    SetCachedClipboardSequence(clipboardSequence);
    g_hookPastePending = TRUE;
    uiPaste(g_hWndMain, WM_DO_PASTE, clipboardSequence, (LPARAM)foregroundWindow);
    assert(!g_hookPastePending && pasteCount == 1);
    g_hookPastePending = g_configCompatibilityPaste = ctrlDown = failTimer = TRUE;
    uiPaste(g_hWndMain, WM_DO_PASTE, clipboardSequence, (LPARAM)foregroundWindow);
    assert(!g_hookPastePending && !g_pasteDeferred && pasteCount == 1);
    failTimer = FALSE;
    g_hookPastePending = TRUE;
    uiPaste(g_hWndMain, WM_DO_PASTE, clipboardSequence, (LPARAM)foregroundWindow);
    assert(g_hookPastePending && g_pasteDeferred && pasteCount == 1);
    failPost = TRUE;
    uiTimer(g_hWndMain, ID_TIMER_DEFERRED_PASTE);
    assert(!g_hookPastePending && !g_pasteDeferred);
    failPost = FALSE;
    g_hookPastePending = TRUE;
    uiPaste(g_hWndMain, WM_DO_PASTE, clipboardSequence, (LPARAM)foregroundWindow);
    ctrlDown = FALSE;
    uiTimer(g_hWndMain, ID_TIMER_DEFERRED_PASTE);
    assert(g_hookPastePending && postCount[WM_DO_PASTE] == 1);
    uiPaste(g_hWndMain, WM_DO_PASTE, g_deferredPasteSequence, (LPARAM)g_deferredPasteTarget);
    assert(!g_hookPastePending && !g_pasteDeferred && pasteCount == 2);
    uiTimer(g_hWndMain, ID_TIMER_DEFERRED_PASTE); // Ignore an obsolete timer tick.
    assert(postCount[WM_DO_PASTE] == 1);
}
static void wake_loop(void) {
    loopWake = TRUE;
    assert(KeyboardHookThreadProc(NULL) == 0);
    assert(waits == 3 && installCount == 1 && registerCount == 1 && unregisterCount == 1);
}
int main(int argc, char **argv) {
    assert(argc == 2);
    if (!strcmp(argv[1], "routing")) routing();
    else if (!strcmp(argv[1], "renewal")) renewal();
    else if (!strcmp(argv[1], "loop")) loop();
    else if (!strcmp(argv[1], "injection")) injection();
    else if (!strcmp(argv[1], "stale")) stale();
    else if (!strcmp(argv[1], "fastpath")) fastpath();
    else if (!strcmp(argv[1], "flood")) capture_flood();
    else if (!strcmp(argv[1], "hotkey")) hotkey();
    else if (!strcmp(argv[1], "hotkey_loop")) hotkey_loop();
    else if (!strcmp(argv[1], "lifecycle")) lifecycle();
    else if (!strcmp(argv[1], "paste_flood")) paste_flood();
    else if (!strcmp(argv[1], "ui_slots")) ui_slots();
    else if (!strcmp(argv[1], "wake_loop")) wake_loop();
    else abort();
    return 0;
}
'''
