"""Exercise production updater calculations/parsing with inert OS boundaries.

Run: python3 -m unittest discover -s tests -v
Requires a host C compiler; never launches the Windows app or updater.
"""
from pathlib import Path
import re
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'main.c').read_text()


def function(name):
    # Match the definition, not a forward declaration of the same name.
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


class UpdaterTests(unittest.TestCase):
    def compile_and_run(self, source):
        with tempfile.TemporaryDirectory(prefix='imagepaster-test-') as directory:
            path = Path(directory)
            (path / 'test.c').write_text(source)
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(path / 'test.c'), '-o', str(path / 'test')], check=True)
            subprocess.run([str(path / 'test')], check=True)

    def test_download_progress_percent(self):
        self.compile_and_run(r'''
#include <assert.h>
#include <stdint.h>
typedef uint32_t DWORD;
typedef uint64_t ULONGLONG;
''' + function('CalculateUpdateProgressPercent') + r'''
int main(void) {
    assert(CalculateUpdateProgressPercent(0, 1000) == 0);
    assert(CalculateUpdateProgressPercent(9, 1000) == 0); // 0.9% is not 1%
    assert(CalculateUpdateProgressPercent(10, 1000) == 1);
    assert(CalculateUpdateProgressPercent(2, 3) == 66); // whole percent, rounded down
    assert(CalculateUpdateProgressPercent(999, 1000) == 99);
    assert(CalculateUpdateProgressPercent(104857599, 104857600) == 99);
    assert(CalculateUpdateProgressPercent(1000, 1000) == 100); // only when complete
    assert(CalculateUpdateProgressPercent(1001, 1000) == 100);
    assert(CalculateUpdateProgressPercent(5, 0) == 0);
    assert(CalculateUpdateProgressPercent(UINT64_MAX - 1, UINT64_MAX) == 99);
    for (ULONGLONG total = 1; total <= 5000; total += 7) {
        for (ULONGLONG received = 0; received <= total; received++) {
            assert(CalculateUpdateProgressPercent(received, total) ==
                   received * 100 / total);
        }
    }
    const ULONGLONG release = 1073178; // a real executable size, 64 KiB reads
    for (ULONGLONG received = 0; received < release; received += 65536) {
        assert(CalculateUpdateProgressPercent(received, release) ==
               received * 100 / release);
    }
    return 0;
}
''')

    def test_stop_is_immediate_and_late_results_are_discarded(self):
        marker = 'static UpdateCheckTask* g_updateReadyTask = NULL;'
        state = SOURCE[SOURCE.index('static volatile LONG g_updateCheckPending'):
                       SOURCE.index(marker) + len(marker)]
        messages = '\n'.join(re.findall(r'^#define WM_APP_UPDATE_\w+\s+.*$',
                                        SOURCE, re.MULTILINE))
        production = '\n'.join(function(name) for name in (
            'CancelUpdateTaskIfRequested', 'PublishUpdateProgress',
            'PublishUpdateTask', 'CfgSendUpdateResult',
            'DiscardPendingUpdateNotice', 'IsUpdateWorkerActive',
            'IsUpdateCheckVisible', 'StartUpdateCheck', 'StartQueuedUpdateCheck',
            'HandleCompletedUpdateCheck', 'CancelUpdateCheck',
            'DiscardPreparedUpdate'))
        # Run the real message handlers rather than copies of their logic.
        window = function('WndProc')
        ui_cases = window[window.index('    case WM_APP_UPDATE_PROGRESS:'):
                          window.index('    case WM_HISTORY_THUMBS_READY:')]
        webview = function('WebViewWndProc')
        destroy_case = webview[webview.index('        case WM_DESTROY:'):
                               webview.index('\n    }\n    return DefWindowProcW(')]
        self.compile_and_run(STOP_PREFIX + messages + '\n' + state + STOP_STUBS +
                             production + r'''
static LRESULT uiUpdate(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    (void)hWnd; (void)wParam; (void)lParam;
    switch (msg) {
''' + ui_cases + r'''
    }
    return -1;
}
static LRESULT webviewDestroy(HWND hwnd, UINT msg) {
    switch (msg) {
''' + destroy_case + r'''
    }
    return -1;
}
''' + STOP_SCENARIOS)

    def test_handoff_choice_is_scoped_to_success(self):
        self.compile_and_run(r'''
#include <assert.h>
#include <stdint.h>
#include <stdlib.h>
#include <wchar.h>
typedef int BOOL;
typedef uint32_t DWORD;
typedef const wchar_t *LPCWSTR;
typedef wchar_t *LPWSTR;
#define TRUE 1
#define FALSE 0
#define ERROR_INVALID_PARAMETER 87
static int argc_mock, apply_calls, cleanup_calls, apply_choice;
static BOOL recognized = TRUE;
static LPWSTR argv_mock[8];
static LPCWSTR GetCommandLineW(void) { return L""; }
static LPWSTR *CommandLineToArgvW(LPCWSTR command, int *count) {
    (void)command;
    *count = argc_mock;
    return argv_mock;
}
static void LocalFree(void *memory) { (void)memory; }
static int RunUpdateApplyHelper(DWORD pid, LPCWSTR event, LPCWSTR target,
                               LPCWSTR staged, BOOL reopen) {
    assert(pid == 12);
    assert(wcscmp(event, L"event") == 0);
    assert(wcscmp(target, L"target") == 0);
    assert(wcscmp(staged, L"staged") == 0);
    apply_choice = reopen;
    apply_calls++;
    return 5; // Simulate a helper failure without filesystem/process effects.
}
static BOOL FinishUpdateCleanup(DWORD helper, DWORD old, LPCWSTR staged,
                                LPCWSTR updater) {
    assert(helper == 12 && old == 34);
    assert(wcscmp(staged, L"staged") == 0);
    assert(wcscmp(updater, L"helper") == 0);
    cleanup_calls++;
    return recognized;
}
''' + function('ParseUpdateProcessId') + '\n' + function('HandleUpdateCommandLine') + r'''
static void check(int count, LPWSTR action, LPWSTR option,
                  BOOL expected_handled, BOOL expected_completed, BOOL expected_reopen) {
    argc_mock = count;
    argv_mock[0] = L"ImagePaster.exe";
    argv_mock[1] = action;
    argv_mock[2] = L"12";
    argv_mock[3] = L"34";
    argv_mock[4] = L"staged";
    argv_mock[5] = L"helper";
    argv_mock[6] = option;
    if (wcscmp(action, L"--apply-update") == 0) {
        argv_mock[3] = L"event";
        argv_mock[4] = L"target";
        argv_mock[5] = L"staged";
    }
    BOOL handled = TRUE, completed = TRUE, reopen = TRUE;
    HandleUpdateCommandLine(&handled, &completed, &reopen);
    assert(handled == expected_handled);
    assert(completed == expected_completed);
    assert(reopen == expected_reopen);
}
int main(void) {
    check(7, L"--apply-update", L"--reopen-settings", TRUE, FALSE, FALSE);
    assert(apply_calls == 1 && apply_choice == TRUE);
    check(6, L"--apply-update", L"", TRUE, FALSE, FALSE);
    assert(apply_calls == 2 && apply_choice == FALSE);
    check(7, L"--finish-update", L"--reopen-settings", FALSE, TRUE, TRUE);
    check(6, L"--finish-update", L"", FALSE, TRUE, FALSE); // legacy/unchecked
    check(7, L"--finish-update-cleanup", L"--reopen-settings", FALSE, FALSE, FALSE);
    check(6, L"--finish-update-cleanup", L"", FALSE, FALSE, FALSE);
    assert(cleanup_calls == 4);
    recognized = FALSE;
    check(7, L"--finish-update", L"--reopen-settings", FALSE, FALSE, FALSE);
    recognized = TRUE;
    int previous_cleanup_calls = cleanup_calls;
    check(7, L"--finish-update", L"--unknown", FALSE, FALSE, FALSE);
    check(8, L"--finish-update", L"--reopen-settings", FALSE, FALSE, FALSE);
    check(2, L"--reopen-settings", L"", FALSE, FALSE, FALSE);
    check(1, L"", L"", FALSE, FALSE, FALSE); // next ordinary launch resets output
    assert(cleanup_calls == previous_cleanup_calls);
    argc_mock = 7;
    argv_mock[1] = L"--finish-update";
    argv_mock[2] = L"invalid";
    argv_mock[6] = L"--reopen-settings";
    BOOL handled, completed, reopen;
    HandleUpdateCommandLine(&handled, &completed, &reopen);
    assert(!handled && !completed && !reopen);
    assert(cleanup_calls == previous_cleanup_calls);
    return 0;
}
''')


STOP_PREFIX = r'''
#include <assert.h>
#include <stddef.h>
#include <stdint.h>
#include <stdlib.h>
#include <wchar.h>
typedef int BOOL;
typedef int32_t LONG;
typedef uint16_t WORD;
typedef uint32_t DWORD, UINT;
typedef uint64_t ULONGLONG;
typedef uintptr_t WPARAM, UINT_PTR;
typedef intptr_t LPARAM, LRESULT;
typedef void *HWND, *HANDLE, *PVOID, *LPVOID;
typedef const wchar_t *LPCWSTR;
typedef DWORD (*LPTHREAD_START_ROUTINE)(LPVOID);
#define WINAPI
#define TRUE 1
#define FALSE 0
#define MAX_PATH 260
#define WM_APP 0x8000
#define WM_DESTROY 0x0002
#define WM_CLOSE 0x0010
#define WM_COMMAND 0x0111
#define ID_TRAY_EXIT 1
#define ID_TIMER_WEBVIEW_SHOW_FALLBACK 2
#define WAIT_OBJECT_0 0
#define WAIT_TIMEOUT 258
#define swprintf_s swprintf
'''

STOP_STUBS = r'''
static HWND g_hWndMain = (HWND)1, g_webviewHwnd = (HWND)2;
static BOOL g_webviewWindowShown, g_configViewReady = TRUE;
static BOOL g_configAutoCheckForUpdates = TRUE;
static BOOL configOpen = TRUE, cancelSignalled, failPost, failThread;
static int publishing; /* 1: a result handoff is expected, 2: none */
static int threadsStarted, sentResults, progressSent, presented, discarded;
static DWORD lastProgress;
static wchar_t lastStatus[32];
static UpdateCheckKind presentedKind;
static BOOL presentedAutomatic;
static UpdateCheckTask *startedTask;
static struct { HWND window; UINT message; } posted[16];
static int queued;

static LONG InterlockedExchange(volatile LONG *target, LONG value) {
    // A worker must hand over its result before reporting itself idle.
    if (publishing == 1 && target == &g_updateCheckPending && value == FALSE) {
        assert(g_updatePostedResult != NULL);
    }
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
static LONG InterlockedIncrement(volatile LONG *target) { return ++*target; }
static PVOID InterlockedExchangePointer(PVOID volatile *target, PVOID value) {
    PVOID previous = *target;
    *target = value;
    return previous;
}
static PVOID InterlockedCompareExchangePointer(PVOID volatile *target,
                                              PVOID value, PVOID comparand) {
    PVOID previous = *target;
    if (previous == comparand) *target = value;
    return previous;
}
static HANDLE CreateEventW(void *attributes, BOOL manualReset, BOOL initial,
                           LPCWSTR name) {
    (void)attributes; (void)initial; (void)name;
    assert(manualReset);
    return (HANDLE)0x55;
}
static BOOL SetEvent(HANDLE event) {
    assert(event == (HANDLE)0x55);
    cancelSignalled = TRUE;
    return TRUE;
}
static BOOL ResetEvent(HANDLE event) {
    assert(event == (HANDLE)0x55);
    cancelSignalled = FALSE;
    return TRUE;
}
static DWORD WaitForSingleObject(HANDLE event, DWORD milliseconds) {
    assert(event == (HANDLE)0x55 && milliseconds == 0);
    return cancelSignalled ? WAIT_OBJECT_0 : WAIT_TIMEOUT;
}
static DWORD WINAPI UpdateCheckThread(LPVOID parameter) {
    (void)parameter;
    return 0;
}
static HANDLE CreateThread(void *attributes, size_t stack,
                           LPTHREAD_START_ROUTINE start, LPVOID parameter,
                           DWORD flags, DWORD *id) {
    (void)attributes; (void)stack; (void)flags; (void)id;
    assert(start == UpdateCheckThread);
    if (failThread) return NULL;
    startedTask = (UpdateCheckTask *)parameter;
    threadsStarted++;
    return (HANDLE)0x77;
}
static BOOL CloseHandle(HANDLE handle) { return handle == (HANDLE)0x77; }
static DWORD GetLastError(void) { return 5; }
static ULONGLONG GetTickCount64(void) { return 1234; }
static DWORD GetCurrentProcessId(void) { return 42; }
static BOOL IsWindow(HWND window) { return window != NULL; }
static BOOL PostMessageW(HWND window, UINT message, WPARAM wParam, LPARAM lParam) {
    (void)wParam; (void)lParam;
    if (message == WM_APP_UPDATE_RESULT) {
        // The UI thread may handle this before the worker runs again.
        assert(g_updateCheckPending == FALSE && g_updatePostedResult != NULL);
    }
    if (failPost) return FALSE;
    assert(queued < 16);
    posted[queued].window = window;
    posted[queued].message = message;
    queued++;
    return TRUE;
}
static LRESULT SendMessageW(HWND window, UINT message, WPARAM wParam,
                            LPARAM lParam) {
    (void)window; (void)message; (void)wParam; (void)lParam;
    return 0;
}
static BOOL KillTimer(HWND window, UINT_PTR id) {
    (void)window; (void)id;
    return TRUE;
}
static void UpdateDebugPrint(const wchar_t *format, ...) { (void)format; }
static void CfgSendUpdateResultWithVersions(LPCWSTR status, LPCWSTR title,
    LPCWSTR message, LPCWSTR currentVersion, LPCWSTR remoteVersion,
    BOOL automatic) {
    (void)title; (void)message; (void)currentVersion; (void)remoteVersion;
    (void)automatic;
    wcsncpy(lastStatus, status, 31);
    sentResults++;
}
static void CfgSendUpdateProgress(DWORD percent) {
    lastProgress = percent;
    progressSent++;
}
static BOOL IsConfigurationViewOpen(void) { return configOpen; }
static BOOL IsIgnoredUpdateVersion(const ExecutableVersion *version) {
    (void)version;
    return FALSE;
}
static void DiscardUpdateTask(UpdateCheckTask *task) {
    if (!task) return;
    discarded++;
    free(task);
}
static void QueueUpdateNotice(UpdateCheckTask *task) {
    presented++;
    presentedKind = task->kind;
    presentedAutomatic = task->automatic;
    DiscardUpdateTask(task);
}
static void ShowWebViewDialog(const char *view, int width, int height) {
    (void)view; (void)width; (void)height;
}
static void DropHistoryThumbnailRequests(void) {}
static void DiscardPreparedUpdate(void);
'''

STOP_SCENARIOS = r'''
static void deliver(void) {
    for (int index = 0; index < queued; index++) {
        uiUpdate(posted[index].window, posted[index].message, 0, 0);
    }
    queued = 0;
}

/* What a worker does once its blocking network call finally returns. */
static void finish(UpdateCheckTask *task, UpdateCheckKind kind) {
    task->kind = kind;
    publishing = task->targetWindow ? 1 : 2;
    PublishUpdateTask(task);
    publishing = 0;
}

static void reset(void) {
    g_updateCheckPending = g_updateCheckAutomatic = FALSE;
    g_updatePostedResult = NULL;
    g_updateNoticeTask = g_updateReadyTask = NULL;
    g_updateCheckAbandoned = g_updateQueuedCheck = g_updateQueuedAutomatic = FALSE;
    g_updateProgressPercent = g_updateProgressPosted = 0;
    g_configViewReady = configOpen = TRUE;
    cancelSignalled = failPost = failThread = FALSE;
    threadsStarted = sentResults = progressSent = presented = discarded = 0;
    queued = 0;
    lastProgress = 0;
    lastStatus[0] = L'\0';
    startedTask = NULL;
}

static void stop_during_download(void) {
    reset();
    StartUpdateCheck(FALSE);
    UpdateCheckTask *first = startedTask;
    assert(threadsStarted == 1 && first && !first->automatic);
    assert(IsUpdateCheckVisible());

    PublishUpdateProgress(first, 42);
    deliver();
    assert(progressSent == 1 && lastProgress == 42);

    PublishUpdateProgress(first, 43); // posted but not displayed yet
    CancelUpdateCheck();
    assert(cancelSignalled && g_updateCheckAbandoned);
    assert(sentResults == 1 && wcscmp(lastStatus, L"cancelled") == 0);
    assert(!IsUpdateCheckVisible());
    deliver();
    assert(progressSent == 1); // stale progress never reaches settings
    PublishUpdateProgress(first, 44); // the worker sees the stop instead
    assert(queued == 0 && first->kind == UPDATE_CHECK_CANCELLED);

    StartUpdateCheck(FALSE); // clicked again while the worker unwinds
    assert(threadsStarted == 1 && sentResults == 1); // queued, no conflict error
    assert(g_updateQueuedCheck && !g_updateQueuedAutomatic);
    assert(IsUpdateCheckVisible());

    finish(first, UPDATE_CHECK_NEWER);
    assert(first->kind == UPDATE_CHECK_CANCELLED);
    deliver();
    assert(presented == 0 && sentResults == 1 && discarded == 1);
    assert(threadsStarted == 2 && !startedTask->automatic);
    assert(!cancelSignalled && !g_updateCheckAbandoned && !g_updateQueuedCheck);
    assert(IsUpdateCheckVisible());

    finish(startedTask, UPDATE_CHECK_NEWER);
    deliver();
    assert(presented == 1 && presentedKind == UPDATE_CHECK_NEWER);
    assert(!presentedAutomatic && !IsUpdateCheckVisible());
}

static void stop_after_result_posted(void) {
    reset();
    StartUpdateCheck(FALSE);
    UpdateCheckTask *task = startedTask;
    finish(task, UPDATE_CHECK_NEWER); // finished, not handled by the UI yet
    assert(!g_updateCheckPending && g_updatePostedResult == task);
    assert(IsUpdateWorkerActive() && IsUpdateCheckVisible());
    CancelUpdateCheck();
    assert(sentResults == 1 && wcscmp(lastStatus, L"cancelled") == 0);
    deliver();
    assert(presented == 0 && discarded == 1 && !IsUpdateCheckVisible());
}

static void idle_and_queued_stops(void) {
    reset();
    CancelUpdateCheck(); // nothing to stop: no stray result
    assert(sentResults == 0 && !g_updateCheckAbandoned && !cancelSignalled);

    StartUpdateCheck(FALSE);
    UpdateCheckTask *first = startedTask;
    CancelUpdateCheck();
    CancelUpdateCheck(); // already stopped
    assert(sentResults == 1);
    StartUpdateCheck(FALSE);
    CancelUpdateCheck(); // a queued request also stops at once
    assert(sentResults == 2 && !g_updateQueuedCheck && !IsUpdateCheckVisible());
    finish(first, UPDATE_CHECK_SAME);
    deliver();
    assert(threadsStarted == 1 && presented == 0 && !IsUpdateWorkerActive());
}

static void queued_request_kinds(void) {
    reset();
    StartUpdateCheck(TRUE);
    UpdateCheckTask *first = startedTask;
    assert(first->automatic);
    CancelUpdateCheck();
    StartUpdateCheck(TRUE); // settings reopened with automatic checks on
    assert(g_updateQueuedCheck && g_updateQueuedAutomatic);
    StartUpdateCheck(FALSE); // a manual request replaces it
    StartUpdateCheck(TRUE); // and is not downgraded again
    assert(g_updateQueuedCheck && !g_updateQueuedAutomatic);
    finish(first, UPDATE_CHECK_OLDER);
    deliver();
    assert(threadsStarted == 2 && !startedTask->automatic && presented == 0);
}

static void closing_settings(void) {
    reset();
    StartUpdateCheck(FALSE);
    UpdateCheckTask *manual = startedTask;
    webviewDestroy(NULL, WM_DESTROY); // stops a manual check at once
    assert(cancelSignalled && g_updateCheckAbandoned && !g_configViewReady);
    assert(!IsUpdateCheckVisible()); // reopened settings start idle...
    StartUpdateCheck(TRUE); // ...and queue their automatic check
    assert(g_updateQueuedCheck && g_updateQueuedAutomatic);
    webviewDestroy(NULL, WM_DESTROY); // closing keeps automatic requests
    assert(g_updateQueuedCheck);
    finish(manual, UPDATE_CHECK_NEWER);
    deliver();
    assert(presented == 0 && threadsStarted == 2 && startedTask->automatic);

    UpdateCheckTask *automatic = startedTask;
    webviewDestroy(NULL, WM_DESTROY); // automatic checks keep running
    assert(!cancelSignalled && !g_updateCheckAbandoned);
    configOpen = FALSE;
    finish(automatic, UPDATE_CHECK_NEWER);
    deliver();
    assert(presented == 1 && presentedAutomatic);

    StartUpdateCheck(FALSE);
    UpdateCheckTask *stopped = startedTask;
    CancelUpdateCheck();
    StartUpdateCheck(FALSE); // a queued manual request...
    webviewDestroy(NULL, WM_DESTROY); // ...is dropped with settings
    assert(!g_updateQueuedCheck);
    finish(stopped, UPDATE_CHECK_SAME);
    deliver();
    assert(threadsStarted == 3 && presented == 1);
}

static void lost_results_leave_updater_idle(void) {
    reset();
    StartUpdateCheck(FALSE);
    UpdateCheckTask *first = startedTask;
    CancelUpdateCheck();
    StartUpdateCheck(FALSE); // queued behind the stopped check
    failPost = TRUE;
    finish(first, UPDATE_CHECK_SAME); // its result cannot be posted
    failPost = FALSE;
    assert(discarded == 1 && !IsUpdateWorkerActive());
    StartUpdateCheck(TRUE); // the next check fulfils the manual request
    assert(threadsStarted == 2 && !startedTask->automatic);
    assert(!g_updateQueuedCheck && !g_updateCheckAbandoned);

    UpdateCheckTask *orphan = startedTask;
    orphan->targetWindow = NULL; // no window left to receive a result
    finish(orphan, UPDATE_CHECK_SAME);
    assert(queued == 0 && discarded == 2 && !IsUpdateWorkerActive());

    failThread = TRUE;
    StartUpdateCheck(FALSE);
    assert(wcscmp(lastStatus, L"error") == 0 && !IsUpdateWorkerActive());
}

int main(void) {
    stop_during_download();
    stop_after_result_posted();
    idle_and_queued_stops();
    queued_request_kinds();
    closing_settings();
    lost_results_leave_updater_idle();
    return 0;
}
'''


if __name__ == '__main__':
    unittest.main()
