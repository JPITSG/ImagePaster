"""Exercise production updater calculations/parsing with inert OS boundaries.

Run: python3 -m unittest discover -s tests -v
Requires a host C compiler; never launches the Windows app or updater.
"""
from pathlib import Path
import subprocess
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT / 'main.c').read_text()


def function(name):
    start = SOURCE.index('static ', SOURCE.index(name + '(') - 20)
    opening = SOURCE.index('{', start)
    depth = 1
    end = opening + 1
    while depth:
        depth += (SOURCE[end] == '{') - (SOURCE[end] == '}')
        end += 1
    return SOURCE[start:end]


class UpdaterTests(unittest.TestCase):
    def compile_and_run(self, source):
        with tempfile.TemporaryDirectory(prefix='imagepaster-test-') as directory:
            path = Path(directory)
            (path / 'test.c').write_text(source)
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(path / 'test.c'), '-o', str(path / 'test')], check=True)
            subprocess.run([str(path / 'test')], check=True)

    def test_received_byte_speed_rounding(self):
        self.compile_and_run(r'''
#include <assert.h>
#include <stdint.h>
typedef uint32_t DWORD;
typedef uint64_t ULONGLONG;
#define MAXLONG INT32_MAX
''' + function('CalculateUpdateSpeedKbps') + r'''
int main(void) {
    assert(CalculateUpdateSpeedKbps(25600, 250) == 100);
    assert(CalculateUpdateSpeedKbps(25472, 250) == 100); // 99.5 rounds up
    assert(CalculateUpdateSpeedKbps(25471, 250) == 99);
    assert(CalculateUpdateSpeedKbps(127, 250) == 0); // no forced minimum
    assert(CalculateUpdateSpeedKbps(128, 250) == 1);
    assert(CalculateUpdateSpeedKbps(25600, 500) == 50);
    assert(CalculateUpdateSpeedKbps(0, 250) == 0);
    assert(CalculateUpdateSpeedKbps(65536, 0) == 0);
    assert(CalculateUpdateSpeedKbps(104857600, 1) == 102400000);
    assert(CalculateUpdateSpeedKbps(UINT64_MAX, 1) == MAXLONG);
    return 0;
}
''')

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


if __name__ == '__main__':
    unittest.main()
