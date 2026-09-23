"""Run main.c's Start with Windows functions against an in-memory registry.

Covers the per-user Run entry and the StartupApproved marker Task Manager uses
to disable it. Run: python3 -m unittest discover -s tests -v
"""
from pathlib import Path
import re
import subprocess
import tempfile
import unittest

from test_updater import SOURCE, function


class StartupTests(unittest.TestCase):
    def test_run_entry_follows_the_toggle_and_task_manager(self):
        definitions = '\n'.join(
            re.findall(r'^#define (?:APP_NAME|STARTUP_\w+)\b(?:[^\n]*\\\n)*[^\n]*$',
                       SOURCE, re.MULTILINE))
        production = '\n'.join(function(name) for name in (
            'GetStartupCommand', 'IsStartWithWindowsEnabled',
            'SetStartWithWindows', 'ApplyStartWithWindows'))
        with tempfile.TemporaryDirectory(prefix='imagepaster-startup-') as directory:
            path = Path(directory)
            (path / 'startup.c').write_text(PREFIX + definitions + '\n' + STUBS +
                                            production + SCENARIO)
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(path / 'startup.c'), '-o', str(path / 'startup')],
                           check=True)
            subprocess.run([str(path / 'startup')], check=True, timeout=10)


PREFIX = r'''
#define _GNU_SOURCE
#include <assert.h>
#include <stdarg.h>
#include <stdint.h>
#include <stdio.h>
#include <string.h>
#include <wchar.h>
typedef int BOOL;
typedef long LONG;
typedef uint8_t BYTE;
typedef uint32_t DWORD;
typedef void *HKEY, *HMODULE;
typedef const wchar_t *LPCWSTR;
#define TRUE 1
#define FALSE 0
#define MAX_PATH 260
#define HKEY_CURRENT_USER ((HKEY)0x80000001)
#define RRF_RT_REG_SZ 0x2
#define RRF_RT_REG_BINARY 0x8
#define REG_SZ 1
#define REG_OPTION_NON_VOLATILE 0
#define KEY_SET_VALUE 0x2
#define ERROR_SUCCESS 0
#define ERROR_FILE_NOT_FOUND 2
#define ERROR_ACCESS_DENIED 5
#define ERROR_BAD_PATHNAME 161
#define ERROR_MORE_DATA 234
#define _wcsicmp wcscasecmp
'''

STUBS = r'''
static wchar_t modulePath[400];
static struct { BOOL present; wchar_t text[400]; } run;
static struct { BOOL present; BYTE bytes[12]; } approved;
static int writes, deletes, failCreate;
static char lastLog[256];

/* MSVCRT's swprintf reads %s as a wide string; glibc spells that %ls. */
static int msvc_swprintf(wchar_t *buffer, size_t count, const wchar_t *format,
                         const wchar_t *text) {
    assert(wcscmp(format, L"\"%s\"") == 0);
    return swprintf(buffer, count, L"\"%ls\"", text);
}
#define swprintf msvc_swprintf

static DWORD GetModuleFileNameW(HMODULE module, wchar_t *buffer, DWORD size) {
    size_t length = wcslen(modulePath);
    assert(module == NULL);
    if (length >= size) { /* truncated, as Windows reports it */
        wmemcpy(buffer, modulePath, size - 1);
        buffer[size - 1] = L'\0';
        return size;
    }
    wcscpy(buffer, modulePath);
    return (DWORD)length;
}

static BOOL isRunKey(LPCWSTR key) { return wcscmp(key, STARTUP_RUN_KEY_W) == 0; }
static BOOL isApprovedKey(LPCWSTR key) {
    return wcscmp(key, STARTUP_APPROVED_RUN_KEY_W) == 0;
}

static LONG RegGetValueW(HKEY root, LPCWSTR key, LPCWSTR name, DWORD flags,
                         DWORD *type, void *data, DWORD *size) {
    assert(root == HKEY_CURRENT_USER && wcscmp(name, APP_NAME) == 0 && !type);
    if (isRunKey(key)) {
        DWORD needed = (DWORD)((wcslen(run.text) + 1) * sizeof(wchar_t));
        assert(flags == RRF_RT_REG_SZ);
        if (!run.present) return ERROR_FILE_NOT_FOUND;
        if (needed > *size) return ERROR_MORE_DATA;
        memcpy(data, run.text, needed);
        *size = needed;
        return ERROR_SUCCESS;
    }
    assert(isApprovedKey(key) && flags == RRF_RT_REG_BINARY);
    if (!approved.present) return ERROR_FILE_NOT_FOUND;
    assert(*size >= sizeof(approved.bytes));
    memcpy(data, approved.bytes, sizeof(approved.bytes));
    *size = sizeof(approved.bytes);
    return ERROR_SUCCESS;
}

static LONG RegCreateKeyExW(HKEY root, LPCWSTR key, DWORD reserved, void *cls,
                            DWORD options, DWORD access, void *security,
                            HKEY *result, DWORD *disposition) {
    (void)reserved; (void)cls; (void)security; (void)disposition;
    assert(root == HKEY_CURRENT_USER && isRunKey(key));
    assert(options == REG_OPTION_NON_VOLATILE && access == KEY_SET_VALUE);
    if (failCreate) return ERROR_ACCESS_DENIED;
    *result = (HKEY)0x77;
    return ERROR_SUCCESS;
}

static LONG RegSetValueExW(HKEY key, LPCWSTR name, DWORD reserved, DWORD type,
                           const BYTE *data, DWORD size) {
    (void)reserved;
    assert(key == (HKEY)0x77 && wcscmp(name, APP_NAME) == 0 && type == REG_SZ);
    assert(size == (wcslen((const wchar_t *)data) + 1) * sizeof(wchar_t));
    memcpy(run.text, data, size);
    run.present = TRUE;
    writes++;
    return ERROR_SUCCESS;
}

static LONG RegCloseKey(HKEY key) {
    assert(key == (HKEY)0x77);
    return ERROR_SUCCESS;
}

static LONG RegDeleteKeyValueW(HKEY root, LPCWSTR key, LPCWSTR name) {
    BOOL *present = isRunKey(key) ? &run.present : &approved.present;
    assert(root == HKEY_CURRENT_USER && wcscmp(name, APP_NAME) == 0);
    assert(isRunKey(key) || isApprovedKey(key));
    deletes++;
    if (!*present) return ERROR_FILE_NOT_FOUND;
    *present = FALSE;
    return ERROR_SUCCESS;
}

static void LogMessage(const char *format, ...) {
    va_list args;
    va_start(args, format);
    vsnprintf(lastLog, sizeof(lastLog), format, args);
    va_end(args);
}
'''

SCENARIO = r'''
static void disable_in_task_manager(BYTE state) {
    approved.present = TRUE;
    memset(approved.bytes, 0, sizeof(approved.bytes));
    approved.bytes[0] = state;
}

int main(void) {
    const wchar_t *expected = L"\"C:\\Program Files\\ImagePaster\\ImagePaster.exe\"";
    int writesBefore, deletesBefore;

    wcscpy(modulePath, L"C:\\Program Files\\ImagePaster\\ImagePaster.exe");
    assert(!IsStartWithWindowsEnabled()); /* off until the user turns it on */

    ApplyStartWithWindows(TRUE); /* the quoted path survives spaces */
    assert(run.present && wcscmp(run.text, expected) == 0);
    assert(IsStartWithWindowsEnabled());
    assert(strcmp(lastLog, "Start with Windows enabled") == 0);

    writesBefore = writes;
    deletesBefore = deletes;
    ApplyStartWithWindows(TRUE); /* an unchanged toggle touches nothing */
    assert(writes == writesBefore && deletes == deletesBefore);

    disable_in_task_manager(3); /* odd first byte: disabled in Task Manager */
    assert(!IsStartWithWindowsEnabled());
    disable_in_task_manager(2); /* even first byte: enabled again */
    assert(IsStartWithWindowsEnabled());
    disable_in_task_manager(3);
    ApplyStartWithWindows(TRUE); /* turning it on clears the disabled marker */
    assert(!approved.present && IsStartWithWindowsEnabled());

    wcscpy(run.text, L"\"c:\\program files\\imagepaster\\IMAGEPASTER.EXE\"");
    assert(IsStartWithWindowsEnabled()); /* Windows paths ignore case */

    disable_in_task_manager(2);
    ApplyStartWithWindows(FALSE); /* off removes the entry and the marker */
    assert(!run.present && !approved.present && !IsStartWithWindowsEnabled());
    assert(strcmp(lastLog, "Start with Windows disabled") == 0);
    assert(SetStartWithWindows(FALSE) == ERROR_SUCCESS); /* already off */

    run.present = TRUE; /* another copy of ImagePaster owns the entry */
    wcscpy(run.text, L"\"D:\\Tools\\ImagePaster.exe\"");
    assert(!IsStartWithWindowsEnabled());
    writesBefore = writes;
    deletesBefore = deletes;
    ApplyStartWithWindows(FALSE); /* saving with the toggle off leaves it alone */
    assert(run.present && writes == writesBefore && deletes == deletesBefore);
    ApplyStartWithWindows(TRUE); /* turning it on points it at this copy */
    assert(wcscmp(run.text, expected) == 0 && IsStartWithWindowsEnabled());

    run.present = FALSE;
    failCreate = 1;
    assert(SetStartWithWindows(TRUE) == ERROR_ACCESS_DENIED);
    ApplyStartWithWindows(TRUE); /* failures are logged, never hidden */
    assert(strcmp(lastLog, "WARNING: Could not enable start with Windows (error 5)") == 0);
    assert(!run.present);
    failCreate = 0;

    for (int i = 0; i < MAX_PATH + 20; i++) modulePath[i] = L'x';
    modulePath[MAX_PATH + 20] = L'\0'; /* a path Windows would truncate */
    assert(!IsStartWithWindowsEnabled());
    assert(SetStartWithWindows(TRUE) == ERROR_BAD_PATHNAME && !run.present);
    return 0;
}
'''


if __name__ == '__main__':
    unittest.main()
