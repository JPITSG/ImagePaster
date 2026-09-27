"""Run the production startup migration and toggle with inert OS boundaries.

Covers per-user task migration, disabled settings, ownership and failures. Run: python3 -m unittest discover -s tests -v
"""
from pathlib import Path
import re
import subprocess
import tempfile
import unittest

from test_updater import SOURCE, function


class StartupTests(unittest.TestCase):
    def test_task_migration_and_toggle(self):
        definitions = '\n'.join(
            re.findall(r'^#define (?:APP_NAME|STARTUP_\w+)\b(?:[^\n]*\\\n)*[^\n]*$',
                       SOURCE, re.MULTILINE))
        production = '\n'.join(function(name) for name in (
            'GetStartupCommand', 'IsLegacyStartWithWindowsEnabled',
            'RemoveLegacyStartupEntry', 'IsStartWithWindowsEnabled',
            'SetStartWithWindows', 'MigrateStartWithWindows', 'ApplyStartWithWindows'))
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
static int deletes, failDelete, taskWrites, failWrite, failRead;
static BOOL taskPresent, taskEnabled;
static LONG ReadStartupTask(BOOL *present, BOOL *enabled) {
    *present = taskPresent;
    *enabled = taskEnabled;
    return failRead ? ERROR_ACCESS_DENIED : ERROR_SUCCESS;
}
static LONG WriteStartupTask(BOOL enable) {
    taskWrites++;
    if (failWrite) return ERROR_ACCESS_DENIED;
    taskPresent = taskEnabled = enable;
    return ERROR_SUCCESS;
}
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

static LONG RegDeleteKeyValueW(HKEY root, LPCWSTR key, LPCWSTR name) {
    BOOL *present = isRunKey(key) ? &run.present : &approved.present;
    assert(root == HKEY_CURRENT_USER && wcscmp(name, APP_NAME) == 0);
    assert(isRunKey(key) || isApprovedKey(key));
    if (failDelete) return ERROR_ACCESS_DENIED;
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
    int before;
    wcscpy(modulePath, L"C:\\Program Files\\ImagePaster\\ImagePaster.exe");
    assert(!IsStartWithWindowsEnabled());
    MigrateStartWithWindows();
    assert(taskWrites == 0); /* initially off stays off */

    run.present = TRUE;
    wcscpy(run.text, L"\"c:\\program files\\imagepaster\\IMAGEPASTER.EXE\"");
    disable_in_task_manager(3);
    MigrateStartWithWindows();
    assert(taskWrites == 0 && run.present); /* respect disabled legacy entry */
    disable_in_task_manager(2);
    assert(IsLegacyStartWithWindowsEnabled()); /* case-insensitive path */
    failWrite = 1;
    MigrateStartWithWindows();
    assert(run.present && !taskPresent && deletes == 0);
    assert(strstr(lastLog, "Could not migrate"));
    failWrite = 0;
    failDelete = 1;
    MigrateStartWithWindows();
    assert(taskEnabled && run.present); /* create before deleting legacy entry */
    failDelete = 0;
    MigrateStartWithWindows(); /* retry cleanup, without recreating the task */
    assert(taskEnabled && !run.present && !approved.present);
    before = taskWrites;
    MigrateStartWithWindows();
    ApplyStartWithWindows(TRUE);
    assert(taskWrites == before && IsStartWithWindowsEnabled());

    ApplyStartWithWindows(FALSE);
    assert(!taskPresent && !IsStartWithWindowsEnabled());
    assert(!strcmp(lastLog, "Start with Windows disabled"));
    assert(SetStartWithWindows(FALSE) == ERROR_SUCCESS);
    ApplyStartWithWindows(TRUE);
    assert(taskEnabled && !strcmp(lastLog, "Start with Windows enabled"));

    taskEnabled = FALSE; /* manually disabled task (or another copy's action) */
    run.present = TRUE;
    before = taskWrites;
    MigrateStartWithWindows();
    ApplyStartWithWindows(FALSE);
    assert(taskWrites == before && taskPresent && run.present);
    ApplyStartWithWindows(TRUE); /* explicit enable repairs the task */
    assert(taskEnabled && !run.present);

    taskPresent = taskEnabled = FALSE;
    run.present = TRUE;
    wcscpy(run.text, L"\"D:\\Tools\\ImagePaster.exe\"");
    before = taskWrites;
    MigrateStartWithWindows();
    ApplyStartWithWindows(FALSE);
    assert(taskWrites == before && run.present); /* preserve another copy */
    ApplyStartWithWindows(TRUE);
    assert(taskEnabled && run.present);

    failRead = 1;
    before = taskWrites;
    ApplyStartWithWindows(FALSE);
    assert(taskWrites == before && taskEnabled && strstr(lastLog, "Could not disable"));
    MigrateStartWithWindows();
    assert(taskWrites == before && strstr(lastLog, "Could not migrate"));
    failRead = 0;
    failWrite = 1;
    ApplyStartWithWindows(FALSE);
    assert(taskEnabled && strstr(lastLog, "Could not disable"));

    for (int i = 0; i < MAX_PATH + 20; i++) modulePath[i] = L'x';
    modulePath[MAX_PATH + 20] = L'\0';
    assert(!IsLegacyStartWithWindowsEnabled()); /* truncated paths never migrate */
    return 0;
}
'''


if __name__ == '__main__':
    unittest.main()
