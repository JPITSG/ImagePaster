"""Exercise the native close gate; check_config_close_ui.py checks edit decisions."""
from pathlib import Path
import subprocess
import tempfile
import unittest

from test_updater import function


class ConfigCloseTests(unittest.TestCase):
    def test_native_close_waits_for_the_config_decision(self):
        source = PREFIX + function('RequestConfigClose') + SCENARIO
        with tempfile.TemporaryDirectory(prefix='imagepaster-close-') as directory:
            path = Path(directory)
            (path / 'test.c').write_text(source)
            subprocess.run(['cc', '-std=c11', '-Wall', '-Wextra', '-Werror',
                            str(path / 'test.c'), '-o', str(path / 'test')], check=True)
            subprocess.run([str(path / 'test')], check=True, timeout=10)

    def test_close_gate_is_wired_before_teardown_and_after_save(self):
        proc = function('WebViewWndProc')
        close = proc.split('case WM_CLOSE:', 1)[1].split('case WM_DESTROY:', 1)[0]
        gate = close.index('if (RequestConfigClose()) return 0;')
        self.assertLess(gate, close.index('->Close('))
        self.assertLess(gate, close.index('DestroyWindow('))
        self.assertIn('g_configCloseApproved = FALSE;',
                      proc.split('case WM_DESTROY:', 1)[1])
        self.assertIn('g_configCloseApproved = FALSE;', function('ShowWebViewDialog'))

        handler = function('MsgReceived_Invoke')
        save = handler.split('strcmp(action, "saveSettings")', 1)[1].split(
            'strcmp(action, "close")', 1)[0]
        approval = save.index('g_configCloseApproved = TRUE;')
        self.assertLess(save.index('SaveConfigToRegistry();'), approval)
        self.assertLess(save.rindex('window.onSaveResult'), approval)
        self.assertLess(approval, save.index('PostMessage(g_webviewHwnd, WM_CLOSE'))
        close = handler.split('strcmp(action, "close")', 1)[1].split(
            'strcmp(action, "clearLog")', 1)[0]
        self.assertLess(close.index('g_configCloseApproved = TRUE;'),
                        close.index('PostMessage(g_webviewHwnd, WM_CLOSE'))

        shutdown = function('WndProc').split('case ID_TRAY_EXIT:', 1)[1]
        self.assertLess(shutdown.index('g_configCloseApproved = TRUE;'),
                        shutdown.index('SendMessage(g_webviewHwnd, WM_CLOSE'))


PREFIX = r'''
#include <assert.h>
#include <string.h>
#include <wchar.h>
typedef int BOOL;
#define TRUE 1
#define FALSE 0
static char g_pendingView[16] = "config";
static BOOL g_configViewReady, g_configCloseApproved, g_updateInstallReady;
static void *g_webviewView;
static int requests;
static void webview_execute_script(const wchar_t *script) {
    assert(wcscmp(script, L"window.onCloseRequested()") == 0);
    requests++;
}
'''

SCENARIO = r'''
int main(void) {
    assert(!RequestConfigClose());  // Loading or failed WebView can still close.
    g_configViewReady = TRUE;
    assert(!RequestConfigClose());  // No WebView, even if marked ready.
    g_configViewReady = FALSE;
    g_webviewView = (void *)1;
    assert(!RequestConfigClose());
    g_configViewReady = TRUE;
    assert(RequestConfigClose());
    assert(RequestConfigClose());  // Repeated X does not bypass the prompt.
    assert(requests == 2);
    g_configCloseApproved = TRUE;
    assert(!RequestConfigClose());  // Save, discard, unchanged, or app shutdown.
    g_configCloseApproved = FALSE;
    g_updateInstallReady = TRUE;
    assert(!RequestConfigClose());  // Updater must finish its handoff.
    g_updateInstallReady = FALSE;
    strcpy(g_pendingView, "history");
    assert(!RequestConfigClose());
    strcpy(g_pendingView, "log");
    assert(!RequestConfigClose());
    assert(requests == 2);
    return 0;
}
'''


if __name__ == '__main__':
    unittest.main()
