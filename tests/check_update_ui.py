"""Check the built UI in Chromium with a recording, inert WebView bridge.

Requires: Python playwright and Chromium. Run after make:
    python3 tests/check_update_ui.py
No native updater, network download, or app restart is invoked.
"""
from pathlib import Path
import re
import shutil
from playwright.sync_api import sync_playwright, expect

ROOT = Path(__file__).resolve().parents[1]
CONFIG = {
    'titleMatch': 'test', 'pasteMethod': 'base64', 'httpMessageTemplate': '{url}',
    'bindIp': '127.0.0.1', 'httpPort': 8080, 'httpAllowList': '',
    'jpegQuality': 80, 'imageHistoryLimit': 10, 'imageStorage': 'memory',
    'imageStorageDir': '', 'compatibilityPaste': False,
    'screenCaptureEnabled': False, 'captureGapFill': 'white',
    'autoCheckForUpdates': False, 'updateCheckPending': False,
    'updatePromptPending': False, 'availableIps': ['127.0.0.1'],
    'bindIpAvailable': True, 'serverStatus': 'Stopped', 'version': '1.0.35',
}

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(
        executable_path=shutil.which('chromium'), args=['--no-sandbox'])
    page = browser.new_page(viewport={'width': 900, 'height': 1000})
    errors = []
    page.on('pageerror', lambda error: errors.append(str(error)))
    page.add_init_script('''
        window.sent = [];
        window.chrome = window.chrome || {};
        window.chrome.webview = {postMessage: message => window.sent.push(JSON.parse(message))};
    ''')
    page.goto((ROOT / 'assets/dist/index.html').as_uri())
    page.wait_for_function('typeof window.onInit === "function"')
    page.evaluate('(config) => window.onInit({view: "config", config})', CONFIG)

    def result(status='newer', automatic=False):
        page.evaluate('(result) => window.onUpdateResult(result)', {
            'status': status, 'automatic': automatic, 'title': 'Update result',
            'message': 'Test result', 'currentVersion': '1.0.35', 'remoteVersion': '1.0.36',
        })

    def last_message(action):
        return page.evaluate('(action) => window.sent.filter(m => m.action === action).at(-1)', action)

    page.get_by_role('button', name='Update', exact=True).click()
    busy = page.get_by_role('button', name='Stop update check and download')
    expect(busy).to_have_text('Checking...')
    expect(busy).to_be_enabled()
    expect(busy).to_have_class(re.compile('bg-red-'))
    for speed, text in [(100, '100'), (12345, '12345'), (99.5, '100'), (0, '0')]:
        page.evaluate('(speed) => window.onUpdateProgress({kilobytesPerSecond: speed})', speed)
        expect(busy).to_have_text(f'Checking ({text}kb/s)...')
    busy.click()
    expect(busy).to_have_text('Stopping...')
    expect(busy).to_be_disabled()
    assert last_message('cancelUpdateCheck') == {'action': 'cancelUpdateCheck'}
    result('cancelled')
    expect(page.get_by_role('button', name='Update', exact=True)).to_be_enabled()

    result()
    modal = page.get_by_role('alertdialog')
    checkbox = modal.get_by_role('checkbox', name='Reopen settings after update', exact=True)
    expect(checkbox).not_to_be_checked()
    checkbox.check()
    modal.get_by_role('button', name='Cancel', exact=True).click()
    result()
    expect(checkbox).not_to_be_checked()
    modal.get_by_role('button', name='Update', exact=True).click()
    assert last_message('installUpdate') == {'action': 'installUpdate', 'reopenSettingsAfterUpdate': 0}
    expect(modal.get_by_role('button', name='Starting...')).to_be_disabled()
    expect(checkbox).to_be_disabled()

    result('error')  # Includes UAC cancellation / helper startup failure.
    expect(modal.get_by_role('checkbox')).to_have_count(0)
    modal.get_by_role('button', name='OK').click()
    result()
    expect(checkbox).not_to_be_checked()
    checkbox.check()
    modal.get_by_role('button', name='Update', exact=True).click()
    assert last_message('installUpdate')['reopenSettingsAfterUpdate'] == 1

    result('same')
    expect(checkbox).not_to_be_checked()
    checkbox.check()
    modal.get_by_role('button', name='Force update').click()
    assert last_message('installUpdate')['reopenSettingsAfterUpdate'] == 1

    result(automatic=True)
    expect(checkbox).not_to_be_checked()
    checkbox.check()
    modal.get_by_role('button', name='Ignore this version').click()
    result(automatic=True)
    expect(checkbox).not_to_be_checked()
    modal.get_by_role('button', name='Cancel', exact=True).click()
    result('older')
    expect(modal.get_by_role('checkbox')).to_have_count(0)
    expect(modal.get_by_role('button', name='Update', exact=True)).to_have_count(0)
    assert not errors, errors
    browser.close()
    print('PASS: speed labels, busy/cancel states, checkbox defaults/resets, install payloads, force update and downgrade UI')
