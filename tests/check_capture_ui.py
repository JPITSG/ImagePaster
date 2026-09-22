"""Check the Print Screen conflict notice in the built settings UI.

Requires: Python playwright and Chromium. Run after make:
    python3 tests/check_capture_ui.py
Uses a recording, inert WebView bridge; no hotkey is registered.
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
    'screenCaptureEnabled': True, 'captureGapFill': 'white',
    'printScreenHotkeyError': 1409,
    'autoCheckForUpdates': False, 'updateCheckPending': False,
    'updatePromptPending': False, 'availableIps': ['127.0.0.1'],
    'bindIpAvailable': True, 'serverStatus': 'Stopped', 'version': '1.0.41',
}
CLAIMED = 'Another program has claimed the Print Screen key.'
IMPACT = "Capture still works, but keyboard-sharing tools can't send the key to another computer."

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

    def init(**changes):
        # Every settings dialog is a fresh page.
        page.goto((ROOT / 'assets/dist/index.html').as_uri())
        page.wait_for_function('typeof window.onInit === "function"')
        page.evaluate('(config) => window.onInit({view: "config", config})', {**CONFIG, **changes})

    def status(error):
        page.evaluate('(error) => window.onPrintScreenStatus({error})', error)

    init()
    notice = page.locator('p', has_text='Print Screen key')
    checkbox = page.get_by_role('checkbox', name='Enable interactive Print Screen capture')
    expect(notice).to_be_visible()
    expect(notice).to_have_text(f'{CLAIMED} {IMPACT}')
    expect(notice).to_have_class(re.compile(r'\btext-red-600\b'))

    # Under the option's own description, in its text column, above the gap fill.
    description = page.get_by_text('Opens a multi-monitor selection overlay.')
    gap_fill = page.get_by_text('Multi-region Gap Fill', exact=True)
    below, box, above = description.bounding_box(), notice.bounding_box(), gap_fill.bounding_box()
    assert below['y'] + below['height'] <= box['y'] < box['y'] + box['height'] <= above['y'], (below, box, above)
    assert abs(box['x'] - below['x']) < 1, (below, box)
    assert box['width'] <= below['width'] + 1, (below, box)
    # As much space below the notice as above it, text to text.
    gaps = page.evaluate('''() => {
        const description = [...document.querySelectorAll('p')]
            .find(p => p.textContent.startsWith('Opens a multi-monitor selection overlay.'));
        const notice = [...document.querySelectorAll('p')]
            .find(p => p.textContent.includes('Print Screen key'));
        const label = [...document.querySelectorAll('label')]
            .find(l => l.textContent === 'Multi-region Gap Fill');
        const top = notice.getBoundingClientRect().top + parseFloat(getComputedStyle(notice).paddingTop);
        return {
            above: top - description.getBoundingClientRect().bottom,
            below: label.getBoundingClientRect().top - notice.getBoundingClientRect().bottom,
        };
    }''')
    assert abs(gaps['above'] - gaps['below']) <= 1.5, gaps

    # Only while the option is ticked: the key matters only for capture.
    checkbox.uncheck()
    expect(notice).to_be_hidden()
    checkbox.check()
    expect(notice).to_be_visible()

    # Live: the other program releases the key, claims it again, or it fails.
    status(0)
    expect(notice).to_be_hidden()
    status(1409)
    expect(notice).to_have_text(f'{CLAIMED} {IMPACT}')
    status(5)
    expect(notice).to_have_text(f'The Print Screen key could not be claimed (error 5). {IMPACT}')

    # Capture off: the probe's result shows as soon as the option is ticked.
    init(screenCaptureEnabled=False)
    expect(notice).to_be_hidden()
    checkbox.check()
    expect(notice).to_have_text(f'{CLAIMED} {IMPACT}')

    # No conflict, no notice.
    init(printScreenHotkeyError=0)
    expect(notice).to_be_hidden()
    init(printScreenHotkeyError=None)
    expect(notice).to_be_hidden()

    assert not errors, errors
    browser.close()

print('Capture notice UI checks passed.')
