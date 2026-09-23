"""Check the Startup section in the built settings UI.

Requires: Python playwright and Chromium. Run after make:
    python3 tests/check_startup_ui.py
Uses a recording, inert WebView bridge; nothing is written to the registry.
The Run entry itself is covered by tests/test_startup.py.
"""
from pathlib import Path
import shutil
from playwright.sync_api import sync_playwright, expect

ROOT = Path(__file__).resolve().parents[1]
CONFIG = {
    'titleMatch': 'xshell', 'pasteMethod': 'base64',
    'httpMessageTemplate': 'image at {URL}', 'bindIp': '127.0.0.1',
    'httpPort': 10444, 'httpAllowList': '', 'jpegQuality': 80,
    'imageHistoryLimit': 10, 'imageStorage': 'memory', 'imageStorageDir': '',
    'compatibilityPaste': False, 'screenCaptureEnabled': True,
    'captureGapFill': 'white', 'printScreenHotkeyError': 0,
    'startWithWindows': False, 'autoCheckForUpdates': False,
    'updateCheckPending': False, 'updatePromptPending': False,
    'availableIps': ['127.0.0.1'], 'bindIpAvailable': True,
    'serverStatus': 'Stopped', 'version': '1.0.42',
}
CAPTION = 'Launches in the tray when you sign in to Windows.'

with sync_playwright() as playwright:
    browser = playwright.chromium.launch(
        executable_path=shutil.which('chromium'), args=['--no-sandbox'])
    errors = []

    # A tall screen keeps one column; a short one splits the dialog in two.
    for layout, screen_height in (('one column', 2400), ('two columns', 700)):
        page = browser.new_page(viewport={'width': 900, 'height': 1400},
                                screen={'width': 1920, 'height': screen_height})
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.add_init_script('''
            window.sent = [];
            window.chrome = window.chrome || {};
            window.chrome.webview = {postMessage: message => window.sent.push(JSON.parse(message))};
        ''')

        def init(**changes):
            page.goto((ROOT / 'assets/dist/index.html').as_uri())
            page.wait_for_function('typeof window.onInit === "function"')
            config = {**CONFIG, **changes}
            for key in [key for key, value in changes.items() if value is None]:
                del config[key]  # an older payload without the field
            page.evaluate('(config) => window.onInit({view: "config", config})', config)

        def saved():
            page.get_by_role('button', name='Save', exact=True).click()
            return page.evaluate(
                '() => window.sent.filter(m => m.action === "saveSettings").at(-1)')

        init()
        startup = page.get_by_text('Startup', exact=True)
        updates = page.get_by_text('Updates', exact=True)
        checkbox = page.get_by_role('checkbox', name='Start with Windows', exact=True)
        caption = page.get_by_text(CAPTION, exact=True)
        expect(startup).to_be_visible()
        expect(caption).to_be_visible()
        expect(checkbox).not_to_be_checked()

        # Its own section: heading, then the toggle and caption, then Updates.
        heading, toggle, below, next_heading = (
            startup.bounding_box(), checkbox.bounding_box(),
            caption.bounding_box(), updates.bounding_box())
        assert heading['y'] < toggle['y'] < below['y'] < next_heading['y'], layout
        assert abs(heading['x'] - next_heading['x']) < 1, (layout, heading, next_heading)
        two_columns = page.locator('.grid-cols-2').count() == 1
        assert two_columns == (layout == 'two columns'), layout

        checkbox.check()
        message = saved()
        assert message['startWithWindows'] == 1, message
        assert message['autoCheckForUpdates'] == 0, message

        init(startWithWindows=True)  # the app reports its entry is active
        expect(checkbox).to_be_checked()
        checkbox.uncheck()
        assert saved()['startWithWindows'] == 0

        init(startWithWindows=None)
        expect(checkbox).not_to_be_checked()
        assert saved()['startWithWindows'] == 0
        page.close()

    assert not errors, errors
    browser.close()
    print('PASS: Startup section above Updates in both layouts, wording, '
          'initial state and saved choice')
