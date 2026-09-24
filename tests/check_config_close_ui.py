"""Check the built settings close flow using a recording, inert WebView bridge.

Run after make: python3 -B tests/check_config_close_ui.py
Requires Python Playwright and Chromium; never saves Windows settings.
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
    'availableIps': ['127.0.0.1', '192.168.1.10'], 'bindIpAvailable': True,
    'serverStatus': 'Stopped', 'version': '1.0.45',
}


def check_layout(browser, width, screen_height):
    page = browser.new_page(viewport={'width': width, 'height': 900},
                            screen={'width': 1920, 'height': screen_height})
    errors = []
    page.on('pageerror', lambda error: errors.append(str(error)))
    page.add_init_script('''
        window.sent = [];
        window.chrome = window.chrome || {};
        window.chrome.webview = {postMessage: message => window.sent.push(JSON.parse(message))};
    ''')

    def reset(**changes):
        page.goto((ROOT / 'assets/dist/index.html').as_uri())
        page.wait_for_function('window.sent.some(m => m.action === "getInit")')
        page.evaluate('(config) => window.onInit({view: "config", config})',
                      {**CONFIG, **changes})
        page.wait_for_function('window.sent.some(m => m.action === "configReady")')

    def actions():
        return page.evaluate('window.sent.filter(m => ["close", "saveSettings"].includes(m.action))')

    def clear_messages():
        page.evaluate('window.sent = []')

    def native_close():
        page.evaluate('window.onCloseRequested()')

    def button(name, modal=False):
        scope = page.get_by_role('alertdialog') if modal else page
        return scope.get_by_role('button', name=name, exact=True)

    def prompt():
        expect(page.get_by_role('alertdialog')).to_have_count(1)
        expect(page.locator('#save-alert-title')).to_have_text('Unsaved changes')
        expect(page.locator('#save-alert-message')).to_have_text('Save changes before closing?')
        assert actions() == []
        assert page.locator('#titleMatch').evaluate('el => !!el.closest("[inert]")')

    def no_prompt():
        expect(page.locator('#save-alert-title')).to_have_count(0)

    def closed():
        page.wait_for_function('window.sent.some(m => m.action === "close")')
        assert actions() == [{'action': 'close'}]
        no_prompt()

    # Unchanged configurations close through every route without interruption.
    for trigger in (lambda: button('Cancel').click(), native_close,
                    lambda: page.keyboard.press('Escape')):
        reset()
        trigger()
        closed()
    assert (page.locator('.grid-cols-2').count() == 1) == (screen_height == 900)

    # Every saved field is tracked, and reverting it makes the form clean again.
    edits = [
        ('titleMatch', 'fill', 'terminal', 'xshell'),
        ('pasteMethod', 'select_option', 'http', 'base64'),
        ('httpMessageTemplate', 'fill', 'new image at {URL}', 'image at {URL}'),
        ('bindIp', 'select_option', '192.168.1.10', '127.0.0.1'),
        ('httpPort', 'fill', '10445', '10444'),
        ('httpAllowList', 'fill', '192.168.1.0/24', ''),
        ('jpegQuality', 'fill', '90', '80'),
        ('imageHistoryLimit', 'fill', '100', '10'),
        ('unlimitedImageHistory', 'set_checked', True, False),
        ('imageStorage', 'select_option', 'disk', 'memory'),
        ('compatibilityPaste', 'set_checked', True, False),
        ('screenCaptureEnabled', 'set_checked', False, True),
        ('captureGapFill', 'select_option', 'black', 'white'),
        ('startWithWindows', 'set_checked', True, False),
        ('autoCheckForUpdates', 'set_checked', True, False),
    ]
    for field, method, changed, original in edits:
        reset()
        getattr(page.locator('#' + field), method)(changed)
        native_close()
        prompt()
        native_close()  # Repeated X/Alt+F4 cannot bypass the question.
        prompt()
        button('Keep editing', True).click()
        no_prompt()
        assert actions() == []
        getattr(page.locator('#' + field), method)(original)
        button('Cancel').click()
        closed()

    # Custom IP and unlimited history are represented by multiple form controls.
    reset(bindIp='10.0.0.1', bindIpAvailable=False)
    page.locator('#customIp').fill('10.0.0.2')
    native_close()
    prompt()
    button('Keep editing', True).click()
    page.locator('#customIp').fill('10.0.0.1')
    native_close()
    closed()
    reset()
    page.locator('#bindIp').select_option('other')
    page.locator('#customIp').fill('10.0.0.2')
    native_close()
    prompt()
    button('Keep editing', True).click()
    page.locator('#bindIp').select_option(CONFIG['bindIp'])
    native_close()
    closed()
    reset(imageHistoryLimit=0)
    page.locator('#unlimitedImageHistory').uncheck()
    page.locator('#imageHistoryLimit').fill('22')
    native_close()
    prompt()
    button('Keep editing', True).click()
    page.locator('#unlimitedImageHistory').check()
    native_close()
    closed()

    # Modal focus, keyboard wrapping, background blocking, Escape, and Discard.
    reset()
    page.locator('#startWithWindows').check()
    button('Cancel').click()
    prompt()
    expect(button('Keep editing', True)).to_be_focused()
    page.keyboard.press('Shift+Tab')
    expect(button('Save', True)).to_be_focused()
    page.keyboard.press('Tab')
    expect(button('Keep editing', True)).to_be_focused()
    page.mouse.click(8, 8)
    prompt()
    expect(button('Keep editing', True)).to_be_focused()
    page.keyboard.press('Escape')
    no_prompt()
    expect(page.locator('#startWithWindows')).to_be_checked()
    expect(button('Cancel')).to_be_focused()
    page.keyboard.press('Escape')
    prompt()
    button('Discard', True).click()
    page.wait_for_function('window.sent.some(m => m.action === "close")')
    assert actions() == [{'action': 'close'}]

    # Prompt Save uses the normal save payload, with no premature close message.
    payloads = []
    for through_prompt in (False, True):
        reset()
        page.locator('#startWithWindows').check()
        page.locator('#imageHistoryLimit').fill('250')
        if through_prompt:
            native_close()
            prompt()
        button('Save', through_prompt).click()
        page.wait_for_function('window.sent.some(m => m.action === "saveSettings")')
        payloads.append(actions())
    assert payloads[0] == payloads[1]
    assert len(payloads[0]) == 1 and payloads[0][0]['action'] == 'saveSettings'
    assert payloads[0][0]['startWithWindows'] == 1
    assert payloads[0][0]['imageHistoryLimit'] == 250

    # Validation and native save errors return to editing with all edits retained.
    for field, value, error in (
        ('httpMessageTemplate', '', 'HTTP paste message must include {URL}.'),
        ('httpPort', '0', 'HTTP port must be between 1 and 65535.'),
        ('httpAllowList', 'invalid', 'Allowed clients must be IPv4 addresses'),
        ('jpegQuality', '101', 'JPEG quality must be between 0 and 100.'),
        ('imageHistoryLimit', '0', 'Image history must be Unlimited or between 1 and 1000.'),
        ('customIp', 'invalid', 'Enter a valid IPv4 bind address.'),
    ):
        reset()
        if field == 'customIp':
            page.locator('#bindIp').select_option('other')
        page.locator('#' + field).fill(value)
        native_close()
        prompt()
        button('Save', True).click()
        no_prompt()
        expect(page.get_by_text(error, exact=False)).to_be_visible()
        expect(page.locator('#' + field)).to_be_focused()
        expect(page.locator('#' + field)).to_have_value(value)
        assert actions() == []
        native_close()
        prompt()
    reset()
    page.locator('#startWithWindows').check()
    native_close()
    button('Save', True).click()
    page.wait_for_function('window.sent.some(m => m.action === "saveSettings")')
    page.evaluate('window.onSaveResult({ok:false, message:"Save failed"})')
    no_prompt()
    expect(page.get_by_text('Save failed', exact=True)).to_be_visible()
    expect(page.locator('#startWithWindows')).to_be_checked()
    clear_messages()
    native_close()
    prompt()

    # Live status is not an edit; arriving updates cannot replace the close prompt.
    reset()
    page.evaluate('window.onPrintScreenStatus({error:1409})')
    native_close()
    closed()
    reset()
    page.locator('#compatibilityPaste').check()
    native_close()
    prompt()
    overlay = page.get_by_role('alertdialog').evaluate('''el => {
        const style = getComputedStyle(el.parentElement);
        return {color: style.backgroundColor, position: style.position,
                width: el.parentElement.clientWidth, height: el.parentElement.clientHeight};
    }''')
    assert overlay == {'color': 'rgba(0, 0, 0, 0.35)', 'position': 'fixed',
                       'width': width, 'height': 900}, overlay
    page.evaluate('''window.onUpdateResult({status:'newer', title:'Update available',
        message:'A newer version is ready to install.', currentVersion:'1.0.44',
        remoteVersion:'1.0.45', automatic:true})''')
    prompt()
    button('Keep editing', True).click()
    expect(page.locator('#update-alert-title')).to_be_visible()
    assert page.get_by_role('alertdialog').evaluate(
        'el => getComputedStyle(el.parentElement).backgroundColor') == overlay['color']
    native_close()
    prompt()
    page.keyboard.press('Escape')
    expect(page.locator('#update-alert-title')).to_be_visible()
    button('Cancel', True).click()
    expect(page.get_by_role('alertdialog')).to_have_count(0)
    expect(page.locator('#compatibilityPaste')).to_be_checked()
    assert actions() == []
    assert not errors, errors
    page.close()


if __name__ == '__main__':
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            executable_path=shutil.which('chromium'), args=['--no-sandbox'])
        try:
            for width, height in ((560, 2400), (900, 900)):
                check_layout(browser, width, height)
        finally:
            browser.close()
    print('PASS: Unsaved edits in both layouts, all close routes, Save/Discard/Keep editing, '
          'keyboard, overlay, validation, and update overlap')
