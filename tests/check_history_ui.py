"""Check the built History view in Chromium with a recording, inert bridge.

Requires: Python playwright and Chromium. Run after make:
    python3 tests/check_history_ui.py
Covers lazy thumbnail requests, scrolling, paging and live refreshes; the
native paging and thumbnail pipeline is covered by tests/test_history.py.
"""
from pathlib import Path
import shutil
from playwright.sync_api import sync_playwright, expect

ROOT = Path(__file__).resolve().parents[1]
PIXEL = ('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0l'
         'EQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=')
TOTAL = 120
PAGE_SIZE = 50


def token(index):
    return f'{index:064x}'


def entry(index, current=False):
    return {
        'token': token(index), 'current': current, 'width': 1920, 'height': 1080,
        'bytes': 250000 + index, 'capturedAt': 1700000000000 + index,
        'storage': 'disk' if index % 2 else 'memory',
        'path': f'C:\\ImagePaster\\{token(index)}.jpg' if index % 2 else '',
        'url': f'http://192.168.1.5:10444/{token(index)}.jpg',
    }


def page_data(page, first_index=1):
    start = page * PAGE_SIZE
    indexes = range(first_index + start, first_index + min(start + PAGE_SIZE, TOTAL))
    return {
        'pasteMethod': 'http', 'historyLimit': 0, 'total': TOTAL,
        'totalBytes': 30000000, 'page': page, 'pageSize': PAGE_SIZE,
        'entries': [entry(i, current=(i == first_index)) for i in indexes],
    }


with sync_playwright() as playwright:
    browser = playwright.chromium.launch(
        executable_path=shutil.which('chromium'), args=['--no-sandbox'])
    page = browser.new_page(viewport={'width': 700, 'height': 900})
    errors = []
    page.on('pageerror', lambda error: errors.append(str(error)))
    page.add_init_script('''
        window.sent = [];
        window.chrome = window.chrome || {};
        window.chrome.webview = {postMessage: message => window.sent.push(JSON.parse(message))};
    ''')
    page.goto((ROOT / 'assets/dist/index.html').as_uri())
    page.wait_for_function('typeof window.onInit === "function"')

    def requests(action):
        return page.evaluate('(a) => window.sent.filter(m => m.action === a)', action)

    # Empty requests (nothing visible is missing) only clear the native queue.
    consumed = [0]

    def next_request():
        page.wait_for_function(
            '(n) => window.sent.filter(m => m.action === "historyThumbs" && m.tokens)'
            '.length > n', arg=consumed[0])
        wanted = [m['tokens'] for m in requests('historyThumbs') if m['tokens']]
        consumed[0] += 1
        return wanted[consumed[0] - 1].split(',')

    def request_with(wanted_token):
        # Intersection updates can arrive in more than one pass.
        for _ in range(10):
            tokens = next_request()
            if wanted_token in tokens:
                return tokens
        raise AssertionError('never requested ' + wanted_token)

    def send_thumbs(tokens, thumb=PIXEL):
        page.evaluate('(items) => window.onHistoryThumbs(items)',
                      [{'token': t, 'thumb': thumb} for t in tokens])

    rows = page.locator('ul > li')
    previews = page.locator('ul > li img[alt$="preview"]')

    page.evaluate('(history) => window.onInit({view: "history", history})', page_data(0))
    expect(rows).to_have_count(PAGE_SIZE)
    expect(previews).to_have_count(0)  # rows appear before any image exists
    first = next_request()
    assert first[0] == token(1), first
    assert 1 < len(first) < 12, f'only near-visible rows are requested: {len(first)}'
    assert first == [token(i) for i in range(1, 1 + len(first))], 'top first'

    send_thumbs(first[:-1])
    send_thumbs(first[-1:], thumb='')  # the worker could not decode this one
    expect(previews).to_have_count(len(first) - 1)
    expect(rows.nth(len(first) - 1)).to_contain_text('No preview')

    page.locator('ul').evaluate('list => list.parentElement.scrollTo({top: 1e6})')
    bottom = request_with(token(PAGE_SIZE))
    assert token(PAGE_SIZE) in bottom and token(1) not in bottom, bottom
    assert len(bottom) < 12, len(bottom)

    pager = page.get_by_role('navigation', name='History pages')
    expect(pager).to_contain_text(f'1–{PAGE_SIZE} of {TOTAL}')
    expect(pager).to_contain_text('Page 1 of 3')
    expect(pager.get_by_role('button', name='First page')).to_be_disabled()
    expect(pager.get_by_role('button', name='Previous page')).to_be_disabled()
    pager.get_by_role('button', name='Next page').click()
    assert requests('historyPage')[-1] == {'action': 'historyPage', 'page': 1}

    page.evaluate('(history) => window.onHistoryData(history)', page_data(1))
    expect(pager).to_contain_text(f'{PAGE_SIZE + 1}–{2 * PAGE_SIZE} of {TOTAL}')
    expect(pager).to_contain_text('Page 2 of 3')
    assert page.locator('ul').evaluate('list => list.parentElement.scrollTop') == 0
    second = request_with(token(PAGE_SIZE + 1))
    assert second[0] == token(PAGE_SIZE + 1), second
    expect(previews).to_have_count(0)  # previews from page 1 were released

    pager.get_by_role('button', name='Last page').click()
    assert requests('historyPage')[-1] == {'action': 'historyPage', 'page': 2}
    page.evaluate('(history) => window.onHistoryData(history)', page_data(2))
    expect(pager).to_contain_text(f'{2 * PAGE_SIZE + 1}–{TOTAL} of {TOTAL}')
    expect(rows).to_have_count(TOTAL - 2 * PAGE_SIZE)
    expect(pager.get_by_role('button', name='Next page')).to_be_disabled()
    expect(pager.get_by_role('button', name='Last page')).to_be_disabled()

    send_thumbs([token(1)])  # arrives late, for a row no longer on the page
    pager.get_by_role('button', name='First page').click()
    assert requests('historyPage')[-1] == {'action': 'historyPage', 'page': 0}
    page.evaluate('(history) => window.onHistoryData(history)', page_data(0))
    expect(pager).to_contain_text('Page 1 of 3')
    expect(previews).to_have_count(0)  # the late preview was not kept
    visible = request_with(token(1))
    assert visible[0] == token(1), visible
    send_thumbs([token(i) for i in range(1, PAGE_SIZE + 1)])
    expect(previews).to_have_count(PAGE_SIZE)

    # A new clipboard image arrives: the page refreshes in place, known
    # previews stay, and only the new row is requested.
    page.evaluate('(history) => window.onHistoryData(history)', page_data(0, first_index=0))
    expect(rows.first).to_contain_text('Current')
    newest = request_with(token(0))
    assert newest == [token(0)], newest
    expect(previews).to_have_count(PAGE_SIZE - 1)  # the oldest row moved on

    small = page_data(0)
    small['total'] = 3
    small['entries'] = small['entries'][:3]
    page.evaluate('(history) => window.onHistoryData(history)', small)
    expect(pager).to_have_count(0)  # one page needs no pager
    expect(rows).to_have_count(3)

    assert not errors, errors
    browser.close()
    print('PASS: instant metadata rows, visible-only thumbnail requests, scrolling, '
          'failed previews, paging, late previews and live refreshes')
