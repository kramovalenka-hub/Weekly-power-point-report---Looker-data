import argparse, shutil, sys
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# Find a chart by its Looker heading and identify the enclosing visualization.
FIND_TILE_JS = """
(config) => {
  // Normalize punctuation and Unicode so small title-format differences still match.
  const normalize = value => (value || '')
    .normalize('NFKC')
    .replace(/[^\\p{L}\\p{N}]+/gu, ' ')
    .trim()
    .toUpperCase();
  const want = normalize(config.text);
  // Looker versions and chart types expose headings through different semantics.
  const selectors = [
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6', '[role="heading"]',
    '[data-testid*="title" i]', '[data-testid*="header" i]',
    '[data-qa*="title" i]', '[data-qa*="header" i]',
    '[class*="title" i]', '[class*="header" i]', '[aria-label]', '[title]'
  ];
  let anchors = Array.from(document.querySelectorAll(selectors.join(',')))
    .filter(el => [
      el.innerText,
      el.textContent,
      el.getAttribute('aria-label'),
      el.getAttribute('title')
    ].some(value => normalize(value) === want))
    .filter(el => !el.closest('svg, canvas, table, [role="grid"], [role="table"]'));
  if (!config.headerOnly) {
    // Legacy mode also checks chart text; header-only mode deliberately skips this.
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      if (normalize(node.textContent) === want) anchors.push(node.parentElement);
    }
  }
  // Walk up from the matching title until its enclosing chart can be captured.
  const visualization = 'table, canvas, svg, iframe, [role="grid"], [role="table"], [class*="visualization" i], [class*="vis-container" i], [data-testid*="visual" i], [data-qa*="viz" i]';
  for (const anchor of anchors) {
    let el = anchor;
    while (el && el !== document.body) {
      const r = el.getBoundingClientRect();
      if (r.height > 200 && r.width > 300 && el.querySelector(visualization)) {
        el.setAttribute('data-smt-capture', '1');
        return true;
      }
      el = el.parentElement;
    }
  }
  return false;
}
"""

# Command-line options supplied by the notebook or weekly deck generator.
ap = argparse.ArgumentParser()
ap.add_argument("--url", help="Dashboard URL (single-chart mode).")
ap.add_argument("--profile", required=True)
ap.add_argument("--channel", default="msedge")
ap.add_argument("--out")
ap.add_argument("--tile")
ap.add_argument("--header-only", action="store_true")
ap.add_argument("--wait", type=int, default=15)
ap.add_argument("--login", action="store_true")
ap.add_argument("--headed", action="store_true")
ap.add_argument("--jobs", help="JSON file listing every chart to capture in ONE browser session.")
a = ap.parse_args()
if not a.jobs and not a.url:
    ap.error("--url is required unless --jobs is given")

CLEAR_MARK_JS = "() => document.querySelectorAll('[data-smt-capture]').forEach(el => el.removeAttribute('data-smt-capture'))"

# Looker marks a tile (and the page loader) with aria-busy="true" and shows a spinner
# (role="progressbar") while its query runs. No busy elements = everything has loaded.
BUSY_COUNT_JS = """() => Array.from(document.querySelectorAll('[aria-busy="true"], [role="progressbar"]'))
  .filter(el => { const r = el.getBoundingClientRect(); return r.width > 0 && r.height > 0; }).length"""

# True while the left navigation sidebar (Home / Explore / Boards ...) is showing.
SIDEBAR_OPEN_JS = """() => Array.from(document.querySelectorAll('a, button, [role="treeitem"], li, span'))
  .some(el => { const t = (el.innerText || '').trim();
                const r = el.getBoundingClientRect();
                return (t === 'Explore' || t === 'Recently viewed') && r.width > 0 && r.left < 240 && r.top > 60; })"""

# Scroll Looker's dashboard container one screen at a time so lazily loaded tiles start their queries.
SCROLL_STEP_JS = """(y) => { const w = document.getElementById('dashboard-layout-wrapper') || document.scrollingElement;
  w.scrollTop = y; return w.scrollHeight; }"""


def wait_until_loaded(page, timeout_seconds=300, quiet_ms=3000):
    """Wait until no tile is loading for quiet_ms in a row. Returns False on timeout."""
    deadline = page.evaluate("Date.now()") + timeout_seconds * 1000
    quiet_since = None
    while page.evaluate("Date.now()") < deadline:
        if page.evaluate(BUSY_COUNT_JS) == 0:
            now = page.evaluate("Date.now()")
            quiet_since = quiet_since or now
            if now - quiet_since >= quiet_ms:
                return True
        else:
            quiet_since = None
        page.wait_for_timeout(500)
    return False


def close_sidebar(page):
    """Collapse Looker's left navigation by clicking the main menu (hamburger) button."""
    if not page.evaluate(SIDEBAR_OPEN_JS):
        return
    menu = page.locator("button.pendo-main-menu-toggle").first
    if menu.count() == 0:
        print("  Note: Looker menu button not found - sidebar left open.", flush=True)
        return
    menu.click()
    page.mouse.move(1400, 5)  # move off the button so its "Main menu" tooltip doesn't show in screenshots
    page.wait_for_timeout(1000)
    if page.evaluate(SIDEBAR_OPEN_JS):
        print("  Note: Looker sidebar still open after clicking the menu.", flush=True)


def open_dashboard(page, url, wait_seconds):
    """Open a dashboard once, close the sidebar, and wait until every chart has finished loading."""
    page.goto(url, wait_until="domcontentloaded", timeout=120_000)
    try:
        page.wait_for_load_state("networkidle", timeout=120_000)
    except PWTimeout:
        pass
    host = url.split("/")[2]
    if host not in page.url or "/login" in page.url:
        return False
    # Minimum settle time, then wait for the loading spinners to clear.
    page.wait_for_timeout(wait_seconds * 1000)
    wait_until_loaded(page)
    close_sidebar(page)
    # Scroll through the whole dashboard so tiles further down load too, then wait again.
    y, height = 0, page.evaluate(SCROLL_STEP_JS, 0)
    while y < height:
        y += 1500
        height = page.evaluate(SCROLL_STEP_JS, y)
        page.wait_for_timeout(500)
        wait_until_loaded(page, timeout_seconds=120)
    page.evaluate(SCROLL_STEP_JS, 0)
    if not wait_until_loaded(page):
        print("  Warning: some Looker charts were still loading after 5 minutes.", flush=True)
    return True


def capture_chart(page, tile, out, header_only):
    """Screenshot one chart on the already-open dashboard. Returns False if not found."""
    page.evaluate(CLEAR_MARK_JS)
    if not page.evaluate(FIND_TILE_JS, {"text": tile, "headerOnly": header_only}):
        return False
    loc = page.locator('[data-smt-capture="1"]').first
    loc.scroll_into_view_if_needed()
    page.wait_for_timeout(1500)  # let charts further down the page finish drawing
    wait_until_loaded(page, timeout_seconds=120)
    # Park the mouse in an empty part of Looker's top bar so no tooltip or hover icons
    # (explore / alert / menu) appear on the chart in the screenshot.
    page.mouse.move(1400, 5)
    page.wait_for_timeout(700)
    out.parent.mkdir(parents=True, exist_ok=True)
    loc.screenshot(path=str(out))
    return True


if a.jobs:
    # Batch mode: open the browser ONCE, open each dashboard ONCE, and take every chart from it.
    import json
    jobs = json.loads(Path(a.jobs).read_text(encoding="utf-8"))
    by_url = {}
    for job in jobs:
        by_url.setdefault(job["url"], []).append(job)
    missing = []
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            a.profile, channel=a.channel, headless=not a.headed,
            viewport={"width": 1920, "height": 2400}, device_scale_factor=2,
            ignore_default_args=["--enable-automation"],
            args=["--disable-blink-features=AutomationControlled"])
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        for url, url_jobs in by_url.items():
            print(f"Opening dashboard ({len(url_jobs)} chart(s)): {url}", flush=True)
            if not open_dashboard(page, url, a.wait):
                ctx.close()
                sys.exit("NOT LOGGED IN - run the login step again.")
            first_out = Path(url_jobs[0]["out"])
            full = first_out.with_name(first_out.stem + "_full_page.png")
            page.screenshot(path=str(full), full_page=True)
            for job in url_jobs:
                out = Path(job["out"])
                if capture_chart(page, job["tile"], out, job.get("header_only", True)):
                    print(f"  Saved {out.name}  <- {job['tile']}", flush=True)
                else:
                    print(f"  LOOKER TILE NOT FOUND: '{job['tile']}'", flush=True)
                    missing.append(job["tile"])
        ctx.close()
    if missing:
        sys.exit("Charts not found on the dashboard: " + "; ".join(missing))
    print(f"All {len(jobs)} screenshot(s) saved.", flush=True)
    sys.exit(0)

# Reuse a persistent browser profile so the Looker sign-in survives between runs.
with sync_playwright() as p:
    ctx = p.chromium.launch_persistent_context(
        a.profile, channel=a.channel, headless=not (a.headed or a.login),
        viewport={"width": 1920, "height": 2400}, device_scale_factor=2,
        ignore_default_args=["--enable-automation"],
        args=["--disable-blink-features=AutomationControlled"])
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(a.url, wait_until="domcontentloaded", timeout=120_000)

    if a.login:
        # Keep a visible browser open so the user can complete interactive sign-in.
        print("Log in in the browser window. When you can see the dashboard, close the browser window to save the session.", flush=True)
        try:
            page.wait_for_event("close", timeout=300_000)
        except Exception:
            print("Browser stayed open; login session remains active for this profile.", flush=True)
        try: ctx.close()
        except Exception: pass
        print("Login saved.")
        sys.exit(0)

    # Looker may keep network requests open, so the explicit wait below is the fallback.
    try:
        page.wait_for_load_state("networkidle", timeout=120_000)
    except PWTimeout:
        pass
    page.wait_for_timeout(a.wait * 1000)

    # Stop before capture if the saved session expired or was redirected to login.
    host = a.url.split("/")[2]
    if host not in page.url or "/login" in page.url:
        ctx.close()
        sys.exit("NOT LOGGED IN - run the login cell again.")

    # Save a full-page reference image even when the requested output is one chart.
    out = Path(a.out); out.parent.mkdir(parents=True, exist_ok=True)
    full = out.with_name(out.stem + "_full_page.png")
    page.screenshot(path=str(full), full_page=True)
    print("Full dashboard:", full)
    if a.tile:
        # A missing requested chart is an error; never substitute a dashboard screenshot.
        if not page.evaluate(FIND_TILE_JS, {"text": a.tile, "headerOnly": a.header_only}):
            ctx.close()
            mode = "header" if a.header_only else "tile text"
            sys.exit(
                f"LOOKER TILE NOT FOUND: '{a.tile}' (searched {mode} elements on {a.url}). "
                "Refusing to save the full dashboard as a chart screenshot."
            )
        loc = page.locator('[data-smt-capture="1"]').first
        loc.scroll_into_view_if_needed()
        loc.screenshot(path=str(out))
        print("Captured tile:", out)
    else:
        # With no chart title, the requested output is the full dashboard screenshot.
        shutil.copy(full, out)
    ctx.close()
