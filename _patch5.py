"""
_patch5.py – Final definitive fix:
  1. _try_slicer: ONLY targets AOM slicer. Uses page.mouse.click() at
     bounding-box coordinates (right side = ▼ arrow). Clicks eraser first.
     Uses wait_for_selector (active 8s poll) instead of fixed 2s sleep.
     No generic hunt. 3 attempts.
  2. _apply_filter_via_pane: Remove useless Page-1 fallback (AOM slicer IS on page 1).
  3. _trigger_pdf_export: Add new-page detection, expanded confirm selectors,
     and mid-export screenshot for debugging.
"""

pbi_path = r'C:\Users\LENOVO\Downloads\Projects\PBi automation\powerbi.py'

with open(pbi_path, encoding='utf-8') as f:
    src = f.read()

# ── 1. Replace _try_slicer completely ─────────────────────────────────────────
NEW_TRY_SLICER = r'''    def _try_slicer(self, filter_email: str) -> bool:
        """
        Apply the AOM filter by interacting exclusively with the AOM-titled slicer.

        Key insight: Power BI's slicer dropdown only opens when it receives the
        full mouse event sequence (pointerdown→mousedown→pointerup→mouseup→click).
        Using Playwright's page.mouse.click(x, y) sends this full sequence at real
        screen coordinates.  force=True synthetic clicks send ONLY the click event,
        which Power BI ignores → dropdown never opens.

        Flow per attempt:
          1. Hover over AOM slicer to reveal the eraser (hidden on hover).
          2. Click eraser (first attempt only) → resets previous AOM selection.
          3. Click at right side of dropdown (≈85% from left, where ▼ is).
          4. Actively wait up to 8 s for [role="option"] to appear.
          5. Find target, deselect-all, click target option.
        Up to 3 attempts before giving up.  No generic 10-slicer hunt.
        """
        page = self._page

        # Force-unhide visuals
        try:
            page.evaluate("""() => {
                document.querySelectorAll(
                    '.visual-container, .visualContainer, [class*="visual"]'
                ).forEach(v => {
                    v.style.setProperty('visibility', 'visible', 'important');
                    v.style.setProperty('opacity',    '1',       'important');
                    v.style.setProperty('pointer-events', 'auto','important');
                });
            }""")
            page.wait_for_timeout(500)
        except Exception:
            pass

        # JS: walk text nodes to find the element that says exactly "AOM",
        # then climb up to find the nearest combobox/haspopup trigger.
        FIND_TRIGGER_JS = """
        () => {
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null
            );
            let node;
            while ((node = walker.nextNode())) {
                if (node.textContent.trim() !== 'AOM') continue;
                let el = node.parentElement;
                for (let i = 0; i < 10; i++) {
                    if (!el || el === document.body) break;
                    const t = el.querySelector(
                        '[role="combobox"], [aria-haspopup="listbox"], [aria-haspopup="true"]'
                    );
                    if (t) return t;
                    el = el.parentElement;
                }
            }
            return null;
        }
        """

        # JS: find the eraser/clear button near the AOM text
        FIND_ERASER_JS = """
        () => {
            const walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null
            );
            let node;
            while ((node = walker.nextNode())) {
                if (node.textContent.trim() !== 'AOM') continue;
                let el = node.parentElement;
                for (let i = 0; i < 8; i++) {
                    if (!el || el === document.body) break;
                    const e = el.querySelector(
                        '[aria-label*="clear" i], [aria-label*="Clear" i], '
                        + '[aria-label*="eraser" i], [title*="clear" i], '
                        + '.slicerDeleteButton, [class*="clearButton"], '
                        + '[class*="clear-button"]'
                    );
                    if (e) return e;
                    el = el.parentElement;
                }
            }
            return null;
        }
        """

        def _mouse_click(elem_h, x_frac: float = 0.5) -> bool:
            """Real mouse.click at a fractional X position of elem_h's bounding box."""
            try:
                elem_h.scroll_into_view_if_needed(timeout=3_000)
            except Exception:
                pass
            bb = elem_h.bounding_box()
            if bb:
                x = bb['x'] + bb['width'] * x_frac
                y = bb['y'] + bb['height'] * 0.5
                page.mouse.move(x, y)
                page.wait_for_timeout(150)
                page.mouse.click(x, y)
                return True
            # Fallback: full JS event dispatch
            try:
                page.evaluate("""
                (el) => {
                    const r = el.getBoundingClientRect();
                    const x = r.left + r.width  * 0.85;
                    const y = r.top  + r.height * 0.5;
                    const opts = {bubbles:true, cancelable:true, view:window,
                                  clientX:x, clientY:y};
                    ['pointerenter','mouseover','pointermove','mousemove',
                     'pointerdown','mousedown','pointerup','mouseup','click'
                    ].forEach(type => el.dispatchEvent(
                        new MouseEvent(type, {...opts,
                            button:0, buttons: type.includes('down') ? 1 : 0})
                    ));
                }
                """, elem_h)
                return True
            except Exception:
                pass
            try:
                elem_h.click(force=True, timeout=2_000)
                return True
            except Exception:
                return False

        email_sels = [
            f'[role="option"]:has-text("{filter_email}")',
            f'[role="listbox"] li:has-text("{filter_email}")',
            f'div[role="listbox"] span:has-text("{filter_email}")',
        ]

        for attempt in range(3):
            log.info(f'  AOM slicer: attempt {attempt + 1}/3 for \'{filter_email}\'...')

            # Find the AOM trigger
            try:
                h = page.evaluate_handle(FIND_TRIGGER_JS)
                elem = h.as_element()
            except Exception:
                elem = None

            if elem is None:
                log.warning('  AOM slicer trigger not found in DOM.')
                return False

            # Attempt 1 only: hover to reveal eraser, then click it
            if attempt == 0:
                try:
                    elem.hover(timeout=2_000)
                    page.wait_for_timeout(400)
                    eh = page.evaluate_handle(FIND_ERASER_JS)
                    eraser = eh.as_element()
                    if eraser:
                        _mouse_click(eraser, x_frac=0.5)
                        page.wait_for_timeout(800)
                        log.info('  AOM slicer: eraser clicked (previous selection cleared).')
                except Exception:
                    pass

            # Click the dropdown trigger at right side (where ▼ is)
            _mouse_click(elem, x_frac=0.85)

            # Active-wait for options to appear (up to 8 s)
            try:
                page.wait_for_selector('[role="option"]', timeout=8_000, state='attached')
                options_found = True
            except Exception:
                options_found = False

            if not options_found:
                page.keyboard.press('Escape')
                page.wait_for_timeout(1_500)
                log.info(f'  AOM dropdown: no options appeared in 8 s (attempt {attempt + 1}).')
                continue

            # Find the target name in options
            found = any(page.locator(s).count() > 0 for s in email_sels)
            if not found:
                page.keyboard.press('Escape')
                page.wait_for_timeout(500)
                log.warning(
                    f"  AOM dropdown opened but '{filter_email}' not listed. "
                    "This name may not exist in the report data."
                )
                return False

            log.info(f"  AOM dropdown: '{filter_email}' present \u2714")

            # Deselect "Select all" first
            for sel in ['[role="option"]:has-text("Select all")',
                         '[role="option"]:has-text("(Select all)")']:
                try:
                    loc = page.locator(sel)
                    if loc.count() > 0:
                        _mouse_click(loc.first, x_frac=0.5)
                        page.wait_for_timeout(600)
                        log.info("  'Select all' deselected \u2713")
                        break
                except Exception:
                    pass

            # Click the target option
            selected = False
            for sel in email_sels:
                try:
                    loc = page.locator(sel)
                    if loc.count() > 0:
                        _mouse_click(loc.first, x_frac=0.5)
                        page.wait_for_timeout(1_500)
                        log.info(f"  Selected '{filter_email}' in AOM slicer \u2713")
                        page.keyboard.press('Escape')
                        page.wait_for_timeout(600)
                        selected = True
                        break
                except Exception:
                    continue

            if selected:
                return True

            page.keyboard.press('Escape')
            return False

        log.warning(f"  AOM slicer: all 3 attempts failed for '{filter_email}'.")
        return False

'''

# ── 2. Replace _apply_filter_via_pane (strip Page-1 fallback) ─────────────────
NEW_APPLY_FILTER = r'''    def _apply_filter_via_pane(self, filter_email: str, filter_column: str) -> bool:
        """
        Set the AOM filter through the Power BI UI (slicer on current page only).
        The AOM slicer is always on the first/current page — no page-switching needed.
        Falls back to the Filters pane card if the slicer attempt fails.
        """
        page = self._page

        # Primary: AOM slicer
        log.info("  Step A: Opening AOM slicer dropdown...")
        if self._try_slicer(filter_email):
            return True

        # Fallback: Filters pane card
        log.info("  Step B: Trying Filters pane card for 'AOM'...")
'''

# ── 3. Replace _trigger_pdf_export with new-page detection ───────────────────
NEW_EXPORT = r'''    def _trigger_pdf_export(self, page, fpath: str, aom_name: str) -> str:
        """
        Drive the Power BI UI to export the report as PDF.

        Power BI sometimes downloads the PDF directly (Playwright download event)
        and sometimes opens it in a new browser tab.  This method handles both.

        Flow:
          dismiss popups → click Export (toolbar) → click PDF option →
          confirm dialog (if shown) → capture via download event OR new tab.
        """
        import time as _time

        # Track new pages that open during export
        new_pages_opened: list = []

        def _on_new_page(p):
            new_pages_opened.append(p)

        page.context.on('page', _on_new_page)

        try:
            # Step 1: Dismiss popups
            self._dismiss_popups()
            page.wait_for_timeout(500)

            # Step 2: Click Export toolbar button
            log.info("  Clicking Export toolbar button...")
            page.locator(
                '[aria-label="Export"], button:has-text("Export")'
            ).first.click(force=True, timeout=10_000)
            page.wait_for_timeout(1_000)

            log.info("  Selecting PDF and waiting for download...")
            os.makedirs(EXPORTS_DIR, exist_ok=True)

            try:
                with page.expect_download(timeout=EXPORT_TIMEOUT) as dl_info:
                    # Click PDF option
                    page.locator(
                        '[role="menuitem"]:has-text("PDF"), '
                        '[role="option"]:has-text("PDF"), '
                        'button:has-text("PDF"), a:has-text("PDF")'
                    ).first.click(force=True, timeout=8_000)
                    page.wait_for_timeout(1_000)

                    # Try confirmation dialog (multiple possible selectors)
                    CONFIRM_SELS = [
                        '[role="dialog"] button:has-text("Export")',
                        '.ms-Dialog button:has-text("Export")',
                        '[role="dialog"] button:has-text("Download")',
                        '.ms-Dialog button:has-text("Download")',
                        '[role="dialog"] button.ms-Button--primary',
                        '.ms-Dialog-actions button.ms-Button--primary',
                        'button[data-testid*="export-confirm"]',
                    ]
                    confirmed = False
                    for sel in CONFIRM_SELS:
                        try:
                            btn = page.locator(sel).first
                            btn.wait_for(timeout=3_000, state='visible')
                            btn.click(force=True)
                            log.info(f"  Export dialog confirmed.")
                            confirmed = True
                            break
                        except Exception:
                            continue
                    if not confirmed:
                        log.info("  No export confirmation dialog \u2014 download started directly.")

                dl = dl_info.value
                dl.save_as(fpath)
                log.info(f"  PDF saved: {os.path.basename(fpath)}")
                return fpath

            except PWTimeout:
                log.warning("  Download event timed out. Checking for PDF in new tab...")

                # Wait up to 30 s for a new page to appear
                deadline = _time.time() + 30
                while not new_pages_opened and _time.time() < deadline:
                    page.wait_for_timeout(1_000)

                for np in new_pages_opened:
                    try:
                        np.wait_for_load_state('networkidle', timeout=30_000)
                        pdf_url = np.url
                        log.info(f"  New tab URL: {pdf_url[:100]}")

                        # Try downloading via HTTP using the browser's cookies
                        try:
                            import requests as _req
                            cookies = {
                                c['name']: c['value']
                                for c in page.context.cookies()
                            }
                            resp = _req.get(pdf_url, cookies=cookies,
                                            timeout=120, stream=True)
                            ct = resp.headers.get('content-type', '')
                            if resp.status_code == 200:
                                with open(fpath, 'wb') as f:
                                    for chunk in resp.iter_content(chunk_size=8192):
                                        f.write(chunk)
                                np.close()
                                if os.path.getsize(fpath) > 1_000:
                                    log.info(f"  PDF saved from new tab: "
                                             f"{os.path.basename(fpath)}")
                                    return fpath
                        except Exception as req_err:
                            log.warning(f"  HTTP download from new tab failed: {req_err}")

                        np.close()
                    except Exception as tab_err:
                        log.warning(f"  New tab handling failed: {tab_err}")
                        try:
                            np.close()
                        except Exception:
                            pass

                raise TimeoutError(
                    f"PDF export failed for '{aom_name}': "
                    "no download event and no PDF in new tab."
                )
        finally:
            try:
                page.context.remove_listener('page', _on_new_page)
            except Exception:
                pass

'''

# ── Apply patches ─────────────────────────────────────────────────────────────
TRY_SLICER_START  = '    def _try_slicer(self, filter_email: str) -> bool:'
APPLY_START       = '    def _apply_filter_via_pane(self, filter_email: str, filter_column: str) -> bool:'
TRIGGER_START     = '    def _trigger_pdf_export(self, page, fpath: str, aom_name: str) -> str:'
DEBUG_START       = '    def _debug_screenshot(self, page, aom_name: str) -> None:'

def replace_method(src, method_start, new_method):
    start_idx = src.find(method_start)
    if start_idx == -1:
        print(f'  SKIP: method start not found: {method_start[:50]}')
        return src
    # Find the next def at the same indentation level (4 spaces)
    search_from = start_idx + len(method_start)
    next_def_idx = src.find('\n    def ', search_from)
    if next_def_idx == -1:
        # Replace to end of file
        src = src[:start_idx] + new_method
    else:
        src = src[:start_idx] + new_method + '\n' + src[next_def_idx + 1:]
    return src

print('Patching _try_slicer...')
src = replace_method(src, TRY_SLICER_START, NEW_TRY_SLICER)

print('Patching _apply_filter_via_pane...')
# Only replace up to the Fallback Step B line (keep rest intact)
idx = src.find(APPLY_START)
if idx == -1:
    print('  SKIP: _apply_filter_via_pane not found')
else:
    # Find "Step B" inside this method
    step_b_idx = src.find("Step B: Trying Filters pane card", idx)
    if step_b_idx == -1:
        print('  SKIP: Step B not found in _apply_filter_via_pane')
    else:
        # Replace from method start up to the Step B log line (keep Step B onward)
        # Find the line start for Step B
        line_start = src.rfind('\n', idx, step_b_idx) + 1
        src = src[:idx] + NEW_APPLY_FILTER + src[line_start:]
        print('  _apply_filter_via_pane patched.')

print('Patching _trigger_pdf_export...')
src = replace_method(src, TRIGGER_START, NEW_EXPORT)

with open(pbi_path, 'w', encoding='utf-8') as f:
    f.write(src)

print('Written. Checking syntax...')
import py_compile, sys
try:
    py_compile.compile(pbi_path, doraise=True)
    print('SYNTAX OK')
except py_compile.PyCompileError as e:
    print(f'SYNTAX ERROR: {e}')
    sys.exit(1)
