from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={'width': 1920, 'height': 1080})
    page.goto('https://login.microsoftonline.com/')
    page.fill('input[type="email"]', 'pragati.panhale@kisna.com')
    page.keyboard.press('Enter')
    page.wait_for_timeout(2000)
    page.fill('input[type="password"]', 'JxQDD6QrPiMp')
    page.keyboard.press('Enter')
    page.wait_for_timeout(3000)
    
    try:
        inter = page.locator('input[type="email"]').first
        if inter.is_visible(timeout=2000):
            inter.fill('pragati.panhale@kisna.com')
            page.keyboard.press('Enter')
            page.wait_for_timeout(3000)
    except:
        pass

    page.goto('https://app.powerbi.com/groups/d5f0d52b-8c11-4092-9b88-d9cfab71c3fe/reports/91d645a5-6cae-45f8-b9c2-b29ad296f652/40867f7e388ce019a7a4', wait_until='domcontentloaded')
    page.wait_for_timeout(10000)
    
    # Run the force unhide script
    page.evaluate("""
        () => {
            document.querySelectorAll('.visual, .visual-container').forEach(el => {
                el.style.setProperty('display', 'block', 'important');
                el.style.setProperty('visibility', 'visible', 'important');
                el.style.setProperty('opacity', '1', 'important');
                el.style.setProperty('z-index', '9999', 'important');
            });
        }
    """)
    page.wait_for_timeout(1000)

    elements = page.locator('.visual [role="combobox"], .visual [aria-haspopup="listbox"], .visual [aria-haspopup="true"], .slicerDropdownMenu').all()
    print(f'Found {len(elements)} combobox elements')
    for i, el in enumerate(elements):
        try:
            html = el.evaluate('node => node.outerHTML')
            print(f'Combobox {i}: {html[:300]}...')
        except:
            pass
            
    slicers = page.locator('.slicer-container, .visual-slicer').all()
    print(f'Found {len(slicers)} slicers')
    
    # See if there are page navigation tabs
    pages = page.locator('button[aria-label*="Page"], .pageNavigation button, .explorationContainer .navigation-node, .navigation-pane').all()
    print(f'Found {len(pages)} page tabs/nav nodes')
    for i, el in enumerate(pages):
        try:
            html = el.evaluate("node => node.outerHTML")
            print(f'Page tab {i}: {html[:200]}...')
        except:
            pass
            
    browser.close()
