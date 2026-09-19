from playwright.sync_api import sync_playwright

def inspect_report():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={'width': 1920, 'height': 1080})
        
        print("Logging in...")
        page.goto('https://login.microsoftonline.com/')
        page.fill('input[type="email"]', 'pragati.panhale@kisna.com')
        page.keyboard.press('Enter')
        page.wait_for_timeout(2000)
        page.fill('input[type="password"]', 'JxQDD6QrPiMp')
        page.keyboard.press('Enter')
        page.wait_for_timeout(3000)
        
        try:
            inter = page.locator('input[type="email"], input[placeholder="Enter email"]').first
            if inter.is_visible(timeout=2000):
                inter.fill('pragati.panhale@kisna.com')
                page.keyboard.press('Enter')
                page.wait_for_timeout(3000)
        except:
            pass

        print("Navigating to report...")
        page.goto('https://app.powerbi.com/groups/d5f0d52b-8c11-4092-9b88-d9cfab71c3fe/reports/91d645a5-6cae-45f8-b9c2-b29ad296f652/40867f7e388ce019a7a4', wait_until='domcontentloaded')
        page.wait_for_timeout(5000)
        
        # Handle identity prompt
        print("Dismissing identity prompt...")
        for _ in range(2):
            try:
                btn = page.locator('button:has-text("Continue"), [aria-label="Continue"]').first
                if btn.is_visible(timeout=3000):
                    btn.click()
                    page.wait_for_timeout(2000)
            except:
                pass
                
        page.wait_for_timeout(5000)
        
        print("Unhiding all visuals...")
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
        page.wait_for_timeout(2000)
        
        print("Looking for emails in DOM...")
        emails = ["aomnorth1@kisna.com", "aomnorth3d@kisna.com"]
        for email in emails:
            els = page.locator(f'text="{email}"').all()
            print(f"Found '{email}' {len(els)} times.")
            for i, el in enumerate(els):
                try:
                    html = el.evaluate('node => node.outerHTML')
                    print(f"  Match {i}: {html[:200]}...")
                except:
                    pass

        print("Looking for comboboxes...")
        combos = page.locator('[role="combobox"]').all()
        print(f"Found {len(combos)} comboboxes.")
        for i, el in enumerate(combos):
            try:
                html = el.evaluate('node => node.outerHTML')
                print(f"  Combobox {i}: {html[:200]}...")
            except:
                pass

        print("Looking for checkboxes/list items...")
        lists = page.locator('[role="checkbox"]').all()
        print(f"Found {len(lists)} checkboxes.")
        
        print("Looking for page tabs...")
        pages = page.locator('button[aria-label*="Page"], .pageNavigation button').all()
        print(f"Found {len(pages)} page tabs.")
        for i, el in enumerate(pages):
            try:
                html = el.evaluate('node => node.outerHTML')
                print(f"  Page {i}: {html[:200]}...")
            except:
                pass

        browser.close()

if __name__ == '__main__':
    inspect_report()
