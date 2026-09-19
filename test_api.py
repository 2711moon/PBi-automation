from playwright.sync_api import sync_playwright

def test_pbi():
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
            inter = page.locator('input[type="email"], input[placeholder="Enter email"]').first
            if inter.is_visible(timeout=2000):
                inter.fill('pragati.panhale@kisna.com')
                page.keyboard.press('Enter')
                page.wait_for_timeout(3000)
        except:
            pass

        page.goto('https://app.powerbi.com/groups/d5f0d52b-8c11-4092-9b88-d9cfab71c3fe/reports/91d645a5-6cae-45f8-b9c2-b29ad296f652/40867f7e388ce019a7a4', wait_until='domcontentloaded')
        page.wait_for_timeout(10000)
        for _ in range(2):
            try:
                btn = page.locator('button:has-text("Continue"), [aria-label="Continue"]').first
                if btn.is_visible(timeout=3000):
                    btn.click()
                    page.wait_for_timeout(2000)
            except: pass
        page.wait_for_timeout(5000)
        
        # Check globals
        print("Globals:", page.evaluate('() => Object.keys(window).filter(k => k.toLowerCase().includes("powerbi"))'))
        print("Frames globals:", [f.evaluate('() => Object.keys(window).filter(k => k.toLowerCase().includes("powerbi"))') for f in page.frames])
        
        browser.close()

if __name__ == '__main__':
    test_pbi()
