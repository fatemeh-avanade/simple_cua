# browser_excel_test.py

from playwright.sync_api import sync_playwright

example_excel_sharepoint_url = "https://avanade-my.sharepoint.com/:x:/r/personal/fatemeh_torabi_asr_avanade_com/Documents/test_data_folder/Book.xlsx?d=w8f930cd74d9641d7a294db6ff4f350db&csf=1&web=1&e=klFwv8"


user_data_dir = r"C:\Users\fatemeh.torabi.asr\AppData\Local\Microsoft\Edge\User Data\Work"

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(user_data_dir, channel='msedge', headless=False, args=['--no-first-run', '--disable-extensions'])
    page = context.new_page()
    page.goto(example_excel_sharepoint_url)
    # This selector is fragile — that's OK for today
    page.wait_for_timeout(5000)
    value = page.evaluate("""
        () => {
            // Try multiple selectors
            let elements = document.querySelectorAll('div, span, td, [role="gridcell"]');
            
            for (const elem of elements) {
                const text = elem.innerText || elem.textContent;
                if (text && text.trim() === '1240.5') {
                    return text.trim();
                }
            }
            
            // If not found, show first few elements to debug
            elements = document.querySelectorAll('*');
            let found = [];
            for (let i = 0; i < Math.min(50, elements.length); i++) {
                const text = (elements[i].innerText || elements[i].textContent || '').trim();
                if (text && text.length > 0 && text.length < 100) {
                    found.push(text);
                }
            }
            return found;
        }
    """)

    print("Value found in Excel page:", value)

    if isinstance(value, str):
        print(f"Found value: {value}")

    else:
        print("Sample of page content:")
        for text in value[:30]:
            if text:
                print(f"  {repr(text)}")


    page.close()
    context.close()


