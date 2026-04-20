
import os, json
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
from openai import OpenAI

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

ORG_HOST = os.getenv("CRM_HOST", "yourorg.crm.dynamics.com")
OPPORTUNITIES_URL = f"https://{ORG_HOST}/main.aspx?app=d365default&pagetype=entitylist&etn=opportunity"
USER_DATA_DIR = os.getenv("USER_DATA_DIR", str(Path.cwd() / ".pw-profile"))

def screenshot_page() -> str:
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(USER_DATA_DIR, headless=False)
        page = ctx.new_page()
        page.goto(OPPORTUNITIES_URL, wait_until="domcontentloaded")

        # slow-scroll to ensure lazy rows render; helpful before full-page screenshots. [8](https://mspowerautomate.com/how-to-scrape-data-from-web-pages-in-microsoft-power-automate-desktop/)
        for _ in range(10):
            page.mouse.wheel(0, 800)
            page.wait_for_timeout(250)

        img_path = Path("opportunities_full.png").as_posix()
        page.screenshot(path=img_path, full_page=True)
        page.close()
        ctx.close()
        return img_path

def extract_from_image(account_keyword: str, img_path: str) -> dict:
    schema = {
        "type": "object",
        "properties": {
            "account": {"type": "string"},
            "deals": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "name":       {"type": "string"},
                        "stage":      {"type": "string"},
                        "amount":     {"type": "number"},
                        "close_date": {"type": "string"},
                        "owner":      {"type": "string"},
                        "account":    {"type": "string"}
                    },
                    "required": ["name"]
                }
            }
        },
        "required": ["account","deals"]
    }

    # Multimodal extraction with structured outputs (vision + schema). [9](https://www.sikich.com/insight/how-to-filter-microsoft-dynamics-365-entities-with-odata/)[6](https://community.dynamics.com/blogs/post/?postid=7d434f5c-ec60-4fc0-87d9-17fd5bca3ddf)
    resp = client.responses.create(
        model="gpt-4o",
        input=[
            {"role":"system","content":"Extract normalized CRM deals (name, stage, amount, close_date, owner, account) from the screenshot."},
            {"role":"user","content":[
                {"type":"input_text","text":f"Account filter: {account_keyword}"},
                {"type":"input_image","image_url":f"file://{Path(img_path).absolute()}"}
            ]}
               ],
        response_format={"type":"json_schema","json_schema":{"name":"DealsFromScreenshot","schema":schema,"strict":True}}
    )
    return json.loads(resp.output_text)

if __name__ == "__main__":
    img = screenshot_page()
    parsed = extract_from_image("Manulife", img)

