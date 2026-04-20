
import os, sys, json, re
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright
import pandas as pd
from openai import OpenAI

# --- env & client ---
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

ORG_HOST = os.getenv("CRM_HOST", "avanade.crm.dynamics.com")
OPPORTUNITIES_URL = os.getenv("OPPORTUNITIES_URL", "https://avanade.crm.dynamics.com")
USER_DATA_DIR = os.getenv("USER_DATA_DIR", r"C:\Users\fatemeh.torabi.asr\AppData\Local\Microsoft\Edge\User Data\Work")
def get_html_with_playwright(account_keyword: str) -> str:
    """Reuse login via persistent context, navigate to Opportunities, return page HTML."""
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(USER_DATA_DIR, channel='msedge', headless=False, args=['--no-first-run', '--disable-extensions'])
        page = context.new_page()
        page.goto(OPPORTUNITIES_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(1500)
        
        # Clear filters
        try:
            clear_button = page.locator("button[aria-label='Clear filters']")
            if clear_button.is_visible():
                clear_button.click()
                page.wait_for_timeout(1000)
        except Exception as e:
            print(f"Could not clear filters: {e}")
        # (Optional) Adjust to your view’s filter/search input; leave no-ops if not needed.
        # Example placeholder:
        # try:
        #     page.locator("input[aria-label='Search this view']").fill(account_keyword)
        #     page.keyboard.press("Enter")
        # except Exception:
        #     pass

        page.wait_for_timeout(1500)
        # gentle scroll to render lazy rows
        for _ in range(4):
            page.mouse.wheel(0, 800)
            page.wait_for_timeout(300)

        html = page.content()
        input("Press Enter to close the browser...")
        page.close()
        context.close()
        return html

def extract_deals_structured_from_html(account_keyword: str, html: str) -> dict:
    """Use OpenAI Structured Outputs to pull normalized deals from the HTML/innerText."""
    text = re.sub(r"<script.*?</script>", "", html, flags=re.S|re.I)
    text = re.sub(r"<style.*?</style>", "", text, flags=re.S|re.I)

    print("Texy before extraction:", text[:500])


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
        "required": ["account", "deals"]
    }

    # OpenAI Structured Outputs guarantees JSON schema adherence (Responses API)
    # See official docs for json_schema response_format and supported model snapshots. [6](https://community.dynamics.com/blogs/post/?postid=7d434f5c-ec60-4fc0-87d9-17fd5bca3ddf)
    resp = client.responses.create(
        model="gpt-4o",
        input=[
            {"role":"system",
             "content":"Extract CRM deal rows from the provided HTML/innerText. "
                       "Only output reliably parsed grid/list content."},
            {"role":"user","content":f"Account filter: {account_keyword}\n\nPAGE:\n{text[:150000]}"}
        ],
        # response_format={"type":"json_schema","json_schema":{"name":"DealsSchema","schema":schema,"strict":True}}
    )
    return json.loads(resp.output_text)

def summarize_totals(account: str, deals: list[dict]) -> dict:
    df = pd.DataFrame(deals)
    if "amount" in df.columns:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    total_all = float(df["amount"].sum()) if "amount" in df.columns else 0.0
    completed_mask = df["stage"].str.lower().str.contains("won|closed", na=False) if "stage" in df.columns else pd.Series(dtype=bool)
    total_completed = float(df.loc[completed_mask, "amount"].sum()) if "amount" in df.columns else 0.0

    # Optional: ask LLM to summarize; numbers computed locally
    summary = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role":"system","content":"You are a senior sales analyst. Be concise and bullet key points."},
            {"role":"user","content":json.dumps({"account":account,"totals":{"all":total_all,"completed":total_completed},"deals_sample":deals[:25]})}
        ],
        temperature=0.2
    ).choices[0].message.content

    return {"account": account, "totals":{"all": total_all, "completed": total_completed}, "summary": summary}

if __name__ == "__main__":
    print("Running omni_ui_agent agent_run.py...")
    account = sys.argv[1] if len(sys.argv) > 1 else "Manulife"
    print(f"Fetching and extracting deals for account filter: {account}")
    html = get_html_with_playwright(account)
    print("Extracting structured deals from HTML...")
    extracted = extract_deals_structured_from_html(account, html)
    print("Summarizing totals...")
    results = summarize_totals(extracted["account"], extracted["deals"])
