
"""
CMD Grid Test - Terminal Screen Grid Extraction with Cursor Detection
========================================================================

Opens CMD/terminal, takes multiple screenshots to catch blinking cursor, 
converts to a character grid, and prints cursor location in grid 
coordinates (row, column) using OCR-based detection.

CONFIGURATION GUIDE
==================

VISION OPTION:
==============
Option 1: Azure OpenAI Vision (Recommended - No extra setup)
  - USE_OPENAI_VISION = True
  - Uses existing GPT-4o deployment (already has vision built-in)
  - No additional Azure services needed
  - Requires: OPENAI_API_KEY in .env

Option 2: Azure Computer Vision API (Standalone)
  - USE_OPENAI_VISION = False
  - Requires separate Computer Vision resource in Azure
  - Better for pure OCR tasks without semantic reasoning
  - Requires: AZURE_VISION_ENDPOINT, AZURE_VISION_API_KEY in .env
  - Uses Azure READ API for better empty line preservation


TERMINAL DIMENSION DETECTION:
=============================
Choose how to detect terminal grid size:

1. "mode_con" (Default, CMD only, fastest)
   - DIMENSION_DETECTION = "mode_con"
   - Queries Windows 'mode con' command
   - Best for Windows CMD
   - Requires active CMD window

2. "screenshot_analysis" (Universal, works for any terminal emulator)
   - DIMENSION_DETECTION = "screenshot_analysis"
   - Analyzes screenshot to infer dimensions empirically
   - Works for mainframe terminals, VT100, xterm, etc.
   - No external commands needed
   - Slower (needs screenshot + vision extraction first)

3. "config" (Manual configuration)
   - DIMENSION_DETECTION = "config"
   - Uses hardcoded TERMINAL_ROWS and TERMINAL_COLS
   - Set environment variables to override:
     TERMINAL_ROWS=25
     TERMINAL_COLS=80


RECOMMENDATIONS:
================
- Windows CMD locally:       Use "mode_con"
- Mainframe/terminal emulator: Use "screenshot_analysis" or "config"
- Quick test with known size:  Use "config" with TERMINAL_ROWS/COLS env vars


CURSOR DETECTION:
=================
Now using pure OCR-based cursor detection with MULTIPLE SCREENSHOTS:
- Takes 2+ screenshots with time delays to catch blinking cursor
- Extracts grid using vision (OCR)
- Always uses Azure OpenAI Vision (GPT-4o) to find cursor (separate from grid extraction method)
- Asks LLM to find cursor visually in all screenshots
- Returns cursor position as (row, col) in grid coordinates
- Works for any terminal type: CMD, mainframe (3270/5250), xterm, etc.
- No pixel-to-grid conversion needed
- Universal across all platforms
"""

import os
import time
import base64
import subprocess
import json
from typing import Optional, List

import pyautogui
import pygetwindow as gw
from PIL import ImageGrab, ImageDraw, Image
from dotenv import load_dotenv
from openai import AzureOpenAI


load_dotenv()

# =====================================================================
# VISION SERVICE CONFIGURATION
# =====================================================================
# Choose one of the following:
# 1. USE_OPENAI_VISION: True  → Use Azure OpenAI GPT-4o (includes vision)
# 2. USE_OPENAI_VISION: False → Use Azure Computer Vision API (separate service)
# =====================================================================

USE_OPENAI_VISION = False  # Set to False to use Computer Vision API instead


# =====================================================================
# TERMINAL DIMENSION DETECTION METHOD
# =====================================================================
# Choose how to detect terminal grid size:
# 1. "mode_con"           → Use Windows 'mode con' command (CMD only)
# 2. "screenshot_analysis"  → Auto-detect from screenshot (universal)
# 3. "config"             → Use user-specified environment variables
# =====================================================================

DIMENSION_DETECTION = "mode_con"  # Options: "mode_con", "screenshot_analysis", "config"

# Config-based dimensions (used when DIMENSION_DETECTION = "config")
# Can override with environment variables
TERMINAL_ROWS = int(os.getenv("TERMINAL_ROWS", "25"))
TERMINAL_COLS = int(os.getenv("TERMINAL_COLS", "80"))


if USE_OPENAI_VISION:
    # Option 1: Azure OpenAI with Vision (GPT-4o)
    AZURE_ENDPOINT = "https://fa-test-openai-instance-canada-east.openai.azure.com/"
    DEPLOYMENT_NAME = "fa-test-gpt-4o"
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
    VISION_MODE = "openai"
    
    client = AzureOpenAI(
        api_key=OPENAI_API_KEY,
        api_version="2024-02-15-preview",
        azure_endpoint=AZURE_ENDPOINT,
    )
else:
    # Option 2: Standalone Azure Computer Vision API
    VISION_ENDPOINT = os.getenv("AZURE_VISION_ENDPOINT", "https://your-region.cognitiveservices.azure.com/")
    VISION_API_KEY = os.getenv("AZURE_VISION_API_KEY")
    VISION_MODE = "computer_vision"
    
    # Will use requests for direct API calls
    import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data")
CMD_SCREENSHOT_PATH = os.path.join(DATA_PATH, "cmd_grid.png")

os.makedirs(DATA_PATH, exist_ok=True)


def open_cmd():
    """Open Command Prompt window"""
    subprocess.Popen("start cmd.exe", shell=True)
    time.sleep(2)

    wins = [
        w for w in gw.getAllWindows()
        if "Command Prompt" in w.title or "cmd" in w.title.lower()
    ]
    if not wins:
        raise RuntimeError("CMD window not found")

    win = wins[0]
    win.activate()
    time.sleep(0.5)
    
    return win


def get_cursor_position():
    """Get current cursor position in screen coordinates"""
    return pyautogui.position()


def screenshot_cmd(win):
    """Take screenshot of CMD window"""
    bbox = (win.left, win.top, win.right, win.bottom)
    img = ImageGrab.grab(bbox=bbox)
    img.save(CMD_SCREENSHOT_PATH)
    print(f">>> Screenshot saved to {CMD_SCREENSHOT_PATH}")
    return img, bbox


def capture_multiple_screenshots(win, num_shots: int = 2, delay_ms: int = 400) -> List[str]:
    """
    Capture multiple screenshots to increase chance of catching blinking cursor.
    
    Args:
        win: Window object
        num_shots: Number of screenshots to capture (default 2)
        delay_ms: Delay between shots in milliseconds
    
    Returns:
        List of screenshot file paths
    """
    print(f"\n>>> Capturing {num_shots} screenshots with {delay_ms}ms interval...")
    screenshot_paths = []
    
    bbox = (win.left, win.top, win.right, win.bottom)
    
    for i in range(num_shots):
        img = ImageGrab.grab(bbox=bbox)
        
        # Save with numbered suffix
        shot_path = CMD_SCREENSHOT_PATH.replace(".png", f"_shot{i}.png")
        img.save(shot_path)
        screenshot_paths.append(shot_path)
        
        print(f">>> Screenshot {i+1}/{num_shots} saved to {shot_path}")
        
        # Wait before next shot (except on last shot)
        if i < num_shots - 1:
            time.sleep(delay_ms / 1000.0)
    
    return screenshot_paths


def extract_grid_with_azure_read_api(image_path: str, width: int, height: int) -> tuple:
    """
    Extract character grid using Azure Computer Vision READ API.
    
    READ API is optimized for:
    - Preserving empty lines and document structure
    - Accurate line-by-line extraction
    - Better for terminals and structured text
    - Includes bounding box coordinates for each line
    
    Uses synchronous polling (no async/await needed).
    Returns: (grid, rows, cols, bounding_boxes)
    """
    import requests
    
    print(">>> [OCR] Using Azure Computer Vision READ API (v3.2)...")
    
    headers = {
        "Ocp-Apim-Subscription-Key": VISION_API_KEY,
    }
    
    # Read image file
    with open(image_path, "rb") as f:
        image_data = f.read()
    
    # Submit image for reading
    read_url = f"{VISION_ENDPOINT}vision/v3.2/read/analyze"
    
    print(">>> [OCR] Submitting image for analysis...")
    response = requests.post(
        read_url,
        headers=headers,
        data=image_data,
        params={"language": "en"}
    )
    response.raise_for_status()
    
    # Get operation location from header
    operation_location = response.headers.get("Operation-Location")
    if not operation_location:
        raise RuntimeError("No Operation-Location header returned from READ API")
    
    print(f">>> [OCR] Operation ID: {operation_location.split('/')[-1]}")
    
    # Poll for completion
    max_attempts = 30
    poll_interval = 1  # seconds
    
    print(">>> [OCR] Polling for results...")
    for attempt in range(max_attempts):
        time.sleep(poll_interval)
        
        result_response = requests.get(
            operation_location,
            headers=headers
        )
        result_response.raise_for_status()
        result_json = result_response.json()
        
        status = result_json.get("status")
        print(f">>> [OCR] Attempt {attempt + 1}/{max_attempts}: status = {status}")
        
        if status == "succeeded":
            # Parse results with bounding boxes
            grid, bounding_boxes = parse_read_api_results(result_json)
            rows = len(grid)
            cols = max(len(line) for line in grid) if grid else 0
            
            print(f">>> [OCR] Read API extraction complete: {rows} rows, {cols} max columns")
            return grid, rows, cols, bounding_boxes
        
        elif status == "failed":
            raise RuntimeError(f"READ API failed: {result_json.get('analyzeResult', {}).get('error')}")
    
    raise RuntimeError(f"READ API polling timeout after {max_attempts} attempts")


def parse_read_api_results(result_json: dict) -> tuple:
    """
    Parse Azure READ API results into character grid with bounding boxes.
    
    Preserves empty lines and document structure.
    Returns: (grid, bounding_boxes) where:
      - grid: list of strings (one per line)
      - bounding_boxes: list of dicts with bbox info per line
    """
    grid = []
    bounding_boxes = []
    
    analyze_result = result_json.get("analyzeResult", {})
    pages = analyze_result.get("readResults", [])
    
    if not pages:
        print(">>> [OCR] Warning: No pages in READ API result")
        return [], []
    
    # Process first page (should be only one for terminal screenshot)
    page = pages[0]
    lines = page.get("lines", [])
    
    print(f">>> [OCR] Found {len(lines)} text lines in READ API result")
    
    # Extract text and bounding boxes from each line
    for line in lines:
        text = line.get("text", "")
        bbox = line.get("boundingBox", [])
        
        grid.append(text)
        bounding_boxes.append(bbox)  # List of [x1, y1, x2, y2, x3, y3, x4, y4] polygon coordinates
    
    return grid, bounding_boxes


def estimate_character_grid(image_path: str, width: int, height: int):
    """
    Extract character grid from terminal screenshot.
    Uses Azure READ OCR API (preserves empty lines + bounding boxes) or GPT-4o vision based on config.
    
    Returns: (grid, rows, cols, bounding_boxes)
      - grid: list of strings (one per line)
      - rows: number of rows
      - cols: max columns
      - bounding_boxes: list of bbox dicts (populated for Azure READ, empty list for Vision)
    """
    if not USE_OPENAI_VISION:
        # Azure READ OCR API - preserves empty lines and includes bounding boxes
        print("    Using Azure READ OCR API")
        grid, rows, cols, bounding_boxes = extract_grid_with_azure_read_api(image_path, width, height)
        return grid, rows, cols, bounding_boxes
    else:
        # Azure OpenAI GPT-4o Vision (no bounding boxes available)
        print("    Using Azure OpenAI GPT-4o Vision")
        import base64
        
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")
        
        response = client.chat.completions.create(
            model=DEPLOYMENT_NAME,
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are analyzing a terminal/command window screenshot.\n"
                        "Extract the visible text content as it appears.\n"
                        "Preserve all whitespace, newlines, and character positions.\n"
                        "Return JSON: { \"grid\": [list of strings, one per line], \"rows\": number, \"cols\": number }"
                    ),
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                f"Extract all visible text from this terminal window as a character grid.\n"
                                f"The window is approximately {width} pixels wide and {height} pixels tall.\n"
                                f"Return the text as a list of lines (rows), preserving spaces and formatting."
                            ),
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{img_b64}"},
                        },
                    ],
                },
            ],
        )
        result = json.loads(response.choices[0].message.content)
        grid = result.get("grid", [])
        
        # For Vision-based extraction, return empty bounding_boxes (Vision API doesn't provide them)
        return grid, result.get("rows", 0), result.get("cols", 0), []


def detect_dimensions_from_screenshot(
    image_path: str, 
    win_height: int, 
    win_width: int
) -> tuple:
    """
    Analyze the screenshot to detect terminal dimensions empirically.
    
    Measures:
    1. Actual character dimensions from extracted text
    2. Grid density (how many lines of text vs window height)
    3. Matches against common terminal sizes
    
    Returns: (estimated_rows, estimated_cols)
    Works for any terminal emulator (mainframe, VT100, etc.)
    """
    print(">>> [Screenshot Analysis] Detecting dimensions from screenshot...")
    
    # First, extract the grid to analyze
    grid_raw, _, _, _ = estimate_character_grid(image_path, win_width, win_height)
    
    actual_lines = len(grid_raw)
    print(f">>> [Screenshot Analysis] Found {actual_lines} non-empty lines in screenshot")
    
    if not grid_raw:
        print(">>> [Screenshot Analysis] No text found, using defaults")
        return 25, 80
    
    # Measure average line width
    avg_width = sum(len(line) for line in grid_raw) / len(grid_raw) if grid_raw else 0
    max_width = max(len(line) for line in grid_raw) if grid_raw else 80
    
    print(f">>> [Screenshot Analysis] Average line width: {avg_width:.1f}, Max width: {max_width}")
    
    # Estimate actual rows based on pixel distribution
    # Each line takes roughly: font_height + line_spacing pixels
    # Standard: ~16-18 pixels per line
    line_height_pixels = 16
    estimated_total_rows = max(25, win_height // line_height_pixels)
    
    # Estimate columns from longest line or window width
    # Each character takes roughly: 8-10 pixels
    char_width_pixels = 8
    estimated_total_cols = max(80, win_width // char_width_pixels)
    
    # Round to common terminal sizes for better accuracy
    # Common sizes: 24/25x80, 27x132, 30x90, 43x132
    common_heights = [24, 25, 27, 30, 43, 50]
    common_widths = [80, 90, 120, 132]
    
    # Find closest common size
    for height in sorted(common_heights, key=lambda h: abs(h - estimated_total_rows)):
        for width in sorted(common_widths, key=lambda w: abs(w - estimated_total_cols)):
            suggested_rows = height
            suggested_cols = width
            break
        break
    
    print(f">>> [Screenshot Analysis] Estimated dimensions: {suggested_rows} rows x {suggested_cols} cols")
    print(f"    (Actual extracted: {actual_lines} lines, Max line length: {max_width} chars)")
    
    return suggested_rows, suggested_cols


def query_cmd_dimensions() -> tuple:
    """
    Query actual CMD terminal dimensions using 'mode con' command.
    Returns: (rows, cols)
    
    On Windows, 'mode con' outputs:
        Lines:          30
        Columns:        80
    """
    try:
        # Run mode con and capture output
        result = subprocess.run(
            "mode con",
            shell=True,
            capture_output=True,
            text=True,
            timeout=5
        )
        
        output = result.stdout
        print(f">>> CMD mode output:\n{output}")
        
        # Parse Lines and Columns
        rows = None
        cols = None
        
        for line in output.split('\n'):
            line = line.strip()
            if "Lines:" in line:
                # Extract number after "Lines:"
                parts = line.split(":")
                if len(parts) > 1:
                    rows = int(parts[1].strip())
            elif "Columns:" in line:
                # Extract number after "Columns:"
                parts = line.split(":")
                if len(parts) > 1:
                    cols = int(parts[1].strip())
        
        if rows and cols:
            print(f">>> Detected terminal dimensions from 'mode con': {rows} rows x {cols} cols")
            return rows, cols
        else:
            print(">>> Could not parse 'mode con' output, falling back to estimation")
            return None, None
            
    except Exception as e:
        print(f">>> Error querying CMD dimensions: {e}")
        return None, None


def estimate_terminal_dimensions(
    win_height: int, 
    win_width: int, 
    image_path: Optional[str] = None
) -> tuple:
    """
    Get terminal dimensions using configured detection method.
    
    Methods:
    - "mode_con": Query Windows 'mode con' command (CMD only, fastest)
    - "screenshot_analysis": Analyze screenshot to infer dimensions (universal)
    - "config": Use user-specified TERMINAL_ROWS/TERMINAL_COLS environment variables
    
    Returns: (rows, cols)
    """
    print(f"\n>>> Terminal dimension detection method: {DIMENSION_DETECTION}")
    
    if DIMENSION_DETECTION == "mode_con":
        rows, cols = query_cmd_dimensions()
        if rows and cols:
            return rows, cols
        print(">>> Falling back to pixel-based estimation...")
        
    elif DIMENSION_DETECTION == "screenshot_analysis":
        if image_path:
            return detect_dimensions_from_screenshot(image_path, win_height, win_width)
        else:
            print(">>> Screenshot path not provided, cannot use screenshot analysis")
            
    elif DIMENSION_DETECTION == "config":
        print(f">>> Using configured dimensions: {TERMINAL_ROWS} rows x {TERMINAL_COLS} cols")
        return TERMINAL_ROWS, TERMINAL_COLS
    
    # Fallback: pixel-based estimation
    print(">>> Falling back to pixel-based estimation...")
    line_height = 16
    char_width = 8
    
    estimated_rows = max(25, win_height // line_height)
    estimated_cols = max(80, win_width // char_width)
    
    print(f">>> Estimated dimensions: {estimated_rows} rows x {estimated_cols} cols")
    return estimated_rows, estimated_cols


def pad_grid_to_terminal_size(grid: list, target_rows: int, target_cols: int) -> list:
    """
    Pad grid with empty lines to reach target terminal size.
    Also pad each line to target column width with spaces.
    """
    padded_grid = []
    
    for i in range(target_rows):
        if i < len(grid):
            line = grid[i]
            # Pad line to target width
            padded_line = line.ljust(target_cols)
            padded_grid.append(padded_line)
        else:
            # Add empty line
            padded_grid.append(" " * target_cols)
    
    return padded_grid
    """
    Use vision to estimate the character grid from CMD screenshot.
    Supports both Azure OpenAI (GPT-4o) and Azure Computer Vision API.
    Returns grid of characters.
    """
    with open(image_path, "rb") as f:
        img_b64 = base64.b64encode(f.read()).decode("utf-8")

    if VISION_MODE == "openai":
        # Option 1: Azure OpenAI GPT-4o (includes vision)
        response = client.chat.completions.create(
            model=DEPLOYMENT_NAME,
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are analyzing a Command Prompt window screenshot.\n"
                        "Extract the visible text content as it appears.\n"
                        "Preserve all whitespace, newlines, and character positions.\n"
                        "Return JSON: { \"grid\": [list of strings, one per line], \"rows\": number, \"cols\": number }"
                    ),
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                f"Extract all visible text from this CMD window as a character grid.\n"
                                f"The window appears to be approximately {width} pixels wide and {height} pixels tall.\n"
                                f"Return the text as a list of lines (rows), preserving spaces and formatting."
                            ),
                        },
                        {
                            "type": "image_url",
                            "image_url": {"url": f"data:image/png;base64,{img_b64}"},
                        },
                    ],
                },
            ],
        )
        result = json.loads(response.choices[0].message.content)
        
    else:
        # Option 2: Azure Computer Vision API (OCR)
        headers = {
            "Ocp-Apim-Subscription-Key": VISION_API_KEY,
            "Content-Type": "application/octet-stream"
        }
        
        # Read image file as bytes
        with open(image_path, "rb") as f:
            image_data = f.read()
        
        # Call Computer Vision OCR endpoint
        ocr_url = f"{VISION_ENDPOINT}vision/v3.2/ocr"
        ocr_response = requests.post(ocr_url, headers=headers, data=image_data)
        ocr_response.raise_for_status()
        ocr_result = ocr_response.json()
        
        # Parse OCR results into grid format
        grid = []
        for region in ocr_result.get("regions", []):
            for line in region.get("lines", []):
                line_text = " ".join([word["text"] for word in line.get("words", [])])
                grid.append(line_text)
        
        result = {
            "grid": grid,
            "rows": len(grid),
            "cols": max(len(line) for line in grid) if grid else 0
        }
    
    return result.get("grid", []), result.get("rows", 0), result.get("cols", 0)


def cursor_to_grid_position(cursor_pos: tuple, window_bbox: tuple, grid: list) -> tuple:
    """
    DEPRECATED: Old pixel-based approach.
    Use detect_cursor_in_grid() instead for OCR-based cursor detection.
    """
    pass


def prepare_cursor_detection_prompt(grid: list, bounding_boxes: Optional[list] = None) -> str:
    """
    Prepare grid visualization for LLM cursor anchor detection.
    Shows line numbers and optionally bounding box pixel coordinates when available.
    
    Args:
        grid: List of text lines
        bounding_boxes: Optional list of bounding box arrays from Azure READ API
    
    Returns: Formatted string showing grid with line numbers and coordinates
    """
    lines = []
    lines.append("Character Grid with Line Numbers:")
    lines.append("=" * 80)
    
    for i, line in enumerate(grid):
        # Show line number and text content
        lines.append(f"[Line {i:3d}] {line}")
        
        # If bounding boxes available, show pixel coordinates
        if bounding_boxes and i < len(bounding_boxes):
            bbox = bounding_boxes[i]
            if bbox and len(bbox) >= 8:  # Valid polygon (4 points with x,y)
                # Extract min/max coordinates from polygon
                xs = [bbox[j] for j in range(0, len(bbox), 2) if j < len(bbox)]
                ys = [bbox[j] for j in range(1, len(bbox), 2) if j < len(bbox)]
                if xs and ys:
                    x_min, x_max = min(xs), max(xs)
                    y_min, y_max = min(ys), max(ys)
                    lines.append(f"          Pixel coords: x={int(x_min)}-{int(x_max)}, y={int(y_min)}-{int(y_max)}")
    
    lines.append("=" * 80)
    return "\n".join(lines)


def find_cursor_by_anchor(grid: list, anchor_line: int, anchor_text: str) -> Optional[tuple]:
    """
    Find cursor position using anchor-based detection.
    Cursor is assumed to be immediately AFTER the specified anchor text on anchor_line.
    
    Args:
        grid: Character grid (list of strings)
        anchor_line: Line number where cursor is located
        anchor_text: Text immediately before cursor position
    
    Returns: (row, col) tuple or None if anchor not found
    """
    if anchor_line < 0 or anchor_line >= len(grid):
        print(f">>> [Cursor] Invalid anchor line: {anchor_line} (grid has {len(grid)} lines)")
        return None
    
    line = grid[anchor_line]
    
    # Find the anchor text in the line
    pos = line.find(anchor_text)
    
    if pos < 0:
        print(f">>> [Cursor] Anchor text not found in line {anchor_line}: '{anchor_text}'")
        print(f"    Line content: '{line}'")
        return None
    
    # Cursor is immediately after the anchor text
    cursor_col = pos + len(anchor_text)
    
    # Clamp to valid range
    cursor_col = min(cursor_col, len(line))
    
    print(f">>> [Cursor] Found anchor at line {anchor_line}, column {pos}")
    print(f"    Anchor text: '{anchor_text}'")
    print(f"    Cursor position: ({anchor_line}, {cursor_col})")
    
    return (anchor_line, cursor_col)


def detect_cursor_in_grid(grid: list, image_paths: List[str], bounding_boxes: Optional[list] = None) -> Optional[tuple]:
    """
    Detect cursor position using anchor-based vision analysis.
    
    Shows LLM the grid with line numbers + optional pixel coordinates (from Azure READ bboxes).
    Asks LLM to identify: which line AND what text is immediately before the cursor.
    
    Args:
        grid: List of strings representing character grid
        image_paths: List of screenshot file paths to analyze
        bounding_boxes: Optional list of bounding box coordinates (from Azure READ API)
    
    Returns: (row, col) in grid coordinates, or None if cursor not found
    Works for any terminal emulator (mainframe, xterm, 5250, etc.)
    """
    print("\n>>> [Cursor Detection] Analyzing screenshots for cursor position...")
    
    # Ensure image_paths is a list
    if isinstance(image_paths, str):
        image_paths = [image_paths]
    
    # Prepare grid visualization with line numbers and optional bounding boxes
    grid_display = prepare_cursor_detection_prompt(grid, bounding_boxes)
    
    # Encode all images
    images_b64 = []
    for img_path in image_paths:
        with open(img_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode("utf-8")
            images_b64.append(img_b64)
    
    print(f">>> [Cursor Detection] Analyzing {len(images_b64)} screenshot(s)")
    
    # Build message content with grid visualization + all images
    content = [
        {
            "type": "text",
            "text": (
                f"You are locating the text cursor in terminal screenshots.\n\n"
                f"Reference Grid (OCR extracted with line numbers):\n{grid_display}\n\n"
                f"Analyze the screenshots to find the text cursor.\n"
                f"The cursor is a VISIBLE DISTINCT MARK:\n"
                f"  - Filled block or box (different color than normal text)\n"
                f"  - Thin vertical or horizontal line\n"
                f"  - Inverted video (reversed colors) or bright highlight\n"
                f"  - NOT: a regular character, underscore in filename, dash, or hyphen\n\n"
                f"When you find the cursor:\n"
                f"1. Identify which line it's on (match to grid line numbers)\n"
                f"2. Identify the text immediately BEFORE the cursor\n"
                f"3. Return this anchor text exactly as it appears in the grid\n\n"
                f"Return value should be the EXACT text that appears right before where the cursor is."
            ),
        }
    ]
    
    # Add all images to the message
    for i, img_b64 in enumerate(images_b64):
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{img_b64}"},
        })
    
    # Ask LLM to identify cursor using structured anchor approach
    response = client.chat.completions.create(
        model=DEPLOYMENT_NAME,
        response_format={"type": "json_object"},
        messages=[
            {
                "role": "system",
                "content": (
                    "You are analyzing terminal screenshots to locate a text cursor.\n\n"
                    "REQUIRED OUTPUT FORMAT (JSON):\n"
                    "{\n"
                    '  "found": true/false,\n'
                    '  "anchor_line": <line number from grid>,\n'
                    '  "anchor_text": "<exact text before cursor>",\n'
                    '  "confidence": "high/medium/low",\n'
                    '  "reason": "<brief explanation>"\n'
                    "}\n\n"
                    'IMPORTANT:\n'
                    "- anchor_line: Use the line number shown in the reference grid\n"
                    "- anchor_text: Copy EXACTLY the text from the grid that appears right before the cursor\n"
                    "- If cursor not found: set found=false, leave anchor_line and anchor_text empty\n"
                    "- The anchor_text is your primary identifier for cursor position\n"
                    "- Be very precise - this text will be searched in the grid to locate the cursor\n"
                ),
            },
            {
                "role": "user",
                "content": content,
            },
        ],
    )

    result = json.loads(response.choices[0].message.content)
    found = result.get("found", False)
    anchor_line = result.get("anchor_line", -1)
    anchor_text = result.get("anchor_text", "")
    confidence = result.get("confidence", "")
    reason = result.get("reason", "")
    
    # Print raw response for debugging
    print(f">>> [Cursor Detection] LLM Response:")
    print(f"    found={found}")
    print(f"    anchor_line={anchor_line}")
    print(f"    anchor_text='{anchor_text}'")
    print(f"    confidence={confidence}")
    print(f"    reason={reason}")
    
    if not found:
        print(f">>> [Cursor Detection] ✗ Cursor NOT found in any screenshot")
        return None
    
    # Use anchor to find cursor position in grid
    cursor_pos = find_cursor_by_anchor(grid, anchor_line, anchor_text)
    
    if cursor_pos is None:
        print(f">>> [Cursor Detection] ✗ Could not map anchor to grid position")
        return None
    
    row, col = cursor_pos
    print(f">>> [Cursor Detection] ✓ Found cursor at grid position ({row}, {col})")
    
    # Clamp to valid range as safety check
    row = max(0, min(row, len(grid) - 1))
    col = max(0, min(col, len(grid[row])))
    
    return (row, col)


def print_grid_with_cursor(grid: list, cursor_row: int, cursor_col: int):
    """Print grid with cursor position marked"""
    print("\n" + "=" * 80)
    print("CMD WINDOW GRID")
    print("=" * 80)
    
    for i, line in enumerate(grid):
        if i == cursor_row:
            # Mark cursor position on this row
            if cursor_col < len(line):
                marked_line = (
                    line[:cursor_col]
                    + "█"  # Cursor indicator
                    + line[cursor_col + 1 :]
                )
            else:
                marked_line = line + "█"
            print(f"[{i:2d}] {marked_line}")
        else:
            print(f"[{i:2d}] {line}")

    print("=" * 80)
    print(f"\n✓ Cursor location: ROW {cursor_row}, COL {cursor_col}")
    print("=" * 80 + "\n")


def print_grid_with_cursor_safe(grid: list, cursor_pos: Optional[tuple]):
    """Print grid without cursor marking if cursor position not found"""
    print("\n" + "=" * 80)
    print("CMD WINDOW GRID (Cursor not detected)")
    print("=" * 80)
    
    for i, line in enumerate(grid):
        print(f"[{i:2d}] {line}")

    print("=" * 80)
    print("\n⚠ Cursor NOT found (possibly blinking or not visible in screenshot)")
    print("=" * 80 + "\n")


def run():
    """Main execution"""
    print(">>> Opening CMD window...")
    win = open_cmd()

    print(f">>> CMD window: {win.left}, {win.top}, {win.right}, {win.bottom}")
    print(f">>> Window size: {win.right - win.left} x {win.bottom - win.top}")

    print("\n>>> Taking screenshots...")
    screenshot_paths = capture_multiple_screenshots(win, num_shots=2, delay_ms=400)
    
    # Use first screenshot for grid extraction
    CMD_SCREENSHOT_PATH_FIRST = screenshot_paths[0]
    bbox = (win.left, win.top, win.right, win.bottom)

    print(f">>> Getting cursor position...")
    # Note: pyautogui.position() gives mouse cursor, not text cursor
    # So we'll detect cursor from grid/image instead
    # cursor_pos = get_cursor_position()
    # print(f">>> Cursor screen position: {cursor_pos}")

    print("\n>>> Extracting character grid using vision...")
    width = bbox[2] - bbox[0]
    height = bbox[3] - bbox[1]
    grid, rows, cols, bounding_boxes = estimate_character_grid(CMD_SCREENSHOT_PATH_FIRST, width, height)

    print(f">>> Raw grid dimensions: {len(grid)} rows extracted from vision")
    
    # Estimate full terminal dimensions and pad grid
    print("\n>>> Estimating full terminal dimensions...")
    estimated_rows, estimated_cols = estimate_terminal_dimensions(
        height, 
        width,
        image_path=CMD_SCREENSHOT_PATH_FIRST  # Pass image path for screenshot analysis
    )
    print(f">>> Estimated terminal size: {estimated_rows} rows x {estimated_cols} cols")
    
    # Pad grid to full terminal size
    padded_grid = pad_grid_to_terminal_size(grid, estimated_rows, estimated_cols)
    print(f">>> Padded grid dimensions: {len(padded_grid)} rows x {estimated_cols} cols")

    print("\n>>> Converting cursor to grid coordinates (Anchor-based, multi-shot)...")
    cursor_result = detect_cursor_in_grid(padded_grid, screenshot_paths, bounding_boxes)  # Pass bounding boxes
    
    if cursor_result is None:
        print(">>> [Warning] Cursor not visible in screenshot (possibly blinking off)")
        cursor_row = None
        cursor_col = None
        print_grid_with_cursor_safe(padded_grid, None)
    else:
        cursor_row, cursor_col = cursor_result
        print_grid_with_cursor(padded_grid, cursor_row, cursor_col)

    return {
        "cursor_pos": None,  # No longer using mouse cursor
        "grid": padded_grid,
        "cursor_row": cursor_row,
        "cursor_col": cursor_col,
        "cursor_found": cursor_result is not None,
        "rows": len(padded_grid),
        "cols": estimated_cols,
    }


if __name__ == "__main__":
    print(f">>> Vision Mode: {VISION_MODE.upper()}")
    if USE_OPENAI_VISION:
        print("    Using Azure OpenAI GPT-4o (includes vision)")
    else:
        print("    Using Azure Computer Vision API (OCR)")
    print()
    
    result = run()
    print("\nResult summary:")
    print(f"  Total grid rows: {result['rows']}")
    print(f"  Total grid cols: {result['cols']}")
    if result['cursor_found']:
        print(f"  Cursor at: ({result['cursor_row']}, {result['cursor_col']})")
    else:
        print(f"  Cursor: NOT FOUND (not visible in screenshot)")
    print("=" * 80)
