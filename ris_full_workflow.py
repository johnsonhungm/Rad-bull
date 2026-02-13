"""
RIS Full Automation Workflow
Combines Search, Image Opening, Extraction, AI Analysis, and Report Entry.

Steps:
1. Search specific X-rays in RIS.
2. Open the first matching Chest X-ray.
3. Wait for PACS viewer and extract image (Ctrl+I, Ctrl+C).
4. Analyze image with MedGemma 1.5 via Hugging Face Inference Endpoint.
5. Enter findings/impression into RIS Report Editor.
"""

# Check for required packages
import sys

def check_packages():
    missing = []
    try:
        import pywinauto
    except ImportError:
        missing.append("pywinauto")
    try:
        from PIL import ImageGrab
    except ImportError:
        missing.append("Pillow")
    try:
        import requests
    except ImportError:
        missing.append("requests")

    if missing:
        print("ERROR: Missing required packages:")
        for pkg in missing:
            print(f"  - {pkg}")
        print("\nInstall them with:")
        print(f"  pip install {' '.join(missing)}")
        input("\nPress Enter to exit...")
        sys.exit(1)

check_packages()

from pywinauto import Desktop, Application
from pywinauto.keyboard import send_keys
from PIL import ImageGrab
import requests
import base64
import time
import os
import re
import ctypes
from datetime import datetime

# --- CONFIGURATION ---
HF_TOKEN = os.environ.get("HF_TOKEN", "").strip('"')
HF_ENDPOINT_URL = os.environ.get("HF_ENDPOINT_URL", "").strip('"').rstrip("/")

# Get the RRG folder path (where this script is located)
# This works regardless of where the script is run from
def get_script_dir():
    """Get the directory containing this script, works in all environments."""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))

SCRIPT_DIR = get_script_dir()

# Ensure the output directory exists
if not os.path.exists(SCRIPT_DIR):
    os.makedirs(SCRIPT_DIR)

# All output files are saved in the same folder as the script (RRG folder)
TEMP_IMAGE_PATH = os.path.join(SCRIPT_DIR, "extracted_xray.png")
REPORT_PATH = os.path.join(SCRIPT_DIR, "report.txt")
LOG_PATH = os.path.join(SCRIPT_DIR, "workflow_log.txt")

# Print paths at startup for debugging
print(f"Working directory: {SCRIPT_DIR}")
print(f"Image will be saved to: {TEMP_IMAGE_PATH}")

# --- LOGGING (without private patient data) ---
def log_message(message, also_print=True):
    """Log message to file and optionally print to console."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {message}"

    # Write to log file
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(log_line + "\n")

    # Also print to console
    if also_print:
        print(message)

def prompt_for_date():
    """Prompt user to enter a date. Press Enter for today's date."""
    today = datetime.now()
    print("\n" + "="*60)
    print("DATE SELECTION")
    print("="*60)
    print(f"Today's date: {today.year}/{today.month:02d}/{today.day:02d}")
    print("\nEnter a date to search (format: YYYY/MM/DD or MM/DD)")
    print("Press Enter to use today's date.")

    user_input = input("\nDate: ").strip()

    if not user_input:
        print(f"Using today's date: {today.year}/{today.month:02d}/{today.day:02d}")
        return today

    # Try parsing the input
    try:
        # Normalize separators: replace dashes with slashes before splitting
        normalized = user_input.replace("-", "/")
        parts = normalized.split("/")

        # Try YYYY/MM/DD format
        if len(parts) == 3:
            year = int(parts[0])
            month = int(parts[1])
            day = int(parts[2])
        # Try MM/DD format (use current year)
        elif len(parts) == 2:
            year = today.year
            month = int(parts[0])
            day = int(parts[1])
        else:
            raise ValueError("Invalid format")

        selected_date = datetime(year, month, day)
        print(f"Using selected date: {selected_date.year}/{selected_date.month:02d}/{selected_date.day:02d}")
        return selected_date
    except Exception as e:
        print(f"Invalid date format: {e}")
        print(f"Falling back to today's date: {today.year}/{today.month:02d}/{today.day:02d}")
        return today

def escape_for_type_keys(text):
    """Escape special characters that pywinauto type_keys interprets as modifiers."""
    # Normalize newlines first so \r\n becomes a single \n
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Use single-pass replacement to avoid corrupting earlier substitutions
    special = {"{": "{{}", "}": "{}}", "+": "{+}", "^": "{^}", "%": "{%}", "(": "{(}", ")": "{)}"}
    result = []
    for ch in text:
        if ch in special:
            result.append(special[ch])
        elif ch == "\n":
            result.append("{ENTER}")
        else:
            result.append(ch)
    return "".join(result)

# Mouse helper
def mouse_click(x, y, double=False):
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(0.05)
    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004
    
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.02)
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    
    if double:
        time.sleep(0.1)
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.02)
        ctypes.windll.user32.mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.15)

# --- STEP 1: SEARCH & OPEN ---
def search_and_open(search_date=None):
    """Search and open X-ray. Uses search_date if provided, otherwise today."""
    log_message("\n=== STEP 1: Search & Open X-ray ===")
    app = Application(backend="win32").connect(title_re=".*放射線資訊管理系統.*主系統.*")
    main_win = app.window(title_re=".*放射線資訊管理系統.*主系統.*")
    main_win.set_focus()
    win_rect = main_win.rectangle()

    # --- 1. FIND ALL CONTROLS ---
    print("Setting Combo Boxes...", flush=True)
    # Find active combo boxes and date pickers
    # Store as list of (initial_text, control) to avoid stale key issues
    combo_list = []
    combos = {}
    date_pickers = []

    for child in main_win.children():
        if "COMBOBOX" in child.class_name() and child.window_text():
            initial_text = child.window_text()
            combo_list.append((initial_text, child))
            combos[initial_text] = child
        elif "DateTimePick" in child.class_name():
            rect = child.rectangle()
            date_pickers.append((child, rect))
    
    # If date pickers not found in immediate children, search descendants
    if not date_pickers:
        print("  Searching descendants for date pickers...", flush=True)
        for ctrl in main_win.descendants():
            try:
                if "DateTimePick" in ctrl.class_name():
                    rect = ctrl.rectangle()
                    date_pickers.append((ctrl, rect))
            except:
                pass
            
    # Define desired settings
    combo_settings = [
        ("所有類別", "一般攝影"),
        ("所有檢查地", "台大總院")
    ]
    
    # Apply category/location settings
    for current_text, desired_value in combo_settings:
        if current_text in combos:
            try:
                combos[current_text].select(desired_value)
                # Send TAB to confirm selection focus change
                combo = combos[current_text]
                combo.type_keys("{TAB}") 
                print(f"[OK] '{current_text}' -> '{desired_value}'")
            except: pass

    # Reset Physician Filters to "All"
    physician_all_values = ["所有報告醫師", "所有撰打住院醫師", "所有執行住院醫師"]
    for initial_text, combo in combo_list:
        if "所有" in initial_text: continue # Already all
        for all_val in physician_all_values:
            try:
                combo.select(all_val)
                if combo.window_text() == all_val:
                    combo.type_keys("{TAB}")
                    print(f"[OK] Reset '{initial_text}' -> '{all_val}'")
                    break
            except: pass

    # --- SET EXAMINATION PART (檢查部位) FILTER ---
    print("\nSetting Examination Part Filter...", flush=True)
    
    exam_part_combo = None
    # Look for combo box containing "檢查部位" (default) or "32001" or "Chest" (already set)
    for text, combo in combos.items():
        if "檢查部位" in text or "32001" in text or "Chest" in text:
            exam_part_combo = combo
            print(f"  Found exam part combo: '{text}'", flush=True)
            break
    
    # If not found in immediate children, search descendants
    if not exam_part_combo:
        for ctrl in main_win.descendants():
            try:
                if "COMBOBOX" in ctrl.class_name():
                    txt = ctrl.window_text()
                    if "檢查部位" in txt or "32001" in txt or "Chest" in txt:
                        exam_part_combo = ctrl
                        print(f"  Found in descendants: '{txt}'", flush=True)
                        break
            except:
                pass
    
    if exam_part_combo:
        try:
            # Click on the combo box to focus it
            exam_part_combo.click_input()
            time.sleep(0.3)
            # Type the examination code
            exam_part_combo.type_keys("32001CXM", with_spaces=True)
            time.sleep(0.3)
            # Press DOWN arrow to select from dropdown
            send_keys("{DOWN}")
            time.sleep(0.3)
            # Press TAB to confirm selection
            send_keys("{TAB}")
            time.sleep(0.3)
            print("[OK] Set Exam Part -> '32001CXM Chest:PA View (Standing)'")
        except Exception as e:
            print(f"[FAIL] Failed to set examination part filter: {e}")
    else:
        print("[WARN] Exam Part combo box not found")

    # --- 2. SET DATES (Keyboard Method with Verification) ---
    print("\nSetting Dates...", flush=True)
    now = search_date if search_date else datetime.now()
    
    print(f"Found {len(date_pickers)} date pickers", flush=True)
    
    for i, (dp, rect) in enumerate(date_pickers):
        old_date = dp.window_text()
        print(f"  Date picker {i}: {old_date}", flush=True)
        
        # Try up to 3 times to set the date correctly
        for attempt in range(3):
            # Click on the year section (left part of the control) - use rect from tuple
            year_x = rect.left + 30
            center_y = (rect.top + rect.bottom) // 2
            print(f"    Attempt {attempt+1}: Clicking year section at ({year_x}, {center_y})...", flush=True)
            mouse_click(year_x, center_y)
            time.sleep(0.5)
            
            # Type year (this selects and replaces the year)
            send_keys(str(now.year), pause=0.1)
            time.sleep(0.3)
            
            # Move to month and type
            send_keys("{RIGHT}", pause=0.1)
            time.sleep(0.3)
            send_keys(str(now.month), pause=0.1)
            time.sleep(0.3)
            
            # Move to day and type
            send_keys("{RIGHT}", pause=0.1)
            time.sleep(0.3)
            send_keys(str(now.day), pause=0.1)
            time.sleep(0.3)
            
            # Press Enter and TAB to confirm
            send_keys("{ENTER}", pause=0.1)
            time.sleep(0.3)
            send_keys("{TAB}", pause=0.1)
            time.sleep(0.5)
            
            # Verify the date was set correctly
            new_date = dp.window_text()
            print(f"    New value: {new_date}", flush=True)
            
            # Check if date matches (use zero-padded values to avoid substring false positives)
            # Check without zero-padding (RIS shows 2/1 not 02/01)
            expected_date = f"{now.year}/{now.month}/{now.day}"
            if expected_date in new_date or new_date == expected_date:
                print(f"    [OK] Date set correctly!", flush=True)
                break
            else:
                print(f"    [FAIL] Date not set correctly, retrying...", flush=True)
                time.sleep(0.5)
        else:
            # All attempts failed
            print(f"    [WARN] WARNING: Could not set date picker {i} to today's date after 3 attempts!", flush=True)
            print(f"    Current value: {dp.window_text()}, Expected: {now.year}/{now.month}/{now.day}", flush=True)
    
    # Final verification - print all date values
    print("\n  Final Date Values:", flush=True)
    for i, (dp, rect) in enumerate(date_pickers):
        print(f"    Date picker {i}: {dp.window_text()}", flush=True)

    # --- 3. CLICK SEARCH ---
    print("Clicking Search...")
    try:
        main_win.child_window(auto_id="cmdSearch").click_input() # Safer with click_input
    except:
        # Fallback to coordinates
        win_rect = main_win.rectangle()
        search_x = win_rect.left + 910
        search_y = win_rect.top + 204
        mouse_click(search_x, search_y)

    print("Search clicked. Waiting for dialog...")
    
    # Handle Dialog
    desktop_win32 = Desktop(backend="win32")
    for _ in range(10):
        try:
            dlg = desktop_win32.window(title_re=".*kReport.*")
            if dlg.exists():
                print("Dialog found. Clicking Yes...")
                dlg.set_focus()
                dlg.type_keys("%y") # Alt+Y
                break
            time.sleep(0.5)
        except: pass
    
    time.sleep(2) # Wait for results

    # --- 4. SELECT FIRST ROW ---
    # Exam filter is already set to 32001CXM, so all results are Chest X-rays.
    # Just double-click the first row in the grid.
    print("Selecting first result...")
    desktop_uia = Desktop(backend="uia")

    # Wait for Grid
    print("Waiting for DataGridView1...")
    grid = None
    for _ in range(20):
        try:
            uia_win = desktop_uia.window(title_re=".*放射線資訊管理系統.*主系統.*")
            grid = uia_win.child_window(auto_id="DataGridView1")
            if grid.exists():
                break
        except: pass
        time.sleep(0.5)

    if not grid or not grid.exists():
        print("Error: DataGridView1 not found via UIA.")
        return False

    grid_rect = grid.rectangle()

    # Click the first data row (offset from top to skip header)
    first_row_x = grid_rect.left + 50
    first_row_y = grid_rect.top + 50
    print(f"Double-clicking first row at ({first_row_x}, {first_row_y})")
    mouse_click(first_row_x, first_row_y, double=True)
    return True

# --- STEP 2: EXTRACT IMAGE ---
def extract_image():
    log_message("\n=== STEP 2: Extract Image ===")
    print("Waiting for PACS viewer...")
    desktop = Desktop(backend="win32")
    pacs_win = None
    
    # Wait up to 15 seconds
    for _ in range(15):
        for w in desktop.windows():
            if w.is_visible() and w.window_text().startswith("[總院]") and "放射線" not in w.window_text():
                pacs_win = w
                break
        if pacs_win: break
        time.sleep(1)
        
    if not pacs_win:
        log_message("Error: PACS viewer not found.")
        return False

    # Print full title to console only (contains patient info - not logged)
    print(f"Found PACS: {pacs_win.window_text()}")
    log_message("Found PACS viewer window", also_print=False)
    
    # Bring window to foreground
    try:
        pacs_win.set_focus()
        time.sleep(0.5)
    except:
        print("  Warning: Could not set focus, proceeding anyway...")

    # Clear clipboard before copying to avoid stale images from previous round
    print("  Clearing clipboard...")
    ctypes.windll.user32.OpenClipboard(0)
    ctypes.windll.user32.EmptyClipboard()
    ctypes.windll.user32.CloseClipboard()

    # Re-focus PACS window after clipboard operation
    try:
        pacs_win.set_focus()
        time.sleep(0.5)
    except:
        pass

    # Send Ctrl+I to anonymize the image
    print("  Sending Ctrl+I (anonymize)...")
    send_keys("^i")
    time.sleep(2.0)

    # Send Ctrl+C to copy the CXR image to clipboard, retry if clipboard is empty
    print("  Sending Ctrl+C (copy CXR)...")
    img = None
    for attempt in range(3):
        # Ensure PACS has focus before sending Ctrl+C (avoid sending to terminal)
        try:
            pacs_win.set_focus()
            time.sleep(0.3)
        except:
            pass
        send_keys("^c")
        time.sleep(2.0)
        img = ImageGrab.grabclipboard()
        if img is not None:
            break
        print(f"  Clipboard empty, retrying ({attempt + 1}/3)...")
        time.sleep(1.0)

    # Grab image from clipboard and save
    try:
        if img is None:
            print("Error: No image found in clipboard after 3 attempts.")
            return False
        img.save(TEMP_IMAGE_PATH, "PNG")
        print(f"  Image saved ({img.width}x{img.height} px).")
        return True
    except Exception as e:
        print(f"Error: Failed to grab image from clipboard: {e}")
        return False

# --- STEP 3: ANALYZE ---
def analyze_image():
    """Analyze image with MedGemma 1.5 via HF Inference Endpoint and return findings only."""
    log_message("\n=== STEP 3: AI Analysis ===")
    if not os.path.exists(TEMP_IMAGE_PATH):
        return None

    # Resize image to reduce payload size (max 1024px on longest side)
    from PIL import Image
    img = Image.open(TEMP_IMAGE_PATH)
    max_side = 1024
    if max(img.size) > max_side:
        img.thumbnail((max_side, max_side), Image.LANCZOS)
        print(f"  Resized image to {img.size[0]}x{img.size[1]}")
    import io
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
    print(f"  Base64 payload size: {len(b64) // 1024} KB")

    headers = {"Authorization": f"Bearer {HF_TOKEN}"}
    prompt = ("Describe the findings in this chest X-ray in plain text. "
              "Ignore any overlaid text such as 'Please refer to arrow(s) in key image(s)' — "
              "that is a software annotation, not part of the X-ray. Do not include it in your response.")

    payload = {
        "inputs": {
            "image": b64,
            "text": prompt
        },
        "parameters": {
            "max_new_tokens": 1000,
            "temperature": 0.2
        }
    }

    try:
        api_url = HF_ENDPOINT_URL
        print(f"Sending to MedGemma 1.5 via HF Endpoint...")
        print(f"  URL: {api_url}")
        res = requests.post(api_url, headers=headers, json=payload, timeout=120)
        res.raise_for_status()
        result = res.json()
        print(f"  Raw response type: {type(result).__name__}")
        print(f"  Raw response (first 300 chars): {str(result)[:300]}")

        # HF Inference API returns either a list or dict
        if isinstance(result, list):
            content = result[0].get("generated_text", "")
        elif isinstance(result, dict):
            content = result.get("generated_text", "") or result.get("text", "")
        else:
            content = str(result)

        # Save raw output before any processing
        raw_log_path = os.path.join(SCRIPT_DIR, "raw_output.txt")
        with open(raw_log_path, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\n")
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}]\n")
            f.write(f"{'='*60}\n")
            f.write(content)
            f.write("\n")
        print(f"  Raw output saved to: {raw_log_path}")

        # Strip the echoed prompt from the beginning of the response
        if content.startswith(prompt):
            content = content[len(prompt):].strip()

        # Take only the first paragraph (model repeats itself after)
        content = content.split("\n\n")[0].strip()

        with open(REPORT_PATH, "w", encoding="utf-8") as f: f.write(content)

        findings = content.strip()
        return findings
    except requests.exceptions.HTTPError as e:
        print(f"API HTTP Error: {e}")
        print(f"  Response body: {res.text[:500]}")
        return None
    except Exception as e:
        print(f"API Error: {e}")
        return None

# --- STEP 4: ENTER REPORT ---
def enter_report(findings):
    """Enter findings into the EXAM box only (skip IMPRESSION)."""
    log_message("\n=== STEP 4: Enter Report ===")

    # Debug: Show what we're about to type
    print(f"\nFindings to enter ({len(findings)} chars):")
    print(f"  {findings[:100]}..." if len(findings) > 100 else f"  {findings}")

    if not findings:
        print("[WARN] WARNING: Findings are empty!")
        return

    # Wait a bit for the report editor window to be ready
    time.sleep(2.0)

    desktop_uia = Desktop(backend="uia")

    app = None
    try:
        print("\nSearching for report editor window...")

        # Get the main RIS window
        main_win = desktop_uia.window(title_re=".*放射線資訊管理系統.*主系統.*")

        # Check if there are date pickers visible (indicates we're on search screen)
        date_pickers_exist = False
        try:
            for child in main_win.descendants():
                if "DateTimePick" in child.class_name() and child.is_visible():
                    date_pickers_exist = True
                    break
        except:
            pass

        if date_pickers_exist:
            print("[WARN] WARNING: Still on search screen (date pickers visible)")
            print("  Waiting for report editor...")
            time.sleep(3.0)

        app = main_win
        app.set_focus()
        time.sleep(0.5)

    except Exception as e:
        print(f"Error finding window: {e}")
        return

    # Enter findings into EXAM box only
    try:
        box = app.child_window(auto_id="EXAM")
        if box.exists():
            print("\nEntering into EXAM box...")
            box.click_input()
            time.sleep(0.3)
            box.type_keys("^{END}")
            time.sleep(0.2)
            box.type_keys("{ENTER}" + escape_for_type_keys(findings), with_spaces=True)
            print("[OK] Findings entered into EXAM box")
        else:
            print("[FAIL] EXAM box not found")

    except Exception as e:
        print(f"[FAIL] Error entering report: {e}")

# --- MAIN ---
if __name__ == "__main__":
    try:
        # Log workflow start
        log_message("\n" + "="*60)
        log_message("WORKFLOW STARTED")
        log_message("="*60)

        # Validate HF credentials before starting any GUI automation
        if not HF_TOKEN:
            print("ERROR: HF_TOKEN environment variable is not set.")
            print("Set it in run.bat or via: $env:HF_TOKEN = 'your-hf-token'")
            input("\nPress Enter to exit...")
            sys.exit(1)
        if not HF_ENDPOINT_URL:
            print("ERROR: HF_ENDPOINT_URL environment variable is not set.")
            print("Deploy MedGemma 1.5 at https://endpoints.huggingface.co/")
            print("Then set it in run.bat or via: $env:HF_ENDPOINT_URL = 'https://your-endpoint.endpoints.huggingface.cloud'")
            input("\nPress Enter to exit...")
            sys.exit(1)

        # Prompt user for date selection
        selected_date = prompt_for_date()
        log_message(f"Search date: {selected_date.year}/{selected_date.month}/{selected_date.day}")

        # Prompt user for number of reports to process
        print("\n" + "="*60)
        print("NUMBER OF REPORTS")
        print("="*60)
        print("How many reports do you want to process?")
        print("(If more than available, it will stop at the last one.)")
        num_input = input("\nNumber of reports: ").strip()
        try:
            num_reports = int(num_input)
            if num_reports < 1:
                num_reports = 1
        except ValueError:
            print("Invalid number. Defaulting to 1.")
            num_reports = 1
        log_message(f"Reports to process: {num_reports}")

        if search_and_open(selected_date):
            all_succeeded = True
            for report_num in range(num_reports):
                log_message(f"\n{'='*60}")
                log_message(f"PROCESSING REPORT {report_num + 1} OF {num_reports}")
                log_message(f"{'='*60}\n")

                # Delete old image to prevent reusing previous patient's image
                if os.path.exists(TEMP_IMAGE_PATH):
                    os.remove(TEMP_IMAGE_PATH)

                if extract_image():
                    findings = analyze_image()
                    if findings:
                        enter_report(findings)
                        log_message(f"\n[OK] Report {report_num + 1} completed!")

                        # If not the last report, press F4 to go to next
                        if report_num < num_reports - 1:
                            print("\n--- Moving to Next Report ---")
                            print("Pressing F4...", flush=True)
                            send_keys("{F4}")
                            time.sleep(3.0)  # Wait for next report to load

                            print("Waiting for PACS viewer to update...", flush=True)
                            time.sleep(2.0)  # Additional wait for PACS to load new image
                    else:
                        log_message(f"[FAIL] Failed to analyze image for report {report_num + 1}")
                        all_succeeded = False
                        break
                else:
                    log_message(f"[FAIL] Failed to extract image for report {report_num + 1}")
                    all_succeeded = False
                    break

            log_message("\n" + "="*60)
            if all_succeeded:
                log_message(f"[OK] ALL {report_num + 1} REPORTS COMPLETED SUCCESSFULLY!")
            else:
                log_message(f"[FAIL] WORKFLOW FINISHED WITH ERRORS (completed {report_num} of {num_reports})")
            log_message("="*60)

    except Exception as e:
        log_message(f"\n{'='*60}")
        log_message(f"ERROR: {e}")
        log_message(f"{'='*60}")
        print("\nMake sure:")
        print("  1. RIS application is open (main window visible)")
        print("  2. You are logged in to the RIS system")
        print("  3. The search screen is visible")
        input("\nPress Enter to exit...")
