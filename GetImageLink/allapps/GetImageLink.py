from datetime import datetime, timedelta
import sys
import os
import time
import subprocess
import socket
from playwright.sync_api import sync_playwright
import pyautogui
import time
import re
from difflib import SequenceMatcher
import pyperclip
from datetime import datetime


class ChromeBrowserController:
    def __init__(self, port=9222, user_data_dir=None):
        """
        Initialize Chrome browser controller
        
        Args:
            port: Port number for Chrome remote debugging (default: 9222)
            user_data_dir: Path to Chrome user data directory for persistent login
        """
        self.port = port
        self.chrome_process = None
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        
        # Setup user data directory for persistent sessions
        if user_data_dir is None:
            self.user_data_dir = r"C:\Amazon_Data"
        else:
            self.user_data_dir = user_data_dir
        
        # Create user data directory if it doesn't exist
        os.makedirs(self.user_data_dir, exist_ok=True)
    
    def is_port_in_use(self, port):
        """Check if a port is already in use"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.connect(('localhost', port))
                return True
            except:
                return False
    
    def find_chrome_path(self):
        """Find Chrome executable path"""
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
        ]
        
        for path in chrome_paths:
            if os.path.exists(path):
                return path
        
        raise FileNotFoundError("Chrome executable not found. Please install Google Chrome.")
    
    def launch_chrome(self, initial_url=None):
        """Launch Chrome with remote debugging enabled (or skip if already running)"""
        # Check if Chrome is already running on this port
        if self.is_port_in_use(self.port):
            print(f"Chrome is already running on port {self.port}, reusing existing instance...")
            return True
        
        chrome_path = self.find_chrome_path()
        
        # Chrome arguments for remote debugging and persistent sessions
        chrome_args = [
            chrome_path,
            f"--remote-debugging-port={self.port}",
            f"--user-data-dir={self.user_data_dir}",
            "--no-first-run",
            "--no-default-browser-check",
        ]

        # If an initial URL is provided, add it so Chrome opens that page on launch
        if initial_url:
            chrome_args.append(initial_url)
        
        print(f"Launching Chrome on port {self.port}...")
        print(f"User data directory: {self.user_data_dir}")
        
        try:
            # Launch Chrome process
            self.chrome_process = subprocess.Popen(
                chrome_args,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
            
            # Wait for Chrome to start
            time.sleep(3)
            print(f"Chrome launched successfully (PID: {self.chrome_process.pid})")
            return True
            
        except Exception as e:
            print(f"Error launching Chrome: {e}")
            return False
    
    def connect_playwright(self):
        """Connect to Chrome using Playwright"""
        try:
            print("Connecting to Chrome via Playwright...")
            self.playwright = sync_playwright().start()
            
            # Connect to existing Chrome instance
            self.browser = self.playwright.chromium.connect_over_cdp(
                f"http://localhost:{self.port}"
            )
            
            # Get the default context (uses Chrome's profile)
            contexts = self.browser.contexts
            if contexts:
                self.context = contexts[0]
            else:
                self.context = self.browser.new_context()
            
            # Get existing page or create new one
            pages = self.context.pages
            if pages:
                self.page = pages[0]
            else:
                self.page = self.context.new_page()
            
            print("Successfully connected to Chrome via Playwright!")
            return True
            
        except Exception as e:
            print(f"Error connecting to Chrome: {e}")
            return False
            
    def disconnect(self):
        """Disconnect Playwright without closing Chrome"""
        if self.browser:
            try: self.browser.disconnect()
            except: pass
        if self.playwright:
            try: self.playwright.stop()
            except: pass
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
    
    def start(self):
        """Start the browser and connect"""
        if not self.launch_chrome():
            return False
        
        if not self.connect_playwright():
            self.cleanup()
            return False
        
        return True

    def wait_and_click_repeat(self,base_xpath):
            if not self.page:
                return False
            
            while True:  # infinite retry loop
                for i in range(0, 20):  # i = 0 to 5
                    xpath = base_xpath.format(i=i)
        
                    locator = self.page.locator(f"xpath={xpath}")
                    try:
                        # Wait VERY shortly to avoid blocking
                        locator.wait_for(state="visible", timeout=100)
                        locator.click()
                        return i
                    except:
                        # Not found → try next i
                        continue

    def wait_and_click(self, xpath):
        """Wait indefinitely for element and click when found"""
        if not self.page:
            return False
        while True:
            try:
                self.page.locator(f"xpath={xpath}").click(timeout=0)
                return True
            except:
                print("Not Found Try Again")

    def wait_for_element(self, xpath):
        """Wait indefinitely for element and click when found"""
        if not self.page:
            return False
        while True:
            try:
                self.page.locator(f"xpath={xpath}")
                return True
            except:
                print("Not Found Try Again")
                time.sleep(0.5)
    
    def wait_and_fill(self, xpath, value, timeout=15, retry_interval=0.5):
        import time
    
        start_time = time.time()
    
        while time.time() - start_time < timeout:
            try:
                success = self.page.evaluate(
                    """(args) => {
                        const { xp, val } = args;
    
                        const el = document.evaluate(
                            xp,
                            document,
                            null,
                            XPathResult.FIRST_ORDERED_NODE_TYPE,
                            null
                        ).singleNodeValue;
    
                        if (!el) return false;
    
                        el.scrollIntoView({ block: "center" });
                        el.focus();
                        el.click();
    
                        el.value = val;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        el.dispatchEvent(new Event('blur', { bubbles: true }));
    
                        return true;
                    }""",
                    {"xp": xpath, "val": value}
                )
    
                if success:
                    return True
    
            except Exception:
                pass
            
            time.sleep(retry_interval)
    
        raise TimeoutError(f"Element not found within {timeout}s: {xpath}")

    def wait_for_text(self, xpath, text):
        """Wait infinitely until the element contains the given text (no timeout)."""

        if not self.page:
            return False

        print(f"Waiting for text '{text}' in element...")
        locator = self.page.locator(f"xpath={xpath}").first

        while True:
            try:
                # Try to ensure it's attached using count()
                try:
                    if locator.count() == 0:
                        print("Element not attached yet, retrying...")
                        continue
                except:
                    print("Element not found, retrying...")
                    continue

                # Try reading text
                try:
                    content = locator.text_content()
                    if content:
                        if text.lower() in content.lower():
                            print(f"✓ Found text: '{text}'")
                            return True
                        else:
                            pass
                    else:
                        print("Element empty, retrying...")

                except Exception as e:
                    print(f"Read error: {e}, retrying...")

            except Exception as e:
                print(f"Locator error: {e}, retrying...")

    def wait_for_user(self):
        """Keep browser open and wait for user interaction"""
        print("\n" + "="*60)
        print("Browser is ready! You can interact with it now.")
        print("Press Ctrl+C to stop and close the browser...")
        print("="*60 + "\n")
        
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nShutting down...")

    def cleanup(self):
        """Clean up resources"""
        print("Cleaning up...")
        
        if self.page:
            try:
                self.page.close()
            except:
                pass
        
        if self.context:
            try:
                self.context.close()
            except:
                pass
        
        if self.browser:
            try:
                self.browser.close()
            except:
                pass
        
        if self.playwright:
            try:
                self.playwright.stop()
            except:
                pass
        
        if self.chrome_process:
            try:
                self.chrome_process.terminate()
                self.chrome_process.wait(timeout=5)
            except:
                try:
                    self.chrome_process.kill()
                except:
                    pass
        
        print("Cleanup complete!")

    def navigate(self, url):
        """Navigate to a URL"""
        if not self.page:
            print("Browser not connected!")
            return False
        
        try:
            print(f"Navigating to: {url}")
            self.page.goto(url, wait_until="domcontentloaded")
            print("Navigation complete!")
            return True
        except Exception as e:
            print(f"Error navigating to {url}: {e}")
            return False
    
def expiry(date_str):
    d = datetime.strptime(date_str, "%d/%m/%Y")
    try: d = d.replace(year=d.year + 15)
    except: d = d.replace(day=28, month=2, year=d.year + 15)
    return (d - timedelta(days=1)).strftime("%d/%m/%Y")

def read_data_dict(path="DATA.txt"):
    """Read DATA.txt into a dict of key->value (stripped)."""
    data = {}
    if not os.path.exists(path):
        return data
    with open(path, 'r', encoding='utf-8') as f:
        for line in f:
            if '=' in line and not line.strip().startswith('-'):
                k, v = line.split('=', 1)
                data[k.strip()] = v.strip().strip('"').strip()
    return data

import sys
import os
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QComboBox, QPushButton, 
                               QLineEdit, QFileDialog, QSystemTrayIcon, QStyle, QMessageBox, QSizePolicy, QSpacerItem, QSpinBox)
from PySide6.QtCore import Qt, QThread, Signal

class ChromeLaunchThread(QThread):
    status_signal = Signal(str)
    notification_signal = Signal(str, str)
    finished_signal = Signal(bool, bool)

    def __init__(self, controller, region_domain=None):
        super().__init__()
        self.controller = controller
        self.region_domain = region_domain

    def run(self):
        try:
            is_running = self.controller.is_port_in_use(self.controller.port)
            
            if is_running:
                self.notification_signal.emit("Already Opened", f"Browser already running...")
            else:
                self.status_signal.emit("Launching Chrome...")
                # If region_domain provided, build sellercentral URL and open it on launch so user can login
                if self.region_domain:
                    try:
                        d = self.region_domain.replace('www.', '')
                        seller_host = f"sellercentral.{d}"
                        seller_url = (f"https://{seller_host}/")
                    except Exception:
                        seller_url = None
                else:
                    seller_url = None

                self.controller.launch_chrome(initial_url=seller_url)
                
            if self.controller.connect_playwright():
                self.status_signal.emit("Browser connected.")
                self.finished_signal.emit(True, is_running)
                self.controller.disconnect()
            else:
                self.notification_signal.emit("Error", "Failed to connect to browser!")
                self.finished_signal.emit(False, is_running)
        except Exception as e:
            self.notification_signal.emit("Error", str(e))
            self.finished_signal.emit(False, False)

class ExcelProcessThread(QThread):
    status_signal = Signal(str)
    notification_signal = Signal(str, str)
    finished_signal = Signal(bool)
    retry_save_signal = Signal()
    login_request_signal = Signal(str, str)

    def __init__(self, controller, excel_path, domain="www.amazon.com", batch_size=5):
        super().__init__()
        self.controller = controller
        self.excel_path = excel_path
        self.domain = domain
        self.stop_requested = False
        self.temp_excel_path = self.excel_path + ".temp.xlsx"
        self.region_domains = [
            "www.amazon.com",
            "www.amazon.in",
            "www.amazon.ca",
            "www.amazon.com.mx",
            "www.amazon.com.br",
            "www.amazon.de",
            "www.amazon.fr",
            "www.amazon.it",
            "www.amazon.es",
            "www.amazon.co.jp",
            "www.amazon.com.au",
            "www.amazon.ae",
            "www.amazon.sa",
            "www.amazon.sg"
        ]
        self.batch_size = batch_size
        self.user_logged_in = False


    def run(self):
        excel = None
        wb = None
        try:
            import time
            import urllib.parse
            import os
            import subprocess
            import win32com.client
            
            abs_excel_path = os.path.abspath(self.excel_path)
            
            # Reconnect Playwright exclusively in this thread
            if not self.controller.connect_playwright():
                self.notification_signal.emit("Error", "Failed to connect to browser!")
                self.finished_signal.emit(False)
                return
            # Ask user to login to Seller Central before proceeding
            try:
                self.notification_signal.emit("Please login", "Please login to Seller Central in the opened browser to continue.")
                # This will trigger the main UI to show a blocking dialog; the connection uses BlockingQueuedConnection
                self.login_request_signal.emit("Please login", "Please login to Seller Central in the opened browser and click OK to continue.")
            except Exception:
                pass

            # Wait until the main UI confirms the user clicked OK (i.e., user logged in)
            wait_start = time.time()
            while not self.user_logged_in and not self.stop_requested:
                time.sleep(0.2)
            if self.stop_requested:
                if excel:
                    try: excel.Quit()
                    except: pass
                self.finished_signal.emit(False)
                return
                
            self.status_signal.emit("Starting Excel application...")
            try:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
            except Exception as e:
                self.notification_signal.emit("Error", f"Could not start Excel COM interface: {e}")
                self.finished_signal.emit(False)
                return
                
            self.status_signal.emit("Locking and opening Excel file...")
            
            # Keep trying to open until not Read-Only and no exceptions
            while not self.stop_requested:
                try:
                    wb = excel.Workbooks.Open(abs_excel_path)
                    if wb.ReadOnly:
                        wb.Close(SaveChanges=False)
                        self.status_signal.emit("Excel file is Read-Only. Waiting for user to close it...")
                        self.retry_save_signal.emit()
                        time.sleep(1)
                        continue
                    sheet = wb.ActiveSheet
                    break
                except Exception as e:
                    self.status_signal.emit("Failed to open Excel file. Waiting for user to close it...")
                    self.retry_save_signal.emit()
                    time.sleep(1)
                    
            if self.stop_requested:
                if excel:
                    try: excel.Quit()
                    except: pass
                self.finished_signal.emit(False)
                return
                
            # Hide the file explicitly as requested
            self.status_signal.emit("Hiding Excel file...")
            try:
                subprocess.run(['attrib', '+h', '+s', abs_excel_path], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            except Exception as e:
                self.status_signal.emit(f"Warning: Could not hide file - {e}")

            self.status_signal.emit("Reading Excel file...")

            max_row = sheet.UsedRange.Rows.Count
            current_row = 2
            batch_size = self.batch_size

            overall_stop = False

            # Read in chunks of 100 rows to avoid large memory/processing stalls
            while current_row <= max_row and not self.stop_requested:
                end_row = min(current_row + 99, max_row)
                rows_to_process = []
                for row_idx in range(current_row, end_row + 1):
                    try:
                        status = sheet.Cells(row_idx, 15).Value
                    except Exception:
                        status = None
                    if status and str(status).strip().lower() == "yes":
                        continue

                    sku = sheet.Cells(row_idx, 2).Value
                    asin = sheet.Cells(row_idx, 3).Value

                    if sku and asin:
                        rows_to_process.append({
                            "row_idx": row_idx,
                            "sku": sku,
                            "asin": asin,
                        })

                # If nothing to process in this chunk, move to next
                if not rows_to_process:
                    current_row = end_row + 1
                    # refresh max_row in case sheet was modified
                    try:
                        max_row = sheet.UsedRange.Rows.Count
                    except Exception:
                        pass
                    continue

                pending = list(rows_to_process)
                opened = []  # currently processing items
                available_pages = []  # pages that can be reused

                while pending or opened:
                    if self.stop_requested:
                        self.status_signal.emit("Stop requested. Cleaning up...")
                        overall_stop = True
                        break

                    # Open new tabs or reuse existing up to batch_size
                    while pending and len(opened) < batch_size:
                        item = pending.pop(0)

                        # Reuse page if available, otherwise create a new one. Keep region fixed.
                        if available_pages:
                            page, last_successful_domain = available_pages.pop()
                            current_domain = last_successful_domain
                            try: page.evaluate("document.body.innerHTML = '';")
                            except: pass
                        else:
                            page = self.controller.context.new_page()
                            current_domain = self.domain

                        opened.append({
                            "page": page,
                            "row_idx": item['row_idx'],
                            "sku": item['sku'],
                            "asin": item['asin'],
                            "current_domain": current_domain,
                            "stage": "navigating",
                            "next_action_time": 0,
                            "start_t": 0,
                            "row_start_t": time.time(),
                            "reload1_done": False,
                            "reload2_done": False
                        })

                    tab_was_freed = False
                    current_time = time.time()

                    for entry in list(opened):
                        page = entry["page"]

                        # 1. Check if user manually closed it
                        try:
                            if page.is_closed():
                                if entry in opened:
                                    opened.remove(entry)
                                    tab_was_freed = True
                                continue
                        except Exception:
                            if entry in opened:
                                try: opened.remove(entry)
                                except: pass
                                tab_was_freed = True
                            continue

                        stage = entry.get("stage")

                        if stage != "delaying_next_row":
                            if current_time - entry.get("row_start_t", current_time) > 30:
                                self.status_signal.emit(f"Row {entry['row_idx']} (ASIN={entry['asin']}) exceeded 30s. Skipping.")
                                if entry in opened:
                                    try: opened.remove(entry)
                                    except: pass
                                    tab_was_freed = True
                                available_pages.append((page, entry["current_domain"]))
                                try: page.evaluate("document.body.innerHTML = '<h2>Skipped - took more than 30s!</h2>';")
                                except: pass
                                continue

                        if stage == "navigating":
                            safe_asin = urllib.parse.quote(str(entry['asin']))
                            # Build sellercentral host from the current_domain (e.g. www.amazon.com -> sellercentral.amazon.com)
                            def make_seller_host(domain):
                                try:
                                    # remove leading www.
                                    d = domain.replace('www.', '')
                                    if d.startswith('amazon.'):
                                        return f"sellercentral.{d}"
                                    # fallback: try to find amazon.* and build sellercentral.amazon.*
                                    idx = d.find('amazon.')
                                    if idx != -1:
                                        return f"sellercentral.{d[idx:]}"
                                    return f"sellercentral.{d}"
                                except:
                                    return f"sellercentral.{domain}"

                            seller_host = make_seller_host(entry['current_domain'])
                            url = (f"https://{seller_host}/myinventory/inventory?fulfilledBy=all&page=1&pageSize=25"
                                   f"&searchField=all&searchTerm={safe_asin}&sort=date_created_desc&status=all")
                            self.status_signal.emit(f"Navigating Tab for Row {entry['row_idx']}: ASIN={entry['asin']} on {seller_host}")
                            try:
                                page.evaluate(f"window.location.href = '{url}';")
                            except: pass
                            entry["stage"] = "waiting_images_link"
                            entry["start_t"] = time.time()
                            continue

                        if stage == "resetting_tab":
                            self.status_signal.emit(f"Resetting Tab for Row {entry['row_idx']}")
                            try:
                                # Clear tab content and wait for it to be ready
                                page.goto("about:blank", wait_until="domcontentloaded", timeout=5000)
                                # Update user on status
                                page.evaluate("document.body.innerHTML = '<h2>Tab cleared. Waiting 2s...</h2>';")
                            except Exception as e:
                                # Log error if tab reset fails, but continue
                                print(f"Could not reset tab for row {entry['row_idx']}: {e}")

                            # Set up delay for the next action
                            entry["stage"] = "delaying_next_row"
                            entry["next_action_time"] = time.time() + 2.0
                            continue

                        if stage == "delaying_next_row":
                            if current_time >= entry["next_action_time"]:
                                # Save the successful domain with the page so the next row can inherit it
                                available_pages.append((page, entry["current_domain"]))
                                if entry in opened:
                                    opened.remove(entry)
                                    tab_was_freed = True
                            continue

                        if stage == "waiting_images_link":
                            # Implement reloads at 10s and 20s, final skip at 30s.
                            elapsed = current_time - entry.get("start_t", current_time)

                            # Final timeout: skip row after 30s
                            if elapsed >= 30:
                                self.status_signal.emit(f"Row {entry['row_idx']} (ASIN={entry['asin']}) exceeded 30s. Skipping.")
                                if entry in opened:
                                    try: opened.remove(entry)
                                    except: pass
                                    tab_was_freed = True
                                available_pages.append((page, entry["current_domain"]))
                                try: page.evaluate("document.body.innerHTML = '<h2>Skipped - took more than 30s!</h2>';")
                                except: pass
                                continue

                            # Second reload at 20s
                            if elapsed >= 20 and not entry.get("reload2_done"):
                                try:
                                    page.evaluate("window.location.reload();")
                                except: pass
                                entry["reload2_done"] = True
                                self.status_signal.emit(f"Reloading (2) Tab for Row {entry['row_idx']}: ASIN={entry['asin']}")
                                continue

                            # First reload at 10s — but first check for page-not-found phrases and skip immediately if found
                            if elapsed >= 10 and not entry.get("reload1_done"):
                                try:
                                    not_found_check = page.evaluate("""() => {
                                        const title = document.title ? document.title.toLowerCase() : '';
                                        const bodyText = document.body ? document.body.innerText.toLowerCase() : '';
                                        const notFoundPhrases = [
                                            'page not found', 'documento no encontrado', 'não foi possível encontrar esta página',
                                            'page introuvable', 'página no encontrada', 'pagina no encontrada',
                                            'impossibile trovare la pagina', 'seite nicht gefunden',
                                            'página não encontrada', 'pagina não encontrada', 'ページが見つかりません'
                                        ];
                                        if (notFoundPhrases.some(phrase => title.includes(phrase))) return true;
                                        if (bodyText.includes('not a functioning page on our site')) return true;
                                        return false;
                                    }""")
                                except:
                                    not_found_check = False

                                if not_found_check:
                                    # Skip immediately if page clearly reports not found
                                    self.status_signal.emit(f"Row {entry['row_idx']} (ASIN={entry['asin']}) reported Page Not Found. Skipping.")
                                    if entry in opened:
                                        try: opened.remove(entry)
                                        except: pass
                                        tab_was_freed = True
                                    available_pages.append((page, entry["current_domain"]))
                                    try: page.evaluate("document.body.innerHTML = '<h2>Not Found - Skipping</h2>';")
                                    except: pass
                                    continue

                                try:
                                    page.evaluate("window.location.reload();")
                                except: pass
                                entry["reload1_done"] = True
                                self.status_signal.emit(f"Reloading (1) Tab for Row {entry['row_idx']}: ASIN={entry['asin']}")
                                continue

                            # For sellercentral flow: try to locate the image using SKU-based xpath
                            result = {"not_found": False, "image_src": None}
                            try:
                                sku_val = entry.get('sku')
                                xpath = f'//*[@id="{sku_val}"]/div/div[4]/div/img'
                                # Try to find element and get src attribute
                                try:
                                    locator = page.locator(f"xpath={xpath}").first
                                    src = locator.get_attribute('src')
                                except Exception:
                                    src = None

                                if src and isinstance(src, str) and src.startswith('http') and 'data:image' not in src:
                                    result['image_src'] = src
                                else:
                                    result['image_src'] = None
                            except Exception:
                                pass

                            if result.get("not_found"):
                                # If page reports not found, skip row immediately
                                self.status_signal.emit(f"Row {entry['row_idx']} (ASIN={entry['asin']}) reported Not Found. Skipping.")
                                if entry in opened:
                                    try: opened.remove(entry)
                                    except: pass
                                    tab_was_freed = True
                                available_pages.append((page, entry["current_domain"]))
                                try: page.evaluate("document.body.innerHTML = '<h2>Not Found - Skipping</h2>';")
                                except: pass
                                continue

                            if result.get("image_src"):
                                src = result["image_src"]
                                
                                # Extract and save
                                try:
                                    import re
                                    img_base = src.split('?', 1)[0]
                                    m = re.match(r'(?P<prefix>.+?)\._[^.]+(?P<ext>\.[a-zA-Z0-9]{2,5})$', img_base)
                                    cleaned_url = m.group('prefix') + m.group('ext') if m else img_base
                                except Exception:
                                    cleaned_url = src

                                self.status_signal.emit(f"Found Image URL for ASIN={entry['asin']}")

                                # Save directly using win32com to avoid overwriting user edits
                                try:
                                    sheet.Cells(entry["row_idx"], 6).Value = cleaned_url
                                    sheet.Cells(entry["row_idx"], 15).Value = "Yes"
                                except Exception as e:
                                    self.status_signal.emit(f"Warning: Failed to update sheet - {e}")
                                
                                entry["stage"] = "resetting_tab"
                                continue

                    if tab_was_freed and len(opened) < batch_size and pending:
                        continue

                    time.sleep(0.05)

                # Close all available pages for this chunk
                for page_tuple in available_pages:
                    try: 
                        if isinstance(page_tuple, tuple):
                            page_tuple[0].close()
                        else:
                            page_tuple.close()
                    except: pass

                # Move to next chunk
                if overall_stop or self.stop_requested:
                    break
                current_row = end_row + 1
                try:
                    max_row = sheet.UsedRange.Rows.Count
                except Exception:
                    pass

            # After processing all chunks (or stop requested)
            try:
                self.controller.disconnect()
            except:
                pass
            if self.stop_requested or overall_stop:
                self.status_signal.emit("Stopped before completing all rows.")
                self.finished_signal.emit(False)
            else:
                self.status_signal.emit("Finished processing Excel file.")
                self.finished_signal.emit(True)
            
        except ImportError:
            self.notification_signal.emit("Dependency Error", "Please install pywin32 (pip install pywin32) to read Excel files.")
            self.finished_signal.emit(False)
        except Exception as e:
            self.notification_signal.emit("Error", f"Excel processing error: {str(e)}")
            self.finished_signal.emit(False)
        finally:
            self.status_signal.emit("Saving and unhiding Excel file...")
            if wb:
                try:
                    subprocess.run(['attrib', '-h', '-s', os.path.abspath(self.excel_path)], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
                except Exception as e:
                    self.status_signal.emit(f"Warning: Could not unhide file - {e}")

                while not getattr(self, "force_quit", False):
                    try:
                        wb.Save()
                        break
                    except Exception as e:
                        self.status_signal.emit("Failed to save. Please close Excel if it is open...")
                        self.retry_save_signal.emit()
                        import time
                        time.sleep(1)

                try: wb.Close(SaveChanges=False)
                except: pass
            
            if excel:
                try: excel.Quit()
                except: pass

            self.status_signal.emit("Finished cleanup.")

    def set_user_logged_in(self):
        """Called by the main UI when the user confirms they have logged into Seller Central."""
        self.user_logged_in = True


os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QLabel, QComboBox, QPushButton, 
                               QLineEdit, QFileDialog, QSystemTrayIcon, QStyle, QMessageBox, QSizePolicy, QSpacerItem, QSpinBox)
from PySide6.QtCore import Qt

# Import the core logic from the existing Instant_Fill script



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Website Launcher")
        self.resize(700, 300)
        self.close_requested_by_user = False
        
        # Center on screen
        try:
            screen = QApplication.primaryScreen().geometry()
            x = (screen.width() - 700) // 2
            y = (screen.height() - 300) // 2
            self.move(x, y)
        except Exception:
            pass
            
        self.controller = ChromeBrowserController(port=9222)
        
        # System Tray Icon for Notifications
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.show()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Add a stretch at the top to push content down
        main_layout.addSpacerItem(QSpacerItem(20, 100, QSizePolicy.Minimum, QSizePolicy.Expanding))
        
        # Dummy font creation since label removed
        font = self.font()
        font.setPointSize(16)

        # --- Region Selection ---
        region_layout = QHBoxLayout()
        region_layout.setAlignment(Qt.AlignCenter)
        region_label = QLabel("Select Amazon Region:")
        region_label.setFont(font)
        
        self.regions = {
            "Amazon US (.com)": "www.amazon.com",
            "Amazon India (.in)": "www.amazon.in",
            "Amazon UK (.co.uk)": "www.amazon.co.uk",
            "Amazon Canada (.ca)": "www.amazon.ca",
            "Amazon Mexico (.com.mx)": "www.amazon.com.mx",
            "Amazon Brazil (.com.br)": "www.amazon.com.br",
            "Amazon Germany (.de)": "www.amazon.de",
            "Amazon France (.fr)": "www.amazon.fr",
            "Amazon Italy (.it)": "www.amazon.it",
            "Amazon Spain (.es)": "www.amazon.es",
            "Amazon Japan (.co.jp)": "www.amazon.co.jp",
            "Amazon Australia (.com.au)": "www.amazon.com.au",
            "Amazon UAE (.ae)": "www.amazon.ae",
            "Amazon Saudi Arabia (.sa)": "www.amazon.sa",
            "Amazon Singapore (.sg)": "www.amazon.sg",
        }
        
        self.region_combo = QComboBox()
        self.region_combo.addItems(list(self.regions.keys()))
        self.region_combo.setFont(font)
        self.region_combo.setMinimumWidth(300)
        self.region_combo.setMinimumHeight(40)
        
        region_layout.addWidget(region_label)
        region_layout.addWidget(self.region_combo)

        # Add tab count next to region dropdown
        tabs_label = QLabel("Tabs:")
        tabs_label.setFont(font)
        self.concurrent_spin = QSpinBox()
        self.concurrent_spin.setFont(font)
        self.concurrent_spin.setMinimum(1)
        self.concurrent_spin.setMaximum(50)
        self.concurrent_spin.setValue(5)
        self.concurrent_spin.setMinimumHeight(40)
        self.concurrent_spin.setMinimumWidth(120)

        region_layout.addWidget(tabs_label)
        region_layout.addWidget(self.concurrent_spin)

        main_layout.addLayout(region_layout)

        # --- File Selection ---
        self.excel_widget = QWidget()
        excel_layout = QVBoxLayout(self.excel_widget)
        excel_layout.setAlignment(Qt.AlignCenter)
        
        file_selection_layout = QHBoxLayout()
        file_selection_layout.setContentsMargins(0, 0, 0, 0)
        
        self.browse_btn = QPushButton("Browse Excel")
        self.browse_btn.setFont(font)
        self.browse_btn.setMinimumHeight(40)
        self.browse_btn.clicked.connect(self.on_browse)
        
        self.excel_path_input = QLineEdit()
        self.excel_path_input.setFont(font)
        self.excel_path_input.setPlaceholderText("Select Excel File First...")
        self.excel_path_input.setReadOnly(True)
        self.excel_path_input.setMinimumWidth(500)
        self.excel_path_input.setMinimumHeight(40)
        
        file_selection_layout.addWidget(self.browse_btn)
        file_selection_layout.addWidget(self.excel_path_input)
        
        excel_layout.addLayout(file_selection_layout)
        main_layout.addWidget(self.excel_widget)
        
        # --- App Control ---
        control_layout = QHBoxLayout()
        control_layout.setAlignment(Qt.AlignCenter)
        
        self.launch_btn = QPushButton("Launch Chrome && Process")
        self.launch_btn.setFont(font)
        self.launch_btn.setMinimumHeight(40)
        self.launch_btn.setEnabled(False) # Disabled until file is selected
        self.launch_btn.clicked.connect(self.on_launch)
        
        self.stop_btn = QPushButton("Stop App")
        self.stop_btn.setFont(font)
        self.stop_btn.setMinimumHeight(40)
        self.stop_btn.setStyleSheet("background-color: #ff4c4c; color: white;")
        self.stop_btn.clicked.connect(self.on_stop)
        
        control_layout.addWidget(self.launch_btn)
        control_layout.addWidget(self.stop_btn)
        
        main_layout.addLayout(control_layout)
        
        
        
        self.status_label = QLabel("")
        self.status_label.setFont(font)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setMinimumHeight(50)
        main_layout.addWidget(self.status_label)
        
        # Add a stretch at the bottom
        main_layout.addSpacerItem(QSpacerItem(20, 100, QSizePolicy.Minimum, QSizePolicy.Expanding))

    def show_coming_soon(self):
        QMessageBox.information(self, "Coming Soon", "This feature is coming soon.")
        
    def on_stop(self):
        self.stop_btn.setEnabled(False)
        self.stop_btn.setText("Stopping...")
        self.stop_btn.setStyleSheet("background-color: #ff9800; color: white;") # Orange for stopping
        if hasattr(self, "process_thread") and self.process_thread.isRunning():
            self.process_thread.stop_requested = True
            self.status_label.setText("Stopping properly and saving Excel... Please wait.")
            # The app will quit when the thread emits finished_signal, see below.
        else:
            self._quit_app()

    def closeEvent(self, event):
        if hasattr(self, "process_thread") and self.process_thread.isRunning():
            self.process_thread.stop_requested = True
            self.status_label.setText("Stopping properly and saving Excel... Please wait.")
            self.close_requested_by_user = True
            event.ignore()
        else:
            self._quit_app()
            event.accept()

    def _quit_app(self, *args):
        # Disconnect Playwright safely and kill the app
        if getattr(self, "controller", None):
            self.controller.disconnect()
        QApplication.quit()

    def on_launch(self):
        self.launch_btn.setEnabled(False)
        self.browse_btn.setEnabled(False)
        
        self.stop_btn.setText("Stop App")
        self.stop_btn.setStyleSheet("background-color: #ff4c4c; color: white;") # Red for Stop App
        self.stop_btn.setEnabled(True)
        try: self.stop_btn.clicked.disconnect()
        except: pass
        self.stop_btn.clicked.connect(self.on_stop)

        # Pass selected region domain so Chrome can open the corresponding Seller Central on launch
        region = self.regions.get(self.region_combo.currentText())
        self.thread = ChromeLaunchThread(self.controller, region_domain=region)
        self.thread.status_signal.connect(self.update_status)
        self.thread.notification_signal.connect(self.show_notification)
        self.thread.finished_signal.connect(self.on_launch_finished)
        self.thread.start()
            
    def update_status(self, text):
        self.status_label.setText(text)
        
    def show_notification(self, title, message):
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.Information, 5000)
        
    def on_launch_finished(self, success, is_running):
        if success:
            excel_path = self.excel_path_input.text()
            domain = self.regions.get(self.region_combo.currentText())
            if excel_path:
                self.process_thread = ExcelProcessThread(self.controller, excel_path, domain, batch_size=self.concurrent_spin.value())
                self.process_thread.status_signal.connect(self.update_status)
                self.process_thread.notification_signal.connect(self.show_notification)
                self.process_thread.retry_save_signal.connect(self.show_save_error_dialog, type=Qt.BlockingQueuedConnection)
                # Connect login request signal to show a blocking login dialog in the UI
                self.process_thread.login_request_signal.connect(self.show_login_dialog, type=Qt.BlockingQueuedConnection)
                self.process_thread.finished_signal.connect(self.on_process_finished)
                self.process_thread.start()
            else:
                self.launch_btn.setEnabled(True)
                self.browse_btn.setEnabled(True)
        else:
            self.launch_btn.setEnabled(True)
            self.browse_btn.setEnabled(True)
            
    def on_browse(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls *.xlsm)")
        if file_path:
            self.excel_path_input.setText(file_path)
            self.launch_btn.setEnabled(True)

    def show_save_error_dialog(self):
        msg = QMessageBox(self)
        msg.setWindowFlags(msg.windowFlags() | Qt.WindowStaysOnTopHint)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Excel File Open")
        msg.setText("The Excel sheet is currently open.\nPlease close the file in Excel and click OK to continue.")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec()

    def show_login_dialog(self, title, message):
        msg = QMessageBox(self)
        msg.setWindowFlags(msg.windowFlags() | Qt.WindowStaysOnTopHint)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.Ok)
        res = msg.exec()
        if res == QMessageBox.Ok:
            try:
                # Inform the worker thread that the user confirmed login
                if hasattr(self, 'process_thread'):
                    self.process_thread.set_user_logged_in()
            except Exception:
                pass

    def on_process_finished(self, success):
        if hasattr(self, "process_thread"):
            self.process_thread.wait() # Ensure thread has completely terminated

        if getattr(self, "close_requested_by_user", False):
            self._quit_app()
            return

        self.launch_btn.setEnabled(True)
        self.browse_btn.setEnabled(True)
        
        self.stop_btn.setText("Close")
        self.stop_btn.setStyleSheet("background-color: #4CAF50; color: white;") # Green for Close
        self.stop_btn.setEnabled(True)
        try: self.stop_btn.clicked.disconnect()
        except: pass
        self.stop_btn.clicked.connect(self._quit_app)

        if success and not getattr(self.process_thread, "stop_requested", False):
            self.show_notification("Process Complete", "Finished processing all valid rows in the Excel file.")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    try:
        import playwright
    except ImportError:
        print("Playwright not found!")
        print("Please install it using: pip install playwright")
        print("Then run: playwright install chromium")
        sys.exit(1)
    
    main()