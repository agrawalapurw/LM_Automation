"""
SAP CRM Lookup helper
- Open SAP CRM
- Navigate to Design Registrations (outer CRMApplicationFrame)
- For each company:
  - Switch to search form frame, click Clear, operator='contains', fill End Customer Name
  - Trigger Search (click + Enter)
  - Switch to results frame and wait for table / 'No result found'
  - Go to last page; pick most recent 'Approved' Sold-to-Party Name; if none, step back
  - If nothing found across candidates, return None (Excel leaves cell empty)
"""
import time
from pathlib import Path
from typing import Optional, List, Dict, Tuple
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

from webdriver_manager.chrome import ChromeDriverManager

SAP_URL = "https://sappc1lb.eu.infineon.com/sap(bD1lbiZjPTEwMCZkPW1pbg==)/bc/bsp/sap/crm_ui_start/default.htm"

MAX_FIELD_LEN = 40
MAX_PAGES_BACKWARD = 3
WAIT_SHORT = 3
WAIT_MED = 12

COMPANY_SUFFIXES = {
    "gmbh", "ag", "se", "ltd", "limited", "inc", "corp", "corporation",
    "llc", "plc", "sa", "srl", "bv", "nv", "kg", "ohg", "gbr",
    "co", "company", "group", "holding", "holdings", "international"
}


class SAPCRMLookup:
    def __init__(self, headless: bool = False):
        self.headless = headless
        self.driver: Optional[WebDriver] = None
        self._cache: Dict[str, Optional[str]] = {}

    def start(self):
        options = webdriver.ChromeOptions()
        if self.headless:
            options.add_argument("--headless=new")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-gpu")
        options.add_argument("--remote-allow-origins=*")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.page_load_strategy = "eager"

        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.maximize_window()

        try:
            caps = self.driver.capabilities
            print(f"Chrome: {caps.get('browserVersion')}, "
                  f"ChromeDriver: {caps.get('chrome', {}).get('chromedriverVersion', '')}")
        except Exception:
            pass

    def stop(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    # ---------------- Navigation ----------------

    def navigate_to_design_registrations(self):
        d = self.driver
        if d is None:
            raise RuntimeError("Driver not started.")

        d.get(SAP_URL)
        WebDriverWait(d, 40).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1.0)

        # Enter outer CRM frame
        if not self._switch_into_crm_root_frame(timeout=30):
            self._save_screenshot("sap_nav_fail_root.png")
            raise RuntimeError("Could not switch into CRMApplicationFrame")

        # If already on search page, keep focus here
        if self._switch_to_frame_with_element_kept(
            ["//input[@title='Enter the value of criterion End Customer Name']"], timeout=8
        ):
            print("SAP CRM: Search page already loaded.")
            return

        # Click Design Registrations directly (preferred)
        if not self._click_link_any_frame_recursive([
            "//a[normalize-space()='Design Registrations']",
            "//a[@aria-label='Design Registration Search' or @title='Design Registration Search']",
            "//a[@id and contains(@id,'HT-DR-SR')]",
            "//a[contains(normalize-space(), 'Design Registrations')]",
        ], timeout=25):
            # Fallback: Sales Center then Design Registrations
            self._click_link_any_frame_recursive([
                "//a[normalize-space()='Sales Center']",
                "//a[@aria-label='Sales Center' or @title='Sales Center']",
                "//a[@id and contains(@id,'HT-CHM-SLS')]",
                "//span[normalize-space()='Sales Center']",
            ], timeout=20)
            if not self._click_link_any_frame_recursive([
                "//a[normalize-space()='Design Registrations']",
                "//a[@aria-label='Design Registration Search' or @title='Design Registration Search']",
                "//a[@id and contains(@id,'HT-DR-SR')]",
                "//a[contains(normalize-space(), 'Design Registrations')]",
            ], timeout=25):
                self._save_screenshot("sap_nav_fail_design_reg.png")
                self._dump_frame_tree("sap_frames.txt")
                raise RuntimeError("Could not find 'Design Registrations' link")

        time.sleep(1.0)

        # Ensure search field input is present; keep focus
        if not self._switch_to_frame_with_element_kept(
            ["//input[@title='Enter the value of criterion End Customer Name']"], timeout=30
        ):
            self._save_screenshot("sap_search_page_fail.png")
            self._dump_frame_tree("sap_frames.txt")
            raise RuntimeError("Could not locate 'End Customer Name' input")

    # ---------------- Frame utilities ----------------

    def _switch_into_crm_root_frame(self, timeout: int = 30) -> bool:
        d = self.driver
        end = time.time() + timeout
        while time.time() < end:
            try:
                d.switch_to.default_content()
                frames = d.find_elements(By.XPATH, "//iframe[@id='CRMApplicationFrame' or @name='CRMApplicationFrame']")
                if not frames:
                    frames = d.find_elements(By.XPATH, "//iframe[contains(@id,'CRMApplicationFrame') or contains(@name,'CRMApplicationFrame')]")
                if frames:
                    d.switch_to.frame(frames[0])
                    return True
            except Exception:
                pass
            time.sleep(0.5)
        return False

    def _reset_to_crm_root(self) -> bool:
        d = self.driver
        try:
            d.switch_to.default_content()
            frames = d.find_elements(By.XPATH, "//iframe[@id='CRMApplicationFrame' or @name='CRMApplicationFrame']")
            if not frames:
                frames = d.find_elements(By.XPATH, "//iframe[contains(@id,'CRMApplicationFrame') or contains(@name,'CRMApplicationFrame')]")
            if frames:
                d.switch_to.frame(frames[0])
                return True
        except Exception:
            pass
        return False

    def _switch_to_frame_with_element_kept(self, xpaths: List[str], timeout: int = 20) -> bool:
        if not self._reset_to_crm_root():
            return False
        end = time.time() + timeout
        while time.time() < end:
            if self._find_frame_with_element_dfs(xpaths):
                return True
            time.sleep(0.4)
        return False

    def _find_frame_with_element_dfs(self, xpaths: List[str]) -> bool:
        d = self.driver
        for xp in xpaths:
            try:
                if d.find_elements(By.XPATH, xp):
                    return True
            except Exception:
                pass

        frames = d.find_elements(By.TAG_NAME, "iframe") + d.find_elements(By.TAG_NAME, "frame")
        for frm in frames:
            try:
                d.switch_to.frame(frm)
                if self._find_frame_with_element_dfs(xpaths):
                    return True
                d.switch_to.parent_frame()
            except Exception:
                self._reset_to_crm_root()
        return False

    def _click_link_any_frame_recursive(self, xpaths: List[str], timeout: int = 25) -> bool:
        end = time.time() + timeout
        while time.time() < end:
            if not self._reset_to_crm_root():
                return False
            if self._try_click_in_current_context(xpaths, wait_seconds=3):
                return True
            if self._recurse_frames_click(xpaths):
                return True
            time.sleep(0.4)
        return False

    def _recurse_frames_click(self, xpaths: List[str]) -> bool:
        d = self.driver
        frames = d.find_elements(By.TAG_NAME, "iframe") + d.find_elements(By.TAG_NAME, "frame")
        for frm in frames:
            try:
                d.switch_to.frame(frm)
                if self._try_click_in_current_context(xpaths, wait_seconds=2):
                    return True
                if self._recurse_frames_click(xpaths):
                    return True
                d.switch_to.parent_frame()
            except Exception:
                self._reset_to_crm_root()
        return False

    def _try_click_in_current_context(self, xpaths: List[str], wait_seconds: int = 2) -> bool:
        d = self.driver
        wait = WebDriverWait(d, wait_seconds)
        for xp in xpaths:
            try:
                elem = wait.until(EC.presence_of_element_located((By.XPATH, xp)))
                d.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", elem)
                try:
                    WebDriverWait(d, 1).until(EC.element_to_be_clickable((By.XPATH, xp)))
                    elem.click()
                except Exception:
                    d.execute_script("arguments[0].click();", elem)
                return True
            except Exception:
                continue
        return False

    # ---------------- Diagnostics ----------------

    def _dump_frame_tree(self, out_file: str):
        d = self.driver
        lines = []

        def dfs(prefix: str):
            try:
                url = d.execute_script("return document.location.href;")
            except Exception:
                url = "n/a"
            try:
                title = d.execute_script("return document.title;")
            except Exception:
                title = "n/a"
            lines.append(f"{prefix}- URL: {url} | Title: {title}")
            frames = d.find_elements(By.TAG_NAME, "iframe") + d.find_elements(By.TAG_NAME, "frame")
            for i, frm in enumerate(frames):
                try:
                    d.switch_to.frame(frm)
                    dfs(prefix + "  ")
                    d.switch_to.parent_frame()
                except Exception:
                    lines.append(f"{prefix}  (frame {i}: switch failed)")
                    self._reset_to_crm_root()

        try:
            self._reset_to_crm_root()
            dfs("")
        except Exception:
            pass
        try:
            Path(out_file).write_text("\n".join(lines), encoding="utf-8")
            print(f"Frame tree written to {out_file}")
        except Exception:
            pass

    def _save_screenshot(self, filename: str):
        try:
            if self.driver:
                self.driver.save_screenshot(str(Path.cwd() / filename))
                print(f"Saved screenshot: {filename}")
        except Exception:
            pass

    def _save_current_frame_html(self, filename: str):
        try:
            if self.driver:
                html = self.driver.page_source
                Path(filename).write_text(html, encoding="utf-8")
                print(f"Saved HTML: {filename}")
        except Exception:
            pass

    # ---------------- Lookup flow ----------------

    def lookup(self, company_name: str) -> Optional[str]:
        key = (company_name or "").strip()
        if not key:
            return None
        if key in self._cache:
            return self._cache[key]

        result = self._lookup_internal(company_name)
        self._cache[key] = result
        return result

    def _lookup_internal(self, company_name: str) -> Optional[str]:
        # Try contains with progressively shorter candidates
        for cand in self._generate_candidates(company_name):
            # Ensure we are in the search form frame
            if not self._switch_to_frame_with_element_kept(
                ["//input[@title='Enter the value of criterion End Customer Name']"], timeout=WAIT_MED
            ):
                return None
            # Clear previous state and fill form
            self._click_clear()
            self._set_operator("contains", "CP")
            self._fill_end_customer_name(cand)
            # Trigger search (click + Enter fallback)
            self._click_search()
            self._press_enter_in_value_input()

            # Switch to results frame
            self._switch_to_results_frame()
            # Wait for results
            self._wait_for_results_ready(timeout=WAIT_MED)

            # Find most recent Approved from last page
            sold_to = self._find_latest_approved_from_last_page(max_pages_back=MAX_PAGES_BACKWARD)
            if sold_to:
                return sold_to
            # If none found, try next shorter candidate
            time.sleep(0.3)

        # Fallback: starts with first 1–3 words
        tokens = self._normalize_company_name(company_name).split()
        for n in range(1, min(3, len(tokens)) + 1):
            cand = " ".join(tokens[:n])[:MAX_FIELD_LEN]
            if not self._switch_to_frame_with_element_kept(
                ["//input[@title='Enter the value of criterion End Customer Name']"], timeout=WAIT_MED
            ):
                return None
            self._click_clear()
            self._set_operator("starts with", "SW")
            self._fill_end_customer_name(cand)
            self._click_search()
            self._press_enter_in_value_input()
            self._switch_to_results_frame()
            self._wait_for_results_ready(timeout=WAIT_MED)
            sold_to = self._find_latest_approved_from_last_page(max_pages_back=MAX_PAGES_BACKWARD)
            if sold_to:
                return sold_to
            time.sleep(0.3)

        # Not found; leave Excel cell empty
        return None

    # ---------------- Page interaction helpers ----------------

    def _set_operator(self, operator_label: str, operator_code: str, timeout: int = 15) -> None:
        d = self.driver
        wait = WebDriverWait(d, timeout)
        op_input = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@title='Choose the operator of criterion End Customer Name']")
            )
        )
        try:
            btn = op_input.find_element(By.XPATH, "../following-sibling::td/a[contains(@id,'OPERATOR-btn')]")
            d.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            try:
                btn.click()
            except Exception:
                d.execute_script("arguments[0].click();", btn)
            time.sleep(0.3)
            opt = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//a[normalize-space()='{operator_label}']")
                )
            )
            try:
                opt.click()
            except Exception:
                d.execute_script("arguments[0].click();", opt)
        except Exception:
            # Fallback: set invisible key + visible value via JS
            try:
                hidden = d.find_element(
                    By.XPATH, "//input[contains(@id,'OPERATOR__key') and @name and contains(@name,'OPERATOR')]"
                )
                d.execute_script("arguments[0].value=arguments[1];", hidden, operator_code)
                d.execute_script("arguments[0].value=arguments[1];", op_input, operator_label)
            except Exception as e:
                raise RuntimeError(f"Could not set operator '{operator_label}': {e}")

    def _fill_end_customer_name(self, name: str, timeout: int = 15) -> None:
        d = self.driver
        wait = WebDriverWait(d, timeout)
        value_input = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@title='Enter the value of criterion End Customer Name']")
            )
        )
        d.execute_script("arguments[0].scrollIntoView({block:'center'});", value_input)
        value_input.clear()
        value_input.send_keys(name[:MAX_FIELD_LEN])

    def _click_search(self, timeout: int = 15) -> None:
        d = self.driver
        wait = WebDriverWait(d, timeout)
        search_btn = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//a[.//b[normalize-space()='Search']]")
            )
        )
        d.execute_script("arguments[0].scrollIntoView({block:'center'});", search_btn)
        try:
            WebDriverWait(d, 2).until(EC.element_to_be_clickable((By.XPATH, "//a[.//b[normalize-space()='Search']]")))
            search_btn.click()
        except Exception:
            d.execute_script("arguments[0].click();", search_btn)

    def _press_enter_in_value_input(self, timeout: int = 5) -> None:
        """
        Some SAP HTMLB pages trigger search on Enter. Use this as a secondary trigger.
        """
        d = self.driver
        try:
            value_input = WebDriverWait(d, timeout).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@title='Enter the value of criterion End Customer Name']")
                )
            )
            value_input.send_keys(Keys.ENTER)
        except Exception:
            pass

    def _click_clear(self, timeout: int = 8) -> None:
        """
        Click 'Clear' to reset prior criteria/results.
        """
        d = self.driver
        try:
            clear_btn = WebDriverWait(d, timeout).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//a[.//b[normalize-space()='Clear']]")
                )
            )
            d.execute_script("arguments[0].scrollIntoView({block:'center'});", clear_btn)
            try:
                clear_btn.click()
            except Exception:
                d.execute_script("arguments[0].click();", clear_btn)
            time.sleep(0.3)
        except Exception:
            # If Clear isn't present or fails, continue
            pass

    # ---------------- Results frame + parsing ----------------

    def _switch_to_results_frame(self):
        """
        After search, results can render in a different inner frame.
        Switch to the frame containing the results table or 'No result found'.
        """
        # Try current frame quick path
        if self._table_present() or self._no_result_present():
            return

        # Search frames for table or result banner
        found = self._switch_to_frame_with_element_kept([
            "//table[.//th[contains(normalize-space(),'Registration Date')]]",
            "//*[contains(normalize-space(),'Result List')]",
            "//*[contains(normalize-space(),'No result found')]",
        ], timeout=WAIT_MED)
        if not found:
            self._save_current_frame_html("sap_results.html")

    def _wait_for_results_ready(self, timeout: int = WAIT_MED) -> None:
        d = self.driver
        end = time.time() + timeout
        while time.time() < end:
            if self._table_present() or self._no_result_present():
                return
            time.sleep(0.3)
        self._save_current_frame_html("sap_results.html")

    def _table_present(self) -> bool:
        try:
            self._find_results_table(timeout=2)
            return True
        except Exception:
            return False

    def _no_result_present(self) -> bool:
        d = self.driver
        try:
            return bool(d.find_elements(By.XPATH, "//*[contains(normalize-space(),'No result found')]"))
        except Exception:
            return False

    def _parse_date(self, text: str) -> Optional[datetime]:
        text = (text or "").strip()
        for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt)
            except Exception:
                continue
        return None

    def _find_results_table(self, timeout: int = 10):
        d = self.driver
        wait = WebDriverWait(d, timeout)
        return wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//table[.//th[contains(normalize-space(),'Registration Date')]]")
            )
        )

    def _normalize_header(self, text: str) -> str:
        return " ".join((text or "").strip().lower().split())

    def _cell_text_js(self, elem) -> str:
        d = self.driver
        try:
            txt = d.execute_script("return arguments[0].innerText || arguments[0].textContent || '';", elem)
            return (txt or "").strip()
        except Exception:
            return (elem.text or "").strip()

    def _get_header_index(self, table, candidates: List[str]) -> Optional[int]:
        headers = table.find_elements(By.XPATH, ".//th")
        normalized = [
            (idx, self._normalize_header(self._cell_text_js(th)))
            for idx, th in enumerate(headers, start=1)
        ]
        cand_norm = [self._normalize_header(c) for c in candidates]
        for idx, txt in normalized:
            for c in cand_norm:
                if txt == c:
                    return idx
        for idx, txt in normalized:
            for c in cand_norm:
                if c in txt or txt in c:
                    return idx
        return None

    def _get_rows_current_page(self) -> List[Tuple[Optional[datetime], str, str]]:
        if self._no_result_present():
            return []
        try:
            table = self._find_results_table(timeout=WAIT_MED)
        except Exception:
            self._save_current_frame_html("sap_results.html")
            return []

        date_col = self._get_header_index(table, ["Registration Date"])
        status_col = self._get_header_index(table, ["Registration Status", "Status"])
        sold_col = self._get_header_index(table, ["Sold-to-Party Name", "Sold-To Party Name", "Sold-To Party", "Sold-to-Party"])
        if not (date_col and status_col and sold_col):
            self._save_current_frame_html("sap_results.html")
            return []

        rows = []
        for tr in table.find_elements(By.XPATH, ".//tbody/tr[td]"):
            try:
                date_td = tr.find_element(By.XPATH, f"./td[{date_col}]")
                status_td = tr.find_element(By.XPATH, f"./td[{status_col}]")
                sold_td = tr.find_element(By.XPATH, f"./td[{sold_col}]")
                date_txt = self._cell_text_js(date_td)
                status_txt = self._cell_text_js(status_td)
                sold_txt = self._cell_text_js(sold_td)
                dt = self._parse_date(date_txt)
                rows.append((dt, status_txt, sold_txt))
            except Exception:
                continue
        return rows

    def _go_to_next_page(self) -> bool:
        d = self.driver
        try:
            fwd = d.find_element(By.XPATH, "//a[normalize-space()='Forward' or normalize-space()='>' or contains(normalize-space(),'Forward')]")
            d.execute_script("arguments[0].scrollIntoView({block:'center'});", fwd)
            try:
                fwd.click()
            except Exception:
                d.execute_script("arguments[0].click();", fwd)
            time.sleep(0.7)
            return True
        except Exception:
            return False

    def _go_to_prev_page(self) -> bool:
        d = self.driver
        try:
            back = d.find_element(By.XPATH, "//a[normalize-space()='Back' or normalize-space()='<']")
            d.execute_script("arguments[0].scrollIntoView({block:'center'});", back)
            try:
                back.click()
            except Exception:
                d.execute_script("arguments[0].click();", back)
            time.sleep(0.7)
            return True
        except Exception:
            return False

    def _paginate_to_last_page(self, max_forwards: int = 50) -> None:
        cnt = 0
        while cnt < max_forwards and self._go_to_next_page():
            cnt += 1

    def _find_latest_approved_from_last_page(self, max_pages_back: int = MAX_PAGES_BACKWARD) -> Optional[str]:
        self._paginate_to_last_page()

        best_date = None
        best_sold = None

        pages_checked = 0
        while pages_checked <= max_pages_back:
            pages_checked += 1
            rows = self._get_rows_current_page()
            for dt, status, sold in rows:
                if status.strip().lower() == "approved" and dt:
                    if best_date is None or dt > best_date:
                        best_date = dt
                        best_sold = sold
            if best_sold:
                return best_sold
            if not self._go_to_prev_page():
                break

        return best_sold

    # ---------------- Candidate generation ----------------

    def _normalize_company_name(self, name: str) -> str:
        if not name:
            return ""
        s = str(name)
        for ch in ("(", ")", "[", "]", "{", "}"):
            s = s.replace(ch, " ")
        s = s.strip()
        tokens = [t for t in s.replace(",", " ").replace(".", " ").split()
                  if t.lower() not in COMPANY_SUFFIXES]
        s = " ".join(tokens)
        s = " ".join(s.split())
        return s

    def _generate_candidates(self, company: str) -> List[str]:
        base = self._normalize_company_name(company)
        if not base:
            return []
        tokens = base.split()
        candidates = [base[:MAX_FIELD_LEN]]
        # First N words (5 -> 1)
        max_n = min(5, len(tokens))
        for n in range(max_n, 0, -1):
            cand = " ".join(tokens[:n]).strip()
            if cand and cand[:MAX_FIELD_LEN] not in candidates:
                candidates.append(cand[:MAX_FIELD_LEN])
        return candidates