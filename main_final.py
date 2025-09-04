import sys
import sqlite3
import csv
import re
import time
import uuid
from datetime import datetime
from urllib.parse import quote_plus, urlparse

# Add these imports for XLS export
import xlwt
from openpyxl import Workbook

from PyQt5.QtWidgets import (
    QApplication, QWidget, QHBoxLayout, QVBoxLayout, QLabel,
    QLineEdit, QPushButton, QTextEdit, QProgressBar, QTableWidget,
    QTableWidgetItem, QFileDialog, QMessageBox, QComboBox, QTabWidget, QSpinBox,
    QGroupBox
)
from PyQt5.QtCore import QUrl, QTimer, QThread, pyqtSignal
from PyQt5.QtWebEngineWidgets import QWebEngineView

# Database files
MAPS_DB_FILE = "businesses.db"
JUSTDIAL_DB_FILE = "justdial_businesses.db"


def normalize_table_name(keyword, location):
    """Turn keyword+location into a safe SQLite table name"""
    safe = f"{keyword}_{location}".lower().replace(" ", "_")
    return re.sub(r"[^a-z0-9_]", "", safe)


class JustDialScraper(QThread):
    progress_signal = pyqtSignal(int)
    status_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()
    data_signal = pyqtSignal(dict)
    
    def __init__(self, web_view, max_listings):
        super().__init__()
        self.web_view = web_view
        self.max_listings = max_listings
        self._is_running = True
        
    def run(self):
        try:
            self.status_signal.emit("Starting JustDial scraping from embedded browser...")
            
            # Execute JavaScript to extract data from the current page
            js_code = """
            (function() {
                // Function to extract data from a single listing
                function extractListingData(listing) {
                    // Extract name
                    let name = "N/A";
                    try {
                        const nameElement = listing.querySelector("h2.resultbox_title, h3.resultbox_title, .resultbox_title_anchor, .complist_title");
                        if (nameElement) name = nameElement.innerText.trim();
                    } catch(e) {}
                    
                    // Skip if no name found
                    if (!name || name === "N/A") return null;
                    
                    // Extract address
                    let address = "N/A";
                    try {
                        const addressSelectors = [
                            ".resultbox_locat_icon + .locatcity",
                            ".cont_fl_addr",
                            ".add_icon_link",
                            "address",
                            ".resultbox_address div",
                            ".locatcity"
                        ];
                        
                        for (const selector of addressSelectors) {
                            const element = listing.querySelector(selector);
                            if (element && element.innerText.trim()) {
                                address = element.innerText.trim();
                                break;
                            }
                        }
                    } catch(e) {}
                    
                    // Extract phone
                    let phone = "N/A";
                    try {
                        const phoneElements = listing.querySelectorAll(".callcontent, .callNowAnchor, .greenfill_animate span");
                        for (const element of phoneElements) {
                            const text = element.innerText.trim();
                            if (/^[\\d\\s\\+\\-\\(\\)]{10,}$/.test(text)) {
                                phone = text;
                                break;
                            }
                        }
                        
                        if (phone === "N/A") {
                            const allElements = listing.querySelectorAll('*');
                            for (const element of allElements) {
                                const text = element.innerText.trim();
                                if (/^[\\d\\s\\+\\-\\(\\)]{10,}$/.test(text) && text.length >= 10) {
                                    phone = text;
                                    break;
                                }
                            }
                        }
                    } catch(e) {}
                    
                    // Extract website
                    let website = "N/A";
                    let website_status = "Unknown";
                    try {
                        const websiteElements = listing.querySelectorAll("a[href*='http']");
                        for (const element of websiteElements) {
                            const href = element.href;
                            if (href && !href.includes('justdial') && !href.includes('tel:') && 
                                !href.includes('mailto:') && !href.startsWith('javascript:')) {
                                website = href;
                                website_status = "Online";
                                break;
                            }
                        }
                    } catch(e) {}
                    
                    // Extract rating
                    let rating = "N/A";
                    try {
                        const ratingElement = listing.querySelector(".resultbox_totalrate, .star_m, .green-box, .rating");
                        if (ratingElement) rating = ratingElement.innerText.trim();
                    } catch(e) {}
                    
                    // Extract votes
                    let votes = "N/A";
                    try {
                        const votesElement = listing.querySelector(".resultbox_countrate, .rt_count, .votes, .review-count");
                        if (votesElement) votes = votesElement.innerText.trim();
                    } catch(e) {}
                    
                    return {
                        name: name,
                        address: address,
                        phone: phone,
                        website: website,
                        website_status: website_status,
                        rating: rating,
                        votes: votes
                    };
                }
                
                // Find all listings on the page
                const listings = document.querySelectorAll("li.cntanr, div.resultbox, section.resultbox_listing");
                const results = [];
                const seenNames = new Set();
                
                for (const listing of listings) {
                    const data = extractListingData(listing);
                    if (data && data.name !== "N/A" && !seenNames.has(data.name)) {
                        seenNames.add(data.name);
                        results.push(data);
                    }
                }
                
                return results;
            })();
            """
            
            # Execute the JavaScript and process results
            def process_results(result):
                if not result or not isinstance(result, list):
                    self.status_signal.emit("No listings found or invalid page structure.")
                    self.finished_signal.emit()
                    return
                
                self.status_signal.emit(f"Found {len(result)} listings. Processing...")
                
                for i, data in enumerate(result):
                    if not self._is_running or i >= self.max_listings:
                        break
                    
                    # Generate unique ID and timestamp
                    business_id = str(uuid.uuid4())
                    scraped_at = datetime.now().isoformat()
                    
                    # Emit the data
                    data_with_meta = {
                        'id': business_id,
                        'name': data.get('name', 'N/A'),
                        'address': data.get('address', 'N/A'),
                        'phone': data.get('phone', 'N/A'),
                        'website': data.get('website', 'N/A'),
                        'website_status': data.get('website_status', 'Unknown'),
                        'rating': data.get('rating', 'N/A'),
                        'votes': data.get('votes', 'N/A'),
                        'scraped_at': scraped_at
                    }
                    
                    self.data_signal.emit(data_with_meta)
                    self.progress_signal.emit(int(((i + 1) / min(len(result), self.max_listings)) * 100))
                    
                    # Small delay to avoid overwhelming the UI
                    time.sleep(0.1)
                
                self.status_signal.emit(f"Finished processing {min(len(result), self.max_listings)} listings.")
                self.finished_signal.emit()
            
            # Execute the JavaScript
            self.web_view.page().runJavaScript(js_code, process_results)
            
        except Exception as e:
            self.status_signal.emit(f"Error in JustDial scraper: {e}")
            self.finished_signal.emit()
            
    def stop(self):
        self._is_running = False


class MapsScraperGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Advanced Google Maps & JustDial Scraper")
        self.setMinimumSize(1100, 650)

        # Main tabs
        self.tabs = QTabWidget()
        self.maps_scraper_tab = QWidget()
        self.justdial_scraper_tab = QWidget()
        self.viewer_tab = QWidget()
        self.tabs.addTab(self.maps_scraper_tab, "Maps Scraper")
        self.tabs.addTab(self.justdial_scraper_tab, "JustDial Scraper")
        self.tabs.addTab(self.viewer_tab, "Data Viewer")

        # --- MAPS SCRAPER TAB ---
        self.web = QWebEngineView()
        self.links = []
        self.links_index = 0
        self.current_keyword = ""
        self.current_location = ""

        # Left controls for Maps
        maps_left_layout = QVBoxLayout()
        maps_left_layout.addWidget(QLabel("<b>Keyword</b>"))
        self.keyword_input = QLineEdit()
        maps_left_layout.addWidget(self.keyword_input)

        maps_left_layout.addWidget(QLabel("<b>Location</b>"))
        self.location_input = QLineEdit()
        maps_left_layout.addWidget(self.location_input)

        self.load_maps_btn = QPushButton("Load Maps")
        self.load_maps_btn.clicked.connect(self.load_maps)
        maps_left_layout.addWidget(self.load_maps_btn)

        self.collect_links_btn = QPushButton("Collect Links")
        self.collect_links_btn.clicked.connect(self.collect_links)
        maps_left_layout.addWidget(self.collect_links_btn)

        maps_left_layout.addWidget(QLabel("Max Links to Extract"))
        self.max_links_input = QSpinBox()
        self.max_links_input.setRange(1, 500)
        self.max_links_input.setValue(30)
        maps_left_layout.addWidget(self.max_links_input)

        self.start_scrape_btn = QPushButton("Start Scrape")
        self.start_scrape_btn.clicked.connect(self.start_scrape)
        maps_left_layout.addWidget(self.start_scrape_btn)

        maps_left_layout.addSpacing(10)
        self.maps_status = QTextEdit()
        self.maps_status.setReadOnly(True)
        self.maps_status.setMaximumHeight(150)
        maps_left_layout.addWidget(self.maps_status)

        self.maps_progress = QProgressBar()
        maps_left_layout.addWidget(self.maps_progress)
        maps_left_layout.addStretch()

        maps_main_layout = QHBoxLayout()
        maps_main_layout.addLayout(maps_left_layout, 1)
        maps_main_layout.addWidget(self.web, 3)
        self.maps_scraper_tab.setLayout(maps_main_layout)

        # --- JUSTDIAL SCRAPER TAB ---
        self.justdial_web = QWebEngineView()
        self.justdial_links = []
        self.justdial_links_index = 0
        self.justdial_keyword = ""
        self.justdial_location = ""
        self.justdial_url = ""
        self.justdial_scraper = None

        # Left controls for JustDial
        justdial_left_layout = QVBoxLayout()
        
        # URL input for direct access (useful for logged-in sessions)
        justdial_left_layout.addWidget(QLabel("<b>JustDial URL (Optional)</b>"))
        self.justdial_url_input = QLineEdit()
        self.justdial_url_input.setPlaceholderText("https://www.justdial.com/Pune/Restaurants/...")
        justdial_left_layout.addWidget(self.justdial_url_input)

        justdial_left_layout.addWidget(QLabel("<b>Keyword</b>"))
        self.justdial_keyword_input = QLineEdit()
        justdial_left_layout.addWidget(self.justdial_keyword_input)

        justdial_left_layout.addWidget(QLabel("<b>Location</b>"))
        self.justdial_location_input = QLineEdit()
        justdial_left_layout.addWidget(self.justdial_location_input)

        self.load_justdial_btn = QPushButton("Load JustDial")
        self.load_justdial_btn.clicked.connect(self.load_justdial)
        justdial_left_layout.addWidget(self.load_justdial_btn)

        # Scroll to load more content
        self.scroll_justdial_btn = QPushButton("Scroll to Load More")
        self.scroll_justdial_btn.clicked.connect(self.scroll_justdial)
        justdial_left_layout.addWidget(self.scroll_justdial_btn)

        justdial_left_layout.addWidget(QLabel("Max Listings to Extract"))
        self.justdial_max_listings_input = QSpinBox()
        self.justdial_max_listings_input.setRange(1, 200)
        self.justdial_max_listings_input.setValue(30)
        justdial_left_layout.addWidget(self.justdial_max_listings_input)

        self.start_justdial_scrape_btn = QPushButton("Extract Data")
        self.start_justdial_scrape_btn.clicked.connect(self.start_justdial_scrape)
        justdial_left_layout.addWidget(self.start_justdial_scrape_btn)
        
        self.stop_justdial_scrape_btn = QPushButton("Stop Extraction")
        self.stop_justdial_scrape_btn.clicked.connect(self.stop_justdial_scrape)
        self.stop_justdial_scrape_btn.setEnabled(False)
        justdial_left_layout.addWidget(self.stop_justdial_scrape_btn)

        justdial_left_layout.addSpacing(10)
        self.justdial_status = QTextEdit()
        self.justdial_status.setReadOnly(True)
        self.justdial_status.setMaximumHeight(150)
        justdial_left_layout.addWidget(self.justdial_status)

        self.justdial_progress = QProgressBar()
        justdial_left_layout.addWidget(self.justdial_progress)
        justdial_left_layout.addStretch()

        justdial_main_layout = QHBoxLayout()
        justdial_main_layout.addLayout(justdial_left_layout, 1)
        justdial_main_layout.addWidget(self.justdial_web, 3)
        self.justdial_scraper_tab.setLayout(justdial_main_layout)

        # --- VIEWER TAB ---
        vlayout = QVBoxLayout()
        vlayout.addWidget(QLabel("<b>Select Data Source</b>"))
        self.source_combo = QComboBox()
        self.source_combo.addItems(["Google Maps", "JustDial"])
        self.source_combo.currentIndexChanged.connect(self.refresh_keyword_tables)
        vlayout.addWidget(self.source_combo)
        
        vlayout.addWidget(QLabel("<b>Select Keyword Table</b>"))
        self.keyword_combo = QComboBox()
        self.keyword_combo.currentIndexChanged.connect(self.load_selected_table)
        vlayout.addWidget(self.keyword_combo)

        self.table = QTableWidget()
        vlayout.addWidget(self.table)

        # Create a horizontal layout for export buttons
        export_layout = QHBoxLayout()
        
        csv_export_btn = QPushButton("Export to CSV")
        csv_export_btn.clicked.connect(self.export_selected_csv)
        export_layout.addWidget(csv_export_btn)
        
        xls_export_btn = QPushButton("Export to XLS")
        xls_export_btn.clicked.connect(self.export_selected_xls)
        export_layout.addWidget(xls_export_btn)
        
        vlayout.addLayout(export_layout)
        self.viewer_tab.setLayout(vlayout)

        # --- MAIN LAYOUT ---
        outer = QVBoxLayout()
        outer.addWidget(self.tabs)
        self.setLayout(outer)

        # State variables
        self.scraping = False
        self.justdial_scraping = False
        self._next_action = None
        self.web.loadFinished.connect(self._on_load_finished)
        self.justdial_web.loadFinished.connect(self._on_justdial_load_finished)

        # Load available tables
        self.refresh_keyword_tables()

    def log(self, msg, scraper="maps"):
        if scraper == "maps":
            self.maps_status.append(msg)
        else:
            self.justdial_status.append(msg)

    def get_current_table(self, scraper="maps"):
        if scraper == "maps":
            return normalize_table_name(self.current_keyword, self.current_location)
        else:
            return normalize_table_name(self.justdial_keyword, self.justdial_location)

    def ensure_table(self, keyword, location, scraper="maps"):
        db_file = MAPS_DB_FILE if scraper == "maps" else JUSTDIAL_DB_FILE
        tname = normalize_table_name(keyword, location)
        con = sqlite3.connect(db_file)
        cur = con.cursor()
        
        if scraper == "maps":
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS {tname} (
                    id INTEGER PRIMARY KEY,
                    name TEXT,
                    address TEXT,
                    phone TEXT,
                    website TEXT,
                    keyword TEXT,
                    location TEXT,
                    scraped_at TEXT
                )
            """)
        else:
            cur.execute(f"""
                CREATE TABLE IF NOT EXISTS {tname} (
                    id TEXT PRIMARY KEY,
                    name TEXT,
                    address TEXT,
                    phone TEXT,
                    website TEXT,
                    website_status TEXT,
                    keyword TEXT,
                    location TEXT,
                    scraped_at TEXT,
                    rating TEXT,
                    votes TEXT
                )
            """)
        con.commit()
        con.close()
        return tname

    # ---- Maps Scraper flow ----
    def load_maps(self):
        k = self.keyword_input.text().strip()
        l = self.location_input.text().strip()
        if not k or not l:
            QMessageBox.warning(self, "Missing Input", "Enter both keyword and location.")
            return
        self.current_keyword = k
        self.current_location = l
        url = f"https://www.google.com/maps/search/{quote_plus(k)}+in+{quote_plus(l)}"
        self.web.load(QUrl(url))
        self.log(f"Loaded Maps for {k} in {l}")

    def collect_links(self):
        js = r"""
        (function(){
          let anchors = Array.from(document.querySelectorAll('a[href*="/maps/place/"]'));
          let hrefs = [];
          anchors.forEach(a=>{ if(a.href && hrefs.indexOf(a.href)===-1) hrefs.push(a.href); });
          return hrefs.slice(0,200);
        })();
        """
        self.web.page().runJavaScript(js, self._got_links)

    def _got_links(self, hrefs):
        if not hrefs:
            self.log("No links found.")
            return
        max_links = self.max_links_input.value()
        self.links = hrefs[:max_links]
        self.links_index = 0
        self.log(f"Collected {len(self.links)} links. Ready to scrape.")

    def start_scrape(self):
        if not self.links:
            QMessageBox.information(self, "No links", "Collect links first.")
            return
        self.ensure_table(self.current_keyword, self.current_location, "maps")
        self.scraping = True
        self.links_index = 0
        self.maps_progress.setValue(0)
        self._process_next_link()

    def _process_next_link(self):
        if self.links_index >= len(self.links):
            self.scraping = False
            self.log("Scraping finished.")
            self.refresh_keyword_tables()
            return
        url = self.links[self.links_index]
        self._next_action = self._extract_place
        self.web.load(QUrl(url))

    def _on_load_finished(self, ok):
        if ok and self._next_action:
            QTimer.singleShot(600, self._next_action)

    def _extract_place(self):
        js = r"""
        (function(){
          let name = document.querySelector("h1")?.innerText.trim() || "";
          let addr = document.querySelector("button[data-item-id='address'] div")?.innerText.trim() || "";
          let phone = document.querySelector("button[data-item-id^='phone'] div")?.innerText.trim() || "";
          let site = document.querySelector("a[data-item-id='authority']")?.href || "";
          return {name:name,address:addr,phone:phone,website:site};
        })();
        """
        self.web.page().runJavaScript(js, self._got_place)

    def _got_place(self, data):
        if not data:
            data = {}
        tname = self.get_current_table("maps")
        con = sqlite3.connect(MAPS_DB_FILE)
        cur = con.cursor()
        
        # Check if business already exists to avoid duplicates
        cur.execute(f"SELECT id FROM {tname} WHERE name = ? AND address = ?", 
                   (data.get("name",""), data.get("address","")))
        existing = cur.fetchone()
        
        if not existing:
            cur.execute(f"""
                INSERT INTO {tname}(name,address,phone,website,keyword,location,scraped_at)
                VALUES (?,?,?,?,?,?,?)
            """, (
                data.get("name",""), data.get("address",""), data.get("phone",""), data.get("website",""),
                self.current_keyword, self.current_location, datetime.utcnow().isoformat()
            ))
            self.log(f"Saved: {data.get('name','')}")
        else:
            self.log(f"Skipped duplicate: {data.get('name','')}")
            
        con.commit()
        con.close()
        self.links_index += 1
        self.maps_progress.setValue(int((self.links_index/len(self.links))*100))
        QTimer.singleShot(800, self._process_next_link)

    # ---- JustDial Scraper flow ----
    def load_justdial(self):
        # Check if a direct URL is provided
        direct_url = self.justdial_url_input.text().strip()
        
        if direct_url:
            # Use the direct URL (for logged-in sessions)
            if not direct_url.startswith("http"):
                direct_url = "https://" + direct_url
                
            if "justdial.com" not in direct_url:
                QMessageBox.warning(self, "Invalid URL", "Please enter a valid JustDial URL")
                return
                
            self.justdial_web.load(QUrl(direct_url))
            self.log(f"Loaded JustDial from direct URL: {direct_url}", "justdial")
            
            # Try to extract keyword and location from URL for table naming
            parsed_url = urlparse(direct_url)
            path_parts = parsed_url.path.split('/')
            
            if len(path_parts) >= 3:
                # Extract location and keyword from URL path
                location = path_parts[1]
                keyword = path_parts[2]
                
                # Clean up keyword (remove any query parameters)
                if '?' in keyword:
                    keyword = keyword.split('?')[0]
                
                self.justdial_keyword_input.setText(keyword.replace('-', ' ').title())
                self.justdial_location_input.setText(location.replace('-', ' ').title())
                
                self.log(f"Extracted keyword: {keyword}, location: {location} from URL", "justdial")
            return
        
        # If no direct URL, use keyword and location
        k = self.justdial_keyword_input.text().strip()
        l = self.justdial_location_input.text().strip()
        
        if not k or not l:
            QMessageBox.warning(self, "Missing Input", "Enter either a direct URL or both keyword and location.")
            return
            
        # Construct the URL based on keyword and location
        search_url = f"https://www.justdial.com/{quote_plus(l)}/{quote_plus(k)}"

        self.justdial_keyword = k
        self.justdial_location = l
        self.justdial_url = search_url
        
        self.justdial_web.load(QUrl(search_url))
        self.log(f"Loaded JustDial: {k} in {l} from generated URL: {search_url}", "justdial")


    def scroll_justdial(self):
        """Scroll the JustDial page to load more content"""
        js_code = """
        (function() {
            // Scroll to bottom to load more content
            window.scrollTo(0, document.body.scrollHeight);
            return "Scrolled to bottom to load more content";
        })();
        """
        
        def scroll_complete(result):
            self.log(f"Scrolled to load more content. Waiting for content to load...", "justdial")
            # Wait a bit for content to load, then you can extract again
            QTimer.singleShot(3000, lambda: self.log("Content should be loaded now. You can extract data.", "justdial"))
        
        self.justdial_web.page().runJavaScript(js_code, scroll_complete)

    def start_justdial_scrape(self):
        # Get keyword and location from inputs
        k = self.justdial_keyword_input.text().strip()
        l = self.justdial_location_input.text().strip()
        max_listings = self.justdial_max_listings_input.value()
        
        if not k or not l:
            QMessageBox.warning(self, "Missing Input", "Enter keyword and location.")
            return
            
        # Ensure the URL is loaded before starting the scrape
        if not self.justdial_web.url() or "justdial.com" not in self.justdial_web.url().toString():
            self.log("Please load a JustDial page first by clicking 'Load JustDial'.", "justdial")
            return

        self.justdial_keyword = k
        self.justdial_location = l
        
        # Ensure table exists
        self.ensure_table(k, l, "justdial")
        
        # Start the scraper
        self.justdial_scraper = JustDialScraper(self.justdial_web, max_listings)
        self.justdial_scraper.status_signal.connect(lambda msg: self.log(msg, "justdial"))
        self.justdial_scraper.progress_signal.connect(self.justdial_progress.setValue)
        self.justdial_scraper.data_signal.connect(self.save_justdial_data)
        self.justdial_scraper.finished_signal.connect(self.justdial_scraping_finished)
        
        self.start_justdial_scrape_btn.setEnabled(False)
        self.stop_justdial_scrape_btn.setEnabled(True)
        self.justdial_scraper.start()
        
    def stop_justdial_scrape(self):
        if self.justdial_scraper and self.justdial_scraper.isRunning():
            self.justdial_scraper.stop()
            self.justdial_scraper.wait()
            self.log("JustDial scraping stopped by user", "justdial")
            
    def justdial_scraping_finished(self):
        self.start_justdial_scrape_btn.setEnabled(True)
        self.stop_justdial_scrape_btn.setEnabled(False)
        self.refresh_keyword_tables()
        
    def save_justdial_data(self, data):
        tname = self.get_current_table("justdial")
        con = sqlite3.connect(JUSTDIAL_DB_FILE)
        cur = con.cursor()
        
        # Check if business already exists to avoid duplicates
        cur.execute(f"SELECT id FROM {tname} WHERE name = ? AND address = ?", 
                   (data.get("name",""), data.get("address","")))
        existing = cur.fetchone()
        
        if not existing:
            cur.execute(f"""
                INSERT INTO {tname}(id, name, address, phone, website, website_status, 
                                  keyword, location, scraped_at, rating, votes)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """, (
                data.get("id"),
                data.get("name",""),
                data.get("address",""),
                data.get("phone",""),
                data.get("website",""),
                data.get("website_status","Unknown"),
                self.justdial_keyword,
                self.justdial_location,
                data.get("scraped_at",""),
                data.get("rating","N/A"),
                data.get("votes","N/A")
            ))
            self.log(f"Saved: {data.get('name','')}", "justdial")
        else:
            self.log(f"Skipped duplicate: {data.get('name','')}", "justdial")
            
        con.commit()
        con.close()

    def _on_justdial_load_finished(self, ok):
        if ok:
            self.log("JustDial page loaded successfully", "justdial")

    # ---- Viewer ----
    def refresh_keyword_tables(self):
        source = self.source_combo.currentText()
        db_file = MAPS_DB_FILE if source == "Google Maps" else JUSTDIAL_DB_FILE
        
        con = sqlite3.connect(db_file)
        cur = con.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = [r[0] for r in cur.fetchall()]
        con.close()
        
        self.keyword_combo.clear()
        self.keyword_combo.addItems(tables)

    def load_selected_table(self):
        tname = self.keyword_combo.currentText()
        if not tname:
            return
            
        source = self.source_combo.currentText()
        db_file = MAPS_DB_FILE if source == "Google Maps" else JUSTDIAL_DB_FILE
        
        con = sqlite3.connect(db_file)
        cur = con.cursor()
        cur.execute(f"SELECT * FROM {tname} ORDER BY id DESC")
        rows = cur.fetchall()
        cols = [desc[0] for desc in cur.description]
        con.close()

        self.table.setRowCount(len(rows))
        self.table.setColumnCount(len(cols))
        self.table.setHorizontalHeaderLabels(cols)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                self.table.setItem(r, c, QTableWidgetItem(str(val)))
        self.table.resizeColumnsToContents()

    def export_selected_csv(self):
        tname = self.keyword_combo.currentText()
        if not tname:
            return
            
        source = self.source_combo.currentText()
        db_file = MAPS_DB_FILE if source == "Google Maps" else JUSTDIAL_DB_FILE
        
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", f"{tname}.csv", "CSV Files (*.csv)")
        if not path:
            return
            
        con = sqlite3.connect(db_file)
        cur = con.cursor()
        cur.execute(f"SELECT * FROM {tname}")
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
        con.close()
        
        with open(path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(cols)
            writer.writerows(rows)
            
        QMessageBox.information(self, "Export", f"Exported {len(rows)} rows to CSV.")

    def export_selected_xls(self):
        tname = self.keyword_combo.currentText()
        if not tname:
            return
            
        source = self.source_combo.currentText()
        db_file = MAPS_DB_FILE if source == "Google Maps" else JUSTDIAL_DB_FILE
        
        path, _ = QFileDialog.getSaveFileName(self, "Save XLS", f"{tname}.xls", "Excel Files (*.xls)")
        if not path:
            return
            
        con = sqlite3.connect(db_file)
        cur = con.cursor()
        cur.execute(f"SELECT * FROM {tname}")
        rows = cur.fetchall()
        cols = [d[0] for d in cur.description]
        con.close()
        
        try:
            # Create a workbook and add a worksheet
            workbook = xlwt.Workbook(encoding='utf-8')
            worksheet = workbook.add_sheet(tname[:31])  # Sheet name max 31 chars
            
            # Write headers
            for col_idx, col_name in enumerate(cols):
                worksheet.write(0, col_idx, col_name)
            
            # Write data rows
            for row_idx, row in enumerate(rows, 1):  # Start from row 1 (after header)
                for col_idx, value in enumerate(row):
                    worksheet.write(row_idx, col_idx, str(value))
            
            # Save the workbook
            workbook.save(path)
            QMessageBox.information(self, "Export", f"Exported {len(rows)} rows to XLS.")
            
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export to XLS: {str(e)}")


def main():
    app = QApplication(sys.argv)
    w = MapsScraperGUI()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()