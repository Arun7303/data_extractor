# data_extractor
A desktop GUI application built with PyQt5 for scraping business listings from Google Maps and JustDial. Features include an embedded browser for interaction, data storage in separate SQLite databases, and export to CSV/XLS formats.

# Advanced Google Maps & JustDial Scraper

This is a powerful desktop application built with Python and PyQt5 that allows you to scrape business listing data from Google Maps and JustDial. It features an intuitive graphical user interface (GUI) with an embedded web browser, making the scraping process interactive and easy to monitor.


## ✨ Key Features

* **Dual Scraper Engines**: Separate, tailored scraping modules for both Google Maps and JustDial.
* **Interactive GUI**: Built with PyQt5, providing a user-friendly interface with tabs for each scraper and a data viewer.
* **Embedded Web Browser**: Uses `QWebEngineView` to load and interact with websites directly within the application. This is crucial for handling JavaScript-heavy sites.
* **Persistent Data Storage**: Scraped data is automatically saved into dedicated SQLite databases (`businesses.db` for Maps, `justdial_businesses.db` for JustDial).
* **Dynamic Table Creation**: Automatically creates new tables for each keyword-location pair to keep data organized.
* **Integrated Data Viewer**: Browse, view, and manage all scraped data from within the app without needing a separate database tool.
* **Data Export**: Export your scraped data easily to **CSV** and **XLS** (Excel) formats.
* **Duplicate Prevention**: The application checks for existing entries (based on name and address) to avoid saving duplicate business listings.

## ⚙️ How It Works

### Google Maps Scraper
The Google Maps scraper uses a two-step process to ensure reliable data collection:
1.  **Load & Collect Links**: First, you enter a keyword and location to perform a search. After the results page loads, you click "Collect Links" to gather the URLs of all the business listings on the page.
2.  **Start Scrape**: The application then automatically navigates to each collected link one by one, waits for the page to load, and extracts the business's **Name**, **Address**, **Phone Number**, and **Website**.

### JustDial Scraper
The JustDial scraper is designed to extract all data from a single search results page, which is more efficient for sites that display a lot of information upfront.
1.  **Load Page**: You can load a page either by entering a keyword and location (which constructs a URL) or by pasting a direct JustDial URL. Using a direct URL is helpful if you need to be logged in to see certain data like phone numbers.
2.  **Scroll to Load More**: JustDial often uses infinite scrolling. The "Scroll to Load More" button scrolls the embedded browser to the bottom, triggering the loading of more listings.
3.  **Extract Data**: Clicking "Extract Data" executes a JavaScript snippet on the currently loaded page to grab all business details at once, including **Name**, **Address**, '**Phone**, **Website**, **Rating**, and **Votes**.

## 🚀 Getting Started

Follow these instructions to get the project up and running on your local machine.

### Prerequisites

You need to have Python 3 installed on your system.

### Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/your-username/your-repository-name.git](https://github.com/your-username/your-repository-name.git)
    cd your-repository-name
    ```

2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3.  **Install the required libraries:**
    A `requirements.txt` file is the best way to handle this. Create a file named `requirements.txt` with the following content:
    ```
    PyQt5
    PyQtWebEngine
    xlwt
    openpyxl
    ```
    Then, install them using pip:
    ```bash
    pip install -r requirements.txt
    ```

### How to Run

Execute the main Python script to launch the application:
```bash
python just_v2.py
