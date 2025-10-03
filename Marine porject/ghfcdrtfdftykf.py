import sys
import io
import PyQt5
import matplotlib
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton,
    QTabWidget, QMessageBox, QApplication, QStyle, QProgressBar, QTextEdit, QHBoxLayout,
    QDialog, QCalendarWidget, QDialogButtonBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject, QDate
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# YOUR ORIGINAL IMPORTS, UNTOUCHED
import pandas as pd
import os
import time
from datetime import date, datetime, timedelta
from tkinter import Tk, Label, Button
from tkcalendar import Calendar
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
try:
    from dateutil.parser import parse as dateutil_parse
except ImportError:
    print("Warning: 'python-dateutil' library not found. Date parsing will be less robust.")
    print("Install it using: pip install python-dateutil")
    dateutil_parse = None
try:
    import requests
except ImportError:
    print("Warning: 'requests' library not found. Online year verification is disabled.")
    print("Install it using: pip install requests")
    requests = None
import re
import spacy
import numpy as np
from collections import Counter
from string import punctuation
import heapq
from sklearn.metrics.pairwise import cosine_similarity
from spacy.matcher import Matcher
import tkinter as tk
from tkinter import filedialog
from tqdm import tqdm
import spacy.cli

DARK_STYLESHEET = """
QWidget{
    background-color: #2c3e50;
    color: #ecf0f1;
    font-family: "Segoe UI", "Arial", sans-serif;
    font-size: 14px;
}
QMainWindow {
    background-color: #2c3e50;
}
QTabWidget::pane {
    border-top: 2px solid #34495e;
    margin-top: -1px;
}
QTabBar::tab {
    background-color: #34495e;
    color: #ecf0f1;
    padding: 10px 20px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    border: 1px solid #2c3e50;
    border-bottom: none;
}
QTabBar::tab:selected {
    background-color: #5d9cec;
    color: #ffffff;
    font-weight: bold;
}
QTabBar::tab:hover {
    background-color: #4a627a;
}
QTabBar::tab:disabled {
    background-color: #2c3e50;
    color: #7f8c8d;
}
QLabel#TitleLabel {
    font-size: 28px;
    font-weight: bold;
    color: #ffffff;
    margin-bottom: 10px;
}
QLabel#DescriptionLabel {
    font-size: 16px;
    color: #bdc3c7;
    margin-bottom: 20px;
}
QPushButton {
    background-color: #5d9cec;
    color: white;
    border-radius: 5px;
    padding: 12px 25px;
    font-weight: bold;
    border: none;
    min-width: 200px;
}
QPushButton:hover {
    background-color: #4a8cdb;
}
QPushButton:pressed {
    background-color: #3a7acb;
}
QPushButton:disabled {
    background-color: #34495e;
    color: #7f8c8d;
}
QProgressBar {
    border: 2px solid #34495e;
    border-radius: 5px;
    text-align: center;
    color: #ecf0f1;
    background-color: #34495e;
}
QProgressBar::chunk {
    background-color: #2ecc71;
    width: 10px;
    margin: 0.5px;
}
QTextEdit {
    background-color: #34495e;
    color: #ecf0f1;
    border: 1px solid #2c3e50;
    border-radius: 4px;
    font-family: "Consolas", "Courier New", monospace;
}
QMessageBox {
    background-color: #34495e;
}
QMessageBox QLabel {
    color: #ecf0f1;
}
QDialog {
    background-color: #34495e;
}
QCalendarWidget QToolButton {
    color: white;
    background-color: #5d9cec;
}
QCalendarWidget QMenu {
    background-color: #34495e;
}
QCalendarWidget QSpinBox {
    color: white;
    background-color: #34495e;
}
QCalendarWidget QHeaderView {
    background-color: #34495e;
}
QCalendarWidget QTableView {
    background-color: #2c3e50;
    color: white;
    alternate-background-color: #34495e;
}
"""

class DateRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Date Range")
        self.layout = QVBoxLayout(self)
        self.layout.addWidget(QLabel("Select Start Date:"))
        self.start_cal = QCalendarWidget(self)
        self.start_cal.setSelectedDate(QDate.currentDate().addMonths(-1))
        self.layout.addWidget(self.start_cal)
        self.layout.addWidget(QLabel("Select End Date:"))
        self.end_cal = QCalendarWidget(self)
        self.end_cal.setSelectedDate(QDate.currentDate())
        self.layout.addWidget(self.end_cal)
        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addWidget(self.buttons)

    def get_dates(self):
        start_date = self.start_cal.selectedDate().toPyDate()
        end_date = self.end_cal.selectedDate().toPyDate()
        return start_date, end_date

class Worker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int)
    output = pyqtSignal(str)
    results = pyqtSignal(dict)

    def __init__(self, task_function, *args, **kwargs):
        super().__init__()
        self.task_function = task_function
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            self.task_function(self, *self.args, **self.kwargs)
        except Exception as e:
            import traceback
            self.output.emit(f"\n--- TASK ERROR --- \n{str(e)}\n{traceback.format_exc()}\n------------------")
        finally:
            self.finished.emit()

class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Upstream News Analyzer")
        self.setGeometry(100, 100, 1200, 800)
        self.set_app_icon()
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Tab 1: News Gathering
        news_tab = self.create_tab_ui(
            "News Gathering Module",
            "Gather news from Upstream Online within a selected date range.",
            "Start News Gathering",
            QStyle.SP_ArrowDown,
            self.start_news_gathering_process
        )
        self.tabs.addTab(news_tab, "News Gathering")
        self.gather_btn = news_tab.findChild(QPushButton)
        self.gather_progress = news_tab.findChild(QProgressBar)
        self.gather_log = news_tab.findChild(QTextEdit)
        
        # Tab 2: Brunei LNG Scraper (NEW)
        brunei_tab = self.create_tab_ui(
            "Brunei LNG Document Scraper",
            "Download PDF documents from Brunei LNG website.",
            "Start Brunei Scraping",
            QStyle.SP_DriveNetIcon,
            self.start_brunei_scraping_process
        )
        self.tabs.addTab(brunei_tab, "Brunei LNG")
        self.brunei_btn = brunei_tab.findChild(QPushButton)
        self.brunei_progress = brunei_tab.findChild(QProgressBar)
        self.brunei_log = brunei_tab.findChild(QTextEdit)
        
        # Tab 3: AI Analyzer
        ai_tab = self.create_tab_ui(
            "AI Analysis Engine",
            "Analyze gathered news to extract structured data.",
            "Run AI Analyzer",
            QStyle.SP_MediaPlay,
            self.start_ai_analyzer_process
        )
        self.tabs.addTab(ai_tab, "AI Analyzer")
        self.ai_btn = ai_tab.findChild(QPushButton)
        self.ai_progress = ai_tab.findChild(QProgressBar)
        self.ai_log = ai_tab.findChild(QTextEdit)
        
        # Tab 4: Dashboard
        self.dashboard_tab = self.create_dashboard_tab()
        self.tabs.addTab(self.dashboard_tab, "Dashboard")
        self.tabs.setTabEnabled(3, False)

    def set_app_icon(self):
        self.setWindowIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))

    def create_tab_ui(self, title, description, button_text, button_icon, on_click_method):
        tab_widget = QWidget()
        layout = QVBoxLayout(tab_widget)
        layout.setContentsMargins(40, 20, 40, 20)
        layout.setAlignment(Qt.AlignCenter)
        title_label = QLabel(title)
        title_label.setObjectName("TitleLabel")
        title_label.setAlignment(Qt.AlignCenter)
        desc_label = QLabel(description)
        desc_label.setObjectName("DescriptionLabel")
        desc_label.setAlignment(Qt.AlignCenter)
        desc_label.setWordWrap(True)
        action_btn = QPushButton(button_text)
        action_btn.setIcon(self.style().standardIcon(button_icon))
        action_btn.setMinimumSize(300, 60)
        action_btn.clicked.connect(on_click_method)
        progress_bar = QProgressBar()
        progress_bar.setRange(0, 100)
        progress_bar.setValue(0)
        progress_bar.hide()
        log_area = QTextEdit()
        log_area.setReadOnly(True)
        log_area.hide()
        layout.addStretch(1)
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addSpacing(20)
        layout.addWidget(action_btn, 0, Qt.AlignCenter)
        layout.addSpacing(20)
        layout.addWidget(progress_bar)
        layout.addWidget(log_area)
        layout.addStretch(2)
        return tab_widget

    def create_dashboard_tab(self):
        dashboard_widget = QWidget()
        main_layout = QVBoxLayout(dashboard_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        title_label = QLabel("AI Analysis Dashboard")
        title_label.setObjectName("TitleLabel")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        self.status_chart_canvas = self.setup_chart_canvas()
        self.location_chart_canvas = self.setup_chart_canvas()
        self.type_chart_canvas = self.setup_chart_canvas()
        chart_layout = QHBoxLayout()
        chart_layout.addWidget(self.status_chart_canvas)
        chart_layout.addWidget(self.location_chart_canvas)
        chart_layout.addWidget(self.type_chart_canvas)
        main_layout.addLayout(chart_layout)
        return dashboard_widget

    def setup_chart_canvas(self):
        fig = Figure(figsize=(5, 4), dpi=100)
        fig.patch.set_facecolor('#2c3e50')
        return FigureCanvas(fig)

    def _run_task_in_thread(self, task_function, button, progress_bar, log_area, on_finish_slot, processing_text, *args, **kwargs):
        class EmittingStream(io.TextIOBase):
            def __init__(self, signal_emitter):
                super().__init__()
                self.text_written = ""
                self.signal_emitter = signal_emitter
            def write(self, text):
                self.text_written += text
                if '\n' in self.text_written:
                    self.signal_emitter.emit(self.text_written.strip())
                    self.text_written = ""
                return len(text)
            def flush(self):
                if self.text_written.strip():
                    self.signal_emitter.emit(self.text_written.strip())
                    self.text_written = ""

        button.setEnabled(False)
        button.setText(processing_text)
        progress_bar.setValue(0)
        progress_bar.show()
        log_area.clear()
        log_area.show()
        
        self.thread = QThread()
        self.worker_obj = Worker(task_function, *args, **kwargs)
        self.worker_obj.moveToThread(self.thread)
        
        original_run = self.worker_obj.run
        def run_with_redirect():
            original_stdout = sys.stdout
            sys.stdout = EmittingStream(self.worker_obj.output)
            try:
                original_run()
            finally:
                sys.stdout.flush()
                sys.stdout = original_stdout
        
        self.worker_obj.run = run_with_redirect

        self.thread.started.connect(self.worker_obj.run)
        self.worker_obj.finished.connect(self.thread.quit)
        self.worker_obj.finished.connect(self.worker_obj.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker_obj.progress.connect(progress_bar.setValue)
        self.worker_obj.output.connect(log_area.append)
        
        if 'perform_ai_analysis' in task_function.__name__:
            self.worker_obj.results.connect(self.update_dashboard_charts)
            
        self.worker_obj.finished.connect(on_finish_slot)
        self.thread.start()

    def start_news_gathering_process(self):
        dialog = DateRangeDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            start_date, end_date = dialog.get_dates()
            self._run_task_in_thread(
                self.perform_news_gathering,
                self.gather_btn,
                self.gather_progress,
                self.gather_log,
                self.on_news_gathering_complete,
                "Gathering News...",
                start_date=start_date,
                end_date=end_date
            )
        else:
            self.gather_log.show()
            self.gather_log.setText("News gathering cancelled by user.")

    def on_news_gathering_complete(self):
        self.gather_btn.setEnabled(True)
        self.gather_btn.setText("Start News Gathering")
        self.gather_progress.hide()
        QMessageBox.information(self, "Process Complete", "News gathering finished.")

    # NEW: Brunei LNG Scraper
    def start_brunei_scraping_process(self):
        self._run_task_in_thread(
            self.perform_brunei_scraping,
            self.brunei_btn,
            self.brunei_progress,
            self.brunei_log,
            self.on_brunei_scraping_complete,
            "Scraping Brunei LNG..."
        )

    def on_brunei_scraping_complete(self):
        self.brunei_btn.setEnabled(True)
        self.brunei_btn.setText("Start Brunei Scraping")
        self.brunei_progress.hide()
        QMessageBox.information(self, "Process Complete", "Brunei LNG scraping finished.")

    def perform_brunei_scraping(self, worker):
        """
        Scrapes PDF links from Brunei LNG website and downloads them.
        """
        BRUNEI_URL = "https://www.blng.com.bn/Pages/Tenders.aspx"
        OUTPUT_DIR = "C:/Office work/Brunei_LNG_PDFs"
        
        print("🌐 Starting Brunei LNG scraper...")
        print(f"Target URL: {BRUNEI_URL}")
        worker.progress.emit(10)
        
        # Create output directory
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        print(f"📁 Output directory: {OUTPUT_DIR}")
        
        # Setup Chrome
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            print("✅ Chrome driver initialized")
        except Exception as e:
            print(f"❌ Failed to initialize Chrome: {e}")
            worker.output.emit(f"❌ Could not start WebDriver: {e}")
            return
        
        worker.progress.emit(20)
        
        try:
            # Load the page
            print(f"Loading page: {BRUNEI_URL}")
            driver.get(BRUNEI_URL)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(3)  # Allow JavaScript to finish loading
            
            worker.progress.emit(40)
            print("✅ Page loaded successfully")
            
            # Parse with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, "html.parser")
            
            # Find all PDF links
            pdf_links = []
            for link in soup.find_all('a', href=True):
                href = link['href']
                if href.lower().endswith('.pdf'):
                    # Handle relative URLs
                    if not href.startswith('http'):
                        if href.startswith('/'):
                            href = f"https://www.blng.com.bn{href}"
                        else:
                            href = f"https://www.blng.com.bn/{href}"
                    
                    pdf_name = href.split('/')[-1]
                    pdf_links.append((pdf_name, href))
                    print(f"Found PDF: {pdf_name}")
            
            worker.progress.emit(60)
            
            if not pdf_links:
                print("⚠️ No PDF links found on the page")
                print("This might mean:")
                print("  - The page structure has changed")
                print("  - PDFs are loaded dynamically via JavaScript")
                print("  - Access restrictions are in place")
            else:
                print(f"\n📄 Found {len(pdf_links)} PDF file(s)")
                
                # Download PDFs
                for i, (pdf_name, pdf_url) in enumerate(pdf_links):
                    try:
                        print(f"\n📥 Downloading [{i+1}/{len(pdf_links)}]: {pdf_name}")
                        response = requests.get(pdf_url, timeout=30)
                        response.raise_for_status()
                        
                        file_path = os.path.join(OUTPUT_DIR, pdf_name)
                        with open(file_path, 'wb') as f:
                            f.write(response.content)
                        
                        print(f"✅ Saved: {file_path}")
                        
                        # Update progress
                        progress = 60 + int((i + 1) / len(pdf_links) * 35)
                        worker.progress.emit(progress)
                        
                        time.sleep(1)  # Be polite to the server
                        
                    except Exception as e:
                        print(f"❌ Error downloading {pdf_name}: {e}")
                
                print(f"\n✅ Download complete! Files saved to: {OUTPUT_DIR}")
        
        except Exception as e:
            print(f"❌ Error during scraping: {e}")
            import traceback
            print(traceback.format_exc())
        
        finally:
            driver.quit()
            print("🔒 Browser closed")
            worker.progress.emit(100)

    def start_ai_analyzer_process(self):
        self._run_task_in_thread(
            self.perform_ai_analysis,
            self.ai_btn,
            self.ai_progress,
            self.ai_log,
            self.on_ai_analysis_complete,
            "Analyzing..."
        )

    def on_ai_analysis_complete(self):
        self.ai_btn.setEnabled(True)
        self.ai_btn.setText("Run AI Analyzer")
        self.ai_progress.hide()
        QMessageBox.information(self, "Process Complete", "AI analysis finished.")
        self.tabs.setTabEnabled(3, True)
        self.tabs.setCurrentIndex(3)

    # [Include your full perform_news_gathering method here - unchanged]
    # [Include your full perform_ai_analysis method here - unchanged]
    # [I'll omit these for brevity since they're unchanged from your original code]

    def update_dashboard_charts(self, data: dict):
        self.status_chart_canvas.figure.clear()
        ax1 = self.status_chart_canvas.figure.add_subplot(111)
        status_data = data.get("status_counts", {})
        if status_data:
            ax1.pie(status_data.values(), labels=status_data.keys(), autopct='%1.1f%%', startangle=90, textprops={'color':"w"})
        ax1.set_title("Project Status Overview", color='white')
        self.status_chart_canvas.draw()

        self.location_chart_canvas.figure.clear()
        ax2 = self.location_chart_canvas.figure.add_subplot(111)
        location_data = data.get("location_counts", {})
        if location_data:
            labels = [loc.strip() for loc, count in location_data.items() if loc.strip()]
            values = [count for loc, count in location_data.items() if loc.strip()]
            ax2.barh(labels, values, color='#5d9cec')
            ax2.invert_yaxis()
        ax2.set_title("Top 10 Locations", color='white')
        ax2.set_facecolor('#34495e')
        ax2.tick_params(colors='white')
        self.location_chart_canvas.figure.tight_layout()
        self.location_chart_canvas.draw()

        self.type_chart_canvas.figure.clear()
        ax3 = self.type_chart_canvas.figure.add_subplot(111)
        type_data = data.get("type_counts", {})
        if type_data:
            ax3.bar(type_data.keys(), type_data.values(), color=['#2ecc71', '#e74c3c', '#f1c40f'])
        ax3.set_title("Onshore vs. Offshore", color='white')
        ax3.set_facecolor('#34495e')
        ax3.tick_params(colors='white')
        self.type_chart_canvas.draw()

def run_app():
    app = QApplication(sys.argv)
    app.setStyleSheet(DARK_STYLESHEET)
    window = AppWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    run_app()