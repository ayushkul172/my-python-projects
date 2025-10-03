import os
import io
import requests
import pdfplumber
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin, unquote
import re
import time
import urllib3
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
from datetime import datetime

# ML/AI imports
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURATION ---
DOWNLOADS_BASE_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
UK_MARINE_FOLDER = os.path.join(DOWNLOADS_BASE_PATH, "UK_Marine_Notices")

# UK Configuration
BASE_URL = "https://msi.admiralty.co.uk"
WEEKLY_URL = "https://msi.admiralty.co.uk/NoticesToMariners/Weekly"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-GB,en;q=0.9',
    'Referer': 'https://msi.admiralty.co.uk/NoticesToMariners/Weekly',
}

# --- AI EXTRACTOR FOR UK NOTICES ---
class UKPlatformExtractor:
    """AI-based extraction for UK Notice to Mariners"""
    
    def __init__(self):
        self.platform_keywords = [
            'platform', 'platfrom', 'platfom', 'rig', 'jack-up', 'semi-submersible',
            'drillship', 'installation', 'monopod', 'wellhead', 'well head',
            'fpso', 'flng', 'fso', 'fsu', 'drilling unit', 'drilling operations',
            'oil platform', 'gas platform', 'offshore structure', 'mopu', 'tlp',
            'spar platform', 'tension leg', 'compliant tower', 'subsea'
        ]
        
        self.action_patterns = {
            'Insert': ['insert', 'establish', 'commence', 'new', 'installed', 'deployed', 'placed', 'positioned', 'construction'],
            'Delete': ['delete', 'remove', 'cancel', 'withdrawn', 'terminated', 'removed', 'ceased'],
            'Amend': ['amend', 'move', 'relocate', 'modify', 'change', 'update', 'moved', 'repositioned']
        }
        
        self.notice_pattern = re.compile(r'(\d{4})\(?\w?\)?/\d{2}')
        
        self.coord_patterns = [
            re.compile(r"(\d{1,2}°\s*\d{1,2}´·?\d{0,2}[NS]\.?),?\s*(\d{1,3}°\s*\d{1,2}´·?\d{0,2}[EW]\.?)"),
            re.compile(r"(\d{1,2}°\s*\d{1,2}'\.?\d{0,2}[NS]\.?),?\s*(\d{1,3}°\s*\d{1,2}'\.?\d{0,2}[EW]\.?)"),
            re.compile(r"(\d{1,2}[°\s-]+\d{1,2}[\.\s'-]*\d{0,2}\s*[\"']?\s*[NS])\s*[,;]?\s*(\d{1,3}[°\s-]+\d{1,2}[\.\s'-]*\d{0,2}\s*[\"']?\s*[EW])")
        ]
        
        self.platform_name_patterns = [
            r"['\"]([A-Z][A-Za-z0-9\s\-]+)['\"]",
            r"(?:platform|rig|installation|fpso|flng|fso)\s+['\"]?([A-Z][A-Za-z0-9\s\-]+)['\"]?",
            r"([A-Z][A-Z0-9\s\-]{2,30})\s+(?:platform|rig|installation|fpso|flng|fso)",
            r"(?:named|called|designated)\s+['\"]?([A-Z][A-Za-z0-9\s\-]+)['\"]?",
        ]
        
        self.vectorizer = TfidfVectorizer(max_features=100, stop_words='english')
        self.trained = False
        
    def train_on_keywords(self):
        """Train AI model"""
        training_texts = self.platform_keywords + [
            'offshore drilling platform north sea',
            'oil rig installation UK sector',
            'platform construction works position',
            'floating production storage offloading',
            'jack-up drilling unit established',
        ]
        self.vectorizer.fit(training_texts)
        self.trained = True
    
    def calculate_relevance_score(self, text):
        """AI scoring for platform relevance"""
        if not self.trained:
            self.train_on_keywords()
        
        try:
            text_vector = self.vectorizer.transform([text.lower()])
            keyword_vectors = self.vectorizer.transform(self.platform_keywords)
            similarities = cosine_similarity(text_vector, keyword_vectors)
            return float(np.max(similarities))
        except:
            text_lower = text.lower()
            matches = sum(1 for kw in self.platform_keywords if kw in text_lower)
            return matches / len(self.platform_keywords)
    
    def extract_platform_name(self, text):
        """Extract platform name"""
        candidates = []
        
        for pattern in self.platform_name_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                name = match.strip()
                if (len(name) > 2 and len(name) < 50 and 
                    not name.lower() in ['the', 'and', 'with', 'from', 'this', 'that', 'located', 
                                        'position', 'north', 'south', 'east', 'west', 'oil', 'gas',
                                        'construction', 'works', 'area'] and
                    not re.match(r'^\d+$', name)):
                    candidates.append(name)
        
        seen = set()
        unique = []
        for name in candidates:
            upper = name.upper()
            if upper not in seen:
                seen.add(upper)
                unique.append(name)
        
        return unique[0] if unique else "Unknown Platform"
    
    def extract_action(self, text):
        """Determine action type"""
        text_lower = text.lower()
        action_scores = {action: sum(1 for kw in keywords if kw in text_lower)
                        for action, keywords in self.action_patterns.items()}
        
        max_action = max(action_scores, key=action_scores.get)
        return max_action if action_scores[max_action] > 0 else "Info"
    
    def extract_notice_number(self, text):
        """Extract UK notice number"""
        match = self.notice_pattern.search(text)
        return match.group(0) if match else ""
    
    def extract_date_from_filename(self, filename):
        """Extract date from filename"""
        match = re.search(r'(\d{2})snii(\d{2})\.pdf', filename, re.IGNORECASE)
        if match:
            week = match.group(1)
            year = f"20{match.group(2)}"
            return f"Week {week}/{year}"
        
        date_match = re.search(r'(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})', filename)
        if date_match:
            return f"{date_match.group(1)}/{date_match.group(2)}/{date_match.group(3)}"
        
        return "Date Unknown"
    
    def extract_country_from_text(self, text):
        """Extract country"""
        countries = {
            'UK': ['england', 'scotland', 'wales', 'northern ireland', 'united kingdom', 'british'],
            'North Sea': ['north sea'],
            'India': ['india', 'indian'],
            'Norway': ['norway', 'norwegian'],
            'USA': ['united states', 'america', 'american', 'usa', 'gulf of mexico'],
        }
        
        text_lower = text.lower()
        for country, keywords in countries.items():
            if any(kw in text_lower for kw in keywords):
                return country
        
        return "Unknown"
    
    def parse_uk_notice_format(self, text, url, source_file):
        """Parse UK Notice format"""
        extracted_records = []
        date = self.extract_date_from_filename(source_file)
        notice_sections = re.split(r'(?=\d{4}\(?\w?\)?/\d{2})', text)
        
        for section in notice_sections:
            if len(section) < 50:
                continue
            
            relevance_score = self.calculate_relevance_score(section)
            
            if relevance_score < 0.05:
                continue
            
            notice_num = self.extract_notice_number(section)
            country = self.extract_country_from_text(section)
            
            coordinates_found = []
            for pattern in self.coord_patterns:
                matches = pattern.findall(section)
                if matches:
                    coordinates_found.extend(matches)
            
            if not coordinates_found:
                continue
            
            platform_name = self.extract_platform_name(section)
            action = self.extract_action(section)
            
            lines = section.split('\n')
            description = ' '.join([l.strip() for l in lines[:5] if l.strip()])[:500]
            
            for lat, lon in coordinates_found:
                lat_clean = re.sub(r'\s+', ' ', lat.strip())
                lon_clean = re.sub(r'\s+', ' ', lon.strip())
                
                validation = "OK" if relevance_score > 0.2 else "Needs Review"
                
                record = {
                    "Date": date,
                    "Country": country,
                    "Source File": os.path.basename(source_file),
                    "Notice Number": notice_num,
                    "Platform Name": platform_name,
                    "Action": action,
                    "Latitude": lat_clean,
                    "Longitude": lon_clean,
                    "Full Description": description,
                    "AI Score": f"{relevance_score:.2f}",
                    "Validation": validation,
                    "PDF Link": url,
                    "Source URL": url
                }
                extracted_records.append(record)
                break
        
        return extracted_records
    
    def parse_notice_text(self, text, url, source_file, relevance_threshold=0.05):
        """Main parsing entry point"""
        if self.notice_pattern.search(text):
            return self.parse_uk_notice_format(text, url, source_file)
        
        date = self.extract_date_from_filename(source_file)
        extracted_records = []
        paragraphs = re.split(r'\n\s*\n', text.strip())
        
        for paragraph in paragraphs:
            relevance_score = self.calculate_relevance_score(paragraph)
            
            if relevance_score < relevance_threshold:
                continue
            
            country = self.extract_country_from_text(paragraph)
            
            for pattern in self.coord_patterns:
                matches = list(pattern.finditer(paragraph.replace('\n', ' ')))
                
                for match in matches:
                    latitude = match.group(1).strip()
                    longitude = match.group(2).strip()
                    
                    platform_name = self.extract_platform_name(paragraph)
                    action = self.extract_action(paragraph)
                    description = ' '.join(paragraph.replace('\n', ' ').split())
                    
                    validation = "OK" if relevance_score > 0.3 else "Needs Review"
                    
                    record = {
                        "Date": date,
                        "Country": country,
                        "Source File": os.path.basename(source_file),
                        "Notice Number": "",
                        "Platform Name": platform_name,
                        "Action": action,
                        "Latitude": re.sub(r'\s+', ' ', latitude),
                        "Longitude": re.sub(r'\s+', ' ', longitude),
                        "Full Description": description[:500],
                        "AI Score": f"{relevance_score:.2f}",
                        "Validation": validation,
                        "PDF Link": url,
                        "Source URL": url
                    }
                    extracted_records.append(record)
                    break
        
        return extracted_records


# --- GUI APPLICATION ---
class UKMarineExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("UK Marine Notice Extractor - AI-Powered")
        self.root.geometry("1400x850")
        
        self.ai_extractor = UKPlatformExtractor()
        self.extracted_data = []
        self.is_running = False
        
        self.setup_ui()
        
    def setup_ui(self):
        """Create GUI"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Tab 1: Download PDFs
        self.download_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.download_frame, text='1. Download PDFs')
        self.setup_download_tab()
        
        # Tab 2: Process Local PDFs
        self.process_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.process_frame, text='2. Process PDFs')
        self.setup_process_tab()
        
        # Tab 3: Results
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text='3. Results')
        self.setup_results_tab()
        
    def setup_download_tab(self):
        """PDF Download tab"""
        info_frame = ttk.LabelFrame(self.download_frame, text="Download UK Weekly Notices", padding=10)
        info_frame.pack(fill='x', padx=10, pady=5)
        
        info_text = "Downloads PDFs from UK Admiralty Maritime Safety Information website\nSaves to: " + UK_MARINE_FOLDER
        ttk.Label(info_frame, text=info_text, justify='left').pack()
        
        button_frame = ttk.Frame(self.download_frame)
        button_frame.pack(fill='x', padx=10, pady=5)
        
        self.download_btn = ttk.Button(button_frame, text="Download Latest Week", command=self.start_download)
        self.download_btn.pack(side='left', padx=5)
        
        self.stop_download_btn = ttk.Button(button_frame, text="Stop", command=self.stop_download, state='disabled')
        self.stop_download_btn.pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Open Folder", command=self.open_download_folder).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Clear Download Log", command=self.clear_download_log).pack(side='left', padx=5)
        
        self.download_progress = ttk.Progressbar(self.download_frame, mode='determinate')
        self.download_progress.pack(fill='x', padx=10, pady=5)
        
        log_frame = ttk.LabelFrame(self.download_frame, text="Download Log", padding=5)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.download_log = scrolledtext.ScrolledText(log_frame, height=20, wrap=tk.WORD)
        self.download_log.pack(fill='both', expand=True)
    
    def setup_process_tab(self):
        """PDF Processing tab"""
        info_frame = ttk.LabelFrame(self.process_frame, text="Process Downloaded PDFs", padding=10)
        info_frame.pack(fill='x', padx=10, pady=5)
        
        info_text = f"Processes PDFs from: {UK_MARINE_FOLDER}\nExtracts platform information using AI"
        ttk.Label(info_frame, text=info_text, justify='left').pack()
        
        button_frame = ttk.Frame(self.process_frame)
        button_frame.pack(fill='x', padx=10, pady=5)
        
        self.process_btn = ttk.Button(button_frame, text="Process All PDFs", command=self.start_processing)
        self.process_btn.pack(side='left', padx=5)
        
        self.stop_process_btn = ttk.Button(button_frame, text="Stop", command=self.stop_processing, state='disabled')
        self.stop_process_btn.pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Clear Process Log", command=self.clear_process_log).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Export Results", command=self.export_results).pack(side='left', padx=5)
        
        self.process_progress = ttk.Progressbar(self.process_frame, mode='determinate')
        self.process_progress.pack(fill='x', padx=10, pady=5)
        
        log_frame = ttk.LabelFrame(self.process_frame, text="Processing Log", padding=5)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.process_log = scrolledtext.ScrolledText(log_frame, height=20, wrap=tk.WORD)
        self.process_log.pack(fill='both', expand=True)
    
    def setup_results_tab(self):
        """Results display"""
        toolbar = ttk.Frame(self.results_frame)
        toolbar.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(toolbar, text="Filter:").pack(side='left', padx=5)
        self.filter_var = tk.StringVar()
        ttk.Entry(toolbar, textvariable=self.filter_var, width=30).pack(side='left', padx=5)
        ttk.Button(toolbar, text="Apply", command=self.apply_filter).pack(side='left', padx=5)
        ttk.Button(toolbar, text="Clear", command=self.clear_filter).pack(side='left', padx=5)
        
        ttk.Label(toolbar, text="Records:").pack(side='right', padx=5)
        self.record_count = ttk.Label(toolbar, text="0", font=('Arial', 10, 'bold'))
        self.record_count.pack(side='right')
        
        tree_frame = ttk.Frame(self.results_frame)
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        self.results_tree = ttk.Treeview(tree_frame, 
                                         columns=('Date', 'Country', 'Notice', 'Platform', 'Action', 'Lat', 'Lon', 'Score'),
                                         show='tree headings',
                                         yscrollcommand=vsb.set,
                                         xscrollcommand=hsb.set)
        
        vsb.config(command=self.results_tree.yview)
        hsb.config(command=self.results_tree.xview)
        
        self.results_tree.heading('#0', text='ID')
        self.results_tree.heading('Date', text='Date')
        self.results_tree.heading('Country', text='Country')
        self.results_tree.heading('Notice', text='Notice #')
        self.results_tree.heading('Platform', text='Platform Name')
        self.results_tree.heading('Action', text='Action')
        self.results_tree.heading('Lat', text='Latitude')
        self.results_tree.heading('Lon', text='Longitude')
        self.results_tree.heading('Score', text='AI Score')
        
        self.results_tree.column('#0', width=40)
        self.results_tree.column('Date', width=90)
        self.results_tree.column('Country', width=90)
        self.results_tree.column('Notice', width=80)
        self.results_tree.column('Platform', width=120)
        self.results_tree.column('Action', width=60)
        self.results_tree.column('Lat', width=100)
        self.results_tree.column('Lon', width=100)
        self.results_tree.column('Score', width=60)
        
        self.results_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        detail_frame = ttk.LabelFrame(self.results_frame, text="Details", padding=5)
        detail_frame.pack(fill='x', padx=5, pady=5)
        
        self.detail_text = scrolledtext.ScrolledText(detail_frame, height=6, wrap=tk.WORD)
        self.detail_text.pack(fill='both', expand=True)
        
        self.results_tree.bind('<<TreeviewSelect>>', self.on_select)
    
    # DOWNLOAD FUNCTIONS
    def log_download(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.download_log.insert('end', f"[{timestamp}] {msg}\n")
        self.download_log.see('end')
        self.root.update_idletasks()
    
    def clear_download_log(self):
        self.download_log.delete('1.0', 'end')
    
    def open_download_folder(self):
        if not os.path.exists(UK_MARINE_FOLDER):
            os.makedirs(UK_MARINE_FOLDER)
        os.startfile(UK_MARINE_FOLDER)
    
    def start_download(self):
        if self.is_running:
            return
        
        self.is_running = True
        self.download_btn.config(state='disabled')
        self.stop_download_btn.config(state='normal')
        
        threading.Thread(target=self.run_download, daemon=True).start()
    
    def stop_download(self):
        self.is_running = False
        self.log_download("Stopping download...")
    
    def run_download(self):
        try:
            # Create folder
            if not os.path.exists(UK_MARINE_FOLDER):
                os.makedirs(UK_MARINE_FOLDER)
                self.log_download(f"Created folder: {UK_MARINE_FOLDER}")
            
            # Setup session
            session = requests.Session()
            session.headers.update(HEADERS)
            
            self.log_download("Fetching PDF list from Admiralty website...")
            
            # Get webpage
            response = session.get(WEEKLY_URL, timeout=30, verify=False)
            response.raise_for_status()
            
            # Extract PDF links
            soup = BeautifulSoup(response.text, 'html.parser')
            pdf_links = []
            
            for link in soup.find_all('a', href=True):
                href = link['href']
                aria_label = link.get('aria-label', '')
                
                if 'DownloadFile' in href and 'Download' in link.get_text():
                    filename = None
                    if 'for' in aria_label.lower():
                        match = re.search(r'for\s+([a-zA-Z0-9_\-]+)', aria_label)
                        if match:
                            filename = match.group(1)
                            if not filename.endswith('.pdf'):
                                filename += '.pdf'
                    
                    if filename:
                        href_clean = href.replace('&amp;', '&')
                        full_url = urljoin(BASE_URL, href_clean)
                        pdf_links.append({'filename': filename, 'url': full_url})
            
            self.log_download(f"Found {len(pdf_links)} PDFs")
            
            if not pdf_links:
                self.log_download("No PDFs found. Check website manually.")
                return
            
            # Download each PDF
            self.root.after(0, lambda: self.download_progress.config(maximum=len(pdf_links)))
            
            success_count = 0
            for i, pdf_info in enumerate(pdf_links, 1):
                if not self.is_running:
                    break
                
                filename = pdf_info['filename']
                url = pdf_info['url']
                filepath = os.path.join(UK_MARINE_FOLDER, filename)
                
                # Skip if exists
                if os.path.exists(filepath):
                    self.log_download(f"SKIP: {filename} (exists)")
                    success_count += 1
                else:
                    try:
                        self.log_download(f"Downloading [{i}/{len(pdf_links)}]: {filename}...")
                        resp = session.get(url, timeout=60, verify=False, stream=True)
                        resp.raise_for_status()
                        
                        with open(filepath, 'wb') as f:
                            for chunk in resp.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                        
                        size = os.path.getsize(filepath) / 1024
                        self.log_download(f"✓ Downloaded: {filename} ({size:.1f} KB)")
                        success_count += 1
                    except Exception as e:
                        self.log_download(f"✗ Error: {filename} - {e}")
                
                self.root.after(0, lambda v=i: self.download_progress.config(value=v))
                time.sleep(1)
            
            self.log_download(f"\nDownload complete: {success_count}/{len(pdf_links)} files")
            
        except Exception as e:
            self.log_download(f"ERROR: {e}")
            import traceback
            self.log_download(traceback.format_exc())
        finally:
            self.root.after(0, self.download_done)
    
    def download_done(self):
        self.is_running = False
        self.download_btn.config(state='normal')
        self.stop_download_btn.config(state='disabled')
        self.download_progress.config(value=0)
    
    # PROCESSING FUNCTIONS
    def log_process(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.process_log.insert('end', f"[{timestamp}] {msg}\n")
        self.process_log.see('end')
        self.root.update_idletasks()
    
    def clear_process_log(self):
        self.process_log.delete('1.0', 'end')
    
    def start_processing(self):
        if self.is_running:
            return
        
        if not os.path.exists(UK_MARINE_FOLDER):
            messagebox.showwarning("No Folder", f"Folder not found:\n{UK_MARINE_FOLDER}\n\nDownload PDFs first!")
            return
        
        self.is_running = True
        self.process_btn.config(state='disabled')
        self.stop_process_btn.config(state='normal')
        self.extracted_data = []
        
        threading.Thread(target=self.run_processing, daemon=True).start()
    
    def stop_processing(self):
        self.is_running = False
        self.log_process("Stopping processing...")
    
    def run_processing(self):
        try:
            self.log_process("Starting PDF processing...")
            self.log_process("Training AI model...")
            self.ai_extractor.train_on_keywords()
            
            # Get all PDFs
            pdf_files = [f for f in os.listdir(UK_MARINE_FOLDER) if f.endswith('.pdf')]
            self.log_process(f"Found {len(pdf_files)} PDF files")
            
            if not pdf_files:
                self.log_process("No PDF files found in folder!")
                return
            
            self.root.after(0, lambda: self.process_progress.config(maximum=len(pdf_files)))
            
            # Process each PDF
            for i, filename in enumerate(pdf_files, 1):
                if not self.is_running:
                    break
                
                filepath = os.path.join(UK_MARINE_FOLDER, filename)
                self.log_process(f"\n[{i}/{len(pdf_files)}] Processing: {filename}")
                
                try:
                    # Extract text from PDF
                    with pdfplumber.open(filepath) as pdf:
                        text = "".join([page.extract_text(x_tolerance=2) or "" for page in pdf.pages])
                    
                    if text.strip():
                        # Parse with AI
                        records = self.ai_extractor.parse_notice_text(text, filepath, filename, 0.05)
                        
                        if records:
                            self.extracted_data.extend(records)
                            self.log_process(f"   ✓ Extracted {len(records)} platform records")
                        else:
                            self.log_process(f"   ℹ No platform data found")
                    else:
                        self.log_process(f"   ✗ Could not extract text")
                    
                except Exception as e:
                    self.log_process(f"   ✗ Error: {e}")
                
                self.root.after(0, lambda v=i: self.process_progress.config(value=v))
            
            self.log_process(f"\n{'='*50}")
            self.log_process(f"Processing complete!")
            self.log_process(f"Total records extracted: {len(self.extracted_data)}")
            self.log_process(f"{'='*50}")
            
            # Update results
            self.root.after(0, self.populate_results)
            
        except Exception as e:
            self.log_process(f"ERROR: {e}")
            import traceback
            self.log_process(traceback.format_exc())
        finally:
            self.root.after(0, self.processing_done)
    
    def processing_done(self):
        self.is_running = False
        self.process_btn.config(state='normal')
        self.stop_process_btn.config(state='disabled')
        self.process_progress.config(value=0)
    
    # RESULTS FUNCTIONS
    def populate_results(self):
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        for i, record in enumerate(self.extracted_data, 1):
            self.results_tree.insert('', 'end', text=str(i), values=(
                record.get('Date', ''),
                record.get('Country', ''),
                record.get('Notice Number', ''),
                record.get('Platform Name', 'Unknown'),
                record.get('Action', ''),
                record.get('Latitude', ''),
                record.get('Longitude', ''),
                record.get('AI Score', '')
            ))
        
        self.record_count.config(text=str(len(self.extracted_data)))
    
    def on_select(self, event):
        sel = self.results_tree.selection()
        if not sel:
            return
        
        idx = int(self.results_tree.item(sel[0])['text']) - 1
        if 0 <= idx < len(self.extracted_data):
            rec = self.extracted_data[idx]
            
            detail = f"Date: {rec.get('Date', 'N/A')}\n"
            detail += f"Country: {rec.get('Country', 'Unknown')}\n"
            detail += f"Notice Number: {rec.get('Notice Number', 'N/A')}\n"
            detail += f"Platform Name: {rec.get('Platform Name', 'Unknown')}\n"
            detail += f"Action: {rec.get('Action', '')}\n"
            detail += f"Coordinates: {rec.get('Latitude', '')}, {rec.get('Longitude', '')}\n"
            detail += f"AI Score: {rec.get('AI Score', '')}\n"
            detail += f"Validation: {rec.get('Validation', '')}\n"
            detail += f"Source File: {rec.get('Source File', '')}\n\n"
            detail += f"Description:\n{rec.get('Full Description', '')}"
            
            self.detail_text.delete('1.0', 'end')
            self.detail_text.insert('1.0', detail)
    
    def apply_filter(self):
        text = self.filter_var.get().lower()
        if not text:
            self.populate_results()
            return
        
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        count = 0
        for i, rec in enumerate(self.extracted_data, 1):
            if any(text in str(v).lower() for v in rec.values()):
                self.results_tree.insert('', 'end', text=str(i), values=(
                    rec.get('Date', ''),
                    rec.get('Country', ''),
                    rec.get('Notice Number', ''),
                    rec.get('Platform Name', 'Unknown'),
                    rec.get('Action', ''),
                    rec.get('Latitude', ''),
                    rec.get('Longitude', ''),
                    rec.get('AI Score', '')
                ))
                count += 1
        
        self.record_count.config(text=f"{count} (filtered)")
    
    def clear_filter(self):
        self.filter_var.set('')
        self.populate_results()
    
    def export_results(self):
        if not self.extracted_data:
            messagebox.showwarning("No Data", "No data to export.")
            return
        
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"UK_Platforms_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if path:
            try:
                df = pd.DataFrame(self.extracted_data)
                
                if path.endswith('.csv'):
                    df.to_csv(path, index=False)
                else:
                    df.to_excel(path, index=False, engine='openpyxl')
                
                messagebox.showinfo("Success", f"Exported {len(self.extracted_data)} records to:\n{path}")
                self.log_process(f"Exported to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")


# --- MAIN ---
if __name__ == "__main__":
    root = tk.Tk()
    app = UKMarineExtractorGUI(root)
    root.mainloop()