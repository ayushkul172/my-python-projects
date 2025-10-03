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
from datetime import datetime, timedelta  # FIXED: was "timedeltaa"

# Selenium imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

import json

# ML/AI imports
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- ENHANCED THEME COLORS ---
BG_COLOR = "#0A0A0A" 
ACCENT_COLOR = "#00BFFF" 
TEXT_COLOR = "#E0E0E0"
DARK_ELEM_COLOR = "#1A1A1A" 
HOVER_COLOR = "#008CBA" 
BORDER_COLOR = "#005080"

# --- CONFIGURATION ---
DOWNLOADS_BASE_PATH = os.path.join(os.path.expanduser("~"), "Downloads")
UK_MARINE_FOLDER = os.path.join(DOWNLOADS_BASE_PATH, "UK_Marine_Notices")
ANGOLA_MARINE_FOLDER = os.path.join(DOWNLOADS_BASE_PATH, "Angola_Marine_Notices")
SETTINGS_FILE = os.path.join(DOWNLOADS_BASE_PATH, "extractor_settings.json")

# UK Configuration
UK_BASE_URL = "https://msi.admiralty.co.uk"
UK_WEEKLY_URL = "https://msi.admiralty.co.uk/NoticesToMariners/Weekly"

# Angola Configuration
ANGOLA_URL = "https://www.sanho.co.za/notices_mariners/navarea_v11_messages.htm"


# Brunei Configuration
BRUNEI_URL = "https://mpabd.gov.bn/notice-to-mariners-2025/"
BRUNEI_MARINE_FOLDER = os.path.join(DOWNLOADS_BASE_PATH, "Brunei_Marine_Notices")


HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'en-GB,en;q=0.9',
}

# --- ENHANCED THEME APPLICATION FUNCTION ---
def apply_dark_blue_theme(root):
    """Configures the Tkinter theme for a sophisticated dark background and electric blue accents."""
    style = ttk.Style()
    
    try:
        style.theme_use('clam')
    except tk.TclError:
        style.theme_use(style.theme_names()[0])

    root.configure(bg=BG_COLOR)
    
    style.configure('TFrame', background=BG_COLOR)
    style.configure('TLabelframe', background=BG_COLOR, foreground=ACCENT_COLOR, 
                    bordercolor=BORDER_COLOR, borderwidth=2, relief='groove') 
    style.configure('TLabelframe.Label', background=BG_COLOR, foreground=ACCENT_COLOR, font=('Inter', 11, 'bold'))

    style.configure('TLabel', background=BG_COLOR, foreground=TEXT_COLOR, font=('Inter', 11))
    
    style.configure('TNotebook', background=DARK_ELEM_COLOR, borderwidth=0) 
    style.configure('TNotebook.Tab', 
                    background=DARK_ELEM_COLOR, 
                    foreground=TEXT_COLOR, 
                    padding=[15, 7],
                    font=('Inter', 11))
    style.map('TNotebook.Tab', 
              background=[('selected', ACCENT_COLOR), ('active', DARK_ELEM_COLOR)],
              foreground=[('selected', BG_COLOR), ('active', TEXT_COLOR)])

    style.configure('TButton', 
                    background=DARK_ELEM_COLOR, 
                    foreground=ACCENT_COLOR, 
                    bordercolor=ACCENT_COLOR,
                    borderwidth=1, 
                    focusthickness=2,
                    focuscolor=ACCENT_COLOR,
                    font=('Inter', 11, 'bold'),
                    padding=10,
                    relief='flat')
    style.map('TButton', 
              background=[('active', ACCENT_COLOR), ('disabled', DARK_ELEM_COLOR)],
              foreground=[('active', BG_COLOR), ('disabled', ACCENT_COLOR)],
              bordercolor=[('active', ACCENT_COLOR)]) 

    style.configure('TEntry', 
                    fieldbackground=DARK_ELEM_COLOR, 
                    foreground=TEXT_COLOR, 
                    bordercolor=BORDER_COLOR,
                    borderwidth=1,
                    insertcolor=ACCENT_COLOR,
                    padding=8)
    style.map('TEntry', bordercolor=[('focus', ACCENT_COLOR)])
    
    style.configure('TProgressbar', 
                    background=ACCENT_COLOR, 
                    troughcolor=DARK_ELEM_COLOR,
                    troughrelief='flat',
                    thickness=12,
                    borderwidth=0)
    
    style.configure('Treeview', 
                    background=DARK_ELEM_COLOR, 
                    foreground=TEXT_COLOR, 
                    fieldbackground=DARK_ELEM_COLOR, 
                    rowheight=30,
                    borderwidth=0,
                    font=('Consolas', 10))
    style.configure('Treeview.Heading', 
                    background=ACCENT_COLOR, 
                    foreground=DARK_ELEM_COLOR,
                    font=('Inter', 11, 'bold'), 
                    padding=[10, 8])
    style.map('Treeview', 
              background=[('selected', HOVER_COLOR)],
              foreground=[('selected', TEXT_COLOR)]) 
    
    style.configure('TScrollbar', 
                    troughcolor=DARK_ELEM_COLOR, 
                    background=BORDER_COLOR, 
                    bordercolor=BG_COLOR, 
                    arrowcolor=ACCENT_COLOR)
    style.map('TScrollbar', background=[('active', ACCENT_COLOR)])

    style.configure('TRadiobutton', 
                    background=BG_COLOR, 
                    foreground=TEXT_COLOR,
                    font=('Inter', 10))
    style.map('TRadiobutton',
              background=[('active', BG_COLOR)],
              foreground=[('active', ACCENT_COLOR)])

# --- AI EXTRACTOR ---
class MarinePlatformExtractor:
    """AI-based extraction for Marine Notices (UK & Angola)"""
    
    def __init__(self):
        self.load_default_settings()
        
        # UK notice pattern (unchanged)
        self.uk_notice_pattern = re.compile(r'(\d{4})\(?\w?\)?/\d{2}')
        
        # Angola notice pattern - IMPROVED
        self.angola_notice_pattern = re.compile(
            r'NAVAREA\s+VII\s+(\d+)\s+OF\s+(\d{4})', 
            re.IGNORECASE
        )
        
          # Brunei notice pattern
        self.brunei_notice_pattern = re.compile(
            r'NM\s+(\d{2})/(\d{4})', 
            re.IGNORECASE
        )
        
        
        
        
        # UK coordinate patterns (unchanged)
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
            r"M/V\s+([A-Z\s]+)",
            r"RIG\s+([A-Z\s]+)",
        ]
        
        self.vectorizer = TfidfVectorizer(max_features=100, stop_words='english')
        self.trained = False
    
    def load_default_settings(self):
        """Load default keyword and action settings"""
        self.platform_keywords = [
            'platform', 'platfrom', 'platfom', 'rig', 'jack-up', 'semi-submersible',
            'drillship', 'installation', 'monopod', 'wellhead', 'well head',
            'fpso', 'flng', 'fso', 'fsu', 'drilling unit', 'drilling operations',
            'oil platform', 'gas platform', 'offshore structure', 'mopu', 'tlp',
            'spar platform', 'tension leg', 'compliant tower', 'subsea',
            'marine mining', 'mining vessel', 'anchor spread', 'm/v', 'r/v',
            'oceanographic', 'survey', 'kizomba', 'deepwater', 'skyros', 'gemini',
            'orion', 'eclipse', 'voyager', 'dalia', 'valaris', 'maersk', 'exclusion zone'
        ]
        
        self.action_patterns = {
            'Insert': ['insert', 'establish', 'commence', 'new', 'installed', 'deployed', 'placed', 'positioned', 'construction'],
            'Delete': ['delete', 'remove', 'cancel', 'withdrawn', 'terminated', 'removed', 'ceased'],
            'Amend': ['amend', 'move', 'relocate', 'modify', 'change', 'update', 'moved', 'repositioned']
        }
        
        self.relevance_threshold = 0.05
        self.high_confidence_threshold = 0.2
    
    def load_settings(self, settings_dict):
        """Load settings from dictionary"""
        if 'platform_keywords' in settings_dict:
            self.platform_keywords = settings_dict['platform_keywords']
        if 'action_patterns' in settings_dict:
            self.action_patterns = settings_dict['action_patterns']
        if 'relevance_threshold' in settings_dict:
            self.relevance_threshold = settings_dict['relevance_threshold']
        if 'high_confidence_threshold' in settings_dict:
            self.high_confidence_threshold = settings_dict['high_confidence_threshold']
        
        self.trained = False
        
    def get_settings(self):
        """Return current settings as dictionary"""
        return {
            'platform_keywords': self.platform_keywords,
            'action_patterns': self.action_patterns,
            'relevance_threshold': self.relevance_threshold,
            'high_confidence_threshold': self.high_confidence_threshold
        }
        
    def train_on_keywords(self):
        """Train AI model"""
        training_texts = self.platform_keywords + [
            'offshore drilling platform north sea',
            'oil rig installation UK sector',
            'platform construction works position',
            'floating production storage offloading',
            'jack-up drilling unit established',
            'marine mining vessel anchor spread',
            'rig list angola namibia south africa'
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
    
    def extract_category(self, text):
        """Identify which platform keyword/category triggered detection"""
        text_lower = text.lower()
        
        categories = {
            'Fixed Platform': ['platform', 'platfrom', 'platfom', 'monopod', 'oil platform', 'gas platform'],
            'Mobile Rig': ['rig', 'jack-up', 'semi-submersible', 'drillship', 'drilling unit', 'drilling operations'],
            'Floating Production': ['fpso', 'flng', 'fso', 'fsu', 'spar platform', 'tension leg', 'tlp', 'kizomba', 'dalia'],
            'Subsea/Wellhead': ['wellhead', 'well head', 'subsea'],
            'Offshore Structure': ['installation', 'offshore structure', 'mopu', 'compliant tower'],
            'Marine Mining': ['marine mining', 'mining vessel', 'anchor spread'],
            'Survey/Research': ['oceanographic', 'survey', 'r/v', 'research vessel']
        }
        
        matched_keywords = []
        for kw in self.platform_keywords:
            if kw in text_lower:
                matched_keywords.append(kw)
        
        for category, keywords in categories.items():
            if any(kw in matched_keywords for kw in keywords):
                matched = [kw for kw in keywords if kw in matched_keywords]
                return f"{category} ({', '.join(matched[:2])})"
        
        if matched_keywords:
            return f"General ({matched_keywords[0]})"
        
        return "Unclassified"
    
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
    
    
      
      
    def extract_notice_number(self, text, source_type='uk'):       
        """Extract notice number based on source type"""
        if source_type == 'angola':
            match = self.angola_notice_pattern.search(text)
            return match.group(0) if match else ""
        elif source_type == 'brunei':
            match = self.brunei_notice_pattern.search(text)
            return match.group(0) if match else ""
        else:
            match = self.uk_notice_pattern.search(text)
            return match.group(0) if match else ""
    
    def extract_publication_date_from_text(self, text, source_type='uk'):
        """Extract publication date from content"""
        if source_type == 'angola':
            # Angola format: (01 SEP 25) or similar
            angola_date_patterns = [
                r'\((\d{1,2}\s+[A-Z]{3}\s+\d{2})\)',
                r'(\d{1,2}\s+[A-Z]{3}\s+\d{2})\s*[\n:]',
            ]
            for pattern in angola_date_patterns:
                match = re.search(pattern, text[:500], re.IGNORECASE)
                if match:
                    try:
                        date_str = match.group(1).strip()
                        parsed_date = datetime.strptime(date_str, "%d %b %y")
                        return parsed_date.strftime("%d %B %Y")
                    except:
                        return match.group(1).strip()
        
        # UK format
        date_pattern = r'Weekly\s+Edition\s+\d+\s+(\d{1,2}\s+\w+\s+\d{4})'
        match = re.search(date_pattern, text, re.IGNORECASE)
        
        if match:
            return match.group(1).strip()
        
        simple_date = r'(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})'
        match = re.search(simple_date, text[:2000], re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        return "Date Unknown"
    
    def extract_country_from_text(self, text):
        """Extract country"""
        countries = {
            'UK': ['england', 'scotland', 'wales', 'northern ireland', 'united kingdom', 'british'],
            'North Sea': ['north sea'],
            'Angola': ['angola', 'angolan'],
            'Namibia': ['namibia', 'namibian'],
            'South Africa': ['south africa', 'rsa', 'south african'],
            'Mozambique': ['mozambique', 'mozambican'],
            'Madagascar': ['madagascar'],
            'India': ['india', 'indian'],
            'Norway': ['norway', 'norwegian'],
            'USA': ['united states', 'america', 'american', 'usa', 'gulf of mexico'],
        }
        
        text_lower = text.lower()
        for country, keywords in countries.items():
            if any(kw in text_lower for kw in keywords):
                return country
        
        return "Unknown"
    
    def generate_ai_observation(self, record, text_snippet):
        """Generate human-readable observation"""
        platform_name = record.get('Platform Name', 'Unknown Platform')
        action = record.get('Action', 'Info')
        category = record.get('Category', 'Unclassified')
        country = record.get('Country', 'Unknown')
        lat = record.get('Latitude', '')
        lon = record.get('Longitude', '')
        ai_score = float(record.get('AI Score', '0'))
        
        category_type = category.split('(')[0].strip()
        
        if action == 'Insert':
            observation = f"NEW {category_type.upper()}: '{platform_name}' has been established"
            if country != 'Unknown':
                observation += f" in {country} waters"
            observation += f" at coordinates {lat}, {lon}."
            observation += " MARINERS: Update your charts and maintain safe distance during transit."
            
            if 'Drilling' in category or 'Rig' in category:
                observation += " Expect possible supply vessel traffic and safety zones."
            elif 'Floating Production' in category or 'FPSO' in category:
                observation += " Permanent installation - verify exclusion zones."
            elif 'Subsea' in category or 'Wellhead' in category:
                observation += " Subsea obstruction - avoid anchoring in vicinity."
            elif 'Mining' in category:
                observation += " Anchor spread deployed - maintain safe distance from unlit buoys."
                
        elif action == 'Delete':
            observation = f"REMOVED: {category_type} '{platform_name}' has been withdrawn"
            if country != 'Unknown':
                observation += f" from {country} waters"
            observation += f" (was at {lat}, {lon})."
            observation += " This position is now clear for navigation."
            
        elif action == 'Amend':
            observation = f"RELOCATED: {category_type} '{platform_name}' has moved"
            if country != 'Unknown':
                observation += f" within {country} waters"
            observation += f" to new position {lat}, {lon}."
            observation += " Update your navigation systems with the new coordinates."
            
        else:
            observation = f"INFO UPDATE: {category_type} '{platform_name}'"
            if country != 'Unknown':
                observation += f" in {country} waters"
            observation += f" at {lat}, {lon}."
            observation += " Verify details before approaching this area."
        
        if ai_score < 0.15:
            observation += " [Low confidence - manual verification recommended]"
        elif ai_score > 0.3:
            observation += " [High confidence detection]"
        
        return observation
    
    # ==================== ANGOLA-SPECIFIC FUNCTIONS (NEW/IMPROVED) ====================
    
    def extract_angola_rig_list_entries(self, text):
        """
        IMPROVED: Extract individual rig entries from Angola rig list format
        Format: A. 06-00.35 S 011-00.15 E KIZOMBA B
        NOTE: Text is typically single-line with no breaks between entries
        """
        rig_entries = []
        
        # FIXED: Pattern for single-line format (no newlines between entries)
        # Lookahead detects next entry by: space + letter + period + space + digits
        rig_pattern = r'([A-Z])\.\s+(\d{2}-\d{2}\.\d{2}\s+[NS])\s+(\d{3}-\d{2}\.\d{2}\s+[EW])\s+(.+?)(?=\s+[A-Z]\.\s+\d{2}-|\s+\d+\.|$)'
        
        matches = re.findall(rig_pattern, text, re.DOTALL)
        
        for letter, lat, lon, name in matches:
            # Clean rig name - remove trailing newlines and extra spaces
            rig_name = ' '.join(name.strip().split())
            
            # Skip if name is too short or looks like garbage
            if len(rig_name) > 1 and not rig_name.isdigit():
                rig_entries.append({
                    'letter': letter,
                    'latitude': lat.strip(),
                    'longitude': lon.strip(),
                    'name': rig_name
                })
        
        return rig_entries
    
    def extract_angola_mining_vessel_entries(self, text):
        """
        IMPROVED: Extract mining vessel entries from Angola messages
        Format: A. 06-00.35 S 011-00.15 E M/V VESSEL NAME (anchor info)
        NOTE: Single-line format with no breaks
        """
        vessel_entries = []
        
        # FIXED: Pattern for single-line mining vessel format
        vessel_pattern = r'([A-Z])\.\s+(\d{2}-\d{2}\.\d{2}\s+[NS])\s+(\d{3}-\d{2}\.\d{2}\s+[EW])\s+(M/V\s+[A-Z\s]+?)(?:\s*\(([^\)]+)\))?(?=\s+[A-Z]\.\s+\d{2}-|\s+\d+\.|$)'
        
        matches = re.findall(vessel_pattern, text, re.IGNORECASE)
        
        for letter, lat, lon, vessel_name, anchor_info in matches:
            vessel_entries.append({
                'letter': letter,
                'latitude': lat.strip(),
                'longitude': lon.strip(),
                'name': vessel_name.strip(),
                'anchor_info': anchor_info.strip() if anchor_info else ''
            })
        
        return vessel_entries
    
    def extract_angola_survey_vessel_entry(self, text):
        """
        IMPROVED: Extract survey vessel information
        Format: R/V VESSEL NAME (C/S CALLSIGN)
        """
        vessel_pattern = r'(R/V\s+[A-Z\s]+)(?:\s*\(C/S\s+([A-Z0-9]+)\))?'
        match = re.search(vessel_pattern, text, re.IGNORECASE)
        
        if match:
            return {
                'name': match.group(1).strip(),
                'callsign': match.group(2) if match.group(2) else ''
            }
        return None
    
    def parse_angola_navarea_message(self, message_text, source_url, source_file, message_date):
        """
        IMPROVED: Parse individual Angola NAVAREA VII message with proper structure handling
        """
        extracted_records = []
        
        # Calculate relevance with Angola-specific boost
        text_lower = message_text.lower()
        relevance_score = self.calculate_relevance_score(message_text)
        
        # Boost score for Angola-specific keywords
        angola_boost_keywords = ['rig list', 'marine mining', 'fpso', 'exclusion zone', 'm/v', 'r/v']
        for keyword in angola_boost_keywords:
            if keyword in text_lower:
                relevance_score = min(relevance_score + 0.1, 1.0)
        
        # Lower threshold for Angola to catch more valid messages
        if relevance_score < 0.03:
            return extracted_records
        
        # Extract notice number
        notice_num = self.extract_notice_number(message_text, source_type='angola')
        country = self.extract_country_from_text(message_text)
        
        # Check message type and handle accordingly
        if 'rig list' in text_lower:
            # RIG LIST MESSAGE - Extract individual rigs
            rig_entries = self.extract_angola_rig_list_entries(message_text)
            
            for entry in rig_entries:
                platform_name = entry['name']
                category = self.extract_category(platform_name)
                action = "Info"  # Rig lists are informational
                
                # Build description
                description = f"Rig list entry: {entry['letter']}. {platform_name} at {entry['latitude']} {entry['longitude']}"
                
                # Check for exclusion zone info
                if 'exclusion zone' in text_lower:
                    zone_match = re.search(r'(\d+)\s*NM\s+EXCLUSION\s+ZONE', message_text, re.IGNORECASE)
                    if zone_match:
                        description += f" [{zone_match.group(1)}NM exclusion zone]"
                
                validation = "OK" if relevance_score > 0.15 else "Needs Review"
                
                record = {
                    "Date": message_date,
                    "Country": country,
                    "Source File": source_file,
                    "Notice Number": notice_num,
                    "Platform Name": platform_name,
                    "Category": category if category != "Unclassified" else "Mobile Rig (rig list)",
                    "Action": action,
                    "Latitude": entry['latitude'],
                    "Longitude": entry['longitude'],
                    "Full Description": description,
                    "AI Score": f"{relevance_score:.2f}",
                    "Validation": validation,
                    "PDF Link": source_url,
                    "Source URL": source_url
                }
                
                record["AI Observation"] = self.generate_ai_observation(record, message_text[:300])
                extracted_records.append(record)
        
        elif 'marine mining' in text_lower:
            # MARINE MINING MESSAGE - Extract individual vessels
            vessel_entries = self.extract_angola_mining_vessel_entries(message_text)
            
            for entry in vessel_entries:
                platform_name = entry['name']
                category = "Marine Mining (mining vessel, anchor spread)"
                action = "Info"
                
                description = f"Mining vessel: {entry['letter']}. {platform_name} at {entry['latitude']} {entry['longitude']}"
                if entry['anchor_info']:
                    description += f" ({entry['anchor_info']})"
                
                validation = "OK" if relevance_score > 0.15 else "Needs Review"
                
                record = {
                    "Date": message_date,
                    "Country": country,
                    "Source File": source_file,
                    "Notice Number": notice_num,
                    "Platform Name": platform_name,
                    "Category": category,
                    "Action": action,
                    "Latitude": entry['latitude'],
                    "Longitude": entry['longitude'],
                    "Full Description": description,
                    "AI Score": f"{relevance_score:.2f}",
                    "Validation": validation,
                    "PDF Link": source_url,
                    "Source URL": source_url
                }
                
                record["AI Observation"] = self.generate_ai_observation(record, message_text[:300])
                extracted_records.append(record)
        
        elif 'oceanographic' in text_lower or 'survey' in text_lower:
            # SURVEY MESSAGE - Extract vessel and coordinates
            vessel_info = self.extract_angola_survey_vessel_entry(message_text)
            
            if vessel_info:
                platform_name = vessel_info['name']
                category = "Survey/Research (oceanographic, survey)"
                action = "Info"
                
                # Extract survey area coordinates (A. B. C. D. format)
                coord_pattern = r'([A-Z])\.\s+(\d{2}-\d{2}(?:\.\d{2})?\s+[NS])\s+(\d{3}-\d{2}(?:\.\d{2})?\s+[EW])'
                coord_matches = re.findall(coord_pattern, message_text)
                
                if coord_matches:
                    # Create one record per coordinate point
                    for letter, lat, lon in coord_matches[:10]:  # Limit to 10 points
                        description = f"Survey vessel {platform_name}"
                        if vessel_info['callsign']:
                            description += f" (C/S {vessel_info['callsign']})"
                        description += f" - Survey area point {letter}"
                        
                        validation = "OK" if relevance_score > 0.15 else "Needs Review"
                        
                        record = {
                            "Date": message_date,
                            "Country": country,
                            "Source File": source_file,
                            "Notice Number": notice_num,
                            "Platform Name": platform_name,
                            "Category": category,
                            "Action": action,
                            "Latitude": lat,
                            "Longitude": lon,
                            "Full Description": description,
                            "AI Score": f"{relevance_score:.2f}",
                            "Validation": validation,
                            "PDF Link": source_url,
                            "Source URL": source_url
                        }
                        
                        record["AI Observation"] = self.generate_ai_observation(record, message_text[:300])
                        extracted_records.append(record)
        
        else:
            # GENERIC MESSAGE - Use fallback coordinate extraction
            # Try Angola dash-separated coordinate format first
            # FIXED: Pattern for single-line entries
            angola_coord_pattern = r'([A-Z])\.\s+(\d{2}-\d{2}(?:\.\d{2})?\s+[NS])\s+(\d{3}-\d{2}(?:\.\d{2})?\s+[EW])\s*(.+?)(?=\s+[A-Z]\.\s+\d{2}-|\s+\d+\.|$)'
            coord_matches = re.findall(angola_coord_pattern, message_text)
            
            if coord_matches:
                # For generic messages, try to extract platform name from context
                platform_name = self.extract_platform_name(message_text)
                
                # If platform name is still generic, try to get it from after coordinates
                if platform_name == "Unknown Platform" and coord_matches:
                    # Use the text after first coordinate as platform name
                    first_match_text = coord_matches[0][3] if len(coord_matches[0]) > 3 else ""
                    if first_match_text:
                        # Clean and use first few words
                        cleaned = ' '.join(first_match_text.strip().split()[:5])
                        if len(cleaned) > 2:
                            platform_name = cleaned
                
                category = self.extract_category(message_text)
                action = self.extract_action(message_text)
                
                lines = message_text.split('\n')
                description = ' '.join([l.strip() for l in lines if l.strip()])[:800]
                
                for letter, lat, lon, *extra in coord_matches[:10]:
                    validation = "OK" if relevance_score > 0.15 else "Needs Review"
                    
                    record = {
                        "Date": message_date,
                        "Country": country,
                        "Source File": source_file,
                        "Notice Number": notice_num,
                        "Platform Name": platform_name,
                        "Category": category,
                        "Action": action,
                        "Latitude": lat,
                        "Longitude": lon,
                        "Full Description": description,
                        "AI Score": f"{relevance_score:.2f}",
                        "Validation": validation,
                        "PDF Link": source_url,
                        "Source URL": source_url
                    }
                    
                    record["AI Observation"] = self.generate_ai_observation(record, message_text[:300])
                    extracted_records.append(record)
        
        return extracted_records
    
    # ==================== MAIN PARSING ENTRY POINT ====================
    
    def parse_notice_text(self, text, url, source_file, source_type='uk'):
        """Main parsing entry point - routes to UK or Angola parser"""
        if source_type == 'angola':
            # Split into individual NAVAREA messages
            messages = re.split(r'(?=NAVAREA\s+VII\s+\d+\s+OF\s+\d{4})', text, flags=re.IGNORECASE)
            
            extracted_records = []
            for message in messages:
                if len(message.strip()) < 50:
                    continue
                
                message_date = self.extract_publication_date_from_text(message, source_type='angola')
                records = self.parse_angola_navarea_message(message, url, source_file, message_date)
                extracted_records.extend(records)
            
            return extracted_records
        
        else:  # UK format - UNCHANGED
            if self.uk_notice_pattern.search(text):
                return self.parse_uk_notice_format(text, url, source_file)
            
            # Generic fallback
            date = self.extract_publication_date_from_text(text, source_type='uk')
            extracted_records = []
            paragraphs = re.split(r'\n\s*\n', text.strip())
            
            for paragraph in paragraphs:
                relevance_score = self.calculate_relevance_score(paragraph)
                
                if relevance_score < self.relevance_threshold:
                    continue
                
                country = self.extract_country_from_text(paragraph)
                category = self.extract_category(paragraph)
                
                for pattern in self.coord_patterns:
                    matches = list(pattern.finditer(paragraph.replace('\n', ' ')))
                    
                    for match in matches:
                        latitude = match.group(1).strip()
                        longitude = match.group(2).strip()
                        
                        platform_name = self.extract_platform_name(paragraph)
                        action = self.extract_action(paragraph)
                        description = ' '.join(paragraph.replace('\n', ' ').split())
                        
                        validation = "OK" if relevance_score > self.high_confidence_threshold else "Needs Review"
                        
                        record = {
                            "Date": date,
                            "Country": country,
                            "Source File": os.path.basename(source_file),
                            "Notice Number": "",
                            "Platform Name": platform_name,
                            "Category": category,
                            "Action": action,
                            "Latitude": re.sub(r'\s+', ' ', latitude),
                            "Longitude": re.sub(r'\s+', ' ', longitude),
                            "Full Description": description[:500],
                            "AI Score": f"{relevance_score:.2f}",
                            "Validation": validation,
                            "PDF Link": url,
                            "Source URL": url
                        }
                        
                        record["AI Observation"] = self.generate_ai_observation(record, paragraph[:300])
                        extracted_records.append(record)
                        break
            
            return extracted_records
    
    def parse_uk_notice_format(self, text, url, source_file):
        """Parse UK Notice format - UNCHANGED"""
        extracted_records = []
        date = self.extract_publication_date_from_text(text, source_type='uk')
        notice_sections = re.split(r'(?=\d{4}\(?\w?\)?/\d{2})', text)
        
        for section in notice_sections:
            if len(section) < 50:
                continue
            
            relevance_score = self.calculate_relevance_score(section)
            if relevance_score < self.relevance_threshold:
                continue
            
            notice_num = self.extract_notice_number(section, source_type='uk')
            country = self.extract_country_from_text(section)
            category = self.extract_category(section)
            
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
                validation = "OK" if relevance_score > self.high_confidence_threshold else "Needs Review"
                
                record = {
                    "Date": date,
                    "Country": country,
                    "Source File": os.path.basename(source_file),
                    "Notice Number": notice_num,
                    "Platform Name": platform_name,
                    "Category": category,
                    "Action": action,
                    "Latitude": lat_clean,
                    "Longitude": lon_clean,
                    "Full Description": description,
                    "AI Score": f"{relevance_score:.2f}",
                    "Validation": validation,
                    "PDF Link": url,
                    "Source URL": url
                }
                
                record["AI Observation"] = self.generate_ai_observation(record, section[:300])
                extracted_records.append(record)
                break
        
        return extracted_records


# --- SETTINGS DIALOG ---
class SettingsDialog:
    def __init__(self, parent, extractor):
        self.result = None
        self.extractor = extractor
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Classification Settings")
        self.dialog.geometry("800x700")
        self.dialog.configure(bg=BG_COLOR)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        apply_dark_blue_theme(self.dialog)
        
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        kw_frame = ttk.LabelFrame(main_frame, text="Platform Keywords (one per line)", padding=10)
        kw_frame.pack(fill='both', expand=True, pady=5)
        
        self.keywords_text = scrolledtext.ScrolledText(kw_frame, height=10, wrap=tk.WORD,
                                                       bg=DARK_ELEM_COLOR, fg=TEXT_COLOR,
                                                       insertbackground=ACCENT_COLOR,
                                                       font=('Consolas', 10))
        self.keywords_text.pack(fill='both', expand=True)
        self.keywords_text.insert('1.0', '\n'.join(extractor.platform_keywords))
        
        action_frame = ttk.LabelFrame(main_frame, text="Action Patterns", padding=10)
        action_frame.pack(fill='both', expand=True, pady=5)
        
        self.action_texts = {}
        for action, keywords in extractor.action_patterns.items():
            af = ttk.Frame(action_frame)
            af.pack(fill='x', pady=2)
            ttk.Label(af, text=f"{action}:", width=10).pack(side='left')
            entry = ttk.Entry(af, width=60)
            entry.pack(side='left', fill='x', expand=True, padx=5)
            entry.insert(0, ', '.join(keywords))
            self.action_texts[action] = entry
        
        thresh_frame = ttk.LabelFrame(main_frame, text="Detection Thresholds", padding=10)
        thresh_frame.pack(fill='x', pady=5)
        
        ttk.Label(thresh_frame, text="Relevance Threshold (0.0-1.0):").grid(row=0, column=0, sticky='w', pady=5)
        self.relevance_entry = ttk.Entry(thresh_frame, width=15)
        self.relevance_entry.grid(row=0, column=1, padx=5, pady=5)
        self.relevance_entry.insert(0, str(extractor.relevance_threshold))
        
        ttk.Label(thresh_frame, text="High Confidence Threshold (0.0-1.0):").grid(row=1, column=0, sticky='w', pady=5)
        self.confidence_entry = ttk.Entry(thresh_frame, width=15)
        self.confidence_entry.grid(row=1, column=1, padx=5, pady=5)
        self.confidence_entry.insert(0, str(extractor.high_confidence_threshold))
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=10)
        
        ttk.Button(btn_frame, text="Save Settings", command=self.save_settings).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Reset to Defaults", command=self.reset_defaults).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.cancel).pack(side='right', padx=5)
        
    def save_settings(self):
        try:
            keywords_text = self.keywords_text.get('1.0', 'end').strip()
            keywords = [k.strip() for k in keywords_text.split('\n') if k.strip()]
            
            action_patterns = {}
            for action, entry in self.action_texts.items():
                keywords_str = entry.get().strip()
                action_patterns[action] = [k.strip() for k in keywords_str.split(',') if k.strip()]
            
            relevance = float(self.relevance_entry.get())
            confidence = float(self.confidence_entry.get())
            
            if not (0 <= relevance <= 1) or not (0 <= confidence <= 1):
                messagebox.showerror("Invalid Input", "Thresholds must be between 0.0 and 1.0")
                return
            
            self.result = {
                'platform_keywords': keywords,
                'action_patterns': action_patterns,
                'relevance_threshold': relevance,
                'high_confidence_threshold': confidence
            }
            
            self.dialog.destroy()
            
        except ValueError as e:
            messagebox.showerror("Invalid Input", f"Please check your input values:\n{e}")
    
    def reset_defaults(self):
        self.extractor.load_default_settings()
        self.keywords_text.delete('1.0', 'end')
        self.keywords_text.insert('1.0', '\n'.join(self.extractor.platform_keywords))
        
        for action, entry in self.action_texts.items():
            entry.delete(0, 'end')
            entry.insert(0, ', '.join(self.extractor.action_patterns[action]))
        
        self.relevance_entry.delete(0, 'end')
        self.relevance_entry.insert(0, str(self.extractor.relevance_threshold))
        
        self.confidence_entry.delete(0, 'end')
        self.confidence_entry.insert(0, str(self.extractor.high_confidence_threshold))
    
    def cancel(self):
        self.dialog.destroy()


# TODO: --- GUI APPLICATION ---
class MarineExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Marine Notice Extractor - UK & Angola")
        self.root.geometry("1600x850")
        

        apply_dark_blue_theme(self.root)
        

        self.ai_extractor = MarinePlatformExtractor()
        self.load_settings()
        

        self.extracted_data = []
        self.is_running = False
        self.source_type = tk.StringVar(value='uk')
        

        self.setup_ui()
        

    def load_settings(self):
        """Load settings from file"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    settings = json.load(f)
                self.ai_extractor.load_settings(settings)
            except:
                pass
    

    def save_settings(self):
        """Save settings to file"""
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(self.ai_extractor.get_settings(), f, indent=2)
        except:
            pass
        

    def setup_ui(self):
        """Create GUI"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Tab 1: Download
        self.download_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.download_frame, text='1. Download Data')
        self.setup_download_tab()
        
        # Tab 2: Process
        self.process_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.process_frame, text='2. Process Data')
        self.setup_process_tab()
        
        # Tab 3: Results
        self.results_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.results_frame, text='3. Results')
        self.setup_results_tab()
    

    def setup_download_tab(self):
        """Data Download tab"""
        source_frame = ttk.LabelFrame(self.download_frame, text="Select Data Source", padding=10)
        source_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Radiobutton(source_frame, text="UK Admiralty (PDF Downloads)", 
                       variable=self.source_type, value='uk',
                       command=self.update_download_info).pack(anchor='w', pady=2)
        ttk.Radiobutton(source_frame, text="Angola NAVAREA VII (HTML Scraping)", 
                       variable=self.source_type, value='angola',
                       command=self.update_download_info).pack(anchor='w', pady=2)
        ttk.Radiobutton(source_frame, text="Brunei MPABD (PDF Downloads)", 
               variable=self.source_type, value='brunei',
               command=self.update_download_info).pack(anchor='w', pady=2)
        
        self.info_frame = ttk.LabelFrame(self.download_frame, text="Source Information", padding=10)
        self.info_frame.pack(fill='x', padx=10, pady=5)
        
        self.info_label = ttk.Label(self.info_frame, text="", justify='left')
        self.info_label.pack()
        
        button_frame = ttk.Frame(self.download_frame)
        button_frame.pack(fill='x', padx=10, pady=5)
        
        self.download_btn = ttk.Button(button_frame, text="Start Download", command=self.start_download)
        self.download_btn.pack(side='left', padx=5)
        
        self.stop_download_btn = ttk.Button(button_frame, text="Stop", command=self.stop_download, state='disabled')
        self.stop_download_btn.pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Open Folder", command=self.open_download_folder).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Clear Log", command=self.clear_download_log).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Open Brunei Site (Manual)",
                   command=self.open_brunei_website).pack(side='left', padx=5)
        
        self.download_progress = ttk.Progressbar(self.download_frame, mode='determinate')
        self.download_progress.pack(fill='x', padx=10, pady=5)
        
        log_frame = ttk.LabelFrame(self.download_frame, text="Download Log", padding=5)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.download_log = scrolledtext.ScrolledText(log_frame, height=20, wrap=tk.WORD,
                                                      bg=DARK_ELEM_COLOR, fg=TEXT_COLOR,
                                                      insertbackground=ACCENT_COLOR,
                                                      borderwidth=0, relief="flat",
                                                      font=('Consolas', 9))
        self.download_log.pack(fill='both', expand=True)
        
        self.update_download_info()
    def update_download_info(self):

        """Update info based on selected source"""
        if self.source_type.get() == 'uk':
            info = f"Source: UK Admiralty Maritime Safety Information\nURL: {UK_WEEKLY_URL}\nSaves to: {UK_MARINE_FOLDER}\nType: PDF files (Weekly bulletins)"
        elif self.source_type.get() == 'brunei':
            info = f"Source: Brunei Maritime Port Authority\nURL: {BRUNEI_URL}\nSaves to: {BRUNEI_MARINE_FOLDER}\nType: PDF files (2025 Notices)"
        else:
            info = f"Source: South African Navy Hydrographic Office\nURL: {ANGOLA_URL}\nSaves to: {ANGOLA_MARINE_FOLDER}\nType: HTML content (Latest month of messages)"
        

        self.info_label.config(text=info)
    def setup_process_tab(self):

 
        """Data Processing tab"""
        info_frame = ttk.LabelFrame(self.process_frame, text="Process Downloaded Data", padding=10)
        info_frame.pack(fill='x', padx=10, pady=5)
        
        info_text = "Processes downloaded files and extracts platform information using AI"
        ttk.Label(info_frame, text=info_text, justify='left').pack()
        
        button_frame = ttk.Frame(self.process_frame)
        button_frame.pack(fill='x', padx=10, pady=5)
        
        self.process_btn = ttk.Button(button_frame, text="Process All Data", command=self.start_processing)
        self.process_btn.pack(side='left', padx=5)
        
        self.stop_process_btn = ttk.Button(button_frame, text="Stop", command=self.stop_processing, state='disabled')
        self.stop_process_btn.pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Settings", command=self.open_settings).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Clear Log", command=self.clear_process_log).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Export Results", command=self.export_results).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Angola Summary", command=self.export_angola_summary).pack(side='left', padx=5)
        
        self.process_progress = ttk.Progressbar(self.process_frame, mode='determinate')
        self.process_progress.pack(fill='x', padx=10, pady=5)
        
        log_frame = ttk.LabelFrame(self.process_frame, text="Processing Log", padding=5)
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.process_log = scrolledtext.ScrolledText(log_frame, height=20, wrap=tk.WORD,
                                                     bg=DARK_ELEM_COLOR, fg=TEXT_COLOR,
                                                     insertbackground=ACCENT_COLOR,
                                                     borderwidth=0, relief="flat",
                                                     font=('Consolas', 9))
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
        self.record_count = ttk.Label(toolbar, text="0", font=('Inter', 10, 'bold'))
        self.record_count.pack(side='right')
        
        tree_frame = ttk.Frame(self.results_frame)
        tree_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        self.results_tree = ttk.Treeview(tree_frame, 
                                         columns=('Date', 'Country', 'Notice', 'Platform', 'Category', 'Action', 'Lat', 'Lon', 'Score', 'Observation'),
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
        self.results_tree.heading('Category', text='Category')
        self.results_tree.heading('Action', text='Action')
        self.results_tree.heading('Lat', text='Latitude')
        self.results_tree.heading('Lon', text='Longitude')
        self.results_tree.heading('Score', text='AI Score')
        self.results_tree.heading('Observation', text='AI Observation')
        
        self.results_tree.column('#0', width=40)
        self.results_tree.column('Date', width=120)
        self.results_tree.column('Country', width=90)
        self.results_tree.column('Notice', width=100)
        self.results_tree.column('Platform', width=120)
        self.results_tree.column('Category', width=180)
        self.results_tree.column('Action', width=60)
        self.results_tree.column('Lat', width=100)
        self.results_tree.column('Lon', width=100)
        self.results_tree.column('Score', width=60)
        self.results_tree.column('Observation', width=500)
        
        self.results_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        detail_frame = ttk.LabelFrame(self.results_frame, text="Full Details", padding=5)
        detail_frame.pack(fill='x', padx=5, pady=5)
        
        self.detail_text = scrolledtext.ScrolledText(detail_frame, height=8, wrap=tk.WORD,
                                                     bg=DARK_ELEM_COLOR, fg=TEXT_COLOR,
                                                     insertbackground=ACCENT_COLOR,
                                                     borderwidth=0, relief="flat",
                                                     font=('Consolas', 9))
        self.detail_text.pack(fill='both', expand=True)
        
        self.results_tree.bind('<<TreeviewSelect>>', self.on_select)
    def open_settings(self):

        """Open settings dialog"""
        dialog = SettingsDialog(self.root, self.ai_extractor)
        self.root.wait_window(dialog.dialog)
        

        if dialog.result:
            self.ai_extractor.load_settings(dialog.result)
            self.save_settings()
            messagebox.showinfo("Settings Saved", "Classification settings have been updated.")

    def fallback_to_manual_brunei(self):
        """Fallback to manual download instructions"""
        import webbrowser
        self.log_download("")
        self.log_download("=" * 70)
        self.log_download("MANUAL DOWNLOAD INSTRUCTIONS:")
        self.log_download(f"1. Opening website in browser: {BRUNEI_URL}")
        self.log_download(f"2. Download PDFs manually to: {BRUNEI_MARINE_FOLDER}")
        self.log_download("3. Then use 'Process Data' tab")
        self.log_download("=" * 70)
        webbrowser.open(BRUNEI_URL)
        messagebox.showinfo("Manual Download Required", "ChromeDriver not found or failed. Please follow the instructions in the log to download files manually.")


    def open_brunei_website(self):
        """Open Brunei website in default browser for manual download"""
        import webbrowser
        webbrowser.open(BRUNEI_URL)
        messagebox.showinfo(
            "Manual Download",
            f"The Brunei website has been opened in your browser.\n\n"
            f"Please download PDF files manually to:\n{BRUNEI_MARINE_FOLDER}\n\n"
            f"The processing tab will then be able to extract data from them."
        )

    # DOWNLOAD FUNCTIONS
    def log_download(self, msg):

        timestamp = datetime.now().strftime("%H:%M:%S")
        self.download_log.insert('end', f"[{timestamp}] {msg}\n")
        self.download_log.see('end')
        self.root.update_idletasks()
    def clear_download_log(self):

        self.download_log.delete('1.0', 'end')
    def open_download_folder(self):

        folder = UK_MARINE_FOLDER if self.source_type.get() == 'uk' else ANGOLA_MARINE_FOLDER
        if not os.path.exists(folder):
            os.makedirs(folder)
        os.startfile(folder)
    def start_download(self):
        if self.is_running:
            return
        
        self.is_running = True
        self.download_btn.config(state='disabled')
        self.stop_download_btn.config(state='normal')
        
        if self.source_type.get() == 'uk':
            threading.Thread(target=self.run_uk_download, daemon=True).start()
        elif self.source_type.get() == 'brunei':
            threading.Thread(target=self.run_brunei_download, daemon=True).start()
        else:
            threading.Thread(target=self.run_angola_download, daemon=True).start()
    def stop_download(self):

        self.is_running = False
        self.log_download("Stopping download...")
    def run_uk_download(self):

        """Download UK PDF files - UNCHANGED"""
        try:
            if not os.path.exists(UK_MARINE_FOLDER):
                os.makedirs(UK_MARINE_FOLDER)
                self.log_download(f"Created folder: {UK_MARINE_FOLDER}")
            
            session = requests.Session()
            session.headers.update(HEADERS)
            
            self.log_download("Fetching PDF list from UK Admiralty website...")
            
            response = session.get(UK_WEEKLY_URL, timeout=30, verify=False)
            response.raise_for_status()
            
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
                        full_url = urljoin(UK_BASE_URL, href_clean)
                        pdf_links.append({'filename': filename, 'url': full_url})
            
            self.log_download(f"Found {len(pdf_links)} PDFs")
            
            if not pdf_links:
                self.log_download("No PDFs found. Check website manually.")
                return
            
            self.root.after(0, lambda: self.download_progress.config(maximum=len(pdf_links)))
            
            success_count = 0
            for i, pdf_info in enumerate(pdf_links, 1):
                if not self.is_running:
                    break
                
                filename = pdf_info['filename']
                url = pdf_info['url']
                filepath = os.path.join(UK_MARINE_FOLDER, filename)
                
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
                        self.log_download(f"Downloaded: {filename} ({size:.1f} KB)")
                        success_count += 1
                    except Exception as e:
                        self.log_download(f"Error: {filename} - {e}")
                
                self.root.after(0, lambda v=i: self.download_progress.config(value=v))
                time.sleep(1)
            
            self.log_download(f"\nDownload complete: {success_count}/{len(pdf_links)} files")
            
        except Exception as e:
            self.log_download(f"ERROR: {e}")
        finally:
            self.root.after(0, self.download_done)
    def run_angola_download(self):

        """Download Angola NAVAREA VII messages - IMPROVED"""
        try:
            if not os.path.exists(ANGOLA_MARINE_FOLDER):
                os.makedirs(ANGOLA_MARINE_FOLDER)
                self.log_download(f"Created folder: {ANGOLA_MARINE_FOLDER}")
            
            self.log_download("Fetching Angola NAVAREA VII messages...")
            
            session = requests.Session()
            session.headers.update(HEADERS)
            
            response = session.get(ANGOLA_URL, timeout=30, verify=False)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Better text extraction - try to find PRE tag first (common for NAVAREA messages)
            content_area = soup.find('pre')
            if content_area:
                page_text = content_area.get_text(separator='\n')
            else:
                page_text = soup.get_text(separator='\n')
            
            # Clean up excessive blank lines while preserving structure
            page_text = re.sub(r'\n\s*\n\s*\n+', '\n\n', page_text)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Save raw HTML
            html_file = os.path.join(ANGOLA_MARINE_FOLDER, f"navarea_vii_{timestamp}.html")
            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(response.text)
            
            self.log_download(f"Saved HTML: {os.path.basename(html_file)}")
            
            # Split into individual NAVAREA messages
            messages = re.split(r'(?=NAVAREA\s+VII\s+\d+\s+OF\s+\d{4})', page_text, flags=re.IGNORECASE)
            
            # Filter for last 30 days
            one_month_ago = datetime.now() - timedelta(days=30)
            recent_messages = []
            
            for msg in messages:
                if len(msg.strip()) < 50:
                    continue
                
                # Try to extract and parse date
                date_match = re.search(r'\((\d{1,2}\s+[A-Z]{3}\s+\d{2})\)', msg[:500], re.IGNORECASE)
                if date_match:
                    try:
                        date_str = date_match.group(1).strip()
                        msg_date = datetime.strptime(date_str, "%d %b %y")
                        
                        if msg_date >= one_month_ago:
                            recent_messages.append(msg)
                    except:
                        # If date parsing fails, include it anyway
                        recent_messages.append(msg)
                else:
                    # No date found, include it
                    recent_messages.append(msg)
            
            # Save filtered messages
            txt_file = os.path.join(ANGOLA_MARINE_FOLDER, f"navarea_vii_latest_{timestamp}.txt")
            with open(txt_file, 'w', encoding='utf-8') as f:
                separator = '\n\n' + '='*80 + '\n\n'
                f.write(separator.join(recent_messages))
            
            self.log_download(f"Processed {len(messages)} total messages")
            self.log_download(f"Saved {len(recent_messages)} recent messages")
            self.log_download(f"Text file: {os.path.basename(txt_file)}")
            self.log_download(f"\nDownload complete!")
            
        except Exception as e:
            self.log_download(f"ERROR: {e}")
            import traceback
            self.log_download(traceback.format_exc())
        finally:
            self.root.after(0, self.download_done)
    def run_brunei_download(self):
        """Download Brunei notices - with comprehensive error handling"""
        try:
            if not os.path.exists(BRUNEI_MARINE_FOLDER):
                os.makedirs(BRUNEI_MARINE_FOLDER)
                self.log_download(f"Created folder: {BRUNEI_MARINE_FOLDER}")

            self.log_download("Fetching Brunei Notice to Mariners list...")

            session = requests.Session()

            # Comprehensive headers
            enhanced_headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
                'Sec-Fetch-Dest': 'document',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'none',
                'Cache-Control': 'max-age=0',
                'DNT': '1'
            }
            session.headers.update(enhanced_headers)

            # Retry strategy
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry

            retry_strategy = Retry(
                total=3,
                backoff_factor=2,
                status_forcelist=[403, 429, 500, 502, 503, 504]
            )
            adapter = HTTPAdapter(max_retries=retry_strategy)
            session.mount("http://", adapter)
            session.mount("https://", adapter)

            # Human-like delay
            time.sleep(2)

            try:
                response = session.get(BRUNEI_URL, timeout=30, verify=True, allow_redirects=True)
                response.raise_for_status()
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 403:
                    self.log_download("Website is blocking automated access.")
                    self.log_download("Please download PDFs manually from:")
                    self.log_download(BRUNEI_URL)
                    self.log_download(f"Save them to: {BRUNEI_MARINE_FOLDER}")
                    messagebox.showinfo(
                        "Manual Download Required",
                        f"The Brunei website is blocking automated downloads.\n\n"
                        f"Please:\n"
                        f"1. Visit {BRUNEI_URL}\n"
                        f"2. Download PDF files manually\n"
                        f"3. Save them to {BRUNEI_MARINE_FOLDER}\n"
                        f"4. Use the Process tab to extract data"
                    )
                    return
                raise

            soup = BeautifulSoup(response.text, 'html.parser')

            # Find all PDF links on the page
            pdf_links = []
            for link in soup.find_all('a', href=True):
                href = link['href']
                if href.lower().endswith('.pdf'):
                    full_url = urljoin(BRUNEI_URL, href)
                    filename = unquote(os.path.basename(href))
                    if not filename.endswith('.pdf'):
                        link_text = link.get_text(strip=True)
                        if link_text:
                            filename = re.sub(r'[^\w\s-]', '', link_text).strip()
                            filename = re.sub(r'[-\s]+', '_', filename) + '.pdf'
                        else:
                            filename = f"brunei_nm_{len(pdf_links)+1}.pdf"

                    pdf_links.append({'filename': filename, 'url': full_url})

            self.log_download(f"Found {len(pdf_links)} PDF files")

            if not pdf_links:
                self.log_download("No PDFs found. Check website manually.")
                return

            # Download PDFs with delays
            self.root.after(0, lambda: self.download_progress.config(maximum=len(pdf_links)))

            success_count = 0
            for i, pdf_info in enumerate(pdf_links, 1):
                if not self.is_running:
                    break

                filename = pdf_info['filename']
                url = pdf_info['url']
                filepath = os.path.join(BRUNEI_MARINE_FOLDER, filename)

                if os.path.exists(filepath):
                    self.log_download(f"SKIP: {filename} (exists)")
                    success_count += 1
                else:
                    try:
                        self.log_download(f"Downloading [{i}/{len(pdf_links)}]: {filename}...")
                        time.sleep(1)  # Polite delay between downloads

                        resp = session.get(url, timeout=60, verify=True, stream=True)
                        resp.raise_for_status()

                        with open(filepath, 'wb') as f:
                            for chunk in resp.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)

                        size = os.path.getsize(filepath) / 1024
                        self.log_download(f"Downloaded: {filename} ({size:.1f} KB)")
                        success_count += 1
                    except Exception as e:
                        self.log_download(f"Error: {filename} - {e}")

                self.root.after(0, lambda v=i: self.download_progress.config(value=v))

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
        
        self.is_running = True
        self.process_btn.config(state='disabled')
        self.stop_process_btn.config(state='normal')
        self.extracted_data = []
        
        threading.Thread(target=self.run_processing, daemon=True).start()
    def stop_processing(self):

        self.is_running = False
        self.log_process("Stopping processing...")
    def run_processing(self):

        """Process both UK PDFs and Angola text files"""
        try:
            self.log_process("Starting data processing...")
            self.log_process("Training AI model...")
            self.ai_extractor.train_on_keywords()
            
            all_files = []
            
            # Get UK PDF files
            if os.path.exists(UK_MARINE_FOLDER):
                uk_pdfs = [(os.path.join(UK_MARINE_FOLDER, f), 'uk') 
                           for f in os.listdir(UK_MARINE_FOLDER) if f.endswith('.pdf')]
                all_files.extend(uk_pdfs)
                self.log_process(f"Found {len(uk_pdfs)} UK PDF files")
            
            # Get Angola text files
            if os.path.exists(ANGOLA_MARINE_FOLDER):
                angola_txts = [(os.path.join(ANGOLA_MARINE_FOLDER, f), 'angola') 
                              for f in os.listdir(ANGOLA_MARINE_FOLDER) if f.endswith('.txt')]
                all_files.extend(angola_txts)
                self.log_process(f"Found {len(angola_txts)} Angola text files")
                
            # Get Brunei PDF files
            if os.path.exists(BRUNEI_MARINE_FOLDER):
                brunei_pdfs = [(os.path.join(BRUNEI_MARINE_FOLDER, f), 'brunei') 
                   for f in os.listdir(BRUNEI_MARINE_FOLDER) if f.endswith('.pdf')]
                all_files.extend(brunei_pdfs)
                self.log_process(f"Found {len(brunei_pdfs)} Brunei PDF files")
    
            
            if not all_files:
                self.log_process("No files found to process!")
                return
            
            self.root.after(0, lambda: self.process_progress.config(maximum=len(all_files)))
            
            for i, (filepath, source_type) in enumerate(all_files, 1):
                if not self.is_running:
                    break
                
                filename = os.path.basename(filepath)
                self.log_process(f"\n[{i}/{len(all_files)}] Processing: {filename} ({source_type.upper()})")
                
                try:
                    if source_type == 'uk':
                        # Process PDF
                        with pdfplumber.open(filepath) as pdf:
                            text = "".join([page.extract_text(x_tolerance=2) or "" for page in pdf.pages])
                    else:
                        # Process text file
                        with open(filepath, 'r', encoding='utf-8') as f:
                            text = f.read()
                    
                    if text.strip():
                        records = self.ai_extractor.parse_notice_text(text, filepath, filename, source_type=source_type)
                        
                        if records:
                            self.extracted_data.extend(records)
                            self.log_process(f"   Extracted {len(records)} platform records")
                            
                            # Show sample info
                            if source_type == 'angola':
                                rigs = [r for r in records if 'rig' in r.get('Category', '').lower()]
                                if rigs:
                                    self.log_process(f"   - Rigs found: {len(rigs)}")
                                    sample_names = list(set([r['Platform Name'] for r in rigs[:3]]))
                                    self.log_process(f"   - Sample: {', '.join(sample_names)}")
                        else:
                            self.log_process(f"   No platform data found")
                    else:
                        self.log_process(f"   Could not extract text")
                    
                except Exception as e:
                    self.log_process(f"   Error: {e}")
                
                self.root.after(0, lambda v=i: self.process_progress.config(value=v))
            
            self.log_process(f"\n{'='*50}")
            self.log_process(f"Processing complete!")
            self.log_process(f"Total records extracted: {len(self.extracted_data)}")
            
            # Show breakdown
            uk_count = len([r for r in self.extracted_data if r['Source File'].endswith('.pdf')])
            angola_count = len([r for r in self.extracted_data if 'navarea' in r['Source File'].lower()])
            brunei_count = len([r for r in self.extracted_data if any(r['Source File'].startswith(f) for f in os.listdir(BRUNEI_MARINE_FOLDER)) if os.path.exists(BRUNEI_MARINE_FOLDER)])
            self.log_process(f"- UK records: {uk_count}")
            self.log_process(f"- Angola records: {angola_count}")
            self.log_process(f"- Brunei records: {brunei_count}")
            self.log_process(f"{'='*50}")
            
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
                record.get('Category', 'Unclassified'),
                record.get('Action', ''),
                record.get('Latitude', ''),
                record.get('Longitude', ''),
                record.get('AI Score', ''),
                record.get('AI Observation', 'No observation available')
            ))
        
        self.record_count.config(text=str(len(self.extracted_data)))
    def on_select(self, event):

        """Handle tree selection to show details"""
        sel = self.results_tree.selection()
        if not sel:
            return
        
        idx = int(self.results_tree.item(sel[0])['text']) - 1
        if 0 <= idx < len(self.extracted_data):
            rec = self.extracted_data[idx]
            
            detail = f"{'='*80}\n"
            detail += f"PLATFORM LOCATION UPDATE - RECORD #{idx+1}\n"
            detail += f"{'='*80}\n\n"
            detail += f"Date: {rec.get('Date', 'N/A')}\n"
            detail += f"Country/Region: {rec.get('Country', 'Unknown')}\n"
            detail += f"Notice Number: {rec.get('Notice Number', 'N/A')}\n"
            detail += f"Platform Name: {rec.get('Platform Name', 'Unknown')}\n"
            detail += f"Category: {rec.get('Category', 'Unclassified')}\n"
            detail += f"Action Type: {rec.get('Action', '')}\n"
            detail += f"Coordinates: {rec.get('Latitude', '')}, {rec.get('Longitude', '')}\n"
            detail += f"AI Confidence Score: {rec.get('AI Score', '')} | Validation: {rec.get('Validation', '')}\n"
            detail += f"Source File: {rec.get('Source File', '')}\n\n"
            detail += f"{'='*80}\n"
            detail += f"AI OBSERVATION (Plain Language Summary):\n"
            detail += f"{'='*80}\n"
            detail += f"{rec.get('AI Observation', 'No observation available')}\n\n"
            detail += f"{'='*80}\n"
            detail += f"ORIGINAL DESCRIPTION:\n"
            detail += f"{'='*80}\n"
            detail += f"{rec.get('Full Description', 'No description available')}"
            
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
                    rec.get('Category', 'Unclassified'),
                    rec.get('Action', ''),
                    rec.get('Latitude', ''),
                    rec.get('Longitude', ''),
                    rec.get('AI Score', ''),
                    rec.get('AI Observation', 'No observation available')
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
            initialfile=f"Marine_Platforms_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if path:
            try:
                df = pd.DataFrame(self.extracted_data)
                
                column_order = ['Date', 'Country', 'Notice Number', 'Platform Name', 'Category', 
                               'Action', 'Latitude', 'Longitude', 'AI Observation', 'AI Score', 
                               'Validation', 'Full Description', 'Source File', 'PDF Link', 'Source URL']
                
                column_order = [col for col in column_order if col in df.columns]
                df = df[column_order]
                
                if path.endswith('.csv'):
                    df.to_csv(path, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(path, index=False, engine='openpyxl')
                
                messagebox.showinfo("Success", f"Exported {len(self.extracted_data)} records to:\n{path}")
                self.log_process(f"Exported to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")
    def export_angola_summary(self):

        """Export Angola-specific formatted summary - placeholder for full implementation"""
        if not self.extracted_data:
            messagebox.showwarning("No Data", "No data to export.")
            return
        
        angola_records = [r for r in self.extracted_data if 'NAVAREA VII' in r.get('Notice Number', '')]
        
        if not angola_records:
            messagebox.showwarning("No Angola Data", "No Angola NAVAREA VII records found.")
            return
        
        # Create simple summary
        summary = "ANGOLA NAVAREA VII SUMMARY\n"
        summary += "="*80 + "\n\n"
        summary += f"Total Records: {len(angola_records)}\n"
        summary += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        
        for rec in angola_records:
            summary += f"\n{rec.get('Notice Number', 'N/A')}\n"
            summary += f"Platform: {rec.get('Platform Name', 'Unknown')}\n"
            summary += f"Position: {rec.get('Latitude', '')}, {rec.get('Longitude', '')}\n"
            summary += f"Category: {rec.get('Category', 'Unclassified')}\n"
            summary += f"Date: {rec.get('Date', 'N/A')}\n"
            summary += "-"*80 + "\n"
        
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"Angola_Summary_{datetime.now().strftime('%Y%m%d')}.txt"
        )
        
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(summary)
                messagebox.showinfo("Success", f"Exported Angola summary to:\n{path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed:\n{e}")


# --- MAIN ---
if __name__ == "__main__":
    root = tk.Tk()
    app = MarineExtractorGUI(root)
    root.mainloop()