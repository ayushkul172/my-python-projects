import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder, StandardScaler, LabelEncoder
from sklearn.ensemble import RandomForestClassifier, IsolationForest
from sklearn.model_selection import cross_val_score, train_test_split
from sklearn.cluster import DBSCAN
from sklearn.feature_extraction.text import TfidfVectorizer
import re
from datetime import datetime, timedelta
import json
import os
import threading
import time
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

CLOSED_STATUSES = {"Executed", "Signed", "Completed", "Closed", "Cancelled"}
PENDING_KEYWORDS = ["pending", "review", "waiting", "approval", "signature", "negotiation", "discussion", "follow-up"]
HIGH_PRIORITY_KEYWORDS = ["urgent", "critical", "deadline", "penalty", "escalate", "asap", "immediate"]

class AdvancedContractAI:
    def __init__(self):
        self.model = None
        self.anomaly_detector = None
        self.text_vectorizer = None
        self.target_col = "DA Comment / Action"
        self.model_accuracy = None
        self.pending_contracts_cache = []
        self.alert_threshold = 0.7
        self.feature_columns = None  # Store feature columns for consistency
        
    def _is_open(self, status: str) -> bool:
        """Check if contract status indicates it's still open."""
        return status not in CLOSED_STATUSES
    
    def _is_pending_critical(self, status: str, comment: str = "") -> tuple:
        """Enhanced pending detection with criticality assessment."""
        status_lower = status.lower()
        comment_lower = comment.lower()
        
        # Check for pending status
        is_pending = any(keyword in status_lower for keyword in PENDING_KEYWORDS)
        
        # Determine criticality level
        criticality = "LOW"
        if any(keyword in status_lower or keyword in comment_lower for keyword in HIGH_PRIORITY_KEYWORDS):
            criticality = "CRITICAL"
        elif any(keyword in status_lower for keyword in ["review", "approval", "signature"]):
            criticality = "HIGH"
        elif any(keyword in status_lower for keyword in ["negotiation", "discussion"]):
            criticality = "MEDIUM"
        
        return is_pending, criticality
    
    def _extract_text_features(self, texts):
        """Extract advanced text features using TF-IDF."""
        if self.text_vectorizer is None:
            self.text_vectorizer = TfidfVectorizer(
                max_features=100,
                stop_words='english',
                ngram_range=(1, 2),
                min_df=2
            )
            return self.text_vectorizer.fit_transform(texts)
        else:
            return self.text_vectorizer.transform(texts)
    
    def _feature_engineer_advanced(self, df: pd.DataFrame) -> pd.DataFrame:
        """Create comprehensive features with advanced AI techniques."""
        df = df.copy()
        
        # Standardize date column name - use the original column name consistently
        date_col = "Date (DD/MM/YYYY)"
        
        # Date processing with better error handling
        if date_col in df.columns:
            df["Date_parsed"] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
        else:
            df["Date_parsed"] = pd.NaT  # Not a Time - pandas null for datetime
        
        ref_date = df["Date_parsed"].min() if not df["Date_parsed"].isna().all() else datetime.now()
        current_date = datetime.now()
        
        # Advanced time-based features
        df["days_since_start"] = (df["Date_parsed"] - ref_date).dt.days.fillna(0)
        df["days_from_today"] = (current_date - df["Date_parsed"]).dt.days.fillna(0)
        df["is_weekend_start"] = df["Date_parsed"].dt.weekday >= 5
        df["month"] = df["Date_parsed"].dt.month.fillna(1)
        df["quarter"] = df["Date_parsed"].dt.quarter.fillna(1)
        df["day_of_week"] = df["Date_parsed"].dt.dayofweek.fillna(0)
        
        # Seasonal analysis
        df["is_year_end"] = df["month"].isin([11, 12])
        df["is_quarter_end"] = df["month"].isin([3, 6, 9, 12])
        
        # Advanced text analysis features
        comment = df[self.target_col].fillna("").str.lower()
        df["comment_length"] = comment.str.len()
        df["word_count"] = comment.str.split().str.len()
        
        # Enhanced keyword detection
        df["has_urgent"] = comment.str.contains(r"urgent|deadline|penalty|asap|critical|emergency", case=False, na=False)
        df["has_negotiation"] = comment.str.contains(r"negotiat|amend|clause|review|discuss|revise", case=False, na=False)
        df["has_legal"] = comment.str.contains(r"legal|compliance|regulation|audit|lawyer|attorney", case=False, na=False)
        df["has_financial"] = comment.str.contains(r"payment|invoice|cost|budget|financial|money|price", case=False, na=False)
        df["has_pending"] = comment.str.contains('|'.join(PENDING_KEYWORDS), case=False, na=False)
        df["has_delay"] = comment.str.contains(r"delay|postpone|extend|late|overdue", case=False, na=False)
        df["has_approval"] = comment.str.contains(r"approval|approve|authorize|sign|signature", case=False, na=False)
        
        # Sentiment analysis (simple rule-based)
        positive_words = ["complete", "approve", "success", "finalize", "ready", "agree"]
        negative_words = ["delay", "issue", "problem", "reject", "deny", "cancel", "dispute"]
        
        df["positive_sentiment"] = comment.str.contains('|'.join(positive_words), case=False, na=False)
        df["negative_sentiment"] = comment.str.contains('|'.join(negative_words), case=False, na=False)
        
        # Priority mapping with better handling
        priority_map = {"High": 3, "Medium": 2, "Low": 1}
        df["priority_num"] = df["Priority"].map(priority_map).fillna(2)
        
        # Contract metadata features
        df["contractor_length"] = df["Contractor"].fillna("").str.len()
        df["project_length"] = df["Field / Project"].fillna("").str.len()
        df["title_length"] = df["Article Title"].fillna("").str.len()
        
        # Contract type analysis
        title = df["Article Title"].fillna("").str.lower()
        df["has_amendment"] = title.str.contains(r"amendment|addendum|modification", case=False, na=False)
        df["has_extension"] = title.str.contains(r"extension|renewal|extend", case=False, na=False)
        df["has_service"] = title.str.contains(r"service|maintenance|support", case=False, na=False)
        df["has_supply"] = title.str.contains(r"supply|procurement|purchase", case=False, na=False)
        
        # Risk scoring
        risk_score = 0
        risk_score += df["days_from_today"] * 0.1  # Age factor
        risk_score += df["priority_num"] * 10  # Priority factor
        risk_score += df["has_urgent"].astype(int) * 20  # Urgency factor
        risk_score += df["has_delay"].astype(int) * 15  # Delay factor
        risk_score += df["negative_sentiment"].astype(int) * 10  # Sentiment factor
        
        df["risk_score"] = risk_score
        
        return df
    
    def train_advanced_model(self, records: list) -> tuple:
        """Train advanced ML models including anomaly detection."""
        if len(records) < 20:
            return False, "Need at least 20 records for advanced AI training."
        
        try:
            df = pd.DataFrame(records)
            if self.target_col not in df.columns:
                return False, f"Target column '{self.target_col}' not found."
            
            df = self._feature_engineer_advanced(df)
            
            # Focus on open contracts
            open_mask = df[self.target_col].apply(self._is_open)
            df_open = df[open_mask]
            
            if df_open.empty:
                return False, "No open contracts for training."
            
            if df_open[self.target_col].nunique() < 2:
                return False, "Need diverse status types for training."
            
            # Enhanced feature selection - store for consistency
            self.feature_columns = [
                "days_since_start", "days_from_today", "priority_num", "risk_score",
                "has_urgent", "has_negotiation", "has_legal", "has_financial", 
                "has_pending", "has_delay", "has_approval",
                "comment_length", "word_count", "contractor_length", "project_length",
                "title_length", "has_amendment", "has_extension", "has_service", "has_supply",
                "is_weekend_start", "month", "quarter", "day_of_week",
                "is_year_end", "is_quarter_end", "positive_sentiment", "negative_sentiment",
                "Country", "Field / Project"
            ]
            
            # Ensure all feature columns exist
            for col in self.feature_columns:
                if col not in df_open.columns:
                    if col in ["Country", "Field / Project"]:
                        df_open[col] = df_open[col].fillna("Unknown") if col in df_open.columns else "Unknown"
                    else:
                        df_open[col] = 0
            
            X = df_open[self.feature_columns]
            y = df_open[self.target_col]
            
            # Advanced preprocessing pipeline
            cat_cols = ["Country", "Field / Project"]
            num_cols = [col for col in self.feature_columns if col not in cat_cols]
            
            preprocessor = ColumnTransformer(
                transformers=[
                    ("cat", OneHotEncoder(handle_unknown="ignore", sparse_output=False), cat_cols),
                    ("num", StandardScaler(), num_cols)
                ]
            )
            
            # Enhanced classifier with better parameters
            classifier = RandomForestClassifier(
                n_estimators=300,
                max_depth=15,
                min_samples_split=3,
                min_samples_leaf=1,
                class_weight="balanced",
                random_state=42,
                n_jobs=-1
            )
            
            self.model = Pipeline([
                ("preprocessor", preprocessor),
                ("classifier", classifier)
            ])
            
            # Train main model
            self.model.fit(X, y)
            
            # Train anomaly detector for outlier detection
            X_processed = preprocessor.fit_transform(X)
            self.anomaly_detector = IsolationForest(
                contamination=0.1,
                random_state=42,
                n_jobs=-1
            )
            self.anomaly_detector.fit(X_processed)
            
            # Cross-validation for accuracy
            try:
                cv_scores = cross_val_score(self.model, X, y, cv=min(5, len(X)), scoring='accuracy')
                self.model_accuracy = cv_scores.mean()
            except:
                self.model_accuracy = None
            
            return True, f"Advanced AI model trained on {len(df_open)} contracts with {self.model_accuracy:.1%} accuracy."
            
        except Exception as e:
            return False, f"Advanced training failed: {str(e)}"
    
    def predict_status_advanced(self, contract: dict) -> dict:
        """Advanced prediction with confidence and anomaly detection."""
        if self.model is None or self.feature_columns is None:
            return {"prediction": None, "confidence": 0.0, "is_anomaly": False, "risk_level": "UNKNOWN"}
        
        try:
            df = pd.DataFrame([contract])
            df = self._feature_engineer_advanced(df)
            
            # Ensure all columns exist with the stored feature columns
            for col in self.feature_columns:
                if col not in df.columns:
                    if col in ["Country", "Field / Project"]:
                        df[col] = "Unknown"
                    else:
                        df[col] = 0
            
            X = df[self.feature_columns]
            
            # Main prediction
            prediction = self.model.predict(X)[0]
            probabilities = self.model.predict_proba(X)[0]
            confidence = max(probabilities)
            
            # Anomaly detection
            is_anomaly = False
            if self.anomaly_detector:
                try:
                    X_processed = self.model.named_steps['preprocessor'].transform(X)
                    anomaly_score = self.anomaly_detector.decision_function(X_processed)[0]
                    is_anomaly = self.anomaly_detector.predict(X_processed)[0] == -1
                except:
                    pass
            
            # Risk assessment
            risk_score = df["risk_score"].iloc[0]
            if risk_score > 50:
                risk_level = "CRITICAL"
            elif risk_score > 30:
                risk_level = "HIGH"
            elif risk_score > 15:
                risk_level = "MEDIUM"
            else:
                risk_level = "LOW"
            
            return {
                "prediction": prediction if confidence >= 0.3 else None,
                "confidence": confidence,
                "is_anomaly": is_anomaly,
                "risk_level": risk_level,
                "risk_score": risk_score
            }
            
        except Exception as e:
            print(f"Advanced prediction error: {e}")
            return {"prediction": None, "confidence": 0.0, "is_anomaly": False, "risk_level": "UNKNOWN"}
    
    def analyze_pending_contracts(self, records: list) -> dict:
        """Advanced analysis of pending contracts with AI insights."""
        analysis = {
            "total_pending": 0,
            "critical_pending": [],
            "high_risk_pending": [],
            "overdue_pending": [],
            "anomalous_contracts": [],
            "insights": []
        }
        
        try:
            for contract in records:
                status = contract.get(self.target_col, "")
                comment = contract.get("Article Title", "")
                
                # Check if pending and get criticality
                is_pending, criticality = self._is_pending_critical(status, comment)
                
                if is_pending:
                    analysis["total_pending"] += 1
                    
                    # Get AI prediction
                    ai_result = self.predict_status_advanced(contract)
                    
                    # Categorize by criticality and AI insights
                    contract_with_ai = contract.copy()
                    contract_with_ai["ai_prediction"] = ai_result["prediction"]
                    contract_with_ai["ai_confidence"] = ai_result["confidence"]
                    contract_with_ai["risk_level"] = ai_result["risk_level"]
                    contract_with_ai["criticality"] = criticality
                    
                    if criticality == "CRITICAL" or ai_result["risk_level"] == "CRITICAL":
                        analysis["critical_pending"].append(contract_with_ai)
                    elif ai_result["risk_level"] in ["HIGH", "MEDIUM"]:
                        analysis["high_risk_pending"].append(contract_with_ai)
                    
                    # Check for overdue contracts - use the correct column name
                    try:
                        date_str = contract.get("Date (DD/MM/YYYY)", "")
                        if date_str:
                            contract_date = datetime.strptime(date_str, "%d/%m/%Y")
                            days_old = (datetime.now() - contract_date).days
                            if days_old > 30:
                                analysis["overdue_pending"].append((contract_with_ai, days_old))
                    except:
                        pass
                    
                    # Check for anomalous contracts
                    if ai_result["is_anomaly"]:
                        analysis["anomalous_contracts"].append(contract_with_ai)
            
            # Generate AI insights
            if analysis["critical_pending"]:
                analysis["insights"].append(f"ALERT: {len(analysis['critical_pending'])} critical pending contracts require immediate attention!")
            
            if analysis["overdue_pending"]:
                avg_days = sum(days for _, days in analysis["overdue_pending"]) / len(analysis["overdue_pending"])
                analysis["insights"].append(f"{len(analysis['overdue_pending'])} contracts are overdue (avg: {avg_days:.1f} days)")
            
            if analysis["anomalous_contracts"]:
                analysis["insights"].append(f"{len(analysis['anomalous_contracts'])} contracts show unusual patterns")
            
            # Performance insights
            if analysis["total_pending"] > 0:
                critical_rate = len(analysis["critical_pending"]) / analysis["total_pending"] * 100
                if critical_rate > 20:
                    analysis["insights"].append(f"High critical rate: {critical_rate:.1f}% of pending contracts are critical")
        
        except Exception as e:
            analysis["insights"].append(f"Analysis error: {str(e)}")
        
        return analysis
    
    def get_advanced_recommendations(self, contract: dict) -> list:
        """AI-powered advanced recommendations."""
        recommendations = []
        
        try:
            # Get AI analysis
            ai_result = self.predict_status_advanced(contract)
            status = contract.get(self.target_col, "")
            
            # Risk-based recommendations
            if ai_result["risk_level"] == "CRITICAL":
                recommendations.append("CRITICAL: Assign C-level executive oversight immediately")
                recommendations.append("Schedule emergency stakeholder meeting within 24 hours")
            elif ai_result["risk_level"] == "HIGH":
                recommendations.append("HIGH RISK: Escalate to department head")
                recommendations.append("Daily monitoring and reporting required")
            
            # AI prediction-based recommendations
            if ai_result["prediction"]:
                recommendations.append(f"AI predicts next status: '{ai_result['prediction']}' (confidence: {ai_result['confidence']:.1%})")
                
                if "signature" in ai_result["prediction"].lower():
                    recommendations.append("Prepare signature package and schedule signing ceremony")
                elif "review" in ai_result["prediction"].lower():
                    recommendations.append("Allocate legal review resources and set timeline")
            
            # Anomaly-based recommendations
            if ai_result["is_anomaly"]:
                recommendations.append("ANOMALY DETECTED: This contract shows unusual patterns - conduct thorough review")
            
            # Time-based recommendations
            try:
                date_str = contract.get("Date (DD/MM/YYYY)", "")
                if date_str:
                    contract_date = datetime.strptime(date_str, "%d/%m/%Y")
                    days_old = (datetime.now() - contract_date).days
                    
                    if days_old > 90:
                        recommendations.append("URGENT: Contract is 90+ days old - immediate resolution required")
                    elif days_old > 60:
                        recommendations.append("Contract aging: 60+ days - prioritize resolution")
            except:
                pass
            
            # Content-based AI recommendations
            status_lower = status.lower()
            if "pending" in status_lower:
                if "approval" in status_lower:
                    recommendations.append("Send approval reminder with deadline to decision makers")
                elif "signature" in status_lower:
                    recommendations.append("Follow up on signature collection with automated reminders")
                elif "review" in status_lower:
                    recommendations.append("Request review status update and expected completion timeline")
            
        except Exception as e:
            recommendations.append(f"AI recommendation error: {str(e)}")
        
        return recommendations


class EnhancedContractTrackerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced AI Contract Tracker - Real-time Analytics")
        self.root.geometry("1600x1000")
        
        # Modern dark theme setup
        self.setup_modern_theme()
        
        # Center window
        self.center_window()
        
        self.data = []
        self.ai = AdvancedContractAI()
        self.current_selection = None
        self.alert_system_active = True
        self.pending_alerts = []
        
        # Define standard columns that will be used consistently
        self.standard_columns = [
            "Date (DD/MM/YYYY)", "Link", "Article Title", "Priority", 
            "Contractor", "Field / Project", "Country", "DA Comment / Action"
        ]
        
        # Auto-save configuration
        self.auto_save_file = "advanced_contract_data.json"
        
        self.create_modern_widgets()
        self.load_auto_save()
        
        # Start background AI monitoring
        self.start_ai_monitoring()
        
        # Auto-save timer
        self.root.after(300000, self.auto_save)
        
        # Bind close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        print("Advanced AI Contract Tracker started!")
    
    def setup_modern_theme(self):
        """Setup modern dark theme with blue accents."""
        self.colors = {
            'bg_primary': '#1a1a2e',        # Dark navy
            'bg_secondary': '#16213e',       # Darker blue
            'bg_card': '#0f3460',           # Deep blue
            'accent_blue': '#00d4ff',       # Bright blue
            'accent_blue_dark': '#0099cc',  # Dark blue
            'text_primary': '#ffffff',      # White
            'text_secondary': '#b0b0b0',    # Light gray
            'success': '#00ff88',           # Green
            'warning': '#ffaa00',           # Orange
            'danger': '#ff4444',            # Red
            'critical': '#ff0066'           # Pink/Red
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['bg_primary'])
        
        # Setup ttk styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure modern styles
        self.style.configure('Modern.TNotebook', 
                           background=self.colors['bg_primary'],
                           borderwidth=0)
        
        self.style.configure('Modern.TNotebook.Tab',
                           background=self.colors['bg_secondary'],
                           foreground=self.colors['text_primary'],
                           padding=[20, 10],
                           focuscolor='none')
        
        self.style.map('Modern.TNotebook.Tab',
                      background=[('selected', self.colors['accent_blue_dark']),
                                ('active', self.colors['bg_card'])])
        
        self.style.configure('Modern.TFrame',
                           background=self.colors['bg_primary'],
                           borderwidth=1,
                           relief='flat')
        
        self.style.configure('Modern.TLabelFrame',
                           background=self.colors['bg_primary'],
                           foreground=self.colors['text_primary'],
                           borderwidth=1,
                           relief='flat')
        
        self.style.configure('Modern.TLabelFrame.Label',
                           background=self.colors['bg_primary'],
                           foreground=self.colors['accent_blue'],
                           font=('Arial', 11, 'bold'))
    
    def center_window(self):
        """Center the window on screen."""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (1600 // 2)
        y = (self.root.winfo_screenheight() // 2) - (1000 // 2)
        self.root.geometry(f"1600x1000+{x}+{y}")
        
        # Force window to front
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(100, lambda: self.root.attributes('-topmost', False))
    
    def create_modern_widgets(self):
        """Create modern UI with advanced features."""
        # Main container with modern styling
        main_container = tk.Frame(self.root, bg=self.colors['bg_primary'])
        main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Modern notebook
        self.notebook = ttk.Notebook(main_container, style='Modern.TNotebook')
        self.notebook.pack(fill="both", expand=True)
        
        # Tabs
        self.main_tab = tk.Frame(self.notebook, bg=self.colors['bg_primary'])
        self.ai_tab = tk.Frame(self.notebook, bg=self.colors['bg_primary'])
        self.alerts_tab = tk.Frame(self.notebook, bg=self.colors['bg_primary'])
        
        self.notebook.add(self.main_tab, text="Contract Management")
        self.notebook.add(self.ai_tab, text="AI Analytics")
        self.notebook.add(self.alerts_tab, text="Smart Alerts")
        
        self.create_main_tab_modern()
        self.create_ai_tab_modern()
        self.create_alerts_tab()
    
    def create_main_tab_modern(self):
        """Create modern main tab with enhanced UI."""
        # Top section - Form
        form_frame = self.create_modern_frame(self.main_tab, "Contract Information")
        form_frame.pack(fill="x", padx=10, pady=10)
        
        # Middle section - Table
        table_frame = self.create_modern_frame(self.main_tab, "Contract Portfolio")
        table_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Bottom section - Status and controls
        control_frame = tk.Frame(self.main_tab, bg=self.colors['bg_primary'])
        control_frame.pack(fill="x", padx=10, pady=10)
        
        self.create_modern_form(form_frame)
        self.create_modern_table(table_frame)
        self.create_modern_controls(control_frame)
    
    def create_modern_frame(self, parent, title):
        """Create a modern styled frame."""
        frame = tk.Frame(parent, bg=self.colors['bg_card'], relief='flat', bd=2)
        
        # Title bar
        title_frame = tk.Frame(frame, bg=self.colors['accent_blue_dark'], height=40)
        title_frame.pack(fill="x")
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(title_frame, text=title, 
                              bg=self.colors['accent_blue_dark'],
                              fg=self.colors['text_primary'],
                              font=('Arial', 12, 'bold'))
        title_label.pack(side="left", padx=15, pady=8)
        
        # Content frame
        content_frame = tk.Frame(frame, bg=self.colors['bg_card'])
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        return content_frame
    
    def create_modern_form(self, parent):
        """Create modern form with enhanced styling."""
        self.entries = {}
        
        # Create form in grid layout using standard columns
        for i, label_text in enumerate(self.standard_columns):
            row = i % 4
            col = i // 4
            
            # Label
            label = tk.Label(parent, text=label_text + ":", 
                           bg=self.colors['bg_card'],
                           fg=self.colors['text_secondary'],
                           font=('Arial', 10))
            label.grid(row=row, column=col*2, padx=10, pady=8, sticky="w")
            
            # Entry/Combobox
            if label_text == "Priority":
                entry = ttk.Combobox(parent, values=["High", "Medium", "Low"], 
                                   state="readonly", font=('Arial', 10))
            elif label_text == "DA Comment / Action":
                values = [
                    "Pending Legal Review", "Pending Approval", "Pending Signature",
                    "Under Negotiation", "Awaiting Information", "In Review",
                    "Executed", "Signed", "Completed", "Closed", "Cancelled"
                ]
                entry = ttk.Combobox(parent, values=values, font=('Arial', 10))
            else:
                entry = tk.Entry(parent, width=25, font=('Arial', 10),
                               bg=self.colors['bg_secondary'],
                               fg=self.colors['text_primary'],
                               insertbackground=self.colors['accent_blue'],
                               relief='flat', bd=5)
            
            entry.grid(row=row, column=col*2 + 1, padx=10, pady=8, sticky="ew")
            self.entries[label_text] = entry
        
        # Configure grid weights
        for i in range(4):
            parent.columnconfigure(i*2 + 1, weight=1)
    
    def create_modern_table(self, parent):
        """Create modern table with enhanced styling."""
        # Create treeview with modern styling using standard columns
        self.tree = ttk.Treeview(parent, columns=self.standard_columns, show='headings', height=15)
        
        # Configure columns
        for col in self.standard_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor="center")
        
        # Modern color tags
        self.tree.tag_configure("closed", background=self.colors['bg_secondary'], 
                              foreground=self.colors['text_secondary'])
        self.tree.tag_configure("pending_critical", background=self.colors['critical'], 
                              foreground=self.colors['text_primary'])
        
        # Scrollbars with modern styling
        v_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(parent, orient="horizontal", command=self.tree.xview)
        
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        
        # Bind events
        self.tree.bind("<Double-1>", self.on_data_tree_double_click)  # <-- FIXED
        self.tree.bind("<ButtonRelease-1>", self.on_single_click)
    
    def create_modern_controls(self, parent):
        """Create modern control buttons and status bar."""
        # Button frame
        button_frame = tk.Frame(parent, bg=self.colors['bg_primary'])
        button_frame.pack(fill="x", pady=10)
        
        # Create modern buttons
        buttons = [
            ("Add Contract", self.add_contract, self.colors['success']),
            ("Update Selected", self.update_selected_contract, self.colors['accent_blue']),
            ("Delete Selected", self.delete_selected_contract, self.colors['danger']),
            ("Clear Form", self.clear_form, self.colors['bg_secondary']),
            ("Save Data", self.save_data, self.colors['accent_blue_dark']),
            ("Load Data", self.load_data, self.colors['accent_blue_dark']),
            ("AI Analysis", self.ai_analyze_selected, self.colors['warning']),
            ("Bulk AI Scan", self.bulk_ai_analysis, self.colors['critical'])
        ]
        
        for i, (text, command, color) in enumerate(buttons):
            btn = tk.Button(button_frame, text=text, command=command,
                           bg=color, fg=self.colors['text_primary'],
                           font=('Arial', 10, 'bold'), relief='flat',
                           padx=15, pady=8, cursor='hand2')
            
            # Hover effects
            def on_enter(e, btn=btn, original_color=color):
                btn.configure(bg=self.lighten_color(original_color))
            
            def on_leave(e, btn=btn, original_color=color):
                btn.configure(bg=original_color)
            
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            
            btn.grid(row=i//4, column=i%4, padx=8, pady=4, sticky="ew")
        
        # Configure button grid
        for i in range(4):
            button_frame.columnconfigure(i, weight=1)
        
        # Status bar
        status_frame = tk.Frame(parent, bg=self.colors['bg_card'], height=40)
        status_frame.pack(fill="x", pady=(10, 0))
        status_frame.pack_propagate(False)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Advanced AI Contract Tracker - Ready for intelligent analysis")
        
        status_label = tk.Label(status_frame, textvariable=self.status_var,
                               bg=self.colors['bg_card'], fg=self.colors['text_secondary'],
                               font=('Arial', 10))
        status_label.pack(side="left", padx=15, pady=10)
        
        # Contract stats
        self.count_var = tk.StringVar()
        self.count_var.set("Contracts: 0 | Pending: 0 | Critical: 0")
        
        count_label = tk.Label(status_frame, textvariable=self.count_var,
                              bg=self.colors['bg_card'], fg=self.colors['accent_blue'],
                              font=('Arial', 10, 'bold'))
        count_label.pack(side="right", padx=15, pady=10)
    
    def create_ai_tab_modern(self):
        """Create modern AI analytics tab."""
        # AI Model Status
        model_frame = self.create_modern_frame(self.ai_tab, "AI Model Status & Performance")
        model_frame.pack(fill="x", padx=10, pady=10)
        
        self.ai_info_text = tk.Text(model_frame, height=4, wrap=tk.WORD,
                                   bg=self.colors['bg_secondary'],
                                   fg=self.colors['text_primary'],
                                   font=('Consolas', 10),
                                   relief='flat', bd=0)
        self.ai_info_text.pack(fill="x", pady=5)
        
        # Real-time Analytics
        analytics_frame = self.create_modern_frame(self.ai_tab, "Real-time Contract Analytics")
        analytics_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.analytics_text = tk.Text(analytics_frame, wrap=tk.WORD,
                                     bg=self.colors['bg_secondary'],
                                     fg=self.colors['text_primary'],
                                     font=('Consolas', 10),
                                     relief='flat', bd=0)
        analytics_scrollbar = ttk.Scrollbar(analytics_frame, orient="vertical", 
                                           command=self.analytics_text.yview)
        self.analytics_text.configure(yscrollcommand=analytics_scrollbar.set)
        
        self.analytics_text.pack(side="left", fill="both", expand=True)
        analytics_scrollbar.pack(side="right", fill="y")
        
        # AI Control Panel
        ai_control_frame = tk.Frame(self.ai_tab, bg=self.colors['bg_primary'])
        ai_control_frame.pack(fill="x", padx=10, pady=10)
        
        ai_buttons = [
            ("Train AI Model", self.train_advanced_ai, self.colors['accent_blue']),
            ("Run Deep Analysis", self.run_deep_analysis, self.colors['warning']),
            ("Generate AI Report", self.generate_ai_report, self.colors['success']),
            ("Export Analytics", self.export_ai_analytics, self.colors['accent_blue_dark'])
        ]
        
        for i, (text, command, color) in enumerate(ai_buttons):
            btn = tk.Button(ai_control_frame, text=text, command=command,
                           bg=color, fg=self.colors['text_primary'],
                           font=('Arial', 11, 'bold'), relief='flat',
                           padx=20, pady=10, cursor='hand2')
            
            # Hover effect
            def on_enter(e, btn=btn, color=color):
                btn.configure(bg=self.lighten_color(color))
            def on_leave(e, btn=btn, color=color):
                btn.configure(bg=color)
            
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            
            btn.pack(side="left", padx=15, pady=5)
    
    def create_alerts_tab(self):
        """Create smart alerts tab with real-time monitoring."""
        # Alert Controls
        control_frame = self.create_modern_frame(self.alerts_tab, "Smart Alert System Controls")
        control_frame.pack(fill="x", padx=10, pady=10)
        
        # Toggle button for alert system
        self.alert_toggle_var = tk.BooleanVar(value=True)
        toggle_btn = tk.Checkbutton(control_frame, text="Enable Real-time AI Monitoring",
                                   variable=self.alert_toggle_var,
                                   command=self.toggle_alert_system,
                                   bg=self.colors['bg_card'],
                                   fg=self.colors['text_primary'],
                                   selectcolor=self.colors['accent_blue'],
                                   font=('Arial', 12, 'bold'))
        toggle_btn.pack(side="left", padx=10, pady=10)
        
        # Alert threshold
        threshold_frame = tk.Frame(control_frame, bg=self.colors['bg_card'])
        threshold_frame.pack(side="right", padx=10, pady=10)
        
        tk.Label(threshold_frame, text="Alert Sensitivity:",
                bg=self.colors['bg_card'], fg=self.colors['text_secondary'],
                font=('Arial', 10)).pack(side="left", padx=5)
        
        self.threshold_var = tk.StringVar(value="High")
        threshold_combo = ttk.Combobox(threshold_frame, textvariable=self.threshold_var,
                                      values=["Low", "Medium", "High", "Critical"],
                                      state="readonly", width=10)
        threshold_combo.pack(side="left", padx=5)
        threshold_combo.bind("<<ComboboxSelected>>", self.update_alert_threshold)
        
        # Active Alerts Display
        alerts_frame = self.create_modern_frame(self.alerts_tab, "Active Smart Alerts")
        alerts_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.alerts_text = tk.Text(alerts_frame, wrap=tk.WORD,
                                  bg=self.colors['bg_secondary'],
                                  fg=self.colors['text_primary'],
                                  font=('Consolas', 10),
                                  relief='flat', bd=0)
        alerts_scrollbar = ttk.Scrollbar(alerts_frame, orient="vertical",
                                        command=self.alerts_text.yview)
        self.alerts_text.configure(yscrollcommand=alerts_scrollbar.set)
        
        self.alerts_text.pack(side="left", fill="both", expand=True)
        alerts_scrollbar.pack(side="right", fill="y")
        
        # Alert action buttons
        action_frame = tk.Frame(self.alerts_tab, bg=self.colors['bg_primary'])
        action_frame.pack(fill="x", padx=10, pady=10)
        
        alert_buttons = [
            ("Refresh Alerts", self.refresh_alerts, self.colors['accent_blue']),
            ("Clear All Alerts", self.clear_alerts, self.colors['warning']),
            ("Export Alert Log", self.export_alerts, self.colors['accent_blue_dark'])
        ]
        
        for text, command, color in alert_buttons:
            btn = tk.Button(action_frame, text=text, command=command,
                           bg=color, fg=self.colors['text_primary'],
                           font=('Arial', 10, 'bold'), relief='flat',
                           padx=15, pady=8, cursor='hand2')
            btn.pack(side="left", padx=10)
    
    def lighten_color(self, color):
        """Lighten a hex color for hover effects."""
        if color.startswith('#'):
            color = color[1:]
        rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
        rgb = tuple(min(255, c + 30) for c in rgb)
        return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
    
    def normalize_contract_data(self, contract_data):
        """Ensure contract data has all required columns in the correct order."""
        normalized = {}
        for col in self.standard_columns:
            normalized[col] = contract_data.get(col, "")
        # Fill missing columns with empty string
        for key in contract_data:
            if key not in normalized:
                normalized[key] = contract_data[key]
        return normalized
    
    def start_ai_monitoring(self):
        """Start background AI monitoring thread."""
        def monitor():
            while self.alert_system_active:
                if self.data:
                    self.update_ai_alerts()
                time.sleep(30)  # Check every 30 seconds
        
        monitoring_thread = threading.Thread(target=monitor, daemon=True)
        monitoring_thread.start()
    
    def update_ai_alerts(self):
        """Update AI alerts based on current data."""
        try:
            if not self.data:
                return
            
            analysis = self.ai.analyze_pending_contracts(self.data)
            
            # Update alerts
            self.pending_alerts.clear()
            
            # Critical alerts
            for contract in analysis["critical_pending"]:
                alert = {
                    "type": "CRITICAL",
                    "title": contract.get("Article Title", "Unknown"),
                    "message": f"Critical pending contract requires immediate attention",
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "risk_level": contract.get("risk_level", "HIGH")
                }
                self.pending_alerts.append(alert)
            
            # Overdue alerts
            for contract, days in analysis["overdue_pending"]:
                alert = {
                    "type": "OVERDUE",
                    "title": contract.get("Article Title", "Unknown"),
                    "message": f"Contract overdue by {days} days",
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "days": days
                }
                self.pending_alerts.append(alert)
            
            # Anomaly alerts
            for contract in analysis["anomalous_contracts"]:
                alert = {
                    "type": "ANOMALY",
                    "title": contract.get("Article Title", "Unknown"),
                    "message": "Unusual contract pattern detected by AI",
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "confidence": contract.get("ai_confidence", 0)
                }
                self.pending_alerts.append(alert)
            
            # Update alerts display
            self.root.after(0, self.display_alerts)
            
        except Exception as e:
            print(f"Alert update error: {e}")
    
    def display_alerts(self):
        """Display current alerts in the alerts tab."""
        self.alerts_text.delete("1.0", "end")
        
        if not self.pending_alerts:
            self.alerts_text.insert("1.0", "No active alerts. All contracts are within normal parameters.")
            return
        
        alert_text = []
        alert_text.append("ACTIVE SMART ALERTS")
        alert_text.append("=" * 50)
        alert_text.append(f"Last Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        alert_text.append(f"Total Alerts: {len(self.pending_alerts)}")
        alert_text.append("")
        
        # Group alerts by type
        critical_alerts = [a for a in self.pending_alerts if a["type"] == "CRITICAL"]
        overdue_alerts = [a for a in self.pending_alerts if a["type"] == "OVERDUE"]
        anomaly_alerts = [a for a in self.pending_alerts if a["type"] == "ANOMALY"]
        
        if critical_alerts:
            alert_text.append("CRITICAL ALERTS:")
            alert_text.append("-" * 30)
            for alert in critical_alerts:
                alert_text.append(f"[CRITICAL] {alert['title']}")
                alert_text.append(f"   {alert['message']}")
                alert_text.append(f"   Risk Level: {alert['risk_level']}")
                alert_text.append(f"   Time: {alert['timestamp']}")
                alert_text.append("")
        
        if overdue_alerts:
            alert_text.append("OVERDUE ALERTS:")
            alert_text.append("-" * 30)
            for alert in sorted(overdue_alerts, key=lambda x: x['days'], reverse=True):
                alert_text.append(f"[OVERDUE] {alert['title']}")
                alert_text.append(f"   {alert['message']}")
                alert_text.append(f"   Time: {alert['timestamp']}")
                alert_text.append("")
        
        if anomaly_alerts:
            alert_text.append("ANOMALY ALERTS:")
            alert_text.append("-" * 30)
            for alert in anomaly_alerts:
                alert_text.append(f"[ANOMALY] {alert['title']}")
                alert_text.append(f"   {alert['message']}")
                alert_text.append(f"   Confidence: {alert['confidence']:.1%}")
                alert_text.append(f"   Time: {alert['timestamp']}")
                alert_text.append("")
        
        self.alerts_text.insert("1.0", "\n".join(alert_text))
    
    # Contract Management Methods
    def add_contract(self):
        """Add new contract with AI validation."""
        new_contract = {}
        for text, entry in self.entries.items():
            if hasattr(entry, 'get'):
                value = entry.get().strip()
            else:
                value = str(entry).strip()
            new_contract[text] = value
        
        # Enhanced validation
        if not new_contract.get("Date (DD/MM/YYYY)"):
            messagebox.showwarning("Validation Error", "Date is required!")
            return
        
        if not new_contract.get("Article Title"):
            messagebox.showwarning("Validation Error", "Article Title is required!")
            return
        
        # Normalize contract data
        new_contract = self.normalize_contract_data(new_contract)
        
        # AI validation
        if self.ai.model:
            ai_result = self.ai.predict_status_advanced(new_contract)
            if ai_result["risk_level"] == "CRITICAL":
                response = messagebox.askyesno("AI Risk Warning", 
                    f"AI has detected this contract as CRITICAL RISK.\n"
                    f"Risk Score: {ai_result['risk_score']:.1f}\n"
                    f"Do you want to proceed?")
                if not response:
                    return
        
        self.data.append(new_contract)
        self.update_treeview_advanced()
        self.clear_form()
        self.update_status("Contract added successfully with AI validation!")
        messagebox.showinfo("Success", "Contract added and analyzed by AI!")
    
    def update_selected_contract(self):
        """Update the selected contract."""
        if not self.current_selection:
            messagebox.showwarning("Selection Error", "Please select a contract to update.")
            return
        
        try:
            # Get the index of the selected item
            index = self.tree.index(self.current_selection)
            
            # Update contract data
            updated_contract = {}
            for text, entry in self.entries.items():
                if hasattr(entry, 'get'):
                    value = entry.get().strip()
                else:
                    value = str(entry).strip()
                updated_contract[text] = value
            
            # Normalize contract data
            updated_contract = self.normalize_contract_data(updated_contract)
            
            self.data[index] = updated_contract
            self.update_treeview_advanced()
            self.update_status("Contract updated successfully!")
            messagebox.showinfo("Success", "Contract updated!")
            
        except Exception as e:
            messagebox.showerror("Update Error", f"Failed to update contract: {str(e)}")
    
    def delete_selected_contract(self):
        """Delete the selected contract."""
        if not self.current_selection:
            messagebox.showwarning("Selection Error", "Please select a contract to delete.")
            return
        
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this contract?"):
            try:
                index = self.tree.index(self.current_selection)
                del self.data[index]
                self.update_treeview_advanced()
                self.clear_form()
                self.current_selection = None
                self.update_status("Contract deleted successfully!")
                messagebox.showinfo("Success", "Contract deleted!")
            except Exception as e:
                messagebox.showerror("Delete Error", f"Failed to delete contract: {str(e)}")
    
    def clear_form(self):
        """Clear all form fields."""
        for entry in self.entries.values():
            if hasattr(entry, 'delete'):
                entry.delete(0, tk.END)
            elif hasattr(entry, 'set'):
                entry.set('')
        self.current_selection = None
    
    def save_data(self):
        """Save contract data to file."""
        if not self.data:
            messagebox.showinfo("Info", "No contracts to save.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.json'):
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(self.data, f, indent=2)
                else:
                    df = pd.DataFrame(self.data)
                    df.to_excel(file_path, index=False)
                
                self.update_status(f"Data saved to {file_path}")
                messagebox.showinfo("Success", f"Data saved to {file_path}")
            except Exception as e:
                messagebox.showerror("Save Error", f"Failed to save data: {str(e)}")
    
    def load_data(self):
        """Load contract data from file with proper normalization."""
        file_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.json'):
                    with open(file_path, 'r', encoding='utf-8') as f:
                        loaded_data = json.load(f)
                else:
                    df = pd.read_excel(file_path)
                    loaded_data = df.to_dict('records')
                
                # Normalize all loaded data to ensure consistent column structure
                self.data = []
                for contract in loaded_data:
                    normalized_contract = self.normalize_contract_data(contract)
                    self.data.append(normalized_contract)
                
                self.update_treeview_advanced()
                self.root.update_idletasks()
                self.update_status(f"Data loaded from {file_path}")
                messagebox.showinfo("Success", f"Loaded {len(self.data)} contracts from {file_path}")
            except Exception as e:
                messagebox.showerror("Load Error", f"Failed to load data: {str(e)}")
        
        # After loading data
        print(f"Loaded data: {self.data}")
    
    def update_treeview_advanced(self):
        """Update treeview with advanced AI-based coloring."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add all contracts with AI-based risk coloring
        critical_count = 0
        pending_count = 0
        
        for contract in self.data:
            # Ensure contract is normalized
            contract = self.normalize_contract_data(contract)
            
            # Get values in the correct order
            values = [contract[col] for col in self.standard_columns]
            
            # Get AI risk assessment
            risk_tag = self.get_ai_risk_tag(contract)
            
            if risk_tag == "pending_critical":
                critical_count += 1
            
            status = contract.get("DA Comment / Action", "")
            if self.ai._is_open(status):
                pending_count += 1
            
            self.tree.insert("", "end", values=values, tags=(risk_tag,))
        
        # Update comprehensive stats
        total = len(self.data)
        closed = total - pending_count
        self.count_var.set(f"Total: {total} | Pending: {pending_count} | Critical: {critical_count} | Closed: {closed}")
        
        # Also update data viewer if it exists
        self.update_data_viewer()

    def update_data_viewer(self):
        """Update the data viewer tab with current data and statistics."""
        if not hasattr(self, 'data_tree'):
            return  # Data viewer not created yet
        
        # Update statistics
        self.update_data_statistics()
        
        # Update data table
        self.refresh_data_viewer()
    
    def update_data_statistics(self):
        """Update data statistics display."""
        if not hasattr(self, 'data_stats_labels') or not self.data:
            return
        
        try:
            # Calculate comprehensive statistics
            total_contracts = len(self.data)
            open_contracts = 0
            closed_contracts = 0
            critical_contracts = 0
            high_priority = 0
            overdue_contracts = 0
            pending_approval = 0
            countries = set()
            
            current_date = datetime.now()
            
            for contract in self.data:
                # Status analysis
                status = contract.get("DA Comment / Action", "")
                if self.ai._is_open(status):
                    open_contracts += 1
                    if "approval" in status.lower() or "pending" in status.lower():
                        pending_approval += 1
                else:
                    closed_contracts += 1
                
                # Priority analysis
                if contract.get("Priority", "") == "High":
                    high_priority += 1
                
                # Risk analysis
                risk_tag = self.get_ai_risk_tag(contract)
                if risk_tag in ["critical", "pending_critical"]:
                    critical_contracts += 1
                
                # Date analysis for overdue contracts
                try:
                    date_str = contract.get("Date (DD/MM/YYYY)", "")
                    if date_str:
                        contract_date = datetime.strptime(date_str, "%d/%m/%Y")
                        days_old = (current_date - contract_date).days
                        if days_old > 30 and self.ai._is_open(status):
                            overdue_contracts += 1
                except:
                    pass
                
                # Country analysis
                country = contract.get("Country", "").strip()
                if country:
                    countries.add(country)
            
            # Update labels
            self.data_stats_labels["total_contracts"].configure(text=str(total_contracts))
            self.data_stats_labels["open_contracts"].configure(text=str(open_contracts))
            self.data_stats_labels["closed_contracts"].configure(text=str(closed_contracts))
            self.data_stats_labels["critical_contracts"].configure(text=str(critical_contracts))
            self.data_stats_labels["high_priority"].configure(text=str(high_priority))
            self.data_stats_labels["overdue_contracts"].configure(text=str(overdue_contracts))
            self.data_stats_labels["pending_approval"].configure(text=str(pending_approval))
            self.data_stats_labels["countries_count"].configure(text=str(len(countries)))
            
            # Color code critical statistics
            if critical_contracts > 0:
                self.data_stats_labels["critical_contracts"].configure(fg=self.colors['danger'])
            if overdue_contracts > 0:
                self.data_stats_labels["overdue_contracts"].configure(fg=self.colors['warning'])
            
        except Exception as e:
            print(f"Error updating statistics: {e}")
    
    def refresh_data_viewer(self):
        """Refresh the data viewer table with current data."""
        if not hasattr(self, 'data_tree'):
            return
        
        # Clear existing items
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)
        
        # Get filtered data
        filtered_data = self.get_filtered_data()
        
        # Add contracts to data viewer
        for contract in filtered_data:
            # Ensure contract is normalized
            contract = self.normalize_contract_data(contract)
            
            # Get values in the correct order
            values = [contract[col] for col in self.standard_columns]
            
            # Get AI risk assessment for coloring
            risk_tag = self.get_ai_risk_tag(contract)
            
            self.data_tree.insert("", "end", values=values, tags=(risk_tag,))
    
    def get_filtered_data(self):
        """Get data filtered by current search and filter criteria."""
        if not hasattr(self, 'search_var'):
            return self.data  # No filters set up yet
        
        filtered_data = []
        search_term = self.search_var.get().lower()
        status_filter = self.status_filter_var.get()
        priority_filter = self.priority_filter_var.get()
        
        for contract in self.data:
            # Apply search filter
            if search_term:
                searchable_text = " ".join([
                    str(contract.get("Article Title", "")),
                    str(contract.get("Contractor", "")),
                    str(contract.get("Field / Project", "")),
                    str(contract.get("Country", "")),
                    str(contract.get("DA Comment / Action", ""))
                ]).lower()
                
                if search_term not in searchable_text:
                    continue
            
            # Apply status filter
            if status_filter != "All":
                status = contract.get("DA Comment / Action", "")
                if status_filter == "Open" and not self.ai._is_open(status):
                    continue
                elif status_filter == "Closed" and self.ai._is_open(status):
                    continue
                elif status_filter == "Critical" and self.get_ai_risk_tag(contract) not in ["critical", "pending_critical"]:
                    continue
                elif status_filter == "Pending" and "pending" not in status.lower():
                    continue
            
            # Apply priority filter
            if priority_filter != "All":
                priority = contract.get("Priority", "")
                if priority != priority_filter:
                    continue
            
            filtered_data.append(contract)
        
        return filtered_data
    
    def filter_data_table(self, *args):
        """Apply filters to data table."""
        self.refresh_data_viewer()
    
    def sort_data_column(self, column):
        """Sort data table by column."""
        if not hasattr(self, 'data_sort_reverse'):
            self.data_sort_reverse = {}
        
        # Toggle sort direction
        reverse = self.data_sort_reverse.get(column, False)
        self.data_sort_reverse[column] = not reverse
        
        # Get current items and sort
        items = [(self.data_tree.set(child, column), child) for child in self.data_tree.get_children()]
        
        # Sort based on column type
        if column == "Date (DD/MM/YYYY)":
            # Sort dates properly
            def date_sort_key(item):
                try:
                    return datetime.strptime(item[0], "%d/%m/%Y")
                except:
                    return datetime.min
            items.sort(key=date_sort_key, reverse=reverse)
        elif column == "Priority":
            # Sort priorities (High > Medium > Low)
            priority_order = {"High": 3, "Medium": 2, "Low": 1}
            items.sort(key=lambda item: priority_order.get(item[0], 0), reverse=reverse)
        else:
            # Default string sort
            items.sort(reverse=reverse)
        
        # Rearrange items in tree
        for index, (_, child) in enumerate(items):
            self.data_tree.move(child, '', index)
    
    def clear_data_filters(self):
        """Clear all data filters."""
        self.search_var.set("")
        self.status_filter_var.set("All")
        self.priority_filter_var.set("All")
        self.refresh_data_viewer()
    
    def export_filtered_data(self):
        """Export currently filtered data."""
        filtered_data = self.get_filtered_data()
        
        if not filtered_data:
            messagebox.showinfo("No Data", "No data matches current filters.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("JSON files", "*.json"), ("CSV files", "*.csv")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.json'):
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(filtered_data, f, indent=2)
                elif file_path.endswith('.csv'):
                    df = pd.DataFrame(filtered_data)
                    df.to_csv(file_path, index=False)
                else:
                    df = pd.DataFrame(filtered_data)
                    df.to_excel(file_path, index=False)
                
                messagebox.showinfo("Export Complete", 
                    f"Exported {len(filtered_data)} contracts to {file_path}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export data: {str(e)}")
    
    def view_selected_details(self):
        """View detailed information of selected contract in data viewer."""
        selected_item = self.data_tree.focus()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select a contract to view details.")
            return
        
        try:
            # Get selected contract data
            values = self.data_tree.item(selected_item, 'values')
            contract = {}
            for i, col in enumerate(self.standard_columns):
                if i < len(values):
                    contract[col] = values[i]
            
            # Create detailed view window
            detail_window = tk.Toplevel(self.root)
            detail_window.title("Contract Details")
            detail_window.geometry("700x600")
            detail_window.configure(bg=self.colors['bg_primary'])
            
            # Center the window
            detail_window.transient(self.root)
            detail_window.grab_set()
            
            # Title frame
            title_frame = tk.Frame(detail_window, bg=self.colors['accent_blue_dark'], height=50)
            title_frame.pack(fill="x")
            title_frame.pack_propagate(False)
            
            title_label = tk.Label(title_frame, text="Contract Details", 
                                  bg=self.colors['accent_blue_dark'], fg=self.colors['text_primary'],
                                  font=('Arial', 16, 'bold'))
            title_label.pack(expand=True)
            
            # Content frame
            content_frame = tk.Frame(detail_window, bg=self.colors['bg_primary'])
            content_frame.pack(fill="both", expand=True, padx=20, pady=20)
            
            # Contract information
            info_frame = tk.Frame(content_frame, bg=self.colors['bg_card'], relief='flat', bd=2)
            info_frame.pack(fill="both", expand=True)
            
            # Add contract details
            detail_text = tk.Text(info_frame, wrap=tk.WORD, font=('Consolas', 11),
                                 bg=self.colors['bg_secondary'], fg=self.colors['text_primary'],
                                 relief='flat', bd=0)
            detail_scrollbar = ttk.Scrollbar(info_frame, orient="vertical", command=detail_text.yview)
            detail_text.configure(yscrollcommand=detail_scrollbar.set)
            
            detail_text.pack(side="left", fill="both", expand=True, padx=15, pady=15)
            detail_scrollbar.pack(side="right", fill="y", padx=(0, 15), pady=15)
            
            # Format contract details
            details = []
            details.append("CONTRACT DETAILED INFORMATION")
            details.append("=" * 50)
            details.append("")
            
            for col, value in contract.items():
                details.append(f"{col}: {value}")
            
            details.append("")
            details.append("AI ANALYSIS:")
            details.append("-" * 30)
            
            # Add AI analysis if model is available
            if self.ai.model:
                ai_result = self.ai.predict_status_advanced(contract)
                details.append(f"Risk Level: {ai_result.get('risk_level', 'Unknown')}")
                details.append(f"Risk Score: {ai_result.get('risk_score', 0):.1f}")
                details.append(f"Anomaly Detected: {'Yes' if ai_result.get('is_anomaly') else 'No'}")
                if ai_result.get('prediction'):
                    details.append(f"Predicted Next Status: {ai_result['prediction']}")
                    details.append(f"Confidence: {ai_result.get('confidence', 0):.1%}")
                
                # Add recommendations
                recommendations = self.ai.get_advanced_recommendations(contract)
                if recommendations:
                    details.append("")
                    details.append("AI RECOMMENDATIONS:")
                    for i, rec in enumerate(recommendations, 1):
                        details.append(f"{i}. {rec}")
            else:
                details.append("AI model not trained - basic information only")
            
            detail_text.insert("1.0", "\n".join(details))
            detail_text.configure(state="disabled")
            
            # Close button
            button_frame = tk.Frame(detail_window, bg=self.colors['bg_primary'])
            button_frame.pack(fill="x", padx=20, pady=(0, 20))
            
            close_btn = tk.Button(button_frame, text="Close", command=detail_window.destroy,
                                 bg=self.colors['accent_blue'], fg=self.colors['text_primary'],
                                 font=('Arial', 11, 'bold'), relief='flat', padx=20, pady=8)
            close_btn.pack(side="right")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to show contract details: {str(e)}")
    
    def on_data_tree_double_click(self, event):
        """Handle double-click in data viewer."""
        self.view_selected_details()
    
    def on_data_tree_select(self, event):
        """Handle selection in data viewer."""
        pass  # Can add functionality later if needed

# Main application entry point
def main():
    """Main function to start the application."""
    try:
        root = tk.Tk()
        app = EnhancedContractTrackerApp(root)
        
        # Set up exception handling
        def handle_exception(exc_type, exc_value, exc_traceback):
            if issubclass(exc_type, KeyboardInterrupt):
                import sys
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return
            
            print(f"Uncaught exception: {exc_type.__name__}: {exc_value}")
            messagebox.showerror("Application Error", 
                f"An unexpected error occurred:\n{exc_type.__name__}: {exc_value}")
        
        import sys
        sys.excepthook = handle_exception
        
        root.mainloop()
        
    except Exception as e:
        print(f"Failed to start application: {e}")
        messagebox.showerror("Startup Error", f"Failed to start application: {e}")

if __name__ == "__main__":
    main()