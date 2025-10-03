import pandas as pd
import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import webbrowser
import tempfile
from datetime import datetime
import re

class IndependentTableDashboard:
    def __init__(self):
        self.raw_df = None
        self.tables = []
        self.root = tk.Tk()
        self.root.withdraw()
        
    def select_file(self):
        print("INDEPENDENT MULTI-TABLE DASHBOARD GENERATOR")
        print("=" * 50)
        
        file_path = filedialog.askopenfilename(
            title="Select your Multi-Table Data File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            print("No file selected. Exiting...")
            return None
            
        return file_path
    
    def load_data(self, file_path):
        try:
            print(f"Loading: {Path(file_path).name}")
            
            if file_path.lower().endswith('.csv'):
                # For CSV, read as strings first to preserve formatting
                self.raw_df = pd.read_csv(file_path, dtype=str, keep_default_na=False)
            else:
                # For Excel, read as strings to preserve formatting
                self.raw_df = pd.read_excel(file_path, dtype=str, keep_default_na=False)
            
            # Replace empty strings with NaN for proper processing
            self.raw_df = self.raw_df.replace('', pd.NA)
            
            print(f"Loaded {len(self.raw_df)} rows, {len(self.raw_df.columns)} columns")
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{str(e)}")
            return False
    
    def is_header_row(self, row):
        """Determine if a row looks like headers"""
        non_null_values = row.dropna()
        
        if len(non_null_values) < 2:  # Need at least 2 non-null values
            return False
        
        # Check if values look like column headers (not numeric, not dates)
        header_indicators = 0
        for val in non_null_values:
            val_str = str(val).strip()
            if val_str:
                # Looks like header if it contains common header words or is non-numeric
                if (any(word in val_str.lower() for word in ['name', 'id', 'status', 'date', 'amount', 'description', 'type', 'category']) or
                    not val_str.replace('.', '').replace('-', '').replace('/', '').isdigit()):
                    header_indicators += 1
        
        # Consider it a header if majority of non-null values look like headers
        return header_indicators >= len(non_null_values) * 0.6
    
    def is_empty_separator(self, row):
        """Check if row is mostly empty (table separator)"""
        non_null_count = row.notna().sum()
        return non_null_count <= 1  # Allow up to 1 non-null value
    
    def detect_independent_tables(self):
        """Detect completely independent tables with their own structures"""
        print("\nDetecting independent tables...")
        
        df = self.raw_df.reset_index(drop=True)
        tables = []
        
        i = 0
        table_counter = 1
        
        while i < len(df):
            # Look for next header row
            header_row_idx = None
            
            # Skip empty rows
            while i < len(df) and self.is_empty_separator(df.iloc[i]):
                i += 1
            
            if i >= len(df):
                break
            
            # Check if current row is a header
            if self.is_header_row(df.iloc[i]):
                header_row_idx = i
                print(f"  Found potential header at row {i}")
                
                # Extract headers from this specific row
                header_row = df.iloc[header_row_idx]
                headers = []
                
                for col_idx, val in enumerate(header_row):
                    if pd.notna(val) and str(val).strip():
                        headers.append(str(val).strip())
                    else:
                        headers.append(f"Column_{col_idx + 1}")
                
                # Find data rows for this table
                data_start = header_row_idx + 1
                data_end = data_start
                
                # Look for data rows until we hit empty separator or another header
                while data_end < len(df):
                    current_row = df.iloc[data_end]
                    
                    if self.is_empty_separator(current_row):
                        break
                    elif data_end > data_start and self.is_header_row(current_row):
                        # Found next table's header
                        break
                    else:
                        data_end += 1
                
                # Extract table data with only the columns that have headers
                if data_start < data_end:
                    # Get data rows
                    table_data = df.iloc[data_start:data_end, :len(headers)].copy()
                    
                    # Set proper column names
                    table_data.columns = headers
                    
                    # Clean the data - remove completely empty rows
                    table_data = table_data.dropna(how='all')
                    
                    if len(table_data) > 0:
                        # Generate table name based on its own content
                        table_name = self.determine_table_name(headers, table_data, table_counter)
                        
                        table_info = {
                            'name': table_name,
                            'data': table_data,
                            'headers': headers,
                            'start_row': header_row_idx,
                            'end_row': data_end,
                            'row_count': len(table_data)
                        }
                        
                        tables.append(table_info)
                        print(f"  ✓ Table: {table_name}")
                        print(f"    - Headers: {headers}")
                        print(f"    - Rows: {len(table_data)}")
                        print(f"    - Columns: {len(headers)}")
                        
                        table_counter += 1
                
                i = data_end
            else:
                i += 1
        
        # If no tables detected using header detection, try alternative approach
        if not tables:
            print("  No clear headers detected, attempting alternative detection...")
            tables = self.detect_tables_alternative()
        
        self.tables = tables
        print(f"\nTotal independent tables detected: {len(tables)}")
        return tables
    
    def detect_tables_alternative(self):
        """Alternative table detection for files without clear headers"""
        df = self.raw_df.reset_index(drop=True)
        tables = []
        
        # Look for data blocks separated by empty rows
        current_block_start = None
        
        for i, row in df.iterrows():
            if not self.is_empty_separator(row) and current_block_start is None:
                # Start of new block
                current_block_start = i
            elif self.is_empty_separator(row) and current_block_start is not None:
                # End of current block
                block_data = df.iloc[current_block_start:i].copy()
                
                if len(block_data) > 1:  # Need at least 2 rows (header + data)
                    # Use first row as headers
                    headers = [str(val).strip() if pd.notna(val) else f"Column_{idx+1}" 
                              for idx, val in enumerate(block_data.iloc[0])]
                    
                    # Data is everything except first row
                    table_data = block_data.iloc[1:].copy()
                    table_data.columns = headers
                    table_data = table_data.dropna(how='all')
                    
                    if len(table_data) > 0:
                        table_name = self.determine_table_name(headers, table_data, len(tables) + 1)
                        
                        tables.append({
                            'name': table_name,
                            'data': table_data,
                            'headers': headers,
                            'start_row': current_block_start,
                            'end_row': i,
                            'row_count': len(table_data)
                        })
                
                current_block_start = None
        
        # Handle last block if file doesn't end with empty row
        if current_block_start is not None:
            block_data = df.iloc[current_block_start:].copy()
            if len(block_data) > 1:
                headers = [str(val).strip() if pd.notna(val) else f"Column_{idx+1}" 
                          for idx, val in enumerate(block_data.iloc[0])]
                
                table_data = block_data.iloc[1:].copy()
                table_data.columns = headers
                table_data = table_data.dropna(how='all')
                
                if len(table_data) > 0:
                    table_name = self.determine_table_name(headers, table_data, len(tables) + 1)
                    
                    tables.append({
                        'name': table_name,
                        'data': table_data,
                        'headers': headers,
                        'start_row': current_block_start,
                        'end_row': len(df),
                        'row_count': len(table_data)
                    })
        
        return tables
    
    def determine_table_name(self, headers, data, table_number):
        """Determine table name based on its specific headers and content"""
        headers_lower = [h.lower() for h in headers]
        headers_text = ' '.join(headers_lower)
        
        # Check for specific patterns in headers
        if any(word in headers_text for word in ['project', 'task', 'activity', 'milestone', 'deliverable']):
            return f"Projects & Tasks"
        elif any(word in headers_text for word in ['budget', 'cost', 'expense', 'amount', 'price', 'financial']):
            return f"Budget & Finance"
        elif any(word in headers_text for word in ['resource', 'team', 'assign', 'owner', 'responsible', 'staff']):
            return f"Resources & Team"
        elif any(word in headers_text for word in ['date', 'schedule', 'timeline', 'deadline', 'due']):
            return f"Timeline & Dates"
        elif any(word in headers_text for word in ['status', 'progress', 'phase', 'stage', 'state']):
            return f"Status Tracking"
        elif any(word in headers_text for word in ['risk', 'issue', 'problem', 'concern', 'challenge']):
            return f"Risks & Issues"
        elif any(word in headers_text for word in ['requirement', 'specification', 'criteria', 'standard']):
            return f"Requirements"
        elif any(word in headers_text for word in ['contact', 'client', 'customer', 'stakeholder']):
            return f"Contacts & Stakeholders"
        elif any(word in headers_text for word in ['invoice', 'payment', 'billing', 'transaction']):
            return f"Billing & Payments"
        elif any(word in headers_text for word in ['inventory', 'asset', 'equipment', 'material']):
            return f"Assets & Inventory"
        else:
            # Use first meaningful header or generic name
            meaningful_header = next((h for h in headers if len(h) > 2 and h.lower() not in ['id', 'no', 'sr']), None)
            if meaningful_header:
                return f"{meaningful_header} Data"
            else:
                return f"Data Table {table_number}"
    
    def analyze_table_independently(self, table_info):
        """Analyze each table based on its own unique structure"""
        data = table_info['data']
        headers = table_info['headers']
        
        analysis = {
            'name': table_info['name'],
            'total_rows': len(data),
            'total_columns': len(headers),
            'columns': headers,
            'sample_data': data.head(3).to_dict('records') if len(data) > 0 else []
        }
        
        # Analyze each column independently
        for col in headers:
            if col in data.columns:
                col_data = data[col].dropna()
                if len(col_data) > 0:
                    # Check what type of data this column contains
                    col_lower = col.lower()
                    
                    # Status analysis
                    if any(word in col_lower for word in ['status', 'state', 'phase', 'condition']):
                        status_counts = col_data.value_counts().to_dict()
                        analysis['status_distribution'] = status_counts
                        
                        # Categorize statuses
                        completed = sum(1 for val in col_data 
                                      if any(keyword in str(val).lower() for keyword in ['complete', 'done', 'finished', 'closed', 'success']))
                        in_progress = sum(1 for val in col_data 
                                        if any(keyword in str(val).lower() for keyword in ['progress', 'active', 'working', 'ongoing', 'started']))
                        
                        analysis['completed'] = completed
                        analysis['in_progress'] = in_progress
                        analysis['completion_rate'] = round((completed / len(col_data)) * 100, 1) if len(col_data) > 0 else 0
                    
                    # Progress analysis
                    elif any(word in col_lower for word in ['progress', 'complete', '%', 'percent']):
                        # Try to convert to numeric
                        numeric_data = pd.to_numeric(col_data, errors='coerce').dropna()
                        if len(numeric_data) > 0:
                            analysis['avg_progress'] = round(numeric_data.mean(), 1)
                            analysis['progress_range'] = f"{numeric_data.min():.1f}% - {numeric_data.max():.1f}%"
                    
                    # Budget/Cost analysis
                    elif any(word in col_lower for word in ['budget', 'cost', 'amount', 'price', 'value', 'expense']):
                        # Clean and convert to numeric
                        cleaned_data = col_data.astype(str).str.replace(r'[^0-9.-]', '', regex=True)
                        numeric_data = pd.to_numeric(cleaned_data, errors='coerce').dropna()
                        if len(numeric_data) > 0:
                            analysis['total_budget'] = numeric_data.sum()
                            analysis['avg_budget'] = round(numeric_data.mean(), 0)
                            analysis['budget_range'] = f"${numeric_data.min():,.0f} - ${numeric_data.max():,.0f}"
                    
                    # Date analysis
                    elif any(word in col_lower for word in ['date', 'deadline', 'due', 'start', 'end']):
                        # Try to parse dates
                        try:
                            date_data = pd.to_datetime(col_data, errors='coerce').dropna()
                            if len(date_data) > 0:
                                current_date = pd.Timestamp.now()
                                future_dates = sum(1 for d in date_data if d > current_date)
                                past_dates = sum(1 for d in date_data if d < current_date)
                                
                                analysis['future_dates'] = future_dates
                                analysis['past_dates'] = past_dates
                                analysis['date_range'] = f"{date_data.min().strftime('%Y-%m-%d')} to {date_data.max().strftime('%Y-%m-%d')}"
                        except:
                            pass
        
        return analysis
    
    def generate_html_dashboard(self, tables_analysis):
        """Generate HTML dashboard with completely independent table handling"""
        
        # Prepare data for each table
        tables_data = []
        overall_stats = {
            'total_tables': len(self.tables),
            'total_rows': sum(len(table['data']) for table in self.tables),
            'unique_columns': 0,
            'total_budget': 0,
            'avg_completion': 0
        }
        
        all_columns = set()
        completion_rates = []
        
        for table_info, analysis in zip(self.tables, tables_analysis):
            # Convert table data to JSON format
            table_data = []
            for _, row in table_info['data'].iterrows():
                row_data = {}
                for col in table_info['headers']:
                    if col in row:
                        value = row[col]
                        if pd.isna(value):
                            row_data[col] = ""
                        else:
                            row_data[col] = str(value).strip()
                    else:
                        row_data[col] = ""
                table_data.append(row_data)
            
            tables_data.append({
                'name': analysis['name'],
                'data': table_data,
                'analysis': analysis,
                'headers': table_info['headers']
            })
            
            # Update overall stats
            all_columns.update(table_info['headers'])
            overall_stats['total_budget'] += analysis.get('total_budget', 0)
            
            if 'completion_rate' in analysis:
                completion_rates.append(analysis['completion_rate'])
        
        overall_stats['unique_columns'] = len(all_columns)
        overall_stats['avg_completion'] = round(sum(completion_rates) / len(completion_rates), 1) if completion_rates else 0
        
        html_content = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Independent Multi-Table Dashboard - {datetime.now().strftime('%Y-%m-%d %H:%M')}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            min-height: 100vh;
            color: #2d3748;
        }}
        
        .container {{
            max-width: 1600px;
            margin: 0 auto;
            padding: 20px;
        }}
        
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 40px;
            border-radius: 20px;
            text-align: center;
            margin-bottom: 30px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }}
        
        .header h1 {{
            font-size: 3rem;
            margin-bottom: 15px;
            font-weight: 700;
        }}
        
        .header p {{
            font-size: 1.2rem;
            opacity: 0.9;
        }}
        
        .overall-stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}
        
        .stat-card {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 8px 25px rgba(0,0,0,0.08);
            border-left: 5px solid #667eea;
            text-align: center;
            transition: transform 0.3s ease;
        }}
        
        .stat-card:hover {{
            transform: translateY(-5px);
        }}
        
        .stat-card h3 {{
            color: #718096;
            margin-bottom: 10px;
            text-transform: uppercase;
            font-size: 0.9rem;
            letter-spacing: 1px;
        }}
        
        .stat-card .number {{
            font-size: 2.5rem;
            font-weight: 800;
            color: #2d3748;
            margin-bottom: 5px;
        }}
        
        .table-navigation {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
            margin-bottom: 30px;
        }}
        
        .nav-title {{
            font-size: 1.4rem;
            color: #2d3748;
            margin-bottom: 20px;
            font-weight: 600;
        }}
        
        .nav-buttons {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
        }}
        
        .nav-button {{
            padding: 15px 20px;
            border: 2px solid #e2e8f0;
            border-radius: 12px;
            background: white;
            color: #4a5568;
            cursor: pointer;
            transition: all 0.3s ease;
            font-weight: 600;
            text-align: center;
            font-size: 0.95rem;
        }}
        
        .nav-button:hover {{
            border-color: #667eea;
            background: #f7fafc;
        }}
        
        .nav-button.active {{
            background: #667eea;
            color: white;
            border-color: #667eea;
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
        }}
        
        .table-section {{
            background: white;
            border-radius: 20px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.08);
            overflow: hidden;
            margin-bottom: 30px;
            display: none;
        }}
        
        .table-section.active {{
            display: block;
        }}
        
        .table-header {{
            background: linear-gradient(135deg, #4a5568 0%, #2d3748 100%);
            color: white;
            padding: 30px;
        }}
        
        .table-header h2 {{
            font-size: 1.8rem;
            margin-bottom: 15px;
            font-weight: 600;
        }}
        
        .table-stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            margin-top: 20px;
        }}
        
        .mini-stat {{
            background: rgba(255,255,255,0.15);
            padding: 15px;
            border-radius: 10px;
            text-align: center;
            backdrop-filter: blur(10px);
        }}
        
        .mini-stat .mini-number {{
            font-size: 1.6rem;
            font-weight: 700;
            margin-bottom: 5px;
        }}
        
        .mini-stat .mini-label {{
            font-size: 0.85rem;
            opacity: 0.9;
        }}
        
        .controls {{
            padding: 20px 30px;
            background: #f7fafc;
            border-bottom: 1px solid #e2e8f0;
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
            align-items: center;
        }}
        
        .search-box {{
            padding: 12px 16px;
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            font-size: 14px;
            width: 300px;
            background: white;
        }}
        
        .search-box:focus {{
            outline: none;
            border-color: #667eea;
        }}
        
        .table-wrapper {{
            overflow-x: auto;
            max-height: 500px;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        
        th {{
            background: #f8fafc;
            padding: 14px 10px;
            text-align: left;
            font-weight: 600;
            color: #4a5568;
            border-bottom: 2px solid #e2e8f0;
            position: sticky;
            top: 0;
            z-index: 10;
            white-space: nowrap;
            font-size: 0.9rem;
        }}
        
        td {{
            padding: 12px 10px;
            border-bottom: 1px solid #e2e8f0;
            vertical-align: top;
            font-size: 0.9rem;
        }}
        
        tr:hover {{
            background-color: #f7fafc;
        }}
        
        .status-badge {{
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: capitalize;
            display: inline-block;
        }}
        
        .status-completed {{ 
            background-color: #c6f6d5; 
            color: #22543d; 
        }}
        .status-progress {{ 
            background-color: #bee3f8; 
            color: #2a4365; 
        }}
        .status-pending {{ 
            background-color: #faf089; 
            color: #744210; 
        }}
        .status-default {{ 
            background-color: #e2e8f0; 
            color: #4a5568; 
        }}
        
        .progress-bar {{
            width: 60px;
            height: 6px;
            background: #e2e8f0;
            border-radius: 3px;
            overflow: hidden;
            margin-bottom: 4px;
        }}
        
        .progress-fill {{
            height: 100%;
            background: linear-gradient(90deg, #48bb78, #38a169);
        }}
        
        .progress-text {{
            font-size: 0.75rem;
            color: #4a5568;
            font-weight: 600;
        }}
        
        .footer {{
            background: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
            color: white;
            text-align: center;
            padding: 30px;
            border-radius: 20px;
            margin-top: 30px;
        }}
        
        @media (max-width: 768px) {{
            .container {{ padding: 10px; }}
            .header h1 {{ font-size: 2rem; }}
            .nav-buttons {{ grid-template-columns: 1fr; }}
            .controls {{ flex-direction: column; align-items: stretch; }}
            .search-box {{ width: 100%; }}
            .table-stats {{ grid-template-columns: repeat(2, 1fr); }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Independent Multi-Table Dashboard</h1>
            <p>Each table processed with its own unique structure and columns</p>
            <p>{overall_stats['total_tables']} independent tables • {overall_stats['total_rows']} total records • {overall_stats['unique_columns']} unique columns</p>
            <p style="font-size: 1rem; margin-top: 10px;">Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>
        
        <div class="overall-stats">
            <div class="stat-card">
                <h3>Independent Tables</h3>
                <div class="number">{overall_stats['total_tables']}</div>
            </div>
            <div class="stat-card">
                <h3>Total Records</h3>
                <div class="number">{overall_stats['total_rows']:,}</div>
            </div>
            <div class="stat-card">
                <h3>Unique Columns</h3>
                <div class="number">{overall_stats['unique_columns']}</div>
            </div>
            <div class="stat-card">
                <h3>Combined Budget</h3>
                <div class="number">${overall_stats['total_budget']:,.0f}</div>
            </div>
        </div>
        
        <div class="table-navigation">
            <div class="nav-title">Navigate Between Independent Tables:</div>
            <div class="nav-buttons">
                {' '.join([f'<button class="nav-button" onclick="showTable({i})">{table["name"]}<br><small>({len(table["data"])} rows)</small></button>' for i, table in enumerate(tables_data)])}
            </div>
        </div>
        
        {''.join([self.generate_table_section_html(i, table) for i, table in enumerate(tables_data)])}
        
        <div class="footer">
            <h3>Independent Multi-Table Dashboard</h3>
            <p>Each of the {overall_stats['total_tables']} tables processed independently with its own unique column structure</p>
            <p>No assumptions made about data format - each table stands alone</p>
        </div>
    </div>

    <script>
        const tablesData = {json.dumps(tables_data, indent=2)};
        
        function showTable(tableIndex) {{
            // Hide all tables
            document.querySelectorAll('.table-section').forEach(section => {{
                section.classList.remove('active');
            }});
            
            // Remove active class from all nav buttons
            document.querySelectorAll('.nav-button').forEach(button => {{
                button.classList.remove('active');
            }});
            
            // Show selected table
            document.getElementById(`table-${{tableIndex}}`).classList.add('active');
            document.querySelectorAll('.nav-button')[tableIndex].classList.add('active');
            
            // Generate table content
            generateTable(tableIndex);
        }}
        
        function generateTable(tableIndex) {{
            const table = tablesData[tableIndex];
            const tbody = document.getElementById(`table-body-${{tableIndex}}`);
            const thead = document.getElementById(`table-header-${{tableIndex}}`);
            
            // Clear existing content
            tbody.innerHTML = '';
            thead.innerHTML = '';
            
            if (table.data.length === 0) {{
                tbody.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 40px;">No data in this table</td></tr>';
                return;
            }}
            
            // Generate header using this table's specific headers
            const headerRow = document.createElement('tr');
            table.headers.forEach(header => {{
                const th = document.createElement('th');
                th.textContent = header;
                headerRow.appendChild(th);
            }});
            thead.appendChild(headerRow);
            
            // Generate rows using this table's specific data
            table.data.forEach(row => {{
                const tr = document.createElement('tr');
                table.headers.forEach(header => {{
                    const td = document.createElement('td');
                    const value = row[header] || '';
                    td.innerHTML = formatCellContent(value, header);
                    tr.appendChild(td);
                }});
                tbody.appendChild(tr);
            }});
        }}
        
        function formatCellContent(value, columnName) {{
            if (!value) return '';
            
            const colLower = columnName.toLowerCase();
            const valueLower = value.toLowerCase();
            
            // Status formatting
            if (colLower.includes('status') || colLower.includes('phase') || colLower.includes('state')) {{
                let statusClass = 'status-default';
                
                if (valueLower.includes('complete') || valueLower.includes('done') || valueLower.includes('finished') || valueLower.includes('success')) {{
                    statusClass = 'status-completed';
                }} else if (valueLower.includes('progress') || valueLower.includes('active') || valueLower.includes('working') || valueLower.includes('started')) {{
                    statusClass = 'status-progress';
                }} else if (valueLower.includes('pending') || valueLower.includes('waiting') || valueLower.includes('planned')) {{
                    statusClass = 'status-pending';
                }}
                
                return `<span class="status-badge ${{statusClass}}">${{value}}</span>`;
            }}
            
            // Progress formatting
            if ((colLower.includes('progress') || colLower.includes('complete')) && /^\d+\.?\d*%?$/.test(value.replace('%', ''))) {{
                const progressValue = parseFloat(value.replace('%', '')) || 0;
                return `
                    <div class="progress-bar">
                        <div class="progress-fill" style="width: ${{Math.min(100, progressValue)}}%"></div>
                    </div>
                    <div class="progress-text">${{progressValue}}%</div>
                `;
            }}
            
            // Currency formatting
            if (colLower.includes('budget') || colLower.includes('cost') || colLower.includes('amount') || colLower.includes('price')) {{
                const numValue = parseFloat(value.toString().replace(/[^0-9.-]/g, ''));
                if (!isNaN(numValue) && numValue > 0) {{
                    return `$$${{numValue.toLocaleString()}}`;
                }}
            }}
            
            // Date formatting
            if (colLower.includes('date')) {{
                try {{
                    const date = new Date(value);
                    if (!isNaN(date.getTime())) {{
                        return date.toLocaleDateString();
                    }}
                }} catch (e) {{
                    // Not a valid date, return as is
                }}
            }}
            
            return value;
        }}
        
        function searchTable(tableIndex) {{
            const searchTerm = document.getElementById(`search-${{tableIndex}}`).value.toLowerCase();
            const rows = document.querySelectorAll(`#table-${{tableIndex}} tbody tr`);
            
            let visibleCount = 0;
            rows.forEach(row => {{
                const text = row.textContent.toLowerCase();
                const isVisible = text.includes(searchTerm);
                row.style.display = isVisible ? '' : 'none';
                if (isVisible) visibleCount++;
            }});
            
            // Update visible count
            const totalSpan = document.querySelector(`#table-${{tableIndex}} .table-info .visible-count`);
            if (totalSpan) {{
                totalSpan.textContent = visibleCount;
            }}
        }}
        
        // Initialize first table
        window.onload = function() {{
            if (tablesData.length > 0) {{
                showTable(0);
            }}
        }};
    </script>
</body>
</html>
        """
        
        return html_content
    
    def generate_table_section_html(self, index, table_data):
        """Generate HTML section for individual independent table"""
        analysis = table_data['analysis']
        
        # Build mini stats based on what's available in this specific table
        mini_stats = []
        mini_stats.append(f'''
            <div class="mini-stat">
                <div class="mini-number">{analysis['total_rows']}</div>
                <div class="mini-label">Rows</div>
            </div>
        ''')
        mini_stats.append(f'''
            <div class="mini-stat">
                <div class="mini-number">{analysis['total_columns']}</div>
                <div class="mini-label">Columns</div>
            </div>
        ''')
        
        if 'completion_rate' in analysis:
            mini_stats.append(f'''
                <div class="mini-stat">
                    <div class="mini-number">{analysis['completion_rate']:.1f}%</div>
                    <div class="mini-label">Completion</div>
                </div>
            ''')
        
        if 'total_budget' in analysis:
            mini_stats.append(f'''
                <div class="mini-stat">
                    <div class="mini-number">${analysis['total_budget']:,.0f}</div>
                    <div class="mini-label">Total Value</div>
                </div>
            ''')
        
        if 'avg_progress' in analysis:
            mini_stats.append(f'''
                <div class="mini-stat">
                    <div class="mini-number">{analysis['avg_progress']:.1f}%</div>
                    <div class="mini-label">Avg Progress</div>
                </div>
            ''')
        
        return f"""
        <div class="table-section" id="table-{index}">
            <div class="table-header">
                <h2>{analysis['name']}</h2>
                <p style="opacity: 0.9; margin-bottom: 15px;">Independent table with {analysis['total_columns']} unique columns</p>
                <div class="table-stats">
                    {''.join(mini_stats)}
                </div>
            </div>
            
            <div class="controls">
                <input type="text" class="search-box" id="search-{index}" placeholder="Search this table..." onkeyup="searchTable({index})">
                <span style="color: #718096;">Showing <span class="visible-count">{analysis['total_rows']}</span> of {analysis['total_rows']} records</span>
            </div>
            
            <div class="table-wrapper">
                <table>
                    <thead id="table-header-{index}"></thead>
                    <tbody id="table-body-{index}"></tbody>
                </table>
            </div>
        </div>
        """
    
    def save_and_open_dashboard(self, html_content):
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
            f.write(html_content)
            temp_path = f.name
        
        print(f"Independent multi-table dashboard generated: {temp_path}")
        print("Opening in your default browser...")
        
        webbrowser.open(f'file://{temp_path}')
        
        current_dir = Path.cwd()
        dashboard_path = current_dir / f"independent_tables_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        
        with open(dashboard_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"Dashboard also saved as: {dashboard_path}")
        return str(dashboard_path)
    
    def run(self):
        try:
            file_path = self.select_file()
            if not file_path:
                return
            
            if not self.load_data(file_path):
                return
            
            # Detect completely independent tables
            tables = self.detect_independent_tables()
            
            if not tables:
                messagebox.showwarning("No Tables Found", "Could not detect any separate data tables in the file.")
                return
            
            # Analyze each table independently
            print("\nAnalyzing each table independently...")
            tables_analysis = []
            for table_info in tables:
                analysis = self.analyze_table_independently(table_info)
                tables_analysis.append(analysis)
            
            print("Generating independent multi-table dashboard...")
            html_content = self.generate_html_dashboard(tables_analysis)
            
            dashboard_path = self.save_and_open_dashboard(html_content)
            
            # Show detailed summary
            summary_msg = f"Independent Multi-Table Dashboard Created!\\n\\n"
            summary_msg += f"✓ Detected {len(tables)} completely independent tables\\n"
            summary_msg += f"✓ Each table processed with its own unique structure\\n"
            summary_msg += f"✓ Total records: {sum(len(t['data']) for t in tables):,}\\n"
            summary_msg += f"✓ Unique column types: {len(set().union(*(t['headers'] for t in tables)))}\\n\\n"
            
            summary_msg += "Tables detected:\\n"
            for i, (table, analysis) in enumerate(zip(tables, tables_analysis), 1):
                summary_msg += f"  {i}. {analysis['name']} ({len(table['data'])} rows, {len(table['headers'])} cols)\\n"
            
            summary_msg += f"\\nSaved as: {Path(dashboard_path).name}\\n"
            summary_msg += "Dashboard is completely offline and shareable!"
            
            messagebox.showinfo("Success!", summary_msg)
            
            print("\n" + "="*60)
            print("SUCCESS! Independent Multi-Table Dashboard Ready!")
            print("="*60)
            print("Key Features:")
            print("- Each table maintains its own unique column structure")
            print("- No cross-table assumptions or formatting applied")
            print("- Independent analysis for each table type")
            print("- Smart table naming based on actual content")
            print("- Professional navigation between different table types")
            print("- Individual search and filtering per table")
            
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            print(error_msg)
            messagebox.showerror("Error", error_msg)
        
        finally:
            self.root.destroy()

if __name__ == "__main__":
    dashboard = IndependentTableDashboard()
    dashboard.run()