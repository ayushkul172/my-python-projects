import pandas as pd
import numpy as np
from datetime import datetime


def preprocess_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Basic preprocessing to normalize common column names and types.
    This is a lightweight stub to allow the Streamlit UI to load for testing.
    """
    df = df.copy()

    # Common renames
    renames = {
        'Drilling Unit Name': 'Rig Name',
        'Drilling Unit': 'Rig Name',
        'Rig': 'Rig Name',
        'Start Date': 'Contract Start Date',
        'End Date': 'Contract End Date',
    }
    df.rename(columns=renames, inplace=True)

    # Parse dates
    for col in ['Contract Start Date', 'Contract End Date']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # Ensure numeric columns
    for col in ['Dayrate ($k)', 'Contract value ($m)', 'Contract Length']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    # Fill Rig Name if missing
    if 'Rig Name' not in df.columns:
        df['Rig Name'] = 'Unknown Rig'

    return df


class AdvancedClimateIntelligence:
    """Stub for climate intelligence helper."""
    def analyze_location(self, location_str):
        # Very simple placeholder scoring: random-ish but deterministic
        if not location_str:
            return 50.0
        return float((sum(ord(c) for c in str(location_str)) % 50) + 50)


class RigEfficiencyCalculator:
    """Lightweight stub for rig efficiency calculations used by the app.
    These heuristics are simple and intended only to let the UI function while
    you plug in your real calculation logic.
    """
    def __init__(self):
        self.name = 'stub_calculator'

    def calculate_comprehensive_efficiency(self, df: pd.DataFrame):
        if df is None or df.empty:
            return None

        # Basic metrics derived from available columns
        metrics = {}

        # Utilization: percent of contracts with Status containing 'active'
        if 'Status' in df.columns:
            active = df['Status'].fillna('').str.contains('active', case=False, na=False).sum()
            total = len(df)
            metrics['contract_utilization'] = float(active / total * 100)
        else:
            metrics['contract_utilization'] = 70.0

        # Dayrate efficiency: normalize mean dayrate to a 0-100 scale using a soft cap
        if 'Dayrate ($k)' in df.columns and not df['Dayrate ($k)'].dropna().empty:
            mean_dayrate = df['Dayrate ($k)'].mean()
            # Assume 300k is excellent, 50k is poor
            metrics['dayrate_efficiency'] = float(np.clip((mean_dayrate - 50) / (300 - 50) * 100, 0, 100))
        else:
            metrics['dayrate_efficiency'] = 65.0

        # Contract stability: inverse of std dev of contract length if available
        if 'Contract Length' in df.columns and not df['Contract Length'].dropna().empty:
            std = df['Contract Length'].std()
            metrics['contract_stability'] = float(np.clip(100 - std, 20, 100))
        else:
            metrics['contract_stability'] = 70.0

        # Location complexity: simple function based on unique locations
        if 'Current Location' in df.columns:
            locations = df['Current Location'].nunique()
            metrics['location_complexity'] = float(np.clip(100 - (locations - 1) * 10, 20, 100))
        else:
            metrics['location_complexity'] = 80.0

        # Climate impact: use simple analyzer if available
        if 'Current Location' in df.columns:
            locs = df['Current Location'].dropna().astype(str)
            if not locs.empty:
                vals = [AdvancedClimateIntelligence().analyze_location(s) for s in locs]
                metrics['climate_impact'] = float(np.mean(vals))
            else:
                metrics['climate_impact'] = 70.0
        else:
            metrics['climate_impact'] = 70.0

        # Contract performance: proxy by dayrate * utilization / 100
        metrics['contract_performance'] = float(np.clip((metrics['dayrate_efficiency'] * metrics['contract_utilization']) / 100, 0, 100))

        # Overall efficiency: weighted average
        weights = {
            'contract_utilization': 0.25,
            'dayrate_efficiency': 0.20,
            'contract_stability': 0.15,
            'location_complexity': 0.15,
            'climate_impact': 0.10,
            'contract_performance': 0.15
        }
        overall = 0.0
        for k, w in weights.items():
            overall += metrics.get(k, 70.0) * w

        metrics['overall_efficiency'] = float(np.clip(overall, 0, 100))

        # Grade
        if metrics['overall_efficiency'] >= 85:
            metrics['efficiency_grade'] = 'A'
        elif metrics['overall_efficiency'] >= 75:
            metrics['efficiency_grade'] = 'B'
        elif metrics['overall_efficiency'] >= 60:
            metrics['efficiency_grade'] = 'C'
        else:
            metrics['efficiency_grade'] = 'D'

        metrics['insights'] = []

        return metrics

    def generate_contract_summary(self, df: pd.DataFrame, metrics: dict):
        if df is None:
            return None

        rig_name = df['Rig Name'].iloc[0] if 'Rig Name' in df.columns and not df['Rig Name'].empty else 'Unknown'
        total_contracts = len(df)
        active_contracts = df['Status'].fillna('').str.contains('active', case=False, na=False).sum() if 'Status' in df.columns else 0
        total_value = float(df['Contract value ($m)'].sum()) if 'Contract value ($m)' in df.columns else 0.0
        avg_dayrate = float(df['Dayrate ($k)'].mean()) if 'Dayrate ($k)' in df.columns else 0.0

        return {
            'rig_name': rig_name,
            'total_contracts': int(total_contracts),
            'active_contracts': int(active_contracts),
            'efficiency_grade': metrics.get('efficiency_grade', 'N/A') if metrics else 'N/A',
            'total_contract_value': total_value,
            'average_dayrate': avg_dayrate,
            'top_strength': 'Balanced Performance',
            'primary_concern': 'Data Quality' if df.isnull().any().any() else 'None'
        }
"""
Rig Efficiency Analysis Backend - COMPLETE VERSION
Core calculation and AI logic with all components
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')


def preprocess_dataframe(df):
    """Preprocess uploaded dataframe"""
    # Standardize column names - handle common variations
    # IMPORTANT: Create 'Rig Name' column from various possible source columns
    if 'Rig Name' not in df.columns:
        if 'Drilling Unit Name' in df.columns:
            df['Rig Name'] = df['Drilling Unit Name']
        elif 'Asset Name' in df.columns:
            df['Rig Name'] = df['Asset Name']
        elif 'Unit Name' in df.columns:
            df['Rig Name'] = df['Unit Name']
        elif 'Contractor' in df.columns:
            df['Rig Name'] = df['Contractor']
    
    # Handle location column variations
    if 'Location' not in df.columns and 'Current Location' in df.columns:
        df['Location'] = df['Current Location']
    
    # Handle dayrate column variations
    if 'Dayrate ($k)' not in df.columns:
        if 'Day Rate' in df.columns:
            df['Dayrate ($k)'] = df['Day Rate']
        elif 'Dayrate' in df.columns:
            df['Dayrate ($k)'] = df['Dayrate']
        elif 'Rate ($k)' in df.columns:
            df['Dayrate ($k)'] = df['Rate ($k)']
    
    # Convert date columns
    date_columns = ['Contract Start Date', 'Contract End Date', 'Award Date', 'TerminationDate']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Clean numeric columns
    numeric_columns = ['Dayrate ($k)', 'Contract value ($m)', 'Contract Length', 'Contract Days Remaining']
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    # Fill missing values
    df = df.fillna({
        'Contract Length': 0,
        'Dayrate ($k)': 0,
        'Contract value ($m)': 0
    })
    
    return df


class AdvancedClimateIntelligence:
    """
    Advanced AI-powered climate analysis engine for rig operations
    Uses multiple algorithms for climate impact prediction and optimization
    """
    
    def __init__(self):
        self.climate_profiles = self._initialize_enhanced_climate_data()
        self.seasonal_patterns = self._initialize_seasonal_patterns()
        self.weather_severity_matrix = self._initialize_severity_matrix()
        
    def _initialize_enhanced_climate_data(self):
        """Enhanced climate data with granular seasonal information"""
        return {
            # Gulf of Mexico - Enhanced
            'gulf of mexico': {
                'climate': 'tropical_storm',
                'risk_months': [6, 7, 8, 9, 10],
                'peak_risk_months': [8, 9],
                'efficiency_factor': 0.75,
                'downtime_risk': 0.25,
                'seasonal_multipliers': {
                    1: 0.95, 2: 0.95, 3: 0.90, 4: 0.85, 5: 0.80,
                    6: 0.70, 7: 0.65, 8: 0.55, 9: 0.60, 10: 0.75,
                    11: 0.90, 12: 0.95
                },
                'weather_events': {
                    'hurricanes': {'probability': 0.40, 'avg_duration_days': 5, 'severity': 0.90},
                    'tropical_storms': {'probability': 0.60, 'avg_duration_days': 3, 'severity': 0.60},
                    'high_seas': {'probability': 0.30, 'avg_duration_days': 2, 'severity': 0.40}
                },
                'optimal_operating_window': [11, 12, 1, 2, 3],
                'description': 'Hurricane season impact with severe weather patterns'
            },
            'us gulf': {
                'climate': 'tropical_storm',
                'risk_months': [6, 7, 8, 9, 10],
                'peak_risk_months': [8, 9],
                'efficiency_factor': 0.75,
                'downtime_risk': 0.25,
                'seasonal_multipliers': {
                    1: 0.95, 2: 0.95, 3: 0.90, 4: 0.85, 5: 0.80,
                    6: 0.70, 7: 0.65, 8: 0.55, 9: 0.60, 10: 0.75,
                    11: 0.90, 12: 0.95
                },
                'weather_events': {
                    'hurricanes': {'probability': 0.40, 'avg_duration_days': 5, 'severity': 0.90},
                    'tropical_storms': {'probability': 0.60, 'avg_duration_days': 3, 'severity': 0.60}
                },
                'optimal_operating_window': [11, 12, 1, 2, 3]
            },
            
            # North Sea - Enhanced
            'north sea': {
                'climate': 'harsh_winter',
                'risk_months': [11, 12, 1, 2, 3],
                'peak_risk_months': [12, 1, 2],
                'efficiency_factor': 0.70,
                'downtime_risk': 0.30,
                'seasonal_multipliers': {
                    1: 0.60, 2: 0.65, 3: 0.70, 4: 0.80, 5: 0.90,
                    6: 0.95, 7: 0.95, 8: 0.95, 9: 0.85, 10: 0.75,
                    11: 0.65, 12: 0.60
                },
                'weather_events': {
                    'winter_storms': {'probability': 0.70, 'avg_duration_days': 4, 'severity': 0.85},
                    'high_winds': {'probability': 0.80, 'avg_duration_days': 3, 'severity': 0.70},
                    'icing': {'probability': 0.50, 'avg_duration_days': 2, 'severity': 0.75}
                },
                'optimal_operating_window': [5, 6, 7, 8, 9],
                'description': 'Severe winter conditions with extreme weather'
            },
            'norway': {
                'climate': 'harsh_winter',
                'risk_months': [11, 12, 1, 2, 3],
                'peak_risk_months': [12, 1],
                'efficiency_factor': 0.70,
                'downtime_risk': 0.30,
                'seasonal_multipliers': {
                    1: 0.55, 2: 0.60, 3: 0.70, 4: 0.80, 5: 0.90,
                    6: 0.95, 7: 0.95, 8: 0.95, 9: 0.85, 10: 0.75,
                    11: 0.65, 12: 0.55
                },
                'weather_events': {
                    'arctic_storms': {'probability': 0.75, 'avg_duration_days': 5, 'severity': 0.90},
                    'extreme_cold': {'probability': 0.60, 'avg_duration_days': 7, 'severity': 0.70}
                },
                'optimal_operating_window': [5, 6, 7, 8]
            },
            'uk': {
                'climate': 'moderate_maritime',
                'risk_months': [11, 12, 1, 2],
                'peak_risk_months': [12, 1],
                'efficiency_factor': 0.80,
                'downtime_risk': 0.20,
                'seasonal_multipliers': {
                    1: 0.75, 2: 0.75, 3: 0.80, 4: 0.85, 5: 0.90,
                    6: 0.95, 7: 0.95, 8: 0.95, 9: 0.90, 10: 0.85,
                    11: 0.80, 12: 0.75
                },
                'weather_events': {
                    'winter_storms': {'probability': 0.50, 'avg_duration_days': 3, 'severity': 0.65},
                    'fog': {'probability': 0.40, 'avg_duration_days': 2, 'severity': 0.40}
                },
                'optimal_operating_window': [4, 5, 6, 7, 8, 9]
            },
            
            # Middle East - Enhanced
            'saudi arabia': {
                'climate': 'desert_stable',
                'risk_months': [6, 7, 8],
                'peak_risk_months': [7, 8],
                'efficiency_factor': 0.95,
                'downtime_risk': 0.05,
                'seasonal_multipliers': {
                    1: 0.98, 2: 0.98, 3: 0.97, 4: 0.96, 5: 0.93,
                    6: 0.90, 7: 0.88, 8: 0.88, 9: 0.92, 10: 0.96,
                    11: 0.98, 12: 0.98
                },
                'weather_events': {
                    'extreme_heat': {'probability': 0.30, 'avg_duration_days': 10, 'severity': 0.30},
                    'sandstorms': {'probability': 0.15, 'avg_duration_days': 1, 'severity': 0.40}
                },
                'optimal_operating_window': [10, 11, 12, 1, 2, 3, 4],
                'description': 'Stable desert climate with minimal weather disruption'
            },
            'uae': {
                'climate': 'desert_stable',
                'risk_months': [6, 7, 8],
                'peak_risk_months': [7, 8],
                'efficiency_factor': 0.90,
                'downtime_risk': 0.10,
                'seasonal_multipliers': {
                    1: 0.96, 2: 0.96, 3: 0.95, 4: 0.93, 5: 0.88,
                    6: 0.85, 7: 0.82, 8: 0.82, 9: 0.88, 10: 0.93,
                    11: 0.96, 12: 0.96
                },
                'weather_events': {
                    'extreme_heat': {'probability': 0.40, 'avg_duration_days': 12, 'severity': 0.35},
                    'shamal_winds': {'probability': 0.25, 'avg_duration_days': 2, 'severity': 0.45}
                },
                'optimal_operating_window': [10, 11, 12, 1, 2, 3]
            },
            'qatar': {
                'climate': 'desert_stable',
                'risk_months': [6, 7, 8],
                'peak_risk_months': [7],
                'efficiency_factor': 0.90,
                'downtime_risk': 0.10,
                'seasonal_multipliers': {
                    1: 0.96, 2: 0.96, 3: 0.95, 4: 0.93, 5: 0.88,
                    6: 0.85, 7: 0.80, 8: 0.85, 9: 0.90, 10: 0.94,
                    11: 0.96, 12: 0.96
                },
                'weather_events': {
                    'extreme_heat': {'probability': 0.45, 'avg_duration_days': 15, 'severity': 0.35}
                },
                'optimal_operating_window': [10, 11, 12, 1, 2, 3, 4]
            },
            
            # Asia Pacific - Enhanced
            'india': {
                'climate': 'monsoon',
                'risk_months': [6, 7, 8, 9],
                'peak_risk_months': [7, 8],
                'efficiency_factor': 0.70,
                'downtime_risk': 0.30,
                'seasonal_multipliers': {
                    1: 0.95, 2: 0.95, 3: 0.93, 4: 0.88, 5: 0.80,
                    6: 0.65, 7: 0.55, 8: 0.55, 9: 0.70, 10: 0.85,
                    11: 0.93, 12: 0.95
                },
                'weather_events': {
                    'monsoon_rains': {'probability': 0.85, 'avg_duration_days': 60, 'severity': 0.75},
                    'cyclones': {'probability': 0.25, 'avg_duration_days': 4, 'severity': 0.85},
                    'rough_seas': {'probability': 0.70, 'avg_duration_days': 3, 'severity': 0.60}
                },
                'optimal_operating_window': [10, 11, 12, 1, 2, 3],
                'description': 'Monsoon season impact with heavy rainfall'
            },
            'indonesia': {
                'climate': 'tropical_monsoon',
                'risk_months': [11, 12, 1, 2, 3],
                'peak_risk_months': [12, 1, 2],
                'efficiency_factor': 0.75,
                'downtime_risk': 0.25,
                'seasonal_multipliers': {
                    1: 0.70, 2: 0.70, 3: 0.75, 4: 0.85, 5: 0.90,
                    6: 0.95, 7: 0.95, 8: 0.95, 9: 0.90, 10: 0.85,
                    11: 0.75, 12: 0.70
                },
                'weather_events': {
                    'tropical_storms': {'probability': 0.50, 'avg_duration_days': 3, 'severity': 0.70},
                    'heavy_rainfall': {'probability': 0.75, 'avg_duration_days': 5, 'severity': 0.50}
                },
                'optimal_operating_window': [5, 6, 7, 8, 9]
            },
            'malaysia': {
                'climate': 'tropical_stable',
                'risk_months': [11, 12],
                'peak_risk_months': [11],
                'efficiency_factor': 0.85,
                'downtime_risk': 0.15,
                'seasonal_multipliers': {
                    1: 0.88, 2: 0.90, 3: 0.92, 4: 0.93, 5: 0.93,
                    6: 0.93, 7: 0.93, 8: 0.92, 9: 0.90, 10: 0.88,
                    11: 0.82, 12: 0.85
                },
                'weather_events': {
                    'monsoon_winds': {'probability': 0.40, 'avg_duration_days': 3, 'severity': 0.50}
                },
                'optimal_operating_window': [2, 3, 4, 5, 6, 7, 8]
            },
            'australia': {
                'climate': 'cyclone_risk',
                'risk_months': [11, 12, 1, 2, 3, 4],
                'peak_risk_months': [1, 2, 3],
                'efficiency_factor': 0.75,
                'downtime_risk': 0.25,
                'seasonal_multipliers': {
                    1: 0.65, 2: 0.65, 3: 0.70, 4: 0.75, 5: 0.85,
                    6: 0.93, 7: 0.95, 8: 0.95, 9: 0.93, 10: 0.88,
                    11: 0.78, 12: 0.70
                },
                'weather_events': {
                    'cyclones': {'probability': 0.35, 'avg_duration_days': 5, 'severity': 0.85},
                    'tropical_lows': {'probability': 0.55, 'avg_duration_days': 3, 'severity': 0.60}
                },
                'optimal_operating_window': [5, 6, 7, 8, 9]
            },
            
            # South America - Enhanced
            'brazil': {
                'climate': 'tropical_variable',
                'risk_months': [1, 2, 3],
                'peak_risk_months': [2],
                'efficiency_factor': 0.80,
                'downtime_risk': 0.20,
                'seasonal_multipliers': {
                    1: 0.75, 2: 0.75, 3: 0.78, 4: 0.85, 5: 0.90,
                    6: 0.93, 7: 0.93, 8: 0.90, 9: 0.88, 10: 0.85,
                    11: 0.82, 12: 0.78
                },
                'weather_events': {
                    'heavy_rainfall': {'probability': 0.60, 'avg_duration_days': 4, 'severity': 0.55},
                    'tropical_storms': {'probability': 0.30, 'avg_duration_days': 2, 'severity': 0.60}
                },
                'optimal_operating_window': [5, 6, 7, 8, 9]
            },
            'argentina': {
                'climate': 'temperate',
                'risk_months': [6, 7, 8],
                'peak_risk_months': [7],
                'efficiency_factor': 0.85,
                'downtime_risk': 0.15,
                'seasonal_multipliers': {
                    1: 0.93, 2: 0.93, 3: 0.90, 4: 0.88, 5: 0.85,
                    6: 0.80, 7: 0.78, 8: 0.80, 9: 0.85, 10: 0.90,
                    11: 0.93, 12: 0.93
                },
                'weather_events': {
                    'winter_storms': {'probability': 0.40, 'avg_duration_days': 3, 'severity': 0.60}
                },
                'optimal_operating_window': [10, 11, 12, 1, 2, 3]
            },
            
            # Africa - Enhanced
            'nigeria': {
                'climate': 'tropical_monsoon',
                'risk_months': [4, 5, 6, 7, 8, 9],
                'peak_risk_months': [6, 7, 8],
                'efficiency_factor': 0.75,
                'downtime_risk': 0.25,
                'seasonal_multipliers': {
                    1: 0.93, 2: 0.93, 3: 0.88, 4: 0.80, 5: 0.72,
                    6: 0.65, 7: 0.65, 8: 0.65, 9: 0.75, 10: 0.88,
                    11: 0.93, 12: 0.93
                },
                'weather_events': {
                    'monsoon_rains': {'probability': 0.80, 'avg_duration_days': 90, 'severity': 0.65},
                    'rough_seas': {'probability': 0.60, 'avg_duration_days': 3, 'severity': 0.55}
                },
                'optimal_operating_window': [10, 11, 12, 1, 2]
            },
            'angola': {
                'climate': 'tropical',
                'risk_months': [11, 12, 1, 2, 3],
                'peak_risk_months': [1, 2],
                'efficiency_factor': 0.80,
                'downtime_risk': 0.20,
                'seasonal_multipliers': {
                    1: 0.75, 2: 0.75, 3: 0.78, 4: 0.85, 5: 0.90,
                    6: 0.93, 7: 0.93, 8: 0.93, 9: 0.90, 10: 0.85,
                    11: 0.78, 12: 0.75
                },
                'weather_events': {
                    'tropical_rains': {'probability': 0.65, 'avg_duration_days': 5, 'severity': 0.55}
                },
                'optimal_operating_window': [5, 6, 7, 8, 9]
            },
            
            # Default
            'default': {
                'climate': 'moderate',
                'risk_months': [],
                'peak_risk_months': [],
                'efficiency_factor': 0.90,
                'downtime_risk': 0.10,
                'seasonal_multipliers': {i: 0.90 for i in range(1, 13)},
                'weather_events': {},
                'optimal_operating_window': list(range(1, 13))
            }
        }
    
    def _initialize_seasonal_patterns(self):
        """Initialize seasonal pattern analysis"""
        return {
            'northern_hemisphere': {
                'winter': [12, 1, 2],
                'spring': [3, 4, 5],
                'summer': [6, 7, 8],
                'autumn': [9, 10, 11]
            },
            'southern_hemisphere': {
                'winter': [6, 7, 8],
                'spring': [9, 10, 11],
                'summer': [12, 1, 2],
                'autumn': [3, 4, 5]
            },
            'tropical': {
                'wet_season': [5, 6, 7, 8, 9, 10],
                'dry_season': [11, 12, 1, 2, 3, 4]
            }
        }
    
    def _initialize_severity_matrix(self):
        """Initialize weather event severity matrix"""
        return {
            'hurricanes': {'downtime_days': 7, 'efficiency_impact': 0.95, 'cost_multiplier': 3.0},
            'tropical_storms': {'downtime_days': 4, 'efficiency_impact': 0.70, 'cost_multiplier': 2.0},
            'winter_storms': {'downtime_days': 5, 'efficiency_impact': 0.80, 'cost_multiplier': 2.5},
            'cyclones': {'downtime_days': 6, 'efficiency_impact': 0.90, 'cost_multiplier': 3.0},
            'monsoon_rains': {'downtime_days': 3, 'efficiency_impact': 0.60, 'cost_multiplier': 1.5},
            'extreme_heat': {'downtime_days': 1, 'efficiency_impact': 0.30, 'cost_multiplier': 1.2},
            'extreme_cold': {'downtime_days': 4, 'efficiency_impact': 0.70, 'cost_multiplier': 2.0},
            'high_winds': {'downtime_days': 2, 'efficiency_impact': 0.50, 'cost_multiplier': 1.5},
            'high_seas': {'downtime_days': 2, 'efficiency_impact': 0.45, 'cost_multiplier': 1.3},
            'fog': {'downtime_days': 1, 'efficiency_impact': 0.35, 'cost_multiplier': 1.1}
        }
    
    def calculate_time_weighted_climate_efficiency(self, location, start_date, end_date):
        """
        Advanced AI Algorithm 1: Time-Weighted Climate Efficiency
        """
        try:
            location_lower = str(location).lower()
            climate_data = self._get_climate_profile(location_lower)
            
            if pd.isna(start_date) or pd.isna(end_date):
                return climate_data['efficiency_factor'] * 100
            
            date_range = pd.date_range(start=start_date, end=end_date, freq='D')
            daily_scores = []
            
            for date in date_range:
                month = date.month
                seasonal_multiplier = climate_data['seasonal_multipliers'].get(month, 0.90)
                
                if month in climate_data.get('peak_risk_months', []):
                    seasonal_multiplier *= 0.85
                elif month in climate_data.get('risk_months', []):
                    seasonal_multiplier *= 0.93
                
                weather_adjustment = self._calculate_weather_event_impact(climate_data, month, date)
                daily_score = seasonal_multiplier * weather_adjustment * 100
                daily_scores.append(daily_score)
            
            efficiency_score = np.mean(daily_scores)
            return min(max(efficiency_score, 0), 100)
            
        except Exception as e:
            return 75.0
    
    def _calculate_weather_event_impact(self, climate_data, month, date):
        """Calculate weather event probability impact"""
        weather_events = climate_data.get('weather_events', {})
        if not weather_events:
            return 1.0
        
        total_impact = 1.0
        for event_name, event_data in weather_events.items():
            probability = event_data.get('probability', 0)
            severity = event_data.get('severity', 0)
            
            if month in climate_data.get('risk_months', []):
                event_impact = 1.0 - (probability * severity * 0.5)
                total_impact *= event_impact
        
        return total_impact
    
    def calculate_predictive_climate_score(self, location, contract_months):
        """
        Advanced AI Algorithm 2: Predictive Climate Scoring
        """
        try:
            location_lower = str(location).lower()
            climate_data = self._get_climate_profile(location_lower)
            
            if not contract_months:
                return climate_data['efficiency_factor'] * 100
            
            features = []
            for month in contract_months:
                seasonal_mult = climate_data['seasonal_multipliers'].get(month, 0.90)
                
                risk_indicator = 1.0
                if month in climate_data.get('peak_risk_months', []):
                    risk_indicator = 0.5
                elif month in climate_data.get('risk_months', []):
                    risk_indicator = 0.75
                
                weather_severity = self._calculate_month_weather_severity(climate_data, month)
                optimal_indicator = 1.0 if month in climate_data.get('optimal_operating_window', []) else 0.7
                
                month_score = (
                    seasonal_mult * 0.35 +
                    risk_indicator * 0.25 +
                    weather_severity * 0.20 +
                    optimal_indicator * 0.20
                ) * 100
                
                features.append(month_score)
            
            if len(features) > 1:
                weights = np.linspace(0.8, 1.2, len(features))
                weights = weights / weights.sum() * len(features)
                predictive_score = np.average(features, weights=weights)
            else:
                predictive_score = np.mean(features)
            
            return min(max(predictive_score, 0), 100)
            
        except Exception as e:
            return 75.0
    
    def _calculate_month_weather_severity(self, climate_data, month):
        """Calculate combined weather severity for a month"""
        weather_events = climate_data.get('weather_events', {})
        
        if not weather_events:
            return 1.0
        
        if month not in climate_data.get('risk_months', []):
            return 0.95
        
        total_severity = 0
        total_weight = 0
        
        for event_name, event_data in weather_events.items():
            probability = event_data.get('probability', 0)
            severity = event_data.get('severity', 0)
            
            weight = probability
            impact = 1.0 - (severity * 0.7)
            
            total_severity += impact * weight
            total_weight += weight
        
        if total_weight > 0:
            return total_severity / total_weight
        
        return 0.85
    
    def calculate_adaptive_climate_efficiency(self, location, start_date, end_date, historical_performance=None):
        """
        Advanced AI Algorithm 3: Adaptive Climate Efficiency with Learning
        """
        try:
            location_lower = str(location).lower()
            climate_data = self._get_climate_profile(location_lower)
            
            base_efficiency = self.calculate_time_weighted_climate_efficiency(location, start_date, end_date)
            
            if historical_performance is None or not historical_performance:
                return base_efficiency
            
            hist_mean = np.mean(historical_performance)
            hist_std = np.std(historical_performance) if len(historical_performance) > 1 else 0
            
            if hist_std > 0:
                confidence_factor = 1.0 - (hist_std / 100) * 0.3
            else:
                confidence_factor = 1.0
            
            adaptive_score = (base_efficiency * 0.6 + hist_mean * 0.4) * confidence_factor
            
            return min(max(adaptive_score, 0), 100)
            
        except Exception as e:
            return 75.0
    
    def calculate_risk_adjusted_climate_score(self, location, contract_duration_days, start_month):
        """
        Advanced AI Algorithm 4: Risk-Adjusted Climate Scoring
        """
        try:
            location_lower = str(location).lower()
            climate_data = self._get_climate_profile(location_lower)
            
            contract_months = []
            current_month = start_month
            days_covered = 0
            
            while days_covered < contract_duration_days:
                contract_months.append(current_month)
                days_covered += 30
                current_month = (current_month % 12) + 1
            
            risk_scores = []
            
            for month in set(contract_months):
                base_risk = climate_data['downtime_risk']
                
                if month in climate_data.get('peak_risk_months', []):
                    month_risk = base_risk * 1.5
                elif month in climate_data.get('risk_months', []):
                    month_risk = base_risk * 1.2
                else:
                    month_risk = base_risk * 0.7
                
                weather_events = climate_data.get('weather_events', {})
                event_risk = 0
                
                for event_name, event_data in weather_events.items():
                    if event_name in self.weather_severity_matrix:
                        severity_data = self.weather_severity_matrix[event_name]
                        probability = event_data.get('probability', 0)
                        expected_downtime = (probability * severity_data['downtime_days'] / 30)
                        event_risk += expected_downtime
                
                total_risk = month_risk + (event_risk * 0.5)
                month_efficiency = (1 - min(total_risk, 0.9)) * 100
                
                risk_scores.append(month_efficiency)
            
            duration_factor = 1.0
            if contract_duration_days > 365:
                duration_factor = 0.95
            elif contract_duration_days > 730:
                duration_factor = 0.90
            
            risk_adjusted_score = np.mean(risk_scores) * duration_factor
            
            return min(max(risk_adjusted_score, 0), 100)
            
        except Exception as e:
            return 75.0
    
    def calculate_optimization_score(self, location, start_month, duration_months):
        """
        Advanced AI Algorithm 5: Optimization Score
        """
        try:
            location_lower = str(location).lower()
            climate_data = self._get_climate_profile(location_lower)
            
            optimal_window = climate_data.get('optimal_operating_window', list(range(1, 13)))
            contract_months = [(start_month - 1 + i) % 12 + 1 for i in range(duration_months)]
            
            optimal_months = sum(1 for m in contract_months if m in optimal_window)
            optimization_ratio = optimal_months / len(contract_months) if contract_months else 0
            
            peak_risk_months = climate_data.get('peak_risk_months', [])
            risk_months = climate_data.get('risk_months', [])
            
            peak_exposure = sum(1 for m in contract_months if m in peak_risk_months)
            risk_exposure = sum(1 for m in contract_months if m in risk_months)
            
            base_score = optimization_ratio * 100
            peak_penalty = (peak_exposure / len(contract_months)) * 30 if contract_months else 0
            risk_penalty = (risk_exposure / len(contract_months)) * 15 if contract_months else 0
            
            optimization_score = base_score - peak_penalty - risk_penalty
            
            if peak_exposure == 0 and risk_exposure == 0:
                optimization_score += 10
            
            return min(max(optimization_score, 0), 100)
            
        except Exception as e:
            return 70.0
    
    def calculate_multi_algorithm_climate_score(self, location, start_date, end_date, 
                                                contract_duration_days, historical_performance=None):
        """
        Advanced AI Algorithm 6: Ensemble Multi-Algorithm Climate Score
        """
        try:
            score1 = self.calculate_time_weighted_climate_efficiency(location, start_date, end_date)
            
            if pd.notna(start_date) and pd.notna(end_date):
                date_range = pd.date_range(start=start_date, end=end_date, freq='M')
                contract_months = [d.month for d in date_range]
            else:
                contract_months = []
            score2 = self.calculate_predictive_climate_score(location, contract_months)
            
            score3 = self.calculate_adaptive_climate_efficiency(location, start_date, end_date, historical_performance)
            
            start_month = start_date.month if pd.notna(start_date) else 1
            score4 = self.calculate_risk_adjusted_climate_score(location, contract_duration_days, start_month)
            
            duration_months = int(contract_duration_days / 30) if contract_duration_days > 0 else 6
            score5 = self.calculate_optimization_score(location, start_month, duration_months)
            
            weights = np.array([0.25, 0.20, 0.20, 0.20, 0.15])
            scores = np.array([score1, score2, score3, score4, score5])
            
            ensemble_score = np.average(scores, weights=weights)
            
            score_variance = np.var(scores)
            confidence_penalty = min(score_variance / 500, 5)
            
            final_score = ensemble_score - confidence_penalty
            
            return min(max(final_score, 0), 100)
            
        except Exception as e:
            return 75.0
    
    def _get_climate_profile(self, location_lower):
        """Get climate profile for location"""
        for key in self.climate_profiles.keys():
            if key in location_lower:
                return self.climate_profiles[key]
        return self.climate_profiles['default']
    
    def get_climate_insights(self, location, start_date, end_date):
        """Generate detailed climate insights"""
        try:
            location_lower = str(location).lower()
            climate_data = self._get_climate_profile(location_lower)
            
            insights = {
                'climate_type': climate_data['climate'],
                'description': climate_data.get('description', 'Standard climate conditions'),
                'risk_assessment': {},
                'recommendations': [],
                'optimal_periods': []
            }
            
            if pd.notna(start_date) and pd.notna(end_date):
                date_range = pd.date_range(start=start_date, end=end_date, freq='M')
                contract_months = [d.month for d in date_range]
                
                peak_risk_months = climate_data.get('peak_risk_months', [])
                risk_months = climate_data.get('risk_months', [])
                optimal_window = climate_data.get('optimal_operating_window', [])
                
                peak_exposure = [m for m in contract_months if m in peak_risk_months]
                risk_exposure = [m for m in contract_months if m in risk_months]
                optimal_coverage = [m for m in contract_months if m in optimal_window]
                
                insights['risk_assessment'] = {
                    'peak_risk_exposure': len(peak_exposure),
                    'general_risk_exposure': len(risk_exposure),
                    'optimal_coverage': len(optimal_coverage),
                    'total_months': len(contract_months)
                }
                
                if peak_exposure:
                    month_names = {1:'Jan', 2:'Feb', 3:'Mar', 4:'Apr', 5:'May', 6:'Jun',
                                  7:'Jul', 8:'Aug', 9:'Sep', 10:'Oct', 11:'Nov', 12:'Dec'}
                    peak_months_str = ', '.join([month_names[m] for m in peak_exposure])
                    insights['recommendations'].append(
                        f"HIGH RISK: Contract operates during peak risk months ({peak_months_str}). "
                        f"Consider enhanced weather monitoring and contingency planning."
                    )
                
                if len(optimal_coverage) == len(contract_months):
                    insights['recommendations'].append(
                        "OPTIMAL: Contract timing aligns perfectly with optimal operating window."
                    )
                elif len(optimal_coverage) / len(contract_months) < 0.5:
                    insights['recommendations'].append(
                        "SUBOPTIMAL: Less than 50% of contract period in optimal operating window."
                    )
            
            return insights
            
        except Exception as e:
            return {
                'climate_type': 'unknown',
                'description': 'Unable to analyze climate data',
                'risk_assessment': {},
                'recommendations': ['Climate analysis unavailable'],
                'optimal_periods': []
            }




class RigEfficiencyCalculator:
    """
    Advanced Rig Efficiency Calculator with Multi-Factor Analysis
    Calculates comprehensive efficiency metrics across 6 key factors
    """
    
    def __init__(self):
        self.climate_ai = AdvancedClimateIntelligence()
        self.efficiency_weights = {
            'contract_utilization': 0.25,
            'dayrate_efficiency': 0.20,
            'contract_stability': 0.15,
            'location_complexity': 0.15,
            'climate_impact': 0.10,
            'contract_performance': 0.15
        }
    
    def calculate_comprehensive_efficiency(self, rig_data):
        """Calculate comprehensive efficiency metrics for a rig"""
        if rig_data.empty:
            return None
        
        try:
            metrics = {}
            
            # Calculate all 6 core metrics
            metrics['contract_utilization'] = self._calculate_contract_utilization(rig_data)
            metrics['dayrate_efficiency'] = self._calculate_dayrate_efficiency(rig_data)
            metrics['contract_stability'] = self._calculate_contract_stability(rig_data)
            metrics['location_complexity'] = self._calculate_location_efficiency(rig_data)
            metrics['climate_impact'] = self._calculate_enhanced_climate_efficiency(rig_data)
            metrics['contract_performance'] = self._calculate_contract_performance(rig_data)
            
            # Calculate additional climate insights
            metrics['climate_insights'] = self._get_detailed_climate_insights(rig_data)
            metrics['climate_optimization'] = self._calculate_climate_optimization_score(rig_data)
            
            # Calculate overall weighted score
            overall_score = sum(
                metrics[key] * self.efficiency_weights[key] 
                for key in self.efficiency_weights.keys()
            )
            
            metrics['overall_efficiency'] = overall_score
            metrics['efficiency_grade'] = self._get_efficiency_grade(overall_score)
            metrics['insights'] = self._generate_detailed_insights(rig_data, metrics)
            metrics['improvement_suggestions'] = self._generate_improvement_suggestions(metrics)
            
            return metrics
            
        except Exception as e:
            print(f"Error calculating efficiency: {str(e)}")
            return None
    
    def _calculate_contract_utilization(self, rig_data):
        """Calculate contract utilization rate"""
        try:
            valid_contracts = rig_data[
                rig_data['Contract Start Date'].notna() & 
                rig_data['Contract End Date'].notna()
            ].copy()
            
            if valid_contracts.empty:
                return 50.0
            
            valid_contracts['contract_days'] = (
                valid_contracts['Contract End Date'] - valid_contracts['Contract Start Date']
            ).dt.days
            
            total_contracted_days = valid_contracts['contract_days'].sum()
            earliest_start = valid_contracts['Contract Start Date'].min()
            latest_end = valid_contracts['Contract End Date'].max()
            total_days = (latest_end - earliest_start).days
            
            if total_days <= 0:
                return 50.0
            
            utilization = (total_contracted_days / total_days) * 100
            return min(utilization, 100.0)
            
        except Exception as e:
            return 50.0
    
    def _calculate_dayrate_efficiency(self, rig_data):
        """Calculate dayrate efficiency based on market rates"""
        try:
            valid_rates = rig_data[rig_data['Dayrate ($k)'].notna()]['Dayrate ($k)']
            
            if valid_rates.empty:
                return 50.0
            
            avg_dayrate = valid_rates.mean()
            
            # Scoring based on dayrate tiers
            if avg_dayrate >= 400:
                score = 95 + min((avg_dayrate - 400) / 100, 5)
            elif avg_dayrate >= 250:
                score = 75 + ((avg_dayrate - 250) / 150) * 20
            elif avg_dayrate >= 150:
                score = 55 + ((avg_dayrate - 150) / 100) * 20
            elif avg_dayrate >= 100:
                score = 35 + ((avg_dayrate - 100) / 50) * 20
            else:
                score = max(10, (avg_dayrate / 100) * 35)
            
            return min(score, 100.0)
            
        except Exception as e:
            return 50.0
    
    def _calculate_contract_stability(self, rig_data):
        """Calculate contract stability based on length and count"""
        try:
            valid_contracts = rig_data[
                rig_data['Contract Start Date'].notna() & 
                rig_data['Contract Length'].notna()
            ].copy()
            
            if valid_contracts.empty:
                return 50.0
            
            avg_length = valid_contracts['Contract Length'].mean()
            
            # Length-based scoring
            if avg_length >= 1095:  # 3+ years
                length_score = 100
            elif avg_length >= 730:  # 2+ years
                length_score = 85
            elif avg_length >= 365:  # 1+ year
                length_score = 70
            elif avg_length >= 180:  # 6+ months
                length_score = 55
            else:
                length_score = 40
            
            # Contract count scoring (fewer is better for stability)
            num_contracts = len(valid_contracts)
            if num_contracts == 1:
                contract_count_score = 100
            elif num_contracts <= 3:
                contract_count_score = 85
            elif num_contracts <= 5:
                contract_count_score = 70
            else:
                contract_count_score = 55
            
            stability_score = (length_score * 0.7) + (contract_count_score * 0.3)
            return stability_score
            
        except Exception as e:
            return 50.0
    
    def _calculate_location_efficiency(self, rig_data):
        """Calculate location complexity and efficiency"""
        try:
            locations = rig_data['Current Location'].dropna()
            
            if locations.empty:
                return 70.0
            
            location_scores = []
            for location in locations:
                location_lower = str(location).lower()
                
                # Complexity-based scoring
                if any(term in location_lower for term in ['deepwater', 'deep water', 'ultra-deep']):
                    location_scores.append(65)  # Higher complexity
                elif any(term in location_lower for term in ['offshore', 'shelf']):
                    location_scores.append(75)  # Medium complexity
                elif any(term in location_lower for term in ['onshore', 'land']):
                    location_scores.append(90)  # Lower complexity
                else:
                    location_scores.append(75)  # Default
            
            return np.mean(location_scores) if location_scores else 75.0
            
        except Exception as e:
            return 70.0
    
    def _calculate_enhanced_climate_efficiency(self, rig_data):
        """Calculate climate efficiency using AI ensemble algorithms"""
        try:
            locations = rig_data['Current Location'].dropna()
            start_dates = pd.to_datetime(rig_data['Contract Start Date'], errors='coerce')
            end_dates = pd.to_datetime(rig_data['Contract End Date'], errors='coerce')
            contract_lengths = rig_data['Contract Length'].fillna(0)
            
            if locations.empty:
                return 80.0
            
            climate_scores = []
            
            for idx, (location, start_date, end_date, duration) in enumerate(
                zip(locations, start_dates, end_dates, contract_lengths)
            ):
                # Use basic climate data if dates are missing
                if pd.isna(start_date) or pd.isna(end_date):
                    location_lower = str(location).lower()
                    climate_data = self.climate_ai._get_climate_profile(location_lower)
                    score = climate_data['efficiency_factor'] * 100
                    climate_scores.append(score)
                    continue
                
                # Use AI ensemble for complete data
                contract_duration_days = duration if duration > 0 else (end_date - start_date).days
                
                ensemble_score = self.climate_ai.calculate_multi_algorithm_climate_score(
                    location=location,
                    start_date=start_date,
                    end_date=end_date,
                    contract_duration_days=contract_duration_days,
                    historical_performance=None
                )
                
                climate_scores.append(ensemble_score)
            
            # Weight by contract duration if available
            if climate_scores:
                if contract_lengths.sum() > 0:
                    weights = contract_lengths / contract_lengths.sum()
                    final_score = np.average(climate_scores, weights=weights)
                else:
                    final_score = np.mean(climate_scores)
            else:
                final_score = 80.0
            
            return min(max(final_score, 0), 100)
            
        except Exception as e:
            return 80.0
    
    def _calculate_climate_optimization_score(self, rig_data):
        """Calculate how well contracts are timed for climate conditions"""
        try:
            locations = rig_data['Current Location'].dropna()
            start_dates = pd.to_datetime(rig_data['Contract Start Date'], errors='coerce')
            contract_lengths = rig_data['Contract Length'].fillna(180)
            
            if locations.empty or start_dates.isna().all():
                return 70.0
            
            optimization_scores = []
            
            for location, start_date, duration in zip(locations, start_dates, contract_lengths):
                if pd.isna(start_date):
                    continue
                
                duration_months = int(duration / 30) if duration > 0 else 6
                
                opt_score = self.climate_ai.calculate_optimization_score(
                    location=location,
                    start_month=start_date.month,
                    duration_months=duration_months
                )
                
                optimization_scores.append(opt_score)
            
            return np.mean(optimization_scores) if optimization_scores else 70.0
            
        except Exception as e:
            return 70.0
    
    def _get_detailed_climate_insights(self, rig_data):
        """Get detailed climate insights for all contracts"""
        try:
            locations = rig_data['Current Location'].dropna()
            start_dates = pd.to_datetime(rig_data['Contract Start Date'], errors='coerce')
            end_dates = pd.to_datetime(rig_data['Contract End Date'], errors='coerce')
            
            all_insights = []
            
            for location, start_date, end_date in zip(locations, start_dates, end_dates):
                if pd.notna(start_date) and pd.notna(end_date):
                    insights = self.climate_ai.get_climate_insights(location, start_date, end_date)
                    insights['location'] = location
                    insights['contract_period'] = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
                    all_insights.append(insights)
            
            return all_insights
            
        except Exception as e:
            return []
    
    def _calculate_contract_performance(self, rig_data):
        """Calculate overall contract performance"""
        try:
            # Check contract status
            status_col = rig_data['Status'].dropna() if 'Status' in rig_data.columns else pd.Series()
            
            if not status_col.empty:
                status_lower = status_col.str.lower()
                active_count = status_lower.str.contains('active|operating', case=False, na=False).sum()
                completed_count = status_lower.str.contains('complete|finished', case=False, na=False).sum()
                total_count = len(status_lower)
                
                if total_count > 0:
                    performance_rate = ((active_count + completed_count) / total_count) * 100
                else:
                    performance_rate = 70.0
            else:
                performance_rate = 70.0
            
            # Factor in contract value
            if 'Contract value ($m)' in rig_data.columns:
                contract_values = rig_data['Contract value ($m)'].dropna()
                if not contract_values.empty:
                    avg_value = contract_values.mean()
                    if avg_value >= 100:
                        value_score = 95
                    elif avg_value >= 50:
                        value_score = 80
                    elif avg_value >= 20:
                        value_score = 65
                    else:
                        value_score = 50
                    
                    performance_rate = (performance_rate * 0.6) + (value_score * 0.4)
            
            return min(performance_rate, 100.0)
            
        except Exception as e:
            return 70.0
    
    def _get_efficiency_grade(self, score):
        """Convert efficiency score to letter grade"""
        if score >= 90:
            return 'A (Excellent)'
        elif score >= 80:
            return 'B (Good)'
        elif score >= 70:
            return 'C (Satisfactory)'
        elif score >= 60:
            return 'D (Fair)'
        else:
            return 'F (Needs Improvement)'
    
    def _generate_detailed_insights(self, rig_data, metrics):
        """Generate detailed insights based on metrics"""
        insights = []
        
        # Overall performance insight
        overall = metrics['overall_efficiency']
        if overall >= 85:
            insights.append({
                'type': 'success',
                'category': 'Overall Performance',
                'message': f'Excellent rig performance with {overall:.1f}% efficiency score.',
                'recommendation': 'Maintain current operational practices and consider this as a benchmark.',
                'priority': 'low'
            })
        elif overall >= 70:
            insights.append({
                'type': 'info',
                'category': 'Overall Performance',
                'message': f'Good rig performance with {overall:.1f}% efficiency score.',
                'recommendation': 'Focus on incremental improvements in lower-scoring areas.',
                'priority': 'medium'
            })
        else:
            insights.append({
                'type': 'warning',
                'category': 'Overall Performance',
                'message': f'Below-average performance at {overall:.1f}% efficiency.',
                'recommendation': 'Conduct comprehensive performance review and address key weaknesses.',
                'priority': 'high'
            })
        
        # Contract utilization insights
        util = metrics['contract_utilization']
        if util < 70:
            insights.append({
                'type': 'warning',
                'category': 'Contract Utilization',
                'message': f'Low utilization at {util:.1f}%. Significant idle time detected.',
                'recommendation': 'Focus on securing back-to-back contracts and reducing gaps between contracts.',
                'priority': 'high'
            })
        elif util < 85:
            insights.append({
                'type': 'info',
                'category': 'Contract Utilization',
                'message': f'Moderate utilization at {util:.1f}%.',
                'recommendation': 'Optimize contract scheduling to improve utilization rate.',
                'priority': 'medium'
            })
        
        # Dayrate insights
        dayrate = metrics['dayrate_efficiency']
        if dayrate < 60:
            insights.append({
                'type': 'warning',
                'category': 'Dayrate Efficiency',
                'message': f'Below-market dayrates ({dayrate:.1f}%).',
                'recommendation': 'Review market rates and consider contract renegotiation or repositioning.',
                'priority': 'high'
            })
        
        # Climate insights
        climate = metrics['climate_impact']
        if climate < 75:
            insights.append({
                'type': 'warning',
                'category': 'Climate Impact',
                'message': f'Challenging climate conditions affecting efficiency ({climate:.1f}%).',
                'recommendation': 'Consider seasonal scheduling optimization and weather contingency planning.',
                'priority': 'medium'
            })
        
        # Contract stability insights
        stability = metrics['contract_stability']
        if stability < 60:
            insights.append({
                'type': 'warning',
                'category': 'Contract Stability',
                'message': f'Low contract stability ({stability:.1f}%).',
                'recommendation': 'Focus on securing longer-term contracts to improve stability.',
                'priority': 'high'
            })
        
        return insights
    
    def _generate_improvement_suggestions(self, metrics):
        """Generate prioritized improvement suggestions"""
        suggestions = []
        
        # Create list of metrics with scores
        metric_scores = [
            ('Contract Utilization', metrics['contract_utilization']),
            ('Dayrate Efficiency', metrics['dayrate_efficiency']),
            ('Contract Stability', metrics['contract_stability']),
            ('Location Complexity', metrics['location_complexity']),
            ('Climate Impact', metrics['climate_impact']),
            ('Contract Performance', metrics['contract_performance'])
        ]
        
        # Sort by score (lowest first = highest priority)
        metric_scores.sort(key=lambda x: x[1])
        
        # Generate suggestions for lowest scoring areas
        for metric_name, score in metric_scores[:3]:  # Top 3 priorities
            if score < 70:
                priority = 'high'
                action = 'immediate'
            elif score < 85:
                priority = 'medium'
                action = 'short-term'
            else:
                continue
            
            suggestions.append({
                'metric': metric_name,
                'current_score': score,
                'priority': priority,
                'action_timeframe': action,
                'potential_impact': 'high' if score < 60 else 'medium'
            })
        
        return suggestions
    
    def generate_contract_summary(self, rig_data, metrics):
        """Generate comprehensive contract summary"""
        try:
            summary = {
                'rig_name': rig_data['Rig Name'].iloc[0] if 'Rig Name' in rig_data.columns else 'Unknown',
                'total_contracts': len(rig_data),
                'active_contracts': len(rig_data[rig_data['Status'].str.contains('active', case=False, na=False)]) if 'Status' in rig_data.columns else 0,
                'total_contract_value': rig_data['Contract value ($m)'].sum() if 'Contract value ($m)' in rig_data.columns else 0,
                'average_dayrate': rig_data['Dayrate ($k)'].mean() if 'Dayrate ($k)' in rig_data.columns else 0,
                'efficiency_grade': metrics['efficiency_grade'],
                'overall_score': metrics['overall_efficiency'],
                'top_strength': self._identify_top_strength(metrics),
                'primary_concern': self._identify_primary_concern(metrics)
            }
            
            return summary
            
        except Exception as e:
            return None
    
    def _identify_top_strength(self, metrics):
        """Identify the highest performing metric"""
        metric_scores = {
            'Contract Utilization': metrics['contract_utilization'],
            'Dayrate Efficiency': metrics['dayrate_efficiency'],
            'Contract Stability': metrics['contract_stability'],
            'Location Efficiency': metrics['location_complexity'],
            'Climate Management': metrics['climate_impact'],
            'Contract Performance': metrics['contract_performance']
        }
        
        top_metric = max(metric_scores.items(), key=lambda x: x[1])
        return f"{top_metric[0]} ({top_metric[1]:.1f}%)"
    
    def _identify_primary_concern(self, metrics):
        """Identify the lowest performing metric"""
        metric_scores = {
            'Contract Utilization': metrics['contract_utilization'],
            'Dayrate Efficiency': metrics['dayrate_efficiency'],
            'Contract Stability': metrics['contract_stability'],
            'Location Efficiency': metrics['location_complexity'],
            'Climate Management': metrics['climate_impact'],
            'Contract Performance': metrics['contract_performance']
        }
        
        lowest_metric = min(metric_scores.items(), key=lambda x: x[1])
        return f"{lowest_metric[0]} ({lowest_metric[1]:.1f}%)"
    
    def compare_rigs(self, rig_data_list):
        """Compare efficiency across multiple rigs"""
        comparisons = []
        
        for rig_data in rig_data_list:
            metrics = self.calculate_comprehensive_efficiency(rig_data)
            if metrics:
                rig_name = rig_data['Rig Name'].iloc[0] if 'Rig Name' in rig_data.columns else 'Unknown'
                comparisons.append({
                    'rig_name': rig_name,
                    'overall_efficiency': metrics['overall_efficiency'],
                    'grade': metrics['efficiency_grade'],
                    'metrics': {
                        'utilization': metrics['contract_utilization'],
                        'dayrate': metrics['dayrate_efficiency'],
                        'stability': metrics['contract_stability'],
                        'location': metrics['location_complexity'],
                        'climate': metrics['climate_impact'],
                        'performance': metrics['contract_performance']
                    }
                })
        
        # Sort by overall efficiency
        comparisons.sort(key=lambda x: x['overall_efficiency'], reverse=True)
        
        return comparisons


# Export main classes
__all__ = ['RigEfficiencyCalculator', 'AdvancedClimateIntelligence', 'preprocess_dataframe']