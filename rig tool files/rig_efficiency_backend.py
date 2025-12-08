"""
Rig Efficiency Analysis Backend - COMPLETE VERSION
Core calculation and AI logic with all components
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from sklearn.preprocessing import MinMaxScaler
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




class RegionalBenchmarkModel:
    """
    Expected Performance Models for each region + geology type
    Provides normalized benchmarks for fair comparison
    """
    def __init__(self):
        self.benchmarks = self._initialize_benchmarks()
    
    def _initialize_benchmarks(self):
        """Initialize regional and geological benchmarks"""
        return {
            'offshore': {'expected_rop': 35, 'expected_npt': 15, 'expected_days_per_well': 45, 'cost_per_meter': 800, 'dayrate_benchmark': 300},
            'onshore': {'expected_rop': 55, 'expected_npt': 8, 'expected_days_per_well': 25, 'cost_per_meter': 350, 'dayrate_benchmark': 150},
            'hpht': {'expected_rop': 20, 'expected_npt': 25, 'expected_days_per_well': 65, 'cost_per_meter': 1200, 'difficulty_multiplier': 1.8},
            'hard_formation': {'expected_rop': 25, 'expected_npt': 18, 'expected_days_per_well': 55, 'cost_per_meter': 950, 'difficulty_multiplier': 1.5},
            'soft_formation': {'expected_rop': 65, 'expected_npt': 7, 'expected_days_per_well': 20, 'cost_per_meter': 400, 'difficulty_multiplier': 0.8},
            'arctic': {'expected_rop': 30, 'expected_npt': 22, 'expected_days_per_well': 60, 'cost_per_meter': 1100, 'difficulty_multiplier': 1.7},
            'desert': {'expected_rop': 50, 'expected_npt': 9, 'expected_days_per_well': 28, 'cost_per_meter': 500, 'difficulty_multiplier': 0.9},
            'tropical': {'expected_rop': 40, 'expected_npt': 14, 'expected_days_per_well': 40, 'cost_per_meter': 700, 'difficulty_multiplier': 1.2},
            'deepwater': {'expected_rop': 30, 'expected_npt': 18, 'expected_days_per_well': 55, 'cost_per_meter': 1000, 'difficulty_multiplier': 1.6},
            'ultra_deepwater': {'expected_rop': 25, 'expected_npt': 22, 'expected_days_per_well': 70, 'cost_per_meter': 1400, 'difficulty_multiplier': 2.0},
            'shallow_water': {'expected_rop': 45, 'expected_npt': 10, 'expected_days_per_well': 30, 'cost_per_meter': 600, 'difficulty_multiplier': 1.0}
        }
    
    def get_benchmark(self, rig_data):
        """Get appropriate benchmark for rig"""
        location = str(rig_data['Current Location'].iloc[0]).lower() if 'Current Location' in rig_data.columns and len(rig_data) > 0 else ''
        categories = []
        
        if any(term in location for term in ['offshore', 'sea', 'platform']):
            categories.append('offshore')
        elif any(term in location for term in ['onshore', 'land']):
            categories.append('onshore')
        
        if any(term in location for term in ['deepwater', 'deep water']):
            categories.append('ultra_deepwater' if 'ultra' in location else 'deepwater')
        
        if any(term in location for term in ['arctic', 'north sea', 'norway']):
            categories.append('arctic')
        elif any(term in location for term in ['middle east', 'saudi', 'uae', 'qatar']):
            categories.append('desert')
        elif any(term in location for term in ['gulf of mexico', 'brazil', 'indonesia']):
            categories.append('tropical')
        
        if not categories:
            categories = ['offshore']
        
        combined_benchmark = {k: 0 for k in ['expected_rop', 'expected_npt', 'expected_days_per_well', 'cost_per_meter', 'difficulty_multiplier']}
        combined_benchmark['difficulty_multiplier'] = 1.0
        
        for cat in categories:
            if cat in self.benchmarks:
                bench = self.benchmarks[cat]
                for key in combined_benchmark:
                    if key == 'difficulty_multiplier':
                        combined_benchmark[key] *= bench.get(key, 1.0)
                    else:
                        combined_benchmark[key] += bench.get(key, 0)
        
        count = len(categories)
        for key in combined_benchmark:
            if key != 'difficulty_multiplier' and count > 0:
                combined_benchmark[key] /= count
        
        combined_benchmark['categories'] = categories
        return combined_benchmark
    
    def calculate_normalized_performance(self, rig_data, actual_metrics=None):
        """
        Calculate performance normalized against benchmark
        
        Parameters:
        - rig_data: DataFrame with contract data
        - actual_metrics: dict with drilling metrics (optional)
                         If None, will generate estimates from available contract data
        
        If actual_metrics not provided, generates reasonable estimates using:
        - Dayrate → Equipment quality → ROP and NPT estimates
        - Contract length → Days per well estimate
        - Industry correlations for cost per meter
        """
        benchmark = self.get_benchmark(rig_data)
        
        # === Generate actual_metrics if not provided ===
        if actual_metrics is None or not isinstance(actual_metrics, dict):
            # Extract available contract-level data
            dayrate = rig_data['Dayrate ($k)'].mean() if 'Dayrate ($k)' in rig_data.columns else 200
            contract_length = rig_data['Contract Length'].mean() if 'Contract Length' in rig_data.columns else 180
            
            # Generate synthetic drilling metrics using industry correlations
            
            # 1. ROP (Rate of Penetration)
            # Higher dayrate = better equipment = faster drilling
            # Base: 35 m/hr (industry average)
            # Multiplier range: 0.7x to 1.5x (70% to 150% of base)
            base_rop = 35
            rop_multiplier = min(1.5, max(0.7, dayrate / 200))
            estimated_rop = base_rop * rop_multiplier
            
            # 2. NPT (Non-Productive Time)
            # Higher dayrate = better equipment = less downtime (inverse relationship)
            # Base: 15% (industry average)
            # Multiplier range: 0.7x to 1.5x
            base_npt = 15
            npt_multiplier = min(1.5, max(0.7, 200 / dayrate))  # Inverse
            estimated_npt = base_npt * npt_multiplier
            
            # 3. Days Per Well
            # Estimate from contract length
            # Assumption: 3-4 wells per contract on average
            estimated_days_per_well = contract_length / 3.5
            
            # 4. Cost Per Meter
            # Direct correlation with dayrate
            # Industry rule of thumb: cost_per_meter ≈ dayrate × 5
            estimated_cost_per_meter = dayrate * 5
            
            # Build synthetic metrics dict
            actual_metrics = {
                'rop': estimated_rop,
                'npt': estimated_npt,
                'days_per_well': estimated_days_per_well,
                'cost_per_meter': estimated_cost_per_meter
            }
        
        # === Existing calculation logic (unchanged) ===
        normalized = {}
        
        if 'rop' in actual_metrics:
            normalized['rop_performance'] = (actual_metrics['rop'] / benchmark['expected_rop'] * 100) if benchmark['expected_rop'] > 0 else 100
        if 'npt' in actual_metrics:
            normalized['npt_performance'] = (benchmark['expected_npt'] / actual_metrics['npt'] * 100) if actual_metrics['npt'] > 0 else 100
        if 'days_per_well' in actual_metrics:
            normalized['time_performance'] = (benchmark['expected_days_per_well'] / actual_metrics['days_per_well'] * 100) if actual_metrics['days_per_well'] > 0 else 100
        if 'cost_per_meter' in actual_metrics:
            normalized['cost_performance'] = (benchmark['cost_per_meter'] / actual_metrics['cost_per_meter'] * 100) if actual_metrics['cost_per_meter'] > 0 else 100
        
        normalized['overall_normalized'] = np.mean([v for v in normalized.values()])
        normalized['benchmark_used'] = benchmark['categories']
        normalized['difficulty_multiplier'] = benchmark['difficulty_multiplier']
        return normalized


class RigWellMatchPredictor:
    """ML Engine for Rig-Well Matching - Predicts execution time, AFE probability, NPT%, risk score, recommended dayrate"""
    
    def __init__(self):
        self.models = {}
        self.feature_scaler = MinMaxScaler()
        self.is_trained = False
        
    def prepare_features(self, rig_data, well_params=None):
        """Prepare feature vector for ML prediction"""
        features = {}
        
        if 'Dayrate ($k)' in rig_data.columns:
            features['avg_dayrate'] = rig_data['Dayrate ($k)'].mean()
        else:
            features['avg_dayrate'] = 200
        
        if 'Contract Length' in rig_data.columns:
            features['avg_contract_length'] = rig_data['Contract Length'].mean()
        else:
            features['avg_contract_length'] = 180
        
        features['region_complexity'] = self._encode_region_complexity(rig_data)
        features['climate_score'] = self._get_climate_score(rig_data)
        
        if 'Water Depth' in rig_data.columns:
            features['water_depth'] = rig_data['Water Depth'].mean()
        else:
            features['water_depth'] = 500
        
        features['contract_success_rate'] = self._calculate_success_rate(rig_data)
        features['utilization_rate'] = self._calculate_utilization(rig_data)
        
        if well_params:
            features['target_depth'] = well_params.get('depth', 3000)
            features['formation_hardness'] = well_params.get('hardness', 5)
            features['temperature'] = well_params.get('temperature', 150)
            features['pressure'] = well_params.get('pressure', 5000)
        else:
            features['target_depth'] = 3000
            features['formation_hardness'] = 5
            features['temperature'] = 150
            features['pressure'] = 5000
        
        return features
    
    def _encode_region_complexity(self, rig_data):
        """Encode region complexity as numeric value"""
        if 'Current Location' not in rig_data.columns:
            return 5
        
        location = str(rig_data['Current Location'].iloc[0]).lower() if len(rig_data) > 0 else ''
        complexity_map = {'ultra-deep': 10, 'deepwater': 8, 'hpht': 9, 'arctic': 8, 'north sea': 7, 
                         'gulf of mexico': 6, 'offshore': 5, 'onshore': 3, 'middle east': 2}
        
        for key, value in complexity_map.items():
            if key in location:
                return value
        return 5
    
    def _get_climate_score(self, rig_data):
        """Get simplified climate score (10=best, 1=worst)"""
        if 'Current Location' not in rig_data.columns:
            return 7
        
        location = str(rig_data['Current Location'].iloc[0]).lower() if len(rig_data) > 0 else ''
        climate_map = {'middle east': 9, 'saudi': 9, 'uae': 9, 'qatar': 9, 'brazil': 7, 
                      'gulf of mexico': 5, 'north sea': 3, 'arctic': 2, 'norway': 3}
        
        for key, value in climate_map.items():
            if key in location:
                return value
        return 7
    
    def _calculate_success_rate(self, rig_data):
        """Calculate historical success rate (0-10)"""
        if 'Status' in rig_data.columns:
            status_col = rig_data['Status'].dropna()
            if not status_col.empty:
                status_lower = status_col.str.lower()
                successful = status_lower.str.contains('complete|active|operating', case=False, na=False).sum()
                total = len(status_lower)
                return (successful / total * 10) if total > 0 else 7
        return 7
    
    def _calculate_utilization(self, rig_data):
        """Calculate utilization score (0-10)"""
        if 'Contract Start Date' not in rig_data.columns or 'Contract End Date' not in rig_data.columns:
            return 7
        
        valid_contracts = rig_data[rig_data['Contract Start Date'].notna() & rig_data['Contract End Date'].notna()].copy()
        if valid_contracts.empty:
            return 7
        
        valid_contracts['Contract Start Date'] = pd.to_datetime(valid_contracts['Contract Start Date'], errors='coerce')
        valid_contracts['Contract End Date'] = pd.to_datetime(valid_contracts['Contract End Date'], errors='coerce')
        valid_contracts['contract_days'] = (valid_contracts['Contract End Date'] - valid_contracts['Contract Start Date']).dt.days
        
        total_contracted_days = valid_contracts['contract_days'].sum()
        earliest_start = valid_contracts['Contract Start Date'].min()
        latest_end = valid_contracts['Contract End Date'].max()
        total_days = (latest_end - earliest_start).days
        
        if total_days <= 0:
            return 7
        
        utilization = (total_contracted_days / total_days) * 10
        return min(utilization, 10)
    
    def predict_well_execution(self, rig_data, well_params=None):
        """Predict well execution outcomes using ML approach"""
        features = self.prepare_features(rig_data, well_params)
        predictions = {}
        
        base_time = features['target_depth'] / 100
        complexity_multiplier = 1 + (features['region_complexity'] / 20)
        climate_multiplier = 1 + ((10 - features['climate_score']) / 20)
        formation_multiplier = 1 + (features['formation_hardness'] / 20)
        capability_factor = features['avg_dayrate'] / 200
        capability_multiplier = 1 / (0.8 + capability_factor * 0.4)
        experience_multiplier = 1 / (0.8 + features['contract_success_rate'] / 25)
        
        expected_time = (base_time * complexity_multiplier * climate_multiplier * formation_multiplier * 
                        capability_multiplier * experience_multiplier)
        predictions['expected_time_days'] = round(expected_time, 1)
        
        base_afe_prob = 70
        capability_bonus = (capability_factor - 1) * 20
        experience_bonus = (features['contract_success_rate'] - 7) * 3
        climate_bonus = (features['climate_score'] - 7) * 2
        complexity_penalty = (features['region_complexity'] - 5) * 2
        afe_probability = base_afe_prob + capability_bonus + experience_bonus + climate_bonus - complexity_penalty
        predictions['afe_probability'] = max(30, min(95, afe_probability))
        
        base_npt = 12
        complexity_npt = (features['region_complexity'] - 5) * 1.5
        climate_npt = (10 - features['climate_score']) * 1.2
        formation_npt = (features['formation_hardness'] - 5) * 0.8
        capability_npt_reduction = (capability_factor - 1) * 3
        experience_npt_reduction = (features['contract_success_rate'] - 7) * 0.8
        expected_npt = base_npt + complexity_npt + climate_npt + formation_npt - capability_npt_reduction - experience_npt_reduction
        predictions['expected_npt_percent'] = max(3, min(30, expected_npt))
        
        risk_components = {
            'complexity_risk': features['region_complexity'] * 5,
            'climate_risk': (10 - features['climate_score']) * 5,
            'formation_risk': features['formation_hardness'] * 4,
            'capability_risk': max(0, (5 - capability_factor) * 10),
            'experience_risk': max(0, (7 - features['contract_success_rate']) * 5)
        }
        total_risk = sum(risk_components.values())
        predictions['risk_score'] = min(100, total_risk)
        predictions['risk_breakdown'] = risk_components
        
        market_base = 200
        complexity_premium = features['region_complexity'] * 15
        formation_premium = features['formation_hardness'] * 10
        climate_adjustment = (10 - features['climate_score']) * 8
        risk_premium = (predictions['risk_score'] / 100) * 50
        
        recommended_dayrate_low = market_base + complexity_premium + formation_premium
        recommended_dayrate_high = recommended_dayrate_low + climate_adjustment + risk_premium
        predictions['recommended_dayrate_range'] = {
            'low': round(recommended_dayrate_low, 0),
            'high': round(recommended_dayrate_high, 0),
            'optimal': round((recommended_dayrate_low + recommended_dayrate_high) / 2, 0)
        }
        
        data_quality_score = 85
        if len(rig_data) < 3:
            data_quality_score -= 15
        elif len(rig_data) < 5:
            data_quality_score -= 8
        if features['region_complexity'] >= 9:
            data_quality_score -= 10
        predictions['confidence_percent'] = max(50, min(95, data_quality_score))
        
        match_score = np.mean([
            self._calculate_capability_match(features),
            features['contract_success_rate'] * 10,
            features['climate_score'] * 10,
            max(0, 100 - (features['region_complexity'] - 5) * 10),
            100 - predictions['risk_score']
        ])
        predictions['match_score'] = round(match_score, 1)
        return predictions
    
    def _calculate_capability_match(self, features):
        """Calculate how well rig capability matches well requirements"""
        ideal_dayrate = 150 + (features['region_complexity'] * 20)
        difference = abs(features['avg_dayrate'] - ideal_dayrate)
        match_score = max(0, 100 - (difference / ideal_dayrate * 100))
        return match_score


class MonteCarloScenarioSimulator:
    """Monte Carlo Simulation for What-If Scenarios"""
    
    def __init__(self, num_simulations=1000):
        self.num_simulations = num_simulations
        self.random_state = np.random.RandomState(42)
    
    def _normalize_params(self, params):
        """
        Normalize and validate basin parameters with comprehensive fallbacks.
        Handles parameter name variations, type conversions, and defaults.
        """
        if not params:
            params = {}
        
        # Define parameter mappings and defaults
        param_mappings = {
            'basin_name': ['basin_name', 'basin'],
            'climate_severity': ['climate_severity', 'climate', 'climate_score'],
            'geology_difficulty': ['geology_difficulty', 'geology', 'difficulty'],
            'water_depth': ['water_depth', 'water_depth_ft', 'depth_ft', 'depth'],
            'typical_dayrate': ['typical_dayrate', 'typical_dayrate_k', 'dayrate', 'rate']
        }
        
        defaults = {
            'basin_name': 'Unknown Basin',
            'climate_severity': 5.0,
            'geology_difficulty': 5.0,
            'water_depth': 2000.0,
            'typical_dayrate': 250.0
        }
        
        normalized = {}
        
        for standard_name, aliases in param_mappings.items():
            value = None
            
            # Try to find parameter by any of its aliases
            for alias in aliases:
                if alias in params:
                    value = params[alias]
                    break
            
            # If not found, use default
            if value is None:
                normalized[standard_name] = defaults[standard_name]
                continue
            
            # Type conversion and validation
            try:
                if standard_name == 'basin_name':
                    normalized[standard_name] = str(value) if value else defaults[standard_name]
                else:
                    # Numeric parameters: convert and clamp
                    num_value = float(value)
                    
                    if standard_name == 'water_depth':
                        # Ensure positive depth; convert from feet to meters if > 100
                        num_value = max(100, abs(num_value))  # Min 100m
                        if num_value > 100:  # Likely in feet, convert to meters
                            num_value = num_value / 3.28084 if num_value > 1000 else num_value
                        normalized[standard_name] = min(num_value, 6000)  # Cap at 6000m
                    else:
                        # Climate/geology: 0-10 scale
                        normalized[standard_name] = max(0, min(10, num_value))
                        
            except (ValueError, TypeError):
                print(f"Warning: Invalid value for {standard_name}: {value}. Using default.")
                normalized[standard_name] = defaults[standard_name]
        
        return normalized
    
    def simulate_basin_transfer(self, rig_data, target_basin_params):
        """Simulate rig performance if moved to different basin - ROBUST VERSION with safe type conversion"""
        
        # === STEP 0: HANDLE EMPTY DATA ===
        if rig_data is None or (hasattr(rig_data, 'empty') and rig_data.empty):
            return {
                'status': 'error',
                'message': 'Empty rig data provided',
                'basin_name': 'Unknown Basin',
                'npt': {'mean': 0, 'std': 0, 'p10': 0, 'p50': 0, 'p90': 0},
                'duration': {'mean': 0, 'std': 0, 'p10': 0, 'p50': 0, 'p90': 0},
                'cost': {'mean': 0, 'std': 0, 'p10': 0, 'p50': 0, 'p90': 0},
                'risk': {'mean': 0, 'std': 0, 'p10': 0, 'p50': 0, 'p90': 0},
                'num_simulations': 0
            }
        
        # === STEP 0B: HANDLE NULL PARAMS ===
        if target_basin_params is None:
            target_basin_params = {}
        
        # === STEP 1: TYPE-SAFE PARAMETER EXTRACTION ===
        def safe_float(value, default):
            """Safely convert to float with comprehensive error handling"""
            try:
                if value is None:
                    return float(default)
                if isinstance(value, str):
                    value = value.strip()
                    if not value:
                        return float(default)
                return float(value)
            except (ValueError, TypeError, AttributeError) as e:
                return float(default)
        
        # Extract and convert parameters with safe type conversion
        climate_severity = safe_float(
            target_basin_params.get('climate_severity', 0.5) if isinstance(target_basin_params, dict) else 0.5,
            0.5
        )
        
        geology_difficulty = safe_float(
            target_basin_params.get('geology_difficulty', 0.5) if isinstance(target_basin_params, dict) else 0.5,
            0.5
        )
        
        water_depth = safe_float(
            target_basin_params.get('water_depth_ft', 5000) if isinstance(target_basin_params, dict) else 5000,
            5000
        )
        
        typical_dayrate = safe_float(
            target_basin_params.get('typical_dayrate_k', 300) if isinstance(target_basin_params, dict) else 300,
            300
        )
        
        basin_name = str(target_basin_params.get('basin_name', 'Unknown Basin')) if isinstance(target_basin_params, dict) else 'Unknown Basin'
        
        # === STEP 2: CLAMP TO VALID RANGES ===
        climate_severity = max(0.0, min(1.0, climate_severity))
        geology_difficulty = max(0.0, min(1.0, geology_difficulty))
        water_depth = max(0.0, water_depth)
        typical_dayrate = max(50.0, min(1000.0, typical_dayrate))
        
        # === STEP 3: EXTRACT BASELINE ===
        baseline = self._extract_baseline_performance(rig_data)
        
        # === STEP 4: RUN SIMULATIONS ===
        npt_results, duration_results, cost_results, risk_results = [], [], [], []
        
        for i in range(self.num_simulations):
            try:
                npt = self._simulate_npt(baseline['avg_npt'], climate_severity, geology_difficulty)
                npt_results.append(npt)
                
                duration = self._simulate_duration(baseline['avg_duration'], climate_severity,
                                                 geology_difficulty, water_depth)
                duration_results.append(duration)
                
                cost = self._simulate_cost(duration, typical_dayrate, npt)
                cost_results.append(cost)
                
                risk = self._simulate_risk(npt, duration, {
                    'climate_severity': climate_severity,
                    'geology_difficulty': geology_difficulty
                })
                risk_results.append(risk)
            except Exception as e:
                continue
        
        # === STEP 5: BUILD RESULTS ===
        results = {
            'status': 'success' if npt_results else 'error',
            'basin_name': basin_name,
            'npt': {
                'mean': float(np.mean(npt_results)) if npt_results else 0,
                'std': float(np.std(npt_results)) if npt_results else 0,
                'p10': float(np.percentile(npt_results, 10)) if npt_results else 0,
                'p50': float(np.percentile(npt_results, 50)) if npt_results else 0,
                'p90': float(np.percentile(npt_results, 90)) if npt_results else 0,
                'distribution': [float(x) for x in npt_results]
            },
            'duration': {
                'mean': float(np.mean(duration_results)) if duration_results else 0,
                'std': float(np.std(duration_results)) if duration_results else 0,
                'p10': float(np.percentile(duration_results, 10)) if duration_results else 0,
                'p50': float(np.percentile(duration_results, 50)) if duration_results else 0,
                'p90': float(np.percentile(duration_results, 90)) if duration_results else 0,
                'distribution': [float(x) for x in duration_results]
            },
            'cost': {
                'mean': float(np.mean(cost_results)) if cost_results else 0,
                'std': float(np.std(cost_results)) if cost_results else 0,
                'p10': float(np.percentile(cost_results, 10)) if cost_results else 0,
                'p50': float(np.percentile(cost_results, 50)) if cost_results else 0,
                'p90': float(np.percentile(cost_results, 90)) if cost_results else 0,
                'distribution': [float(x) for x in cost_results]
            },
            'risk': {
                'mean': float(np.mean(risk_results)) if risk_results else 0,
                'std': float(np.std(risk_results)) if risk_results else 0,
                'p10': float(np.percentile(risk_results, 10)) if risk_results else 0,
                'p50': float(np.percentile(risk_results, 50)) if risk_results else 0,
                'p90': float(np.percentile(risk_results, 90)) if risk_results else 0,
                'distribution': [float(x) for x in risk_results]
            },
            'num_simulations': len(npt_results),
            'parameters_used': {
                'climate_severity': climate_severity,
                'geology_difficulty': geology_difficulty,
                'water_depth': water_depth,
                'typical_dayrate': typical_dayrate,
                'basin_name': basin_name
            }
        }
        
        return results
    
    def _extract_baseline_performance(self, rig_data):
        """Extract baseline performance metrics"""
        baseline = {'avg_npt': 12, 'avg_duration': 40, 'avg_dayrate': 200}
        if 'Contract Length' in rig_data.columns:
            baseline['avg_duration'] = rig_data['Contract Length'].mean() / 3
        if 'Dayrate ($k)' in rig_data.columns:
            baseline['avg_dayrate'] = rig_data['Dayrate ($k)'].mean()
        return baseline
    
    def _simulate_npt(self, baseline_npt, climate_severity, geology_difficulty):
        """Simulate NPT with variability"""
        climate_impact = self.random_state.normal(climate_severity * 0.8, climate_severity * 0.3)
        geology_impact = self.random_state.normal(geology_difficulty * 0.6, geology_difficulty * 0.2)
        random_factor = self.random_state.normal(1.0, 0.15)
        npt = (baseline_npt + climate_impact + geology_impact) * random_factor
        return max(2, min(40, npt))
    
    def _simulate_duration(self, baseline_duration, climate_severity, geology_difficulty, water_depth):
        """Simulate well duration"""
        climate_delay = self.random_state.normal(climate_severity * 1.5, climate_severity * 0.5)
        geology_time = self.random_state.normal(geology_difficulty * 1.2, geology_difficulty * 0.4)
        depth_factor = 1 + (water_depth / 2000)
        random_factor = self.random_state.normal(1.0, 0.2)
        duration = (baseline_duration + climate_delay + geology_time) * depth_factor * random_factor
        return max(15, min(120, duration))
    
    def _simulate_cost(self, duration, dayrate, npt_percent):
        """Simulate total cost"""
        operating_days = duration
        npt_cost_multiplier = 1 + (npt_percent / 100) * 0.5
        total_cost = operating_days * dayrate * npt_cost_multiplier
        random_factor = self.random_state.normal(1.0, 0.1)
        return total_cost * random_factor
    
    def _simulate_risk(self, npt, duration, basin_params):
        """Simulate overall risk score"""
        # Handle both old and new parameter formats
        try:
            climate_sev = float(basin_params.get('climate_severity', basin_params.get('climate', 5.0)))
        except (ValueError, TypeError):
            climate_sev = 5.0
        
        try:
            geology_diff = float(basin_params.get('geology_difficulty', basin_params.get('geology', 5.0)))
        except (ValueError, TypeError):
            geology_diff = 5.0
        
        npt_risk = npt * 1.5
        duration_risk = (duration - 30) * 0.8 if duration > 30 else 0
        climate_risk = climate_sev * 4
        geology_risk = geology_diff * 3.5
        total_risk = (npt_risk + duration_risk + climate_risk + geology_risk)
        random_factor = self.random_state.normal(1.0, 0.15)
        return max(0, min(100, total_risk * random_factor))


class ContractorPerformanceAnalyzer:
    """Analyze contractor performance consistency"""
    
    def __init__(self):
        self.consistency_weights = {'rop_variance': 0.25, 'npt_variance': 0.25, 'schedule_variance': 0.20,
                                   'delivery_reliability': 0.20, 'crew_stability': 0.10}
    
    def analyze_contractor_consistency(self, contractor_data):
        """Comprehensive contractor consistency analysis"""
        if contractor_data.empty or len(contractor_data) < 2:
            return {'overall_consistency': 50, 'grade': 'Insufficient Data', 'note': 'Need at least 2 contracts'}
        
        metrics = {}
        metrics['rop_consistency'] = self._analyze_rop_variance(contractor_data)
        metrics['npt_consistency'] = self._analyze_npt_variance(contractor_data)
        metrics['schedule_consistency'] = self._analyze_schedule_variance(contractor_data)
        metrics['delivery_reliability'] = self._analyze_delivery_reliability(contractor_data)
        metrics['crew_stability'] = self._analyze_crew_stability(contractor_data)
        
        weights = [0.25, 0.25, 0.20, 0.20, 0.10]
        scores = [metrics['rop_consistency'], metrics['npt_consistency'], metrics['schedule_consistency'],
                 metrics['delivery_reliability'], metrics['crew_stability']]
        overall = sum(s * w for s, w in zip(scores, weights))
        
        sample_size_factor = min(1.0, len(contractor_data) / 10)
        confidence_adjusted_score = overall * (0.7 + 0.3 * sample_size_factor)
        
        metrics['overall_consistency'] = confidence_adjusted_score
        metrics['consistency_grade'] = self._get_consistency_grade(confidence_adjusted_score)
        metrics['sample_size'] = len(contractor_data)
        return metrics
    
    def _analyze_rop_variance(self, data):
        if 'Contract Length' in data.columns:
            lengths = data['Contract Length'].dropna()
            if len(lengths) >= 2:
                mean_length = lengths.mean()
                std_length = lengths.std()
                cv = (std_length / mean_length) * 100 if mean_length > 0 else 50
                consistency_score = max(40, 100 - cv)
                return min(100, consistency_score)
        return 70
    
    def _analyze_npt_variance(self, data):
        npt_col = 'NPT %' if 'NPT %' in data.columns else ('NPT_Percent' if 'NPT_Percent' in data.columns else None)
        if npt_col:
            npt_values = data[npt_col].dropna()
            if len(npt_values) >= 2:
                mean_npt = npt_values.mean()
                std_npt = npt_values.std()
                variance_score = 95 if std_npt < 3 else (85 if std_npt < 5 else (70 if std_npt < 8 else 55))
                if mean_npt > 20:
                    variance_score *= 0.8
                return variance_score
        return 70
    
    def _analyze_schedule_variance(self, data):
        if 'Contract Length' not in data.columns:
            return 70
        lengths = data['Contract Length'].dropna()
        if len(lengths) < 2:
            return 70
        mean_length = lengths.mean()
        std_length = lengths.std()
        cv = (std_length / mean_length) * 100 if mean_length > 0 else 50
        return 90 if cv < 15 else (75 if cv < 25 else 60)
    
    def _analyze_delivery_reliability(self, data):
        if 'Status' in data.columns:
            status_col = data['Status'].dropna()
            if not status_col.empty:
                status_lower = status_col.str.lower()
                successful = status_lower.str.contains('complete|successful|finished|active|operating', na=False).sum()
                failed = status_lower.str.contains('terminated|cancelled|suspended|failed', na=False).sum()
                total = len(status_col)
                if total > 0:
                    success_rate = (successful / total) * 100
                    failure_penalty = (failed / total) * 20
                    return max(0, min(100, success_rate - failure_penalty))
        return 75
    
    def _analyze_crew_stability(self, data):
        if 'Contract Start Date' not in data.columns or len(data) < 3:
            return 70
        sorted_data = data.sort_values('Contract Start Date')
        start_dates = pd.to_datetime(sorted_data['Contract Start Date'], errors='coerce').dropna()
        if len(start_dates) < 2:
            return 70
        gaps = [(start_dates.iloc[i] - start_dates.iloc[i-1]).days for i in range(1, len(start_dates))]
        avg_gap = np.mean(gaps)
        return 95 if avg_gap < 30 else (85 if avg_gap < 90 else 70)
    
    def _get_consistency_grade(self, score):
        if score >= 90:
            return 'A+ (Highly Consistent)'
        elif score >= 80:
            return 'A (Very Consistent)'
        elif score >= 70:
            return 'B (Consistent)'
        return 'C (Moderately Consistent)'


class LearningCurveAnalyzer:
    """Analyze and visualize rig learning curves"""
    
    def __init__(self):
        pass
    
    def calculate_learning_curve(self, rig_data):
        """Calculate learning curve parameters using power law"""
        if len(rig_data) < 3:
            return {'status': 'INSUFFICIENT_DATA', 'message': 'Need at least 3 data points'}
        
        if 'Contract Start Date' in rig_data.columns:
            sorted_data = rig_data.sort_values('Contract Start Date').reset_index(drop=True)
        else:
            sorted_data = rig_data.reset_index(drop=True)
        
        if 'Contract Length' not in sorted_data.columns:
            return {'status': 'NO_TIME_DATA', 'message': 'No time-based data available'}
        
        times = sorted_data['Contract Length'].dropna().values
        if len(times) < 3:
            return {'status': 'INSUFFICIENT_DATA', 'message': 'Need at least 3 time measurements'}
        
        n = np.arange(1, len(times) + 1)
        log_times = np.log(times)
        log_n = np.log(n)
        
        slope, intercept = np.polyfit(log_n, log_times, 1)
        k = -slope
        T1 = np.exp(intercept)
        
        predicted_log_times = intercept + slope * log_n
        ss_res = np.sum((log_times - predicted_log_times) ** 2)
        ss_tot = np.sum((log_times - np.mean(log_times)) ** 2)
        r_squared = 1 - (ss_res / ss_tot) if ss_tot > 0 else 0
        
        future_n = np.arange(1, len(times) + 6)
        predicted_times = T1 * (future_n ** -k)
        
        improvement_percent = ((times[0] - times[-1]) / times[0] * 100) if times[0] > 0 else 0
        
        return {
            'status': 'SUCCESS',
            'learning_rate_k': k,
            'initial_time_T1': T1,
            'r_squared': r_squared,
            'actual_times': times.tolist(),
            'predicted_times': predicted_times.tolist(),
            'improvement_percent': improvement_percent,
            'n_contracts': len(times),
            'current_efficiency': times[-1] if len(times) > 0 else 0,
            'projected_efficiency': predicted_times[-1] if len(predicted_times) > 0 else 0
        }


class InvisibleLostTimeDetector:
    """AI-powered pattern mining to detect invisible lost time (ILT)"""
    
    def __init__(self):
        pass
    
    def detect_ilt(self, rig_data):
        """Detect invisible lost time from available data"""
        ilt_findings = []
        total_ilt_days = 0
        
        if 'Contract Length' in rig_data.columns:
            lengths = rig_data['Contract Length'].dropna()
            if len(lengths) >= 3:
                mean_length = lengths.mean()
                std_length = lengths.std()
                if std_length > 0:
                    ilt_findings.append({'type': 'Duration Variance', 'severity': 'MEDIUM'})
                    total_ilt_days += std_length * 0.5
        
        # Baseline ILT estimate (industry typical 7-10%)
        total_days = rig_data['Contract Length'].sum() if 'Contract Length' in rig_data.columns else 0
        base_ilt_rate = 0.07
        
        if 'Current Location' in rig_data.columns:
            location = str(rig_data['Current Location'].iloc[0]).lower() if len(rig_data) > 0 else ''
            if any(term in location for term in ['deepwater', 'hpht', 'arctic']):
                base_ilt_rate += 0.03
        
        base_ilt_days = total_days * base_ilt_rate
        total_ilt_days += base_ilt_days
        
        ilt_percentage = (total_ilt_days / total_days * 100) if total_days > 0 else 0
        avg_dayrate = rig_data['Dayrate ($k)'].mean() if 'Dayrate ($k)' in rig_data.columns else 200
        cost_impact = total_ilt_days * avg_dayrate
        
        return {
            'total_ilt_days': total_ilt_days,
            'ilt_percentage': ilt_percentage,
            'cost_impact_$k': cost_impact,
            'findings': ilt_findings,
            'severity': 'LOW' if ilt_percentage < 5 else ('MODERATE' if ilt_percentage < 10 else 'HIGH')
        }


class RigEfficiencyCalculator:
    """
    Advanced Rig Efficiency Calculator with Multi-Factor Analysis
    Calculates comprehensive efficiency metrics across 6 key factors
    """
    
    def __init__(self):
        self.climate_ai = AdvancedClimateIntelligence()
        self.benchmark_model = RegionalBenchmarkModel()
        self.ml_predictor = RigWellMatchPredictor()
        self.monte_carlo = MonteCarloScenarioSimulator(num_simulations=1000)
        self.contractor_analyzer = ContractorPerformanceAnalyzer()
        self.learning_analyzer = LearningCurveAnalyzer()
        self.ilt_detector = InvisibleLostTimeDetector()
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


class RigAvailabilitySearchEngine:
    """
    Search rigs based on available contract data
    Works with existing columns without lithology/pressure
    """
    
    def __init__(self, climate_intelligence):
        """
        Initialize the search engine with climate intelligence
        
        Args:
            climate_intelligence: Instance of AdvancedClimateIntelligence
        """
        self.climate_ai = climate_intelligence
        
        # Location-based lithology inference (optional)
        self.location_lithology_map = {
            'US Gulf': ['Sandstone', 'Shale'],
            'North Sea': ['Sandstone', 'Carbonate'],
            'Middle East': ['Carbonate', 'Limestone'],
            'Brazil': ['Carbonate', 'Sandstone'],
            'West Africa': ['Sandstone'],
            'Southeast Asia': ['Sandstone', 'Carbonate']
        }
        
        # Rig type inference from contractor/rig name (if available)
        self.rig_type_keywords = {
            'jackup': ['Jackup', 'Jack-up', 'JU'],
            'semi': ['Semi', 'Semisubmersible', 'SS'],
            'drillship': ['Drillship', 'DS'],
            'platform': ['Platform', 'Fixed']
        }
        
        # Regional location groupings for fuzzy matching
        self.regional_groups = {
            'Gulf of Mexico': ['gulf of mexico', 'us gulf', 'mexico gulf', 'gom'],
            'North Sea': ['north sea', 'uk north sea', 'norway', 'norwegian sea'],
            'Middle East': ['saudi arabia', 'uae', 'qatar', 'kuwait', 'iraq', 'oman'],
            'West Africa': ['nigeria', 'angola', 'ghana', 'congo', 'equatorial guinea'],
            'Southeast Asia': ['malaysia', 'indonesia', 'thailand', 'vietnam', 'brunei'],
            'Brazil': ['brazil', 'santos basin', 'campos basin'],
            'Australia': ['australia', 'timor sea', 'browse basin']
        }
    
    def search_available_rigs(self, df, filters):
        """
        Main search function using only available data
        
        Parameters:
        -----------
        df : DataFrame
            Your rig contract data
        filters : dict
            {
                'location': str or list,
                'region': str or list,
                'dayrate_min': float,
                'dayrate_max': float,
                'climate_preference': str,
                'availability_status': str,
                'contractor': str (optional)
            }
        
        Returns:
        --------
        DataFrame with available rigs and match scores
        """
        
        results = df.copy()
        
        # 1. FILTER BY LOCATION
        if filters.get('location') and filters['location'] != 'All':
            if isinstance(filters['location'], list):
                results = results[results['Current Location'].isin(filters['location'])]
            else:
                results = results[results['Current Location'] == filters['location']]
        
        # 2. FILTER BY REGION
        if filters.get('region') and filters['region'] != 'All':
            results = results[results['Region'] == filters['region']]
        
        # 3. FILTER BY DAY RATE RANGE
        if filters.get('dayrate_min') is not None:
            results = results[results['Dayrate ($k)'] >= filters['dayrate_min']]
        if filters.get('dayrate_max') is not None:
            results = results[results['Dayrate ($k)'] <= filters['dayrate_max']]
        
        # 4. FILTER BY AVAILABILITY (based on contract dates)
        if filters.get('availability_status'):
            results = self._filter_by_availability(results, filters['availability_status'])
        
        # 5. FILTER BY STATUS
        if filters.get('status'):
            results = results[results['Status'].isin(filters['status'])]
        
        # 6. CALCULATE MATCH SCORES
        results = self._calculate_match_scores(results, filters)
        
        # 7. ADD CLIMATE COMPATIBILITY
        results = self._add_climate_scores(results, filters.get('climate_preference'))
        
        # 8. SORT BY MATCH SCORE
        results = results.sort_values('Match_Score', ascending=False)
        
        return results
    
    def _filter_by_availability(self, df, status):
        """Filter based on contract availability"""
        today = pd.Timestamp.now()
        
        if status == 'Available Now':
            # Contracts ended or ending within 7 days
            mask = (
                (df['Contract End Date'].notna()) & 
                (df['Contract End Date'] <= today + pd.Timedelta(days=7))
            ) | (df['Status'].isin(['Available', 'Idle', 'Stacked']))
            return df[mask]
        
        elif status == 'Available Soon':
            # Contracts ending within 30 days
            mask = (
                (df['Contract End Date'].notna()) & 
                (df['Contract End Date'] <= today + pd.Timedelta(days=30)) &
                (df['Contract End Date'] > today)
            )
            return df[mask]
        
        elif status == 'Available <90 days':
            mask = (
                (df['Contract End Date'].notna()) & 
                (df['Contract End Date'] <= today + pd.Timedelta(days=90))
            )
            return df[mask]
        
        else:  # All
            return df
    
    def _calculate_match_scores(self, df, filters):
        """Calculate how well each rig matches criteria"""
        
        def calculate_row_score(row):
            score = 0
            max_score = 100
            
            # Location match (40 points)
            if filters.get('location') and filters['location'] != 'All':
                if row['Current Location'] == filters['location']:
                    score += 40
                elif self._is_nearby_location(row['Current Location'], filters['location']):
                    score += 20  # Partial credit for nearby locations
            else:
                score += 40  # Full points if no location specified
            
            # Day rate match (30 points)
            if filters.get('dayrate_min') and filters.get('dayrate_max'):
                rate = row['Dayrate ($k)']
                mid_range = (filters['dayrate_min'] + filters['dayrate_max']) / 2
                rate_diff = abs(rate - mid_range)
                max_diff = (filters['dayrate_max'] - filters['dayrate_min']) / 2
                
                if max_diff > 0:
                    rate_score = 30 * (1 - min(rate_diff / max_diff, 1))
                    score += rate_score
                else:
                    score += 30
            else:
                score += 30
            
            # Availability (30 points)
            if pd.notna(row.get('Contract Days Remaining')):
                days_remaining = row['Contract Days Remaining']
                if days_remaining <= 0:
                    score += 30  # Immediately available
                elif days_remaining <= 30:
                    score += 25  # Available soon
                elif days_remaining <= 90:
                    score += 15  # Available within quarter
                else:
                    score += 5   # Future availability
            else:
                # No contract info - assume available
                score += 30
            
            return score
        
        df['Match_Score'] = df.apply(calculate_row_score, axis=1)
        return df
    
    def _add_climate_scores(self, df, climate_preference):
        """Add climate compatibility scores"""
        
        def get_climate_score(location):
            if not location or pd.isna(location):
                return 5  # Neutral score
            
            location_lower = str(location).lower()
            
            # Use existing climate intelligence
            climate_profile = None
            for key in self.climate_ai.climate_profiles.keys():
                if key in location_lower or location_lower in key:
                    climate_profile = self.climate_ai.climate_profiles[key]
                    break
            
            if climate_profile:
                # Score based on efficiency factor (higher is better)
                efficiency = climate_profile.get('efficiency_factor', 0.8)
                return efficiency * 10  # Convert to 0-10 scale
            
            return 5  # Default neutral score
        
        df['Climate_Score'] = df['Current Location'].apply(get_climate_score)
        
        # Adjust match score based on climate
        if climate_preference and climate_preference != 'Any':
            df['Match_Score'] = df['Match_Score'] * (0.9 + df['Climate_Score'] / 100)
        
        return df
    
    def _is_nearby_location(self, loc1, loc2):
        """Check if locations are in same region"""
        # Simple region matching - can be enhanced
        regions_map = {
            'US Gulf': ['Gulf of Mexico', 'US Gulf', 'GoM'],
            'North Sea': ['North Sea', 'Norway', 'UK', 'Netherlands'],
            'Middle East': ['Saudi Arabia', 'UAE', 'Qatar', 'Kuwait'],
            'Brazil': ['Brazil', 'South America'],
            'West Africa': ['Nigeria', 'Angola', 'Ghana', 'West Africa']
        }
        
        for region, locations in regions_map.items():
            if (any(l.lower() in str(loc1).lower() for l in locations) and 
                any(l.lower() in str(loc2).lower() for l in locations)):
                return True
        
        return False
    
    def infer_rig_capabilities(self, df):
        """
        Infer lithology and pressure capabilities from location and rig name
        (Optional enhancement)
        """
        
        def infer_lithology(row):
            location = str(row.get('Current Location', '')).lower()
            
            for region, lithos in self.location_lithology_map.items():
                if region.lower() in location:
                    return ', '.join(lithos)
            
            return 'Mixed'  # Default
        
        def infer_rig_type(row):
            rig_name = str(row.get('Drilling Unit Name', '')).lower()
            contractor = str(row.get('Contractor', '')).lower()
            
            for rig_type, keywords in self.rig_type_keywords.items():
                for keyword in keywords:
                    if keyword.lower() in rig_name or keyword.lower() in contractor:
                        return rig_type.title()
            
            return 'Unknown'
        
        df['Inferred_Lithology'] = df.apply(infer_lithology, axis=1)
        df['Inferred_Rig_Type'] = df.apply(infer_rig_type, axis=1)
        
        return df


# Export main classes
__all__ = [
    'RigEfficiencyCalculator',
    'AdvancedClimateIntelligence',
    'RegionalBenchmarkModel',
    'RigWellMatchPredictor',
    'MonteCarloScenarioSimulator',
    'ContractorPerformanceAnalyzer',
    'LearningCurveAnalyzer',
    'InvisibleLostTimeDetector',
    'RigAvailabilitySearchEngine',
    'preprocess_dataframe'
]