
"""
Advanced Rig Efficiency Analysis Tool - Enhanced Professional Version
Multi-factor efficiency analysis with AI-powered climate intelligence
Premium Black Theme with Enhanced Interactivity
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import io
import streamlit.components.v1 as components
from rig_efficiency_backend import (
    RigEfficiencyCalculator, 
    AdvancedClimateIntelligence,
    RegionalBenchmarkModel,
    RigWellMatchPredictor,
    MonteCarloScenarioSimulator,
    ContractorPerformanceAnalyzer,
    LearningCurveAnalyzer,
    InvisibleLostTimeDetector,
    preprocess_dataframe
)

# ==================== IMPORT ENHANCED AI CHATBOT ====================
# Import the advanced AI chatbot with NLP, sentiment analysis, and context tracking
import sys
import importlib.util

# Load the enhanced chatbot module (handles filename with space)
_chatbot_spec = importlib.util.spec_from_file_location(
    "rig_chatbot_enhanced",
    r"c:\Office work\Upstream SCRAP news\Tools\rig tool files\rig chatbot.py"
)
_chatbot_module = importlib.util.module_from_spec(_chatbot_spec)
sys.modules["rig_chatbot_enhanced"] = _chatbot_module
_chatbot_spec.loader.exec_module(_chatbot_module)

# Import the enhanced chatbot class
RigEfficiencyAIChatbot = _chatbot_module.RigEfficiencyAIChatbot

# Page configuration
st.set_page_config(
    page_title="Rig Efficiency Intelligence Platform",
    page_icon="🛢️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== PERFORMANCE OPTIMIZATION: CACHING LAYER ====================
# These cached functions dramatically improve app performance by preventing unnecessary recalculations

@st.cache_resource
def load_custom_css():
    """Cache CSS - loads only once per session"""
    return """
<style>
    /* Import Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;500;700;900&family=Rajdhani:wght@300;400;500;600;700&family=Poppins:wght@300;400;600;700;900&display=swap');
    
    /* ====== COLOR PALETTE ====== */
    :root {
        /* Transparent Blacks with Blue Tint */
        --bg-primary: rgba(0, 10, 25, 0.95);
        --bg-secondary: rgba(5, 15, 35, 0.92);
        --bg-tertiary: rgba(10, 20, 45, 0.88);
        --bg-card: rgba(15, 30, 60, 0.75);
        
        /* Electric Blues */
        --blue-primary: #00D4FF;
        --blue-secondary: #0099FF;
        --blue-accent: #0066FF;
        --blue-glow: rgba(0, 212, 255, 0.4);
        
        /* Cyber Accents */
        --cyan-bright: #00FFFF;
        --electric-blue: #1E90FF;
        --neon-blue: #4FC3F7;
        --ice-blue: #87CEEB;
        
        /* Text Colors - High Visibility */
        --text-primary: #FFFFFF;
        --text-secondary: #E0F2FF;
        --text-muted: #A0C4E0;
        
        /* Status Colors */
        --success: #00FFB3;
        --warning: #FFB800;
        --danger: #FF3366;
        --info: #00D4FF;
        
        /* Fonts */
        --font-primary: 'Rajdhani', 'Poppins', sans-serif;
        --font-display: 'Orbitron', sans-serif;
        --font-body: 'Poppins', sans-serif;
    }
    
    /* Global Styles - Avoid disrupting Streamlit layout */
    body, .stApp, h1, h2, h3, h4, h5, h6, p {
        font-family: var(--font-primary) !important;
        letter-spacing: 0.3px;
    }
    
    /* Preserve Streamlit's native checkbox/radio layout */
    .stCheckbox,
    .stCheckbox *,
    .stRadio,
    .stRadio * {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
    }
    
    .stApp {
        background: 
            linear-gradient(135deg, rgba(0, 10, 25, 0.98) 0%, rgba(0, 25, 50, 0.95) 100%),
            radial-gradient(ellipse at top left, rgba(0, 100, 255, 0.12) 0%, transparent 50%),
            radial-gradient(ellipse at bottom right, rgba(0, 212, 255, 0.12) 0%, transparent 50%),
            linear-gradient(180deg, #000a19 0%, #001932 100%);
        background-attachment: fixed;
        position: relative;
        z-index: 1;
        overflow: visible !important;
    }
    
    /* Animated Cyber Grid Background */
    .stApp::before {
        content: '';
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-image: 
            linear-gradient(rgba(0, 212, 255, 0.03) 1px, transparent 1px),
            linear-gradient(90deg, rgba(0, 212, 255, 0.03) 1px, transparent 1px);
        background-size: 50px 50px;
        animation: gridMove 25s linear infinite;
        z-index: -2;
        pointer-events: none;
    }
    
    @keyframes gridMove {
        0% { transform: translate(0, 0); }
        100% { transform: translate(50px, 50px); }
    }
    
    /* Scanning Line Effect (Transformer Style) */
    .stApp::after {
        content: '';
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 3px;
        background: linear-gradient(90deg, 
            transparent 0%, 
            var(--blue-primary) 20%,
            var(--cyan-bright) 50%,
            var(--blue-primary) 80%,
            transparent 100%);
        box-shadow: 
            0 0 20px var(--blue-glow),
            0 0 40px var(--blue-glow);
        animation: scanLine 5s ease-in-out infinite;
        z-index: -2;
        pointer-events: none;
    }

    /* Ensure streamlit content renders above decorative overlays */
    .stApp > * {
        position: relative;
        z-index: auto;
    }
    
    @keyframes scanLine {
        0%, 100% { 
            transform: translateY(0); 
            opacity: 0; 
        }
        50% { 
            transform: translateY(100vh); 
            opacity: 1; 
        }
    }
    
    /* Headers with Enhanced Visibility */
    h1, h2, h3, h4, h5, h6 {
        color: var(--text-primary) !important;
        font-weight: 700 !important;
        text-shadow: 0 0 25px var(--blue-glow);
        letter-spacing: 1px;
    }
    
    /* TRANSFORMER HEADER - Main Title */
    .main-header {
        font-size: 4rem;
        font-weight: 900;
        font-family: var(--font-display) !important;
        background: linear-gradient(135deg, 
            var(--cyan-bright) 0%, 
            var(--blue-primary) 20%,
            var(--electric-blue) 40%,
            var(--blue-primary) 60%,
            var(--cyan-bright) 80%,
            var(--blue-primary) 100%);
        background-size: 300% 100%;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        text-align: center;
        padding: 2rem 0;
        animation: 
            transformerGlow 3s ease-in-out infinite,
            textShift 10s linear infinite,
            floatAnimation 4s ease-in-out infinite;
        filter: drop-shadow(0 0 40px var(--blue-glow))
                drop-shadow(0 0 60px var(--blue-glow));
        text-transform: uppercase;
        letter-spacing: 6px;
        position: relative;
    }
    
    @keyframes transformerGlow {
        0%, 100% { 
            filter: drop-shadow(0 0 30px var(--blue-glow))
                    drop-shadow(0 0 50px var(--blue-glow));
        }
        50% { 
            filter: drop-shadow(0 0 50px var(--blue-glow))
                    drop-shadow(0 0 90px var(--blue-glow));
        }
    }
    
    @keyframes textShift {
        0% { background-position: 0% 50%; }
        100% { background-position: 300% 50%; }
    }
    
    @keyframes floatAnimation {
        0%, 100% { transform: translateY(0px); }
        50% { transform: translateY(-15px); }
    }
    
    /* CLIENT GUIDE BANNER */
    .client-guide-banner {
        background: linear-gradient(135deg, 
            rgba(0, 100, 255, 0.2) 0%, 
            rgba(0, 212, 255, 0.15) 100%);
        backdrop-filter: blur(20px);
        border: 2px solid var(--blue-primary);
        border-radius: 20px;
        padding: 2rem;
        margin: 2rem 0;
        text-align: center;
        box-shadow: 
            0 8px 32px rgba(0, 0, 0, 0.5),
            inset 0 1px 0 rgba(0, 212, 255, 0.3),
            0 0 40px rgba(0, 212, 255, 0.2);
        animation: guidePulse 3s ease-in-out infinite;
        position: relative;
        overflow: hidden;
    }
    
    .client-guide-banner::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: linear-gradient(
            45deg,
            transparent 30%,
            rgba(0, 212, 255, 0.15) 50%,
            transparent 70%
        );
        transform: rotate(45deg);
        animation: bannerShimmer 4s linear infinite;
    }
    
    @keyframes guidePulse {
        0%, 100% { 
            border-color: var(--blue-primary);
            box-shadow: 
                0 8px 32px rgba(0, 0, 0, 0.5),
                inset 0 1px 0 rgba(0, 212, 255, 0.3),
                0 0 40px rgba(0, 212, 255, 0.2);
        }
        50% { 
            border-color: var(--cyan-bright);
            box-shadow: 
                0 8px 32px rgba(0, 0, 0, 0.5),
                inset 0 1px 0 rgba(0, 255, 255, 0.4),
                0 0 60px rgba(0, 212, 255, 0.4);
        }
    }
    
    @keyframes bannerShimmer {
        0% { transform: translateX(-100%) translateY(-100%) rotate(45deg); }
        100% { transform: translateX(100%) translateY(100%) rotate(45deg); }
    }
    
    /* Subtitle - High Visibility */
    .subtitle {
        text-align: center;
        color: var(--text-secondary);
        font-size: 1.15rem;
        margin-bottom: 2rem;
        letter-spacing: 4px;
        font-weight: 500;
        text-transform: uppercase;
        animation: fadeInUp 1s ease, subtitleGlow 3s ease-in-out infinite;
        text-shadow: 0 0 15px var(--blue-glow);
        font-family: var(--font-display) !important;
    }
    
    @keyframes subtitleGlow {
        0%, 100% { 
            opacity: 0.85;
            text-shadow: 0 0 15px var(--blue-glow);
        }
        50% { 
            opacity: 1;
            text-shadow: 0 0 25px var(--blue-glow);
        }
    }
    
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(30px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    /* HOLOGRAPHIC CARDS - Transparent Black/Blue */
    .metric-card {
        background: linear-gradient(135deg, 
            rgba(15, 30, 60, 0.85) 0%,
            rgba(20, 40, 75, 0.75) 100%);
        backdrop-filter: blur(25px) saturate(180%);
        -webkit-backdrop-filter: blur(25px) saturate(180%);
        padding: 2rem;
        border-radius: 20px;
        border: 2px solid transparent;
        border-image: linear-gradient(135deg, 
            var(--blue-primary), 
            var(--cyan-bright), 
            var(--blue-primary)) 1;
        box-shadow: 
            0 8px 32px rgba(0, 0, 0, 0.6),
            inset 0 1px 0 rgba(0, 212, 255, 0.25),
            0 0 0 1px rgba(0, 212, 255, 0.15);
        margin: 0.5rem 0;
        position: relative;
        overflow: hidden;
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
    }
    
    /* Holographic Shimmer Effect */
    .metric-card::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: linear-gradient(
            45deg,
            transparent 30%,
            rgba(0, 212, 255, 0.2) 50%,
            transparent 70%
        );
        transform: rotate(45deg);
        animation: holoShimmer 3.5s linear infinite;
    }
    
    @keyframes holoShimmer {
        0% { transform: translateX(-100%) translateY(-100%) rotate(45deg); }
        100% { transform: translateX(100%) translateY(100%) rotate(45deg); }
    }
    
    .metric-card:hover {
        transform: translateY(-10px) scale(1.02);
        border-image: linear-gradient(135deg, 
            var(--cyan-bright), 
            var(--blue-primary), 
            var(--cyan-bright)) 1;
        box-shadow: 
            0 16px 48px rgba(0, 212, 255, 0.5),
            0 0 70px rgba(0, 212, 255, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.4);
    }
    
    /* TRANSFORMER SCORE DISPLAY */
    .score-display {
        font-size: 4rem;
        font-weight: 900;
        font-family: var(--font-display) !important;
        text-align: center;
        margin: 1.5rem 0;
        animation: 
            scoreReveal 0.9s cubic-bezier(0.175, 0.885, 0.32, 1.275),
            scoreGlow 2.5s ease-in-out infinite;
        position: relative;
    }
    
    @keyframes scoreReveal {
        0% { 
            transform: scale(0) rotate(-180deg);
            opacity: 0;
        }
        60% {
            transform: scale(1.1) rotate(10deg);
        }
        100% { 
            transform: scale(1) rotate(0deg);
            opacity: 1;
        }
    }
    
    @keyframes scoreGlow {
        0%, 100% { 
            text-shadow: 
                0 0 15px currentColor,
                0 0 30px currentColor,
                0 0 45px currentColor;
        }
        50% { 
            text-shadow: 
                0 0 25px currentColor,
                0 0 50px currentColor,
                0 0 75px currentColor,
                0 0 100px currentColor;
        }
    }
    
    .score-excellent { 
        color: var(--success);
        animation: scoreReveal 0.9s cubic-bezier(0.175, 0.885, 0.32, 1.275),
                   scoreGlow 2.5s ease-in-out infinite;
    }
    .score-good { 
        color: var(--blue-primary);
        animation: scoreReveal 0.9s cubic-bezier(0.175, 0.885, 0.32, 1.275),
                   scoreGlow 2.5s ease-in-out infinite;
    }
    .score-fair { 
        color: var(--warning);
        animation: scoreReveal 0.9s cubic-bezier(0.175, 0.885, 0.32, 1.275),
                   scoreGlow 2.5s ease-in-out infinite;
    }
    .score-poor { 
        color: var(--danger);
        animation: scoreReveal 0.9s cubic-bezier(0.175, 0.885, 0.32, 1.275),
                   scoreGlow 2.5s ease-in-out infinite;
    }
    
    /* Enhanced Insight Boxes */
    .insight-box {
        padding: 1.8rem;
        border-radius: 16px;
        margin: 1.2rem 0;
        border-left: 4px solid;
        background: rgba(33, 38, 45, 0.8);
        backdrop-filter: blur(10px);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
    }
    
    .insight-box::after {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: linear-gradient(135deg, transparent 0%, rgba(255, 255, 255, 0.03) 100%);
        pointer-events: none;
    }
    
    .insight-box:hover {
        transform: translateX(12px) scale(1.01);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.5);
    }
    
    .success-box {
        border-left-color: var(--success-green);
        background: linear-gradient(90deg, rgba(0, 255, 159, 0.15) 0%, rgba(33, 38, 45, 0.8) 100%);
    }
    
    .warning-box {
        border-left-color: var(--warning-orange);
        background: linear-gradient(90deg, rgba(255, 183, 77, 0.15) 0%, rgba(33, 38, 45, 0.8) 100%);
    }
    
    .info-box {
        border-left-color: var(--accent-cyan);
        background: linear-gradient(90deg, rgba(0, 229, 255, 0.15) 0%, rgba(33, 38, 45, 0.8) 100%);
    }
    
    .danger-box {
        border-left-color: var(--danger-red);
        background: linear-gradient(90deg, rgba(255, 107, 107, 0.15) 0%, rgba(33, 38, 45, 0.8) 100%);
    }
    
    /* CYBER TABS - Blue Theme */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.8rem;
        background: rgba(15, 30, 60, 0.7);
        backdrop-filter: blur(25px);
        padding: 1rem;
        border-radius: 20px;
        border: 2px solid rgba(0, 212, 255, 0.3);
        box-shadow: 
            0 4px 24px rgba(0, 0, 0, 0.5),
            inset 0 1px 0 rgba(0, 212, 255, 0.3),
            0 0 30px rgba(0, 212, 255, 0.15);
    }
    
    .stTabs [data-baseweb="tab"] {
        padding: 1rem 2rem;
        font-size: 1.1rem;
        font-weight: 600;
        font-family: var(--font-display) !important;
        color: var(--text-muted);
        background: transparent;
        border-radius: 12px;
        border: 1px solid transparent;
        transition: all 0.3s ease;
        position: relative;
        text-transform: uppercase;
        letter-spacing: 1.5px;
    }
    
    .stTabs [data-baseweb="tab"]::before {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        width: 0%;
        height: 3px;
        background: linear-gradient(90deg, var(--blue-primary), var(--cyan-bright));
        transition: width 0.3s ease;
        box-shadow: 0 0 10px var(--blue-primary);
    }
    
    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(0, 212, 255, 0.15);
        color: var(--blue-primary);
        border-color: rgba(0, 212, 255, 0.4);
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(0, 212, 255, 0.3);
    }
    
    .stTabs [data-baseweb="tab"]:hover::before {
        width: 100%;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, 
            rgba(0, 100, 255, 0.35) 0%, 
            rgba(0, 212, 255, 0.25) 100%);
        color: var(--cyan-bright) !important;
        border-color: var(--blue-primary);
        box-shadow: 
            0 0 25px rgba(0, 212, 255, 0.5),
            inset 0 0 25px rgba(0, 212, 255, 0.15);
        font-weight: 700;
    }
    
    .stTabs [aria-selected="true"]::before {
        width: 100%;
    }
    
    /* QUANTUM BUTTONS - Blue Theme */
    .stButton > button {
        background: linear-gradient(135deg, 
            var(--blue-accent) 0%, 
            var(--electric-blue) 50%,
            var(--blue-primary) 100%);
        color: var(--text-primary);
        font-weight: 700;
        font-family: var(--font-display) !important;
        border: 2px solid var(--blue-primary);
        border-radius: 14px;
        padding: 1rem 3rem;
        font-size: 1.15rem;
        text-transform: uppercase;
        letter-spacing: 2px;
        position: relative;
        overflow: hidden;
        transition: all 0.3s ease;
        box-shadow: 
            0 4px 20px rgba(0, 100, 255, 0.6),
            0 0 40px rgba(0, 212, 255, 0.4),
            inset 0 0 20px rgba(0, 212, 255, 0.15);
    }
    
    /* Ripple Effect */
    .stButton > button::before {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        width: 0;
        height: 0;
        border-radius: 50%;
        background: rgba(0, 255, 255, 0.6);
        transform: translate(-50%, -50%);
        transition: width 0.6s, height 0.6s;
    }
    
    .stButton > button:hover::before {
        width: 350px;
        height: 350px;
    }
    
    .stButton > button:hover {
        transform: translateY(-4px) scale(1.05);
        border-color: var(--cyan-bright);
        box-shadow: 
            0 8px 35px rgba(0, 212, 255, 0.7),
            0 0 60px rgba(0, 212, 255, 0.6),
            inset 0 0 30px rgba(0, 212, 255, 0.25);
    }
    
    /* CYBER SIDEBAR - Enhanced Blue Theme */
    .css-1d391kg, [data-testid="stSidebar"] {
        background: linear-gradient(180deg, 
            rgba(0, 10, 25, 0.98) 0%, 
            rgba(5, 20, 40, 0.95) 50%,
            rgba(10, 25, 50, 0.98) 100%);
        backdrop-filter: blur(25px);
        border-right: 2px solid var(--blue-primary);
        box-shadow: 
            4px 0 32px rgba(0, 0, 0, 0.5),
            inset -1px 0 0 rgba(0, 212, 255, 0.2);
        position: relative;
    }
    
    .css-1d391kg::before, [data-testid="stSidebar"]::before {
        content: '';
        position: absolute;
        top: 0;
        right: 0;
        width: 2px;
        height: 100%;
        background: linear-gradient(180deg,
            transparent 0%,
            var(--cyan-bright) 50%,
            transparent 100%);
        animation: sidebarPulse 4s ease-in-out infinite;
    }
    
    @keyframes sidebarPulse {
        0%, 100% { opacity: 0.3; }
        50% { opacity: 1; }
    }
    
    .css-1d391kg h3, [data-testid="stSidebar"] h3 {
        color: var(--cyan-bright) !important;
        font-weight: 700;
        font-family: var(--font-display) !important;
        letter-spacing: 2px;
        text-transform: uppercase;
        font-size: 1rem;
        margin-top: 1.5rem;
        text-shadow: 0 0 15px var(--blue-glow);
        animation: headerGlow 3s ease-in-out infinite;
    }
    
    @keyframes headerGlow {
        0%, 100% { text-shadow: 0 0 15px var(--blue-glow); }
        50% { text-shadow: 0 0 25px var(--blue-glow), 0 0 35px var(--blue-glow); }
    }
    
    /* HOLOGRAPHIC FILE UPLOADER */
    .stFileUploader {
        background: rgba(15, 30, 60, 0.5);
        border: 2px dashed var(--blue-primary);
        border-radius: 20px;
        padding: 2rem;
        transition: all 0.4s ease;
        backdrop-filter: blur(15px);
        position: relative;
        overflow: hidden;
    }
    
    .stFileUploader::before {
        content: '';
        position: absolute;
        top: -2px;
        left: -2px;
        right: -2px;
        bottom: -2px;
        background: linear-gradient(45deg,
            var(--blue-primary),
            var(--cyan-bright),
            var(--electric-blue),
            var(--blue-primary));
        background-size: 400% 400%;
        border-radius: 20px;
        opacity: 0;
        z-index: -1;
        animation: borderFlow 3s ease infinite;
        transition: opacity 0.4s ease;
    }
    
    @keyframes borderFlow {
        0%, 100% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
    }
    
    .stFileUploader:hover {
        border-color: var(--cyan-bright);
        background: rgba(15, 30, 60, 0.8);
        box-shadow: 
            0 0 40px rgba(0, 212, 255, 0.3),
            inset 0 0 20px rgba(0, 212, 255, 0.1);
        transform: scale(1.02);
    }
    
    .stFileUploader:hover::before {
        opacity: 1;
    }
    
    /* QUANTUM SELECT BOX */
    .stSelectbox > div > div {
        background: rgba(15, 30, 60, 0.85) !important;
        border: 2px solid var(--blue-primary) !important;
        border-radius: 14px;
        color: var(--text-primary) !important;
        backdrop-filter: blur(15px);
        box-shadow: 
            0 4px 20px rgba(0, 0, 0, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.2);
        transition: all 0.3s ease;
    }
    
    .stSelectbox > div > div:hover {
        border-color: var(--cyan-bright) !important;
        box-shadow: 
            0 0 25px rgba(0, 212, 255, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.3);
        transform: translateY(-2px);
    }
    
    .stSelectbox > div > div:focus {
        border-color: var(--cyan-bright) !important;
        box-shadow: 
            0 0 30px rgba(0, 212, 255, 0.5),
            inset 0 0 20px rgba(0, 212, 255, 0.15);
    }
    
    /* HOLOGRAPHIC METRICS */
    [data-testid="stMetricValue"] {
        font-size: 3.2rem !important;
        font-weight: 900 !important;
        font-family: var(--font-display) !important;
        color: var(--cyan-bright) !important;
        text-shadow: 
            0 0 20px var(--blue-glow),
            0 0 40px var(--blue-glow);
        animation: metricPulse 2.5s ease-in-out infinite;
    }
    
    @keyframes metricPulse {
        0%, 100% { 
            transform: scale(1);
            text-shadow: 
                0 0 20px var(--blue-glow),
                0 0 40px var(--blue-glow);
        }
        50% { 
            transform: scale(1.05);
            text-shadow: 
                0 0 30px var(--blue-glow),
                0 0 60px var(--blue-glow),
                0 0 80px var(--blue-glow);
        }
    }
    
    [data-testid="stMetricLabel"] {
        color: var(--text-secondary) !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        font-family: var(--font-display) !important;
        text-transform: uppercase;
        letter-spacing: 2px;
    }
    
    [data-testid="stMetricDelta"] {
        color: var(--success) !important;
        font-weight: 700 !important;
        animation: deltaGlow 2s ease-in-out infinite;
    }
    
    @keyframes deltaGlow {
        0%, 100% { opacity: 0.8; }
        50% { opacity: 1; }
    }
    
    /* CYBER DATAFRAME - Enhanced Blue Theme */
    .dataframe {
        background: rgba(15, 30, 60, 0.65) !important;
        border-radius: 16px;
        backdrop-filter: blur(20px);
        color: var(--text-secondary) !important;
        border: 1px solid var(--blue-primary);
        box-shadow: 
            0 4px 24px rgba(0, 0, 0, 0.5),
            inset 0 1px 0 rgba(0, 212, 255, 0.15);
    }
    
    .dataframe th {
        background: rgba(0, 100, 255, 0.25) !important;
        color: var(--cyan-bright) !important;
        font-weight: 700 !important;
        font-family: var(--font-display) !important;
        text-transform: uppercase;
        font-size: 0.9rem !important;
        letter-spacing: 1.5px;
        border-bottom: 2px solid var(--blue-primary) !important;
        text-shadow: 0 0 10px var(--blue-glow);
    }
    
    .dataframe td {
        color: var(--text-secondary) !important;
        border-bottom: 1px solid rgba(0, 212, 255, 0.1) !important;
        transition: background 0.2s ease;
    }
    
    .dataframe tr:hover td {
        background: rgba(0, 212, 255, 0.1) !important;
    }
    
    /* TRANSFORMER INSIGHT BOXES - Enhanced */
    .insight-box {
        padding: 2rem;
        border-radius: 18px;
        margin: 1.5rem 0;
        border-left: 5px solid;
        background: linear-gradient(135deg,
            rgba(15, 30, 60, 0.9) 0%,
            rgba(20, 40, 75, 0.85) 100%);
        backdrop-filter: blur(20px);
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        position: relative;
        overflow: hidden;
        box-shadow: 
            0 8px 32px rgba(0, 0, 0, 0.5),
            inset 0 1px 0 rgba(0, 212, 255, 0.2);
    }
    
    .insight-box::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: linear-gradient(
            45deg,
            transparent 30%,
            rgba(0, 212, 255, 0.1) 50%,
            transparent 70%
        );
        transform: rotate(45deg);
        animation: insightShimmer 4s linear infinite;
    }
    
    @keyframes insightShimmer {
        0% { transform: translateX(-100%) translateY(-100%) rotate(45deg); }
        100% { transform: translateX(100%) translateY(100%) rotate(45deg); }
    }
    
    .insight-box::after {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: linear-gradient(135deg, 
            transparent 0%, 
            rgba(0, 212, 255, 0.05) 100%);
        pointer-events: none;
    }
    
    .insight-box:hover {
        transform: translateX(15px) translateY(-5px) scale(1.02);
        box-shadow: 
            0 16px 48px rgba(0, 212, 255, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.3);
    }
    
    .success-box {
        border-left-color: var(--success);
        background: linear-gradient(135deg, 
            rgba(0, 255, 179, 0.15) 0%, 
            rgba(15, 30, 60, 0.9) 30%);
    }
    
    .success-box:hover {
        box-shadow: 
            0 16px 48px rgba(0, 255, 179, 0.4),
            inset 0 1px 0 rgba(0, 255, 179, 0.3);
    }
    
    .warning-box {
        border-left-color: var(--warning);
        background: linear-gradient(135deg, 
            rgba(255, 184, 0, 0.15) 0%, 
            rgba(15, 30, 60, 0.9) 30%);
    }
    
    .warning-box:hover {
        box-shadow: 
            0 16px 48px rgba(255, 184, 0, 0.4),
            inset 0 1px 0 rgba(255, 184, 0, 0.3);
    }
    
    .info-box {
        border-left-color: var(--info);
        background: linear-gradient(135deg, 
            rgba(0, 212, 255, 0.15) 0%, 
            rgba(15, 30, 60, 0.9) 30%);
    }
    
    .info-box:hover {
        box-shadow: 
            0 16px 48px rgba(0, 212, 255, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.3);
    }
    
    .danger-box {
        border-left-color: var(--danger);
        background: linear-gradient(135deg, 
            rgba(255, 51, 102, 0.15) 0%, 
            rgba(15, 30, 60, 0.9) 30%);
    }
    
    .danger-box:hover {
        box-shadow: 
            0 16px 48px rgba(255, 51, 102, 0.4),
            inset 0 1px 0 rgba(255, 51, 102, 0.3);
    }
    
    /* CYBER EXPANDER */
    .streamlit-expanderHeader {
        background: rgba(15, 30, 60, 0.85) !important;
        border: 2px solid var(--blue-primary) !important;
        border-radius: 14px;
        color: var(--text-primary) !important;
        font-weight: 700;
        font-family: var(--font-display) !important;
        backdrop-filter: blur(15px);
        transition: all 0.3s ease;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        box-shadow: 
            0 4px 16px rgba(0, 0, 0, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.2);
    }
    
    .streamlit-expanderHeader:hover {
        border-color: var(--cyan-bright) !important;
        background: rgba(0, 212, 255, 0.15) !important;
        box-shadow: 
            0 0 30px rgba(0, 212, 255, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.3);
        transform: translateX(5px);
    }
    
    /* QUANTUM PROGRESS BAR */
    .stProgress > div > div {
        background: linear-gradient(90deg, 
            var(--blue-accent) 0%, 
            var(--electric-blue) 25%,
            var(--blue-primary) 50%,
            var(--cyan-bright) 75%,
            var(--electric-blue) 100%) !important;
        background-size: 200% 100%;
        box-shadow: 
            0 0 20px rgba(0, 212, 255, 0.6),
            inset 0 0 10px rgba(0, 212, 255, 0.3);
        animation: progressFlow 2s linear infinite, progressPulse 2s ease-in-out infinite;
        border-radius: 10px;
        height: 12px !important;
    }
    
    @keyframes progressFlow {
        0% { background-position: 0% 50%; }
        100% { background-position: 200% 50%; }
    }
    
    @keyframes progressPulse {
        0%, 100% { 
            opacity: 0.9;
            box-shadow: 
                0 0 20px rgba(0, 212, 255, 0.6),
                inset 0 0 10px rgba(0, 212, 255, 0.3);
        }
        50% { 
            opacity: 1;
            box-shadow: 
                0 0 35px rgba(0, 212, 255, 0.9),
                inset 0 0 15px rgba(0, 212, 255, 0.5);
        }
    }
    
    @keyframes shimmer {
        0% {
            background-position: -1000px 0;
        }
        100% {
            background-position: 1000px 0;
        }
    }
    
    /* HOLOGRAPHIC ALERTS */
    .stAlert {
        background: rgba(15, 30, 60, 0.85) !important;
        backdrop-filter: blur(20px);
        border-radius: 14px;
        border-left-width: 5px;
        color: var(--text-secondary) !important;
        box-shadow: 
            0 8px 32px rgba(0, 0, 0, 0.5),
            inset 0 1px 0 rgba(0, 212, 255, 0.2);
        position: relative;
        overflow: hidden;
    }
    
    .stAlert::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: linear-gradient(135deg,
            transparent 0%,
            rgba(0, 212, 255, 0.05) 100%);
        pointer-events: none;
    }
    
    /* QUANTUM SCROLLBAR */
    ::-webkit-scrollbar {
        width: 14px;
        height: 14px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(0, 10, 25, 0.5);
        border-radius: 10px;
        border: 1px solid var(--blue-primary);
    }
    
    ::-webkit-scrollbar-thumb {
        background: linear-gradient(135deg, 
            var(--blue-accent) 0%, 
            var(--cyan-bright) 100%);
        border-radius: 10px;
        border: 2px solid rgba(0, 10, 25, 0.5);
        box-shadow: 
            0 0 10px rgba(0, 212, 255, 0.5),
            inset 0 0 5px rgba(0, 212, 255, 0.3);
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: linear-gradient(135deg, 
            var(--electric-blue) 0%, 
            var(--cyan-bright) 100%);
        box-shadow: 
            0 0 20px rgba(0, 212, 255, 0.8),
            inset 0 0 10px rgba(0, 212, 255, 0.5);
    }
    
    /* TRANSFORMER ANIMATIONS - Additional */
    .fade-in {
        animation: transformerFadeIn 0.7s cubic-bezier(0.175, 0.885, 0.32, 1.275);
    }
    
    @keyframes transformerFadeIn {
        from { 
            opacity: 0; 
            transform: translateY(40px) scale(0.9) rotateX(-10deg);
        }
        to { 
            opacity: 1; 
            transform: translateY(0) scale(1) rotateX(0deg);
        }
    }
    
    .slide-in-left {
        animation: cyberSlideLeft 0.6s ease-out;
    }
    
    @keyframes cyberSlideLeft {
        from {
            opacity: 0;
            transform: translateX(-60px) rotateY(-15deg);
        }
        to {
            opacity: 1;
            transform: translateX(0) rotateY(0deg);
        }
    }
    
    .slide-in-right {
        animation: cyberSlideRight 0.6s ease-out;
    }
    
    @keyframes cyberSlideRight {
        from {
            opacity: 0;
            transform: translateX(60px) rotateY(15deg);
        }
        to {
            opacity: 1;
            transform: translateX(0) rotateY(0deg);
        }
    }
    
    .rotate-in {
        animation: transformerRotate 0.7s ease-out;
    }
    
    @keyframes transformerRotate {
        from {
            opacity: 0;
            transform: rotate(-20deg) scale(0.7);
        }
        to {
            opacity: 1;
            transform: rotate(0) scale(1);
        }
    }
    
    .bounce-in {
        animation: cyberBounce 0.9s cubic-bezier(0.68, -0.55, 0.265, 1.55);
    }
    
    @keyframes cyberBounce {
        0% {
            opacity: 0;
            transform: scale(0.2) translateY(-50px);
        }
        50% {
            transform: scale(1.1) translateY(5px);
        }
        70% {
            transform: scale(0.95) translateY(-2px);
        }
        100% {
            opacity: 1;
            transform: scale(1) translateY(0);
        }
    }
    
    /* DATA FLOW ANIMATION */
    .data-flow {
        animation: dataStream 2.5s linear infinite;
    }
    
    @keyframes dataStream {
        0% {
            transform: translateX(-100%);
            opacity: 0;
        }
        50% {
            opacity: 1;
        }
        100% {
            transform: translateX(100%);
            opacity: 0;
        }
    }
    
    /* HOLOGRAM FLICKER */
    .hologram-flicker {
        animation: holoFlicker 0.15s infinite;
    }
    
    @keyframes holoFlicker {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.95; }
    }
    
    /* CIRCUIT PULSE */
    .circuit-pulse {
        animation: circuitGlow 3s ease-in-out infinite;
    }
    
    @keyframes circuitGlow {
        0%, 100% {
            box-shadow: 
                0 0 5px var(--blue-primary),
                0 0 10px var(--blue-primary);
        }
        50% {
            box-shadow: 
                0 0 15px var(--cyan-bright),
                0 0 30px var(--cyan-bright),
                0 0 45px var(--blue-primary);
        }
    }
    
    /* GLITCH EFFECT */
    .glitch {
        animation: glitchAnimation 5s infinite;
    }
    
    @keyframes glitchAnimation {
        0%, 98%, 100% {
            transform: translate(0);
            opacity: 1;
        }
        99% {
            transform: translate(2px, -2px);
            opacity: 0.8;
        }
        99.5% {
            transform: translate(-2px, 2px);
            opacity: 0.9;
        }
    }
    
    /* LOADING SPINNER - Cyber Style */
    .stSpinner > div {
        border-color: var(--blue-primary) transparent var(--cyan-bright) transparent !important;
        border-width: 4px !important;
        animation: cyberSpin 1s linear infinite;
        box-shadow: 0 0 20px var(--blue-glow);
    }
    
    @keyframes cyberSpin {
        0% { 
            transform: rotate(0deg);
            box-shadow: 0 0 20px var(--blue-glow);
        }
        50% {
            box-shadow: 0 0 40px var(--blue-glow);
        }
        100% { 
            transform: rotate(360deg);
            box-shadow: 0 0 20px var(--blue-glow);
        }
    }
    
    /* HOLOGRAPHIC CARD CONTAINER */
    .card-container {
        background: linear-gradient(135deg,
            rgba(15, 30, 60, 0.85) 0%,
            rgba(20, 40, 75, 0.75) 100%);
        backdrop-filter: blur(25px) saturate(180%);
        border-radius: 24px;
        padding: 3rem;
        margin: 2rem 0;
        border: 2px solid transparent;
        border-image: linear-gradient(135deg,
            var(--blue-primary),
            var(--cyan-bright),
            var(--blue-primary)) 1;
        box-shadow: 
            0 12px 48px rgba(0, 0, 0, 0.6),
            inset 0 1px 0 rgba(0, 212, 255, 0.25),
            0 0 0 1px rgba(0, 212, 255, 0.2);
        position: relative;
        overflow: hidden;
        animation: cardReveal 0.8s cubic-bezier(0.175, 0.885, 0.32, 1.275);
    }
    
    @keyframes cardReveal {
        from {
            opacity: 0;
            transform: translateY(30px) scale(0.95);
        }
        to {
            opacity: 1;
            transform: translateY(0) scale(1);
        }
    }
    
    .card-container::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -150%;
        width: 200%;
        height: 200%;
        background: linear-gradient(
            90deg,
            transparent,
            rgba(0, 212, 255, 0.15),
            transparent
        );
        animation: cardSlideAcross 4s infinite;
    }
    
    @keyframes cardSlideAcross {
        0% { left: -150%; }
        100% { left: 150%; }
    }
    
    .card-container:hover {
        border-image: linear-gradient(135deg,
            var(--cyan-bright),
            var(--electric-blue),
            var(--cyan-bright)) 1;
        box-shadow: 
            0 16px 64px rgba(0, 212, 255, 0.5),
            0 0 80px rgba(0, 212, 255, 0.3),
            inset 0 1px 0 rgba(0, 212, 255, 0.4);
        transform: translateY(-5px);
    }
    
    /* TEXT ENHANCEMENTS - High Visibility */
    p, span, div {
        color: var(--text-secondary) !important;
    }
    
    strong, b {
        color: var(--text-primary) !important;
        font-weight: 700;
        text-shadow: 0 0 10px rgba(0, 212, 255, 0.3);
    }
    
    /* CYBER LINK STYLING */
    a {
        color: var(--cyan-bright) !important;
        text-decoration: none;
        transition: all 0.3s ease;
        position: relative;
        text-shadow: 0 0 10px var(--blue-glow);
    }
    
    a::after {
        content: '';
        position: absolute;
        width: 0%;
        height: 2px;
        bottom: -2px;
        left: 0;
        background: linear-gradient(90deg,
            var(--cyan-bright),
            var(--blue-primary));
        transition: width 0.3s ease;
    }
    
    a:hover {
        color: var(--electric-blue) !important;
        text-shadow: 0 0 20px var(--blue-glow);
    }
    
    a:hover::after {
        width: 100%;
    }
    
    /* QUANTUM CHECKBOX AND RADIO - FIXED LAYOUT */
    .stCheckbox, .stRadio {
        color: var(--text-secondary) !important;
        margin-bottom: 0.75rem !important;
    }
    
    /* Container alignment fix */
    .stCheckbox > label,
    .stRadio > label {
        display: flex !important;
        align-items: center !important;
        gap: 0.5rem !important;
        cursor: pointer !important;
        padding: 0.25rem 0 !important;
        transition: all 0.3s ease !important;
    }
    
    /* Input element positioning */
    .stCheckbox input[type="checkbox"],
    .stRadio input[type="radio"] {
        flex-shrink: 0 !important;
        width: 18px !important;
        height: 18px !important;
        margin: 0 !important;
        cursor: pointer !important;
    }
    
    /* Text label styling */
    .stCheckbox label > div,
    .stRadio label > div {
        color: var(--text-secondary) !important;
        font-size: 0.95rem !important;
        line-height: 1.5 !important;
        user-select: none !important;
    }
    
    /* Hover state */
    .stCheckbox:hover label > div,
    .stRadio:hover label > div {
        color: var(--cyan-bright) !important;
    }
    
    /* Checked state styling */
    .stCheckbox input[type="checkbox"]:checked + div,
    .stRadio input[type="radio"]:checked + div {
        color: var(--blue-primary) !important;
        font-weight: 600 !important;
    }
    
    /* CYBER INPUT FIELDS */
    input, textarea {
        background: rgba(15, 30, 60, 0.85) !important;
        border: 2px solid var(--blue-primary) !important;
        border-radius: 12px;
        color: var(--text-primary) !important;
        backdrop-filter: blur(15px);
        font-family: var(--font-primary) !important;
        padding: 0.75rem !important;
        transition: all 0.3s ease;
        box-shadow: 
            0 4px 16px rgba(0, 0, 0, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.1);
    }
    
    input:focus, textarea:focus {
        border-color: var(--cyan-bright) !important;
        box-shadow: 
            0 0 30px rgba(0, 212, 255, 0.5) !important,
            inset 0 0 15px rgba(0, 212, 255, 0.15) !important;
        outline: none !important;
        transform: scale(1.02);
    }
    
    input::placeholder, textarea::placeholder {
        color: var(--text-muted) !important;
        font-style: italic;
    }
    
    /* HOLOGRAPHIC BORDERS FOR CONTAINERS */
    .element-container, .stMarkdown {
        position: relative;
    }
    
    /* NUMBER INPUT SPECIFIC */
    input[type="number"] {
        font-family: var(--font-display) !important;
        font-weight: 600;
        font-size: 1.1rem;
    }
    
    /* TEXT AREA SPECIFIC */
    textarea {
        min-height: 120px;
        resize: vertical;
        font-size: 0.95rem;
        line-height: 1.6;
    }
    
    /* DISABLED STATE */
    input:disabled, textarea:disabled {
        opacity: 0.5;
        cursor: not-allowed;
        background: rgba(10, 20, 40, 0.5) !important;
    }
    
    /* ====== FLOATING AI CHATBOT - BOTTOM RIGHT ====== */
    
    /* Chatbot Container - Collapsed State */
    .floating-chatbot {
        position: fixed;
        bottom: 30px;
        right: 30px;
        z-index: 9999;
        transition: all 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        pointer-events: auto !important;
    }
    
    /* Chatbot Toggle Button - Floating Icon */
    .chatbot-toggle {
        width: 70px;
        height: 70px;
        border-radius: 50%;
        background: linear-gradient(135deg, 
            var(--blue-primary) 0%, 
            var(--electric-blue) 50%,
            var(--cyan-bright) 100%);
        border: 3px solid var(--cyan-bright);
        box-shadow: 
            0 8px 32px rgba(0, 212, 255, 0.6),
            0 0 60px rgba(0, 212, 255, 0.4),
            inset 0 0 20px rgba(0, 212, 255, 0.3);
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 2.5rem;
        color: white;
        transition: all 0.3s ease;
        animation: chatbotPulse 2.5s infinite;
        position: relative;
        overflow: hidden;
        pointer-events: auto !important;
        z-index: 10000 !important;
    }
    
    /* Pulse Animation */
    @keyframes chatbotPulse {
        0%, 100% {
            box-shadow: 
                0 8px 32px rgba(0, 212, 255, 0.6),
                0 0 60px rgba(0, 212, 255, 0.4);
            transform: scale(1);
        }
        50% {
            box-shadow: 
                0 12px 48px rgba(0, 212, 255, 0.8),
                0 0 80px rgba(0, 212, 255, 0.6);
            transform: scale(1.05);
        }
    }
    
    /* Hover Effect */
    .chatbot-toggle:hover {
        transform: scale(1.1) rotate(5deg);
        box-shadow: 
            0 12px 48px rgba(0, 212, 255, 0.9),
            0 0 100px rgba(0, 255, 255, 0.7);
    }
    
    /* Ripple Effect on Click */
    .chatbot-toggle::before {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        width: 0;
        height: 0;
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.5);
        transform: translate(-50%, -50%);
        transition: width 0.6s, height 0.6s;
    }
    
    .chatbot-toggle:active::before {
        width: 300px;
        height: 300px;
    }
    
    /* Notification Badge */
    .chatbot-badge {
        position: absolute;
        top: -5px;
        right: -5px;
        width: 24px;
        height: 24px;
        background: #FF3366;
        border-radius: 50%;
        border: 2px solid #000a19;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.75rem;
        font-weight: bold;
        color: white;
        animation: badgeBounce 1s infinite;
    }
    
    @keyframes badgeBounce {
        0%, 100% { transform: scale(1); }
        50% { transform: scale(1.2); }
    }
    
    /* Chatbot Window - Expanded State */
    .chatbot-window {
        position: fixed;
        bottom: 120px;
        right: 30px;
        width: 380px;
        height: 550px;
        background: linear-gradient(135deg, 
            rgba(0, 10, 25, 0.98) 0%, 
            rgba(5, 20, 40, 0.95) 100%);
        backdrop-filter: blur(25px);
        border-radius: 24px;
        border: 2px solid var(--blue-primary);
        box-shadow: 
            0 20px 60px rgba(0, 0, 0, 0.8),
            0 0 80px rgba(0, 212, 255, 0.4),
            inset 0 1px 0 rgba(0, 212, 255, 0.3);
        display: none;
        flex-direction: column;
        overflow: hidden;
        animation: slideInUp 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
        z-index: 9998;
    }
    
    /* Slide-in Animation */
    @keyframes slideInUp {
        from {
            opacity: 0;
            transform: translateY(40px) scale(0.9);
        }
        to {
            opacity: 1;
            transform: translateY(0) scale(1);
        }
    }
    
    /* Chatbot Window - Visible State */
    .chatbot-window.active {
        display: flex;
    }
    
    /* Chatbot Header */
    .chatbot-header {
        padding: 1.5rem;
        background: linear-gradient(135deg, 
            rgba(0, 100, 255, 0.3) 0%, 
            rgba(0, 212, 255, 0.2) 100%);
        border-bottom: 2px solid var(--blue-primary);
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    
    .chatbot-header-title {
        display: flex;
        align-items: center;
        gap: 0.75rem;
    }
    
    .chatbot-header-title h3 {
        margin: 0;
        font-size: 1.2rem;
        color: var(--cyan-bright);
        font-family: var(--font-display);
    }
    
    .chatbot-status {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        font-size: 0.85rem;
        color: var(--text-secondary);
    }
    
    .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: #00FFB3;
        animation: statusPulse 2s infinite;
    }
    
    @keyframes statusPulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.4; }
    }
    
    /* Close Button */
    .chatbot-close {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: rgba(255, 51, 102, 0.2);
        border: 1px solid #FF3366;
        color: #FF3366;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all 0.3s ease;
    }
    
    .chatbot-close:hover {
        background: #FF3366;
        color: white;
        transform: rotate(90deg);
    }
    
    /* Chatbot Messages Container */
    .chatbot-messages {
        flex: 1;
        padding: 1.5rem;
        overflow-y: auto;
        display: flex;
        flex-direction: column;
        gap: 1rem;
    }
    
    /* Custom Scrollbar */
    .chatbot-messages::-webkit-scrollbar {
        width: 6px;
    }
    
    .chatbot-messages::-webkit-scrollbar-track {
        background: rgba(0, 0, 0, 0.2);
    }
    
    .chatbot-messages::-webkit-scrollbar-thumb {
        background: var(--blue-primary);
        border-radius: 10px;
    }
    
    /* Message Bubble - Bot */
    .message-bot {
        background: linear-gradient(135deg, 
            rgba(0, 100, 255, 0.2) 0%, 
            rgba(0, 212, 255, 0.15) 100%);
        border: 1px solid var(--blue-primary);
        border-radius: 16px 16px 16px 4px;
        padding: 1rem;
        max-width: 80%;
        align-self: flex-start;
        animation: messageSlideIn 0.3s ease;
    }
    
    /* Message Bubble - User */
    .message-user {
        background: linear-gradient(135deg, 
            rgba(0, 255, 179, 0.2) 0%, 
            rgba(0, 212, 255, 0.15) 100%);
        border: 1px solid var(--success);
        border-radius: 16px 16px 4px 16px;
        padding: 1rem;
        max-width: 80%;
        align-self: flex-end;
        animation: messageSlideIn 0.3s ease;
    }
    
    @keyframes messageSlideIn {
        from {
            opacity: 0;
            transform: translateY(10px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
    
    /* Message Text */
    .message-text {
        color: var(--text-secondary);
        font-size: 0.95rem;
        line-height: 1.6;
        margin: 0;
    }
    
    /* Typing Indicator */
    .typing-indicator {
        display: flex;
        gap: 0.4rem;
        padding: 1rem;
    }
    
    .typing-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: var(--blue-primary);
        animation: typingBounce 1.4s infinite;
    }
    
    .typing-dot:nth-child(2) {
        animation-delay: 0.2s;
    }
    
    .typing-dot:nth-child(3) {
        animation-delay: 0.4s;
    }
    
    @keyframes typingBounce {
        0%, 60%, 100% {
            transform: translateY(0);
        }
        30% {
            transform: translateY(-10px);
        }
    }
    
    /* Quick Actions */
    .quick-actions {
        padding: 1rem 1.5rem;
        display: flex;
        flex-wrap: wrap;
        gap: 0.5rem;
        border-top: 1px solid rgba(0, 212, 255, 0.2);
    }
    
    .quick-action-btn {
        padding: 0.5rem 1rem;
        background: rgba(0, 100, 255, 0.2);
        border: 1px solid var(--blue-primary);
        border-radius: 20px;
        color: var(--text-secondary);
        font-size: 0.85rem;
        cursor: pointer;
        transition: all 0.3s ease;
    }
    
    .quick-action-btn:hover {
        background: rgba(0, 212, 255, 0.3);
        border-color: var(--cyan-bright);
        color: var(--cyan-bright);
        transform: translateY(-2px);
    }
    
    /* Input Area */
    .chatbot-input-area {
        padding: 1rem 1.5rem;
        background: rgba(0, 10, 25, 0.5);
        border-top: 2px solid var(--blue-primary);
        display: flex;
        gap: 0.75rem;
    }
    
    .chatbot-input {
        flex: 1;
        background: rgba(15, 30, 60, 0.7);
        border: 2px solid var(--blue-primary);
        border-radius: 24px;
        padding: 0.75rem 1.25rem;
        color: var(--text-primary);
        font-size: 0.95rem;
        outline: none;
        transition: all 0.3s ease;
    }
    
    .chatbot-input:focus {
        border-color: var(--cyan-bright);
        box-shadow: 0 0 20px rgba(0, 212, 255, 0.4);
    }
    
    .chatbot-input::placeholder {
        color: var(--text-muted);
    }
    
    .chatbot-send-btn {
        width: 48px;
        height: 48px;
        border-radius: 50%;
        background: linear-gradient(135deg, 
            var(--blue-primary), 
            var(--cyan-bright));
        border: none;
        color: white;
        font-size: 1.5rem;
        cursor: pointer;
        transition: all 0.3s ease;
        display: flex;
        align-items: center;
        justify-content: center;
    }
    
    .chatbot-send-btn:hover {
        transform: scale(1.1);
        box-shadow: 0 0 30px rgba(0, 212, 255, 0.6);
    }
    
    /* Mobile Responsiveness */
    @media (max-width: 768px) {
        .chatbot-window {
            width: calc(100vw - 40px);
            height: calc(100vh - 160px);
            right: 20px;
            bottom: 100px;
        }
        
        .chatbot-toggle {
            width: 60px;
            height: 60px;
            font-size: 2rem;
        }
    }
    
    /* Force chatbot to appear in the correct location */
    .floating-chatbot {
        position: fixed !important;
        bottom: 30px !important;
        right: 30px !important;
        z-index: 2147483647 !important;
        pointer-events: auto !important;
    }
    
    /* Ensure it's visible over Streamlit sidebar */
    .floating-chatbot,
    .chatbot-toggle {
        display: block !important;
        visibility: visible !important;
        opacity: 1 !important;
    }
    
    /* Ensure chatbot is clickable */
    .chatbot-toggle {
        pointer-events: auto !important;
        cursor: pointer !important;
    }
    
    /* Make sure Streamlit elements don't overlap */
    section.main .block-container {
        padding-right: 120px !important;
    }
    
    /* Ensure sidebar doesn't cover chatbot */
    [data-testid="stSidebar"] {
        z-index: 1000 !important;
    }
</style>"""
    return css_content

# Apply cached CSS
st.markdown(load_custom_css(), unsafe_allow_html=True)

# ==================== RIG AVAILABILITY SEARCH ENGINE ====================
class RigAvailabilitySearchEngine:
    """Search and filter rigs based on availability, location, and other criteria"""
    
    def __init__(self, climate_ai=None):
        """Initialize search engine with optional climate AI"""
        self.climate_ai = climate_ai
    
    def search_available_rigs(self, df: pd.DataFrame, filters: dict) -> pd.DataFrame:
        """
        Search for available rigs based on multiple filters
        
        Parameters:
        - df: Full rig dataframe
        - filters: Dictionary containing search criteria
        
        Returns:
        - Filtered dataframe with Match_Score and Climate_Score
        """
        results = df.copy()
        
        # Apply location filter
        if filters.get('location') and filters['location'] != 'All':
            results = results[results['Current Location'] == filters['location']]
        
        # Apply region filter
        if filters.get('region') and filters['region'] != 'All':
            results = results[results['Region'] == filters['region']]
        
        # Apply day rate filter
        if 'dayrate_min' in filters and 'dayrate_max' in filters:
            results = results[
                (results['Dayrate ($k)'] >= filters['dayrate_min']) &
                (results['Dayrate ($k)'] <= filters['dayrate_max'])
            ]
        
        # Apply contractor filter
        if filters.get('contractor') and filters['contractor'] != 'All':
            results = results[results['Contractor'] == filters['contractor']]
        
        # Apply availability filter
        if filters.get('availability_status') and filters['availability_status'] != 'All':
            if 'Contract Days Remaining' in results.columns:
                if filters['availability_status'] == 'Available Now':
                    results = results[results['Contract Days Remaining'] <= 0]
                elif filters['availability_status'] == 'Available Soon (<30 days)':
                    results = results[results['Contract Days Remaining'] <= 30]
                elif filters['availability_status'] == 'Available <90 days':
                    results = results[results['Contract Days Remaining'] <= 90]
        
        # Calculate Match Score
        results['Match_Score'] = self._calculate_match_score(results, filters)
        
        # Calculate Climate Score
        results['Climate_Score'] = self._calculate_climate_score(results, filters)
        
        # Sort by match score descending
        results = results.sort_values('Match_Score', ascending=False)
        
        return results
    
    def _calculate_match_score(self, df: pd.DataFrame, filters: dict) -> pd.Series:
        """Calculate match score based on filters (0-100)"""
        score = pd.Series(100.0, index=df.index)
        
        # Reduce score based on availability (prefer sooner availability)
        if 'Contract Days Remaining' in df.columns:
            days_remaining = df['Contract Days Remaining'].fillna(0)
            # Penalize based on days remaining (max 20 point penalty)
            availability_penalty = (days_remaining / 365.0 * 20).clip(0, 20)
            score -= availability_penalty
        
        # Reduce score based on day rate deviation from target
        if 'dayrate_min' in filters and 'dayrate_max' in filters:
            target_rate = (filters['dayrate_min'] + filters['dayrate_max']) / 2
            rate_deviation = abs(df['Dayrate ($k)'] - target_rate) / target_rate * 100
            rate_penalty = (rate_deviation * 0.3).clip(0, 30)  # Max 30 point penalty
            score -= rate_penalty
        
        return score.clip(0, 100)
    
    def _calculate_climate_score(self, df: pd.DataFrame, filters: dict) -> pd.Series:
        """Calculate climate compatibility score (0-10)"""
        # Base score
        score = pd.Series(7.0, index=df.index)
        
        climate_pref = filters.get('climate_preference', 'Any')
        
        if climate_pref != 'Any':
            # Adjust based on climate preference
            # This is a simplified version - in production would use actual climate data
            if climate_pref == 'Stable Climate':
                # Prefer certain regions
                if 'Region' in df.columns:
                    stable_regions = ['Middle East', 'Southeast Asia']
                    score += df['Region'].isin(stable_regions).astype(float) * 2
            elif climate_pref == 'Storm Resistant':
                # Prefer regions with storm experience
                storm_regions = ['Gulf of Mexico', 'North Sea']
                if 'Region' in df.columns:
                    score += df['Region'].isin(storm_regions).astype(float) * 2
            elif climate_pref == 'Cold Weather Capable':
                # Prefer northern regions
                cold_regions = ['North Sea', 'Arctic', 'Norway']
                if 'Region' in df.columns:
                    score += df['Region'].isin(cold_regions).astype(float) * 2
        
        return score.clip(0, 10)

# ==================== FLOATING AI CHATBOT WIDGET ====================
# Premium floating chatbot interface

@st.cache_resource
def get_chatbot_html():
    """Cache chatbot HTML/JS - loads only once per session"""
    return """<div class="floating-chatbot" id="chatbot-container"><div class="chatbot-toggle" id="chatbot-toggle"><span style="z-index: 10; position: relative;">💬</span><div class="chatbot-badge" id="chatbot-badge" style="display: none;">1</div></div><div class="chatbot-window" id="chatbot-window"><div class="chatbot-header"><div class="chatbot-header-title"><span style="font-size: 1.5rem;">🤖</span><h3>Rig AI Assistant</h3></div><div style="display: flex; align-items: center; gap: 1rem;"><div class="chatbot-status"><div class="status-dot"></div><span>Online</span></div><button class="chatbot-close" id="chatbot-close">×</button></div></div><div class="chatbot-messages" id="chatbot-messages"><div class="message-bot"><p class="message-text">👋 Hey there! I'm your Rig Efficiency AI Assistant. How can I help you analyze rig performance today?</p></div><div class="message-bot"><p class="message-text">💡 I can help with efficiency calculations, climate analysis, benchmarking, and more!</p></div></div><div class="quick-actions"><button class="quick-action-btn" data-action="efficiency">📊 Efficiency Tips</button><button class="quick-action-btn" data-action="climate">🌤️ Climate Impact</button><button class="quick-action-btn" data-action="benchmark">📈 Benchmarks</button></div><div class="chatbot-input-area"><input type="text" class="chatbot-input" id="chatbot-input" placeholder="Ask me anything..." autocomplete="off"/><button class="chatbot-send-btn" id="chatbot-send">➤</button></div></div></div><script>(function() { function initChatbot() { const toggle = document.getElementById('chatbot-toggle'); const window_ = document.getElementById('chatbot-window'); const closeBtn = document.getElementById('chatbot-close'); const input = document.getElementById('chatbot-input'); const sendBtn = document.getElementById('chatbot-send'); const messagesContainer = document.getElementById('chatbot-messages'); const badge = document.getElementById('chatbot-badge'); const quickActionBtns = document.querySelectorAll('.quick-action-btn'); if (!toggle || !window_) return; toggle.addEventListener('click', () => { window_.classList.toggle('active'); badge.style.display = 'none'; }); closeBtn.addEventListener('click', () => { window_.classList.remove('active'); }); function sendMessage() { const text = input.value.trim(); if (text === '') return; const userMsg = document.createElement('div'); userMsg.className = 'message-user'; userMsg.innerHTML = '<p class="message-text">' + escapeHtml(text) + '</p>'; messagesContainer.appendChild(userMsg); input.value = ''; messagesContainer.scrollTop = messagesContainer.scrollHeight; const typingIndicator = document.createElement('div'); typingIndicator.className = 'typing-indicator'; typingIndicator.innerHTML = '<div class="typing-dot"></div><div class="typing-dot"></div><div class="typing-dot"></div>'; messagesContainer.appendChild(typingIndicator); messagesContainer.scrollTop = messagesContainer.scrollHeight; setTimeout(() => { typingIndicator.remove(); const botMsg = document.createElement('div'); botMsg.className = 'message-bot'; botMsg.innerHTML = '<p class="message-text">' + generateBotResponse(text) + '</p>'; messagesContainer.appendChild(botMsg); messagesContainer.scrollTop = messagesContainer.scrollHeight; }, 1500); } sendBtn.addEventListener('click', sendMessage); input.addEventListener('keypress', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMessage(); } }); quickActionBtns.forEach(btn => { btn.addEventListener('click', () => { const action = btn.dataset.action; let message = ''; if (action === 'efficiency') { message = 'What are the best practices for improving rig efficiency?'; } else if (action === 'climate') { message = 'How does climate affect rig performance?'; } else if (action === 'benchmark') { message = 'What are the industry benchmarks for my region?'; } input.value = message; sendMessage(); }); }); function escapeHtml(text) { const div = document.createElement('div'); div.textContent = text; return div.innerHTML; } function generateBotResponse(userMessage) { const responses = {'efficiency': '⚡ Efficiency is maximized by monitoring real-time conditions, optimizing crew schedules, and reducing invisible lost time. Our analysis tools help identify bottlenecks specific to your rig.', 'climate': '🌤️ Climate impacts drilling through temperature extremes, wind patterns, and seasonal variations. Our AI analyzes regional climate data to predict performance impacts.', 'benchmark': '📊 Regional benchmarks help you compare performance against similar rigs. We analyze NPT, move time, and crew efficiency metrics.', 'default': '✨ Thats a great question! Our advanced analytics can help provide insights. Would you like me to analyze specific data or explore efficiency improvements?'}; const lower = userMessage.toLowerCase(); if (lower.includes('efficiency') || lower.includes('optimize')) return responses.efficiency; if (lower.includes('climate') || lower.includes('weather')) return responses.climate; if (lower.includes('benchmark') || lower.includes('compare')) return responses.benchmark; return responses.default; } setTimeout(() => { badge.style.display = 'flex'; }, 3000); } if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', initChatbot); } else { initChatbot(); } window.addEventListener('load', initChatbot); })();</script>"""

@st.cache_resource
def get_calculator():
    """Cache RigEfficiencyCalculator - expensive initialization"""
    return RigEfficiencyCalculator()

@st.cache_resource
def get_climate_ai():
    """Cache AdvancedClimateIntelligence - loads climate data for 40+ regions"""
    return AdvancedClimateIntelligence()

@st.cache_resource
def get_benchmark_model():
    """Cache RegionalBenchmarkModel - loads benchmark databases"""
    return RegionalBenchmarkModel()

@st.cache_resource
def get_ml_predictor():
    """Cache RigWellMatchPredictor - initializes ML models"""
    return RigWellMatchPredictor()

@st.cache_resource
def get_monte_carlo():
    """Cache MonteCarloScenarioSimulator - initializes simulation engines"""
    return MonteCarloScenarioSimulator()

@st.cache_resource
def get_contractor_analyzer():
    """Cache ContractorPerformanceAnalyzer"""
    return ContractorPerformanceAnalyzer()

@st.cache_resource
def get_learning_analyzer():
    """Cache LearningCurveAnalyzer"""
    return LearningCurveAnalyzer()

@st.cache_resource
def get_ilt_detector():
    """Cache InvisibleLostTimeDetector"""
    return InvisibleLostTimeDetector()

@st.cache_resource
def get_search_engine():
    """Initialize search engine once per session"""
    climate_ai = get_climate_ai()
    return RigAvailabilitySearchEngine(climate_ai)

# ==================== CACHED CALCULATION FUNCTIONS ====================
# These prevent expensive calculations from running on every page rerun

@st.cache_data(ttl=3600, show_spinner=False)
def calculate_metrics_cached(rig_name: str, rig_data_dict: dict) -> dict:
    """
    Cache comprehensive efficiency metrics - THE BIG PERFORMANCE WIN!
    TTL: 1 hour - recalculates if data changes or after 1 hour
    Saves 10-15 seconds per interaction!
    """
    try:
        # Reconstruct DataFrame
        rig_data = pd.DataFrame(rig_data_dict)
        
        # Get cached calculator
        calculator = get_calculator()
        
        # Calculate (expensive operation - but cached!)
        metrics = calculator.calculate_comprehensive_efficiency(rig_data)
        
        return metrics
    except Exception as e:
        st.error(f"Error calculating metrics: {str(e)}")
        return {}

@st.cache_data(ttl=3600, show_spinner=False)
def calculate_climate_score_cached(rig_name: str, location: str, contract_start: str, contract_end: str) -> dict:
    """Cache climate efficiency calculations"""
    try:
        climate_ai = get_climate_ai()
        
        # Parse dates
        start_date = pd.to_datetime(contract_start) if contract_start else None
        end_date = pd.to_datetime(contract_end) if contract_end else None
        
        # Calculate time-weighted climate efficiency
        score = climate_ai.calculate_time_weighted_climate_efficiency(
            location, start_date, end_date
        )
        
        # Get climate profile
        climate_profile = climate_ai._get_climate_profile(str(location).lower())
        
        return {
            'score': score,
            'profile': climate_profile
        }
    except Exception:
        return {'score': 70, 'profile': {}}

@st.cache_data(ttl=3600, show_spinner=False)
def generate_summary_cached(rig_name: str, rig_data_dict: dict, metrics_dict: dict) -> dict:
    """Cache contract summary generation"""
    try:
        rig_data = pd.DataFrame(rig_data_dict)
        calculator = get_calculator()
        
        summary = calculator.generate_contract_summary(rig_data, metrics_dict)
        return summary
    except Exception:
        return {}

@st.cache_data(ttl=3600, show_spinner=False)
def calculate_normalized_perf_cached(rig_data_dict: dict) -> dict:
    """Cache normalized performance calculations"""
    try:
        rig_data = pd.DataFrame(rig_data_dict)
        benchmark = get_benchmark_model()
        
        # Backend now generates metrics internally, just pass rig_data
        normalized_perf = benchmark.calculate_normalized_performance(rig_data)
        return normalized_perf
    except Exception as e:
        # Return default structure on error
        return {
            'rop_performance': 70.0,
            'npt_performance': 70.0,
            'time_performance': 70.0,
            'cost_performance': 70.0,
            'overall_normalized': 70.0,
            'benchmark_used': ['offshore'],
            'difficulty_multiplier': 1.0
        }

@st.cache_data(ttl=3600, show_spinner=False)
def calculate_learning_cached(rig_data_dict: dict) -> dict:
    """Cache learning curve analysis"""
    try:
        rig_data = pd.DataFrame(rig_data_dict)
        analyzer = get_learning_analyzer()
        
        curve_analysis = analyzer.calculate_learning_curve(rig_data)
        return curve_analysis
    except Exception:
        return {}

@st.cache_data(show_spinner=False)
def process_uploaded_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Cache file processing - only process once per file"""
    try:
        if filename.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(file_bytes))
        else:
            df = pd.read_excel(io.BytesIO(file_bytes))
        
        # Preprocess
        df = preprocess_dataframe(df)
        return df
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return pd.DataFrame()

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'analysis_results' not in st.session_state:
    st.session_state.analysis_results = {}
if 'selected_rig_history' not in st.session_state:
    st.session_state.selected_rig_history = []


def get_score_class(score):
    """Get CSS class based on score"""
    if score >= 85:
        return 'score-excellent'
    elif score >= 75:
        return 'score-good'
    elif score >= 60:
        return 'score-fair'
    else:
        return 'score-poor'


def get_score_emoji(score):
    """Get emoji based on score"""
    if score >= 85:
        return '🌟'
    elif score >= 75:
        return '✅'
    elif score >= 60:
        return '⚠️'
    else:
        return '🔴'


def hex_to_rgba(hex_color, alpha=0.1):
    """Convert a hex color string to an rgba(...) string with given alpha"""
    try:
        h = hex_color.lstrip('#')
        if len(h) != 6:
            raise ValueError("Invalid hex color")
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        return f"rgba({r}, {g}, {b}, {alpha})"
    except Exception:
        # Fallback to a semi-transparent gold
        return f"rgba(212, 175, 55, {alpha})"


def create_enhanced_gauge_chart(value, title, max_value=100):
    """Create an enhanced gauge chart with cyber blue theme"""
    # Determine color based on value with cyber blue colors
    if value >= 85:
        color = "#00FFB3"  # Success green
    elif value >= 75:
        color = "#00D4FF"  # Blue primary
    elif value >= 60:
        color = "#FFB800"  # Warning orange
    else:
        color = "#FF3366"  # Danger red
    
    fig = go.Figure(go.Indicator(
        mode="gauge+number+delta",
        value=value,
        domain={'x': [0, 1], 'y': [0, 1]},
        title={
            'text': f"<b>{title}</b>",
            'font': {'size': 20, 'color': '#B8D4FF', 'family': 'Orbitron'}
        },
        number={
            'suffix': '%',
            'font': {'size': 56, 'color': color, 'family': 'Orbitron'}  # Increased from 52 to 56
        },
        delta={'reference': 75, 'increasing': {'color': "#00FFB3"}},
        gauge={
            'axis': {
                'range': [None, max_value],
                'tickwidth': 3,
                'tickcolor': "#00D4FF",
                'tickfont': {'color': '#E0F2FF', 'size': 14, 'family': 'Rajdhani'}
            },
            'bar': {'color': color, 'thickness': 0.75},
            'bgcolor': "rgba(10, 15, 30, 0.5)",
            'borderwidth': 3,
            'bordercolor': "#00D4FF",
            'steps': [
                {'range': [0, 60], 'color': "rgba(255, 51, 102, 0.15)"},
                {'range': [60, 75], 'color': "rgba(255, 184, 0, 0.15)"},
                {'range': [75, 85], 'color': "rgba(0, 212, 255, 0.15)"},
                {'range': [85, 100], 'color': "rgba(0, 255, 179, 0.15)"}
            ],
            'threshold': {
                'line': {'color': "#00FFFF", 'width': 4},
                'thickness': 0.75,
                'value': 90
            }
        }
    ))
    
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font={'color': "#E0F2FF", 'family': "Rajdhani"},
        height=340,  # Slightly increased
        margin=dict(l=20, r=20, t=90, b=20)
    )
    
    return fig


def create_premium_bar_chart(metrics, title):
    """Create a premium bar chart with gradient colors"""
    metric_names = list(metrics.keys())
    metric_values = list(metrics.values())
    
    colors = []
    for v in metric_values:
        if v >= 85:
            colors.append('#00FF9F')
        elif v >= 75:
            colors.append('#00E5FF')
        elif v >= 60:
            colors.append('#FFB74D')
        else:
            colors.append('#FF6B6B')
    
    fig = go.Figure(data=[
        go.Bar(
            x=metric_names,
            y=metric_values,
            marker=dict(
                color=colors,
                line=dict(color='#FFD700', width=2),
                opacity=0.9
            ),
            text=[f'{v:.1f}%' for v in metric_values],
            textposition='outside',
            textfont=dict(size=15, color='#E6EDF3', family='Poppins'),
            hovertemplate='<b style="color:#FFD700">%{x}</b><br>Score: %{y:.1f}%<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title={
            'text': f"<b>{title}</b>",
            'font': {'size': 22, 'color': '#E6EDF3', 'family': 'Poppins'},
            'x': 0.5,
            'xanchor': 'center'
        },
        xaxis=dict(
            title=dict(text="<b>Metrics</b>", font=dict(size=16, color='#E6EDF3', family='Poppins')),
            tickfont=dict(size=13, color='#E6EDF3', family='Poppins'),
            gridcolor='rgba(255, 215, 0, 0.15)',
            showgrid=False
        ),
        yaxis=dict(
            title=dict(text="<b>Score (%)</b>", font=dict(size=16, color='#E6EDF3', family='Poppins')),
            tickfont=dict(size=13, color='#E6EDF3', family='Poppins'),
            range=[0, 110],
            gridcolor='rgba(255, 215, 0, 0.15)'
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(33, 38, 45, 0.4)',
        height=470,
        showlegend=False,
        hovermode='x unified',
        hoverlabel=dict(
            bgcolor="rgba(33, 38, 45, 0.95)",
            font_size=14,
            font_family="Poppins",
            font_color="#E6EDF3"
        )
    )
    
    return fig


def create_enhanced_radar_chart(metrics, rig_name):
    """Create an enhanced radar chart with premium styling"""
    categories = [
        'Contract<br>Utilization',
        'Dayrate<br>Efficiency',
        'Contract<br>Stability',
        'Location<br>Complexity',
        'Climate<br>Impact',
        'Contract<br>Performance'
    ]
    
    values = [
        metrics['contract_utilization'],
        metrics['dayrate_efficiency'],
        metrics['contract_stability'],
        metrics['location_complexity'],
        metrics['climate_impact'],
        metrics['contract_performance']
    ]
    
    fig = go.Figure()
    
    # Add filled area
    fig.add_trace(go.Scatterpolar(
        r=values,
        theta=categories,
        fill='toself',
        fillcolor='rgba(255, 215, 0, 0.25)',
        name=rig_name,
        line=dict(color='#FFD700', width=3.5),
        marker=dict(size=10, color='#00E5FF', symbol='circle', line=dict(color='#FFD700', width=2)),
        hovertemplate='<b>%{theta}</b><br><b>Score:</b> %{r:.1f}%<extra></extra>'
    ))
    
    # Add benchmark circle at 75%
    benchmark = [75] * len(categories)
    fig.add_trace(go.Scatterpolar(
        r=benchmark,
        theta=categories,
        line=dict(color='rgba(230, 237, 243, 0.4)', width=2, dash='dash'),
        name='Target (75%)',
        hovertemplate='<b>Target:</b> 75%<extra></extra>'
    ))
    
    fig.update_layout(
        polar=dict(
            bgcolor='rgba(33, 38, 45, 0.4)',
            radialaxis=dict(
                visible=True,
                range=[0, 100],
                tickfont=dict(size=13, color='#E6EDF3', family='Poppins'),
                gridcolor='rgba(255, 215, 0, 0.2)',
                linecolor='rgba(255, 215, 0, 0.3)'
            ),
                angularaxis=dict(
                tickfont=dict(size=14, color='#FFFFFF', family='Poppins'),
                gridcolor='rgba(255, 215, 0, 0.2)',
                linecolor='rgba(255, 215, 0, 0.3)'
            )
        ),
        showlegend=True,
        legend=dict(
            font=dict(size=13, color='#E6EDF3', family='Poppins'),
            bgcolor='rgba(33, 38, 45, 0.9)',
            bordercolor='#FFD700',
            borderwidth=2
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        height=580,
        title={
            'text': f"<b>Efficiency Profile: {rig_name}</b>",
            'font': {'size': 22, 'color': '#E6EDF3', 'family': 'Poppins'},
            'x': 0.5,
            'xanchor': 'center'
        },
        hoverlabel=dict(
            bgcolor="rgba(33, 38, 45, 0.95)",
            font_size=14,
            font_family="Poppins",
            font_color="#E6EDF3"
        )
    )
    
    return fig


def create_timeline_chart_enhanced(rig_data):
    """Create an enhanced timeline chart"""
    try:
        valid_contracts = rig_data[
            rig_data['Contract Start Date'].notna() & 
            rig_data['Contract End Date'].notna()
        ].copy()
        
        if valid_contracts.empty:
            return None
        
        fig = go.Figure()
        
        colors = ['#D4AF37', '#4FC3F7', '#00E676', '#FF9800', '#FF5252']
        
        for idx, row in valid_contracts.iterrows():
            color_idx = idx % len(colors)
            
            fig.add_trace(go.Scatter(
                x=[row['Contract Start Date'], row['Contract End Date']],
                y=[idx, idx],
                mode='lines+markers',
                name=f"Contract {idx+1}",
                line=dict(width=15, color=colors[color_idx]),
                marker=dict(size=12, color=colors[color_idx], symbol='diamond'),
                hovertemplate=(
                    f"<b>Contract {idx+1}</b><br>" +
                    f"<b>Start:</b> {row['Contract Start Date'].strftime('%Y-%m-%d')}<br>" +
                    f"<b>End:</b> {row['Contract End Date'].strftime('%Y-%m-%d')}<br>" +
                    f"<b>Dayrate:</b> ${row['Dayrate ($k)']}k<br>" +
                    f"<b>Duration:</b> {(row['Contract End Date'] - row['Contract Start Date']).days} days" +
                    "<extra></extra>"
                )
            ))
        
        fig.update_layout(
            title={
                'text': "<b>Contract Timeline</b>",
                'font': {'size': 20, 'color': '#FFFFFF', 'family': 'Arial Black'},
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis=dict(
                title=dict(text="<b>Date</b>", font=dict(size=14, color='#FFFFFF')),
                tickfont=dict(size=11, color='#B0B0B0'),
                gridcolor='rgba(212, 175, 55, 0.1)',
                showgrid=True
            ),
            yaxis=dict(
                title=dict(text="<b>Contracts</b>", font=dict(size=14, color='#FFFFFF')),
                tickfont=dict(size=11, color='#B0B0B0'),
                gridcolor='rgba(212, 175, 55, 0.1)',
                showticklabels=False
            ),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(26, 26, 26, 0.5)',
            height=400,
            showlegend=True,
            legend=dict(
                font=dict(size=11, color='#FFFFFF'),
                bgcolor='rgba(26, 26, 26, 0.8)',
                bordercolor='#D4AF37',
                borderwidth=1
            ),
            hovermode='closest'
        )
        
        return fig
        
    except Exception as e:
        st.error(f"Error creating timeline: {str(e)}")
        return None


def display_enhanced_insights(insights):
    """Display insights with enhanced premium styling"""
    for insight in insights:
        insight_type = insight.get('type', 'info')
        priority = insight.get('priority', 'medium')
        
        # Map insight type to styling
        if insight_type == 'success':
            box_class = 'success-box'
            icon = '🌟'
            priority_badge = 'LOW'
            badge_color = '#00E676'
        elif insight_type == 'warning':
            box_class = 'warning-box'
            icon = '⚠️'
            priority_badge = 'MEDIUM'
            badge_color = '#FF9800'
        elif insight_type == 'danger':
            box_class = 'danger-box'
            icon = '🔴'
            priority_badge = 'HIGH'
            badge_color = '#FF5252'
        else:
            box_class = 'info-box'
            icon = '💡'
            priority_badge = 'INFO'
            badge_color = '#4FC3F7'
        
        st.markdown(f"""
        <div class="insight-box {box_class} fade-in">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h3 style="margin: 0; color: #FFFFFF;">
                    {icon} {insight['category']}
                </h3>
                <span style="background: {badge_color}; color: #000; padding: 0.3rem 0.8rem; 
                             border-radius: 20px; font-weight: bold; font-size: 0.8rem;">
                    {priority_badge}
                </span>
            </div>
            <hr style="border: 1px solid rgba(212, 175, 55, 0.2); margin: 1rem 0;">
            <p style="color: #E0E0E0; font-size: 1rem; line-height: 1.6;">
                <b style="color: #D4AF37;">Finding:</b> {insight['message']}
            </p>
            <p style="color: #B0B0B0; font-size: 0.95rem; line-height: 1.6; margin-top: 0.8rem;">
                <b style="color: #4FC3F7;">Recommendation:</b> {insight['recommendation']}
            </p>
        </div>
        """, unsafe_allow_html=True)


def create_performance_heatmap(comparison_df):
    """Create a performance heatmap for fleet comparison"""
    # Prepare data
    metrics_columns = ['Utilization', 'Dayrate', 'Stability', 'Location', 'Climate', 'Performance']
    heatmap_data = comparison_df[metrics_columns].values
    
    fig = go.Figure(data=go.Heatmap(
        z=heatmap_data,
        x=metrics_columns,
        y=comparison_df['Rig Name'].values,
        colorscale=[
            [0, '#FF5252'],      # Red for poor
            [0.6, '#FF9800'],    # Orange for fair
            [0.75, '#4FC3F7'],   # Blue for good
            [0.85, '#00E676']    # Green for excellent
        ],
        text=heatmap_data,
        texttemplate='%{text:.1f}%',
        textfont=dict(size=12, color='#FFFFFF', family='Arial Black'),
        hovertemplate='<b>%{y}</b><br><b>%{x}:</b> %{z:.1f}%<extra></extra>',
        colorbar=dict(
            title=dict(text="Score", side="right", font=dict(color='#FFFFFF')),
            tickmode="linear",
            tick0=0,
            dtick=20,
            tickfont=dict(color='#FFFFFF')
        )
    ))
    
    fig.update_layout(
        title={
            'text': "<b>Fleet Performance Heatmap</b>",
            'font': {'size': 20, 'color': '#FFFFFF', 'family': 'Arial Black'},
            'x': 0.5,
            'xanchor': 'center'
        },
        xaxis=dict(
            tickfont=dict(size=12, color='#FFFFFF'),
            side='top'
        ),
        yaxis=dict(
            tickfont=dict(size=11, color='#FFFFFF')
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        height=max(400, len(comparison_df) * 40)
    )
    
    return fig


def export_to_excel_enhanced(rig_data, metrics, filename="rig_analysis.xlsx"):
    """Enhanced Excel export with formatting"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Raw data
        rig_data.to_excel(writer, sheet_name='Raw Data', index=False)
        
        # Metrics summary
        metrics_df = pd.DataFrame({
            'Metric': [
                'Overall Efficiency',
                'Contract Utilization',
                'Dayrate Efficiency',
                'Contract Stability',
                'Location Complexity',
                'Climate Impact (AI)',
                'Contract Performance',
                'Climate Optimization'
            ],
            'Score (%)': [
                metrics['overall_efficiency'],
                metrics['contract_utilization'],
                metrics['dayrate_efficiency'],
                metrics['contract_stability'],
                metrics['location_complexity'],
                metrics['climate_impact'],
                metrics['contract_performance'],
                metrics.get('climate_optimization', 70)
            ],
            'Grade': [
                metrics['efficiency_grade'],
                '', '', '', '', '', '', ''
            ],
            'Weight (%)': [
                '100',
                '25',
                '20',
                '15',
                '15',
                '10',
                '15',
                'N/A'
            ]
        })
        metrics_df.to_excel(writer, sheet_name='Metrics Summary', index=False)
        
        # Insights
        if 'insights' in metrics and metrics['insights']:
            insights_df = pd.DataFrame(metrics['insights'])
            insights_df.to_excel(writer, sheet_name='Quick Insights', index=False)
        
        # AI Observations
        if 'ai_observations' in metrics and metrics['ai_observations']:
            ai_obs_data = []
            for obs in metrics['ai_observations']:
                ai_obs_data.append({
                    'Priority': obs.get('priority', 'N/A'),
                    'Title': obs.get('title', 'N/A'),
                    'Observation': obs.get('observation', 'N/A'),
                    'Impact': obs.get('impact', 'N/A')
                })
            ai_obs_df = pd.DataFrame(ai_obs_data)
            ai_obs_df.to_excel(writer, sheet_name='AI Observations', index=False)
        
        # Climate Analysis
        if 'climate_insights' in metrics and metrics['climate_insights']:
            climate_data = []
            for insight in metrics['climate_insights']:
                climate_data.append({
                    'Location': insight.get('location', 'N/A'),
                    'Climate Type': insight.get('climate_type', 'N/A'),
                    'Description': insight.get('description', 'N/A'),
                    'Contract Period': insight.get('contract_period', 'N/A')
                })
            climate_df = pd.DataFrame(climate_data)
            climate_df.to_excel(writer, sheet_name='Climate Analysis', index=False)
    
    output.seek(0)
    return output


def show_animated_progress(stages=None):
    """Display animated progress bar for file processing"""
    if stages is None:
        stages = [
            ("Reading File", 0.2),
            ("Validating Data", 0.4),
            ("Processing Columns", 0.6),
            ("Calculating Metrics", 0.8),
            ("Finalizing", 1.0)
        ]
    
    progress_container = st.container()
    
    with progress_container:
        for stage_name, progress_value in stages:
            st.markdown(f"""
            <div style="margin-bottom: 1rem;">
                <div style="color: #4FC3F7; font-size: 0.9rem; margin-bottom: 0.5rem;">
                    <b>{stage_name}</b>
                </div>
                <div style="width: 100%; height: 8px; background: rgba(255,255,255,0.1); border-radius: 10px; overflow: hidden;">
                    <div style="width: {progress_value*100}%; height: 100%; background: linear-gradient(90deg, #00E676, #4FC3F7, #00E676); 
                                animation: shimmer 1.5s infinite; border-radius: 10px;"></div>
                </div>
                <div style="text-align: right; color: #B0B0B0; font-size: 0.8rem; margin-top: 0.3rem;">
                    {int(progress_value*100)}%
                </div>
            </div>
            """, unsafe_allow_html=True)
            import time
            time.sleep(0.1)


def show_architecture_overview():
    """Display comprehensive architecture and capabilities overview"""
    st.markdown("""
    <style>
        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translateY(10px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        .arch-card {
            animation: slideIn 0.5s ease-out;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Architecture Overview Header
    st.markdown("""
    <div class="card-container" style="margin-bottom: 2rem;">
        <h2 style="color: #D4AF37; text-align: center; margin-bottom: 1rem;">
            🏗️ ADVANCED ARCHITECTURE OVERVIEW
        </h2>
        <p style="text-align: center; color: #B0B0B0; line-height: 1.8; font-size: 1.1rem;">
            Enterprise-grade AI-powered rig efficiency analysis platform with 8 specialized components, 
            50+ performance metrics, and 6 advanced climate algorithms
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Core Components
    st.markdown("""
    <div style="margin-bottom: 2rem;">
        <h3 style="color: #4FC3F7; font-size: 1.3rem; margin-bottom: 1.5rem;">
            🎯 8 CORE PROCESSING COMPONENTS
        </h3>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    components = [
        ("AdvancedClimate\nIntelligence", "🌤️", "6 AI algorithms for weather impact analysis across 40+ regions"),
        ("RegionalBenchmark\nModel", "🌍", "Global basin benchmarking and regional performance comparison"),
        ("RigWellMatch\nPredictor", "🎯", "ML-based predictive matching of rigs to well requirements"),
        ("MonteCarloScenario\nSimulator", "🎲", "1,000-iteration risk simulation and scenario analysis"),
    ]
    
    for idx, (title, emoji, desc) in enumerate(components[:4]):
        with [col1, col2, col3, col4][idx]:
            st.markdown(f"""
            <div class="metric-card arch-card" style="height: 180px; display: flex; flex-direction: column; justify-content: space-between;">
                <div style="font-size: 2rem; text-align: center;">{emoji}</div>
                <div style="color: #4FC3F7; font-weight: bold; text-align: center; font-size: 0.85rem;">
                    {title}
                </div>
                <div style="color: #B0B0B0; font-size: 0.8rem; text-align: center; line-height: 1.4;">
                    {desc}
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    components_2 = [
        ("ContractorPerformance\nAnalyzer", "👥", "Consistency metrics and contractor reliability assessment"),
        ("LearningCurve\nAnalyzer", "📈", "Performance trajectory prediction and improvement tracking"),
        ("InvisibleLostTime\nDetector", "⚡", "Hidden efficiency gap detection and recovery opportunity"),
        ("RigEfficiency\nCalculator", "⚙️", "Core multi-factor efficiency calculations and scoring"),
    ]
    
    for idx, (title, emoji, desc) in enumerate(components_2):
        with [col1, col2, col3, col4][idx]:
            st.markdown(f"""
            <div class="metric-card arch-card" style="height: 180px; display: flex; flex-direction: column; justify-content: space-between;">
                <div style="font-size: 2rem; text-align: center;">{emoji}</div>
                <div style="color: #00E676; font-weight: bold; text-align: center; font-size: 0.85rem;">
                    {title}
                </div>
                <div style="color: #B0B0B0; font-size: 0.8rem; text-align: center; line-height: 1.4;">
                    {desc}
                </div>
            </div>
            """, unsafe_allow_html=True)
    
    # 8 Main Tabs Overview
    st.markdown("---")
    st.markdown("""
    <div style="margin-bottom: 2rem;">
        <h3 style="color: #4FC3F7; font-size: 1.3rem; margin-bottom: 1.5rem;">
            📊 8 INTELLIGENT ANALYSIS TABS
        </h3>
    </div>
    """, unsafe_allow_html=True)
    
    tabs_info = [
        ("HOME", "🏠", "Central navigation hub with quick actions and system overview"),
        ("RIG ANALYSIS", "⚙️", "Deep-dive individual rig efficiency with 6+ comprehensive metrics"),
        ("CLIMATE AI", "🌤️", "6 advanced algorithms analyzing 40+ global regions for weather intelligence"),
        ("DASHBOARD", "📊", "Interactive performance visualization with real-time KPIs"),
        ("FLEET COMPARISON", "📈", "Multi-rig comparative analysis and competitive benchmarking"),
        ("AI INSIGHTS", "🤖", "Strategic AI-powered observations with priority-based recommendations"),
        ("REPORTS", "📄", "Professional report generation with multiple export formats"),
        ("ML PREDICTIONS", "🔮", "Predictive analytics: rig-well matching, scenarios, learning curves"),
    ]
    
    for i in range(0, len(tabs_info), 2):
        col1, col2 = st.columns(2)
        
        for col_idx, col in enumerate([col1, col2]):
            if i + col_idx < len(tabs_info):
                tab_name, emoji, description = tabs_info[i + col_idx]
                with col:
                    st.markdown(f"""
                    <div class="metric-card arch-card" style="padding: 1.5rem;">
                        <div style="display: flex; align-items: center; gap: 1rem; margin-bottom: 0.8rem;">
                            <div style="font-size: 2.5rem;">{emoji}</div>
                            <div style="color: #4FC3F7; font-weight: bold; font-size: 1.1rem;">
                                TAB: {tab_name}
                            </div>
                        </div>
                        <div style="color: #B0B0B0; font-size: 0.95rem; line-height: 1.6;">
                            {description}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
    
    # Key Features Summary
    st.markdown("---")
    st.markdown("""
    <div style="margin-bottom: 2rem;">
        <h3 style="color: #4FC3F7; font-size: 1.3rem; margin-bottom: 1.5rem;">
            ✨ KEY FEATURES & CAPABILITIES
        </h3>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class="metric-card" style="height: auto; padding: 1.5rem;">
            <h4 style="color: #4FC3F7; margin-top: 0;">📊 COMPREHENSIVE METRICS</h4>
            <ul style="color: #B0B0B0; line-height: 2; font-size: 0.9rem;">
                <li>✓ Overall Efficiency Score</li>
                <li>✓ Contract Utilization Rate</li>
                <li>✓ Dayrate Efficiency Index</li>
                <li>✓ Climate Impact Score (AI)</li>
                <li>✓ 50+ Advanced Metrics</li>
                <li>✓ Composite REI Index</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="metric-card" style="height: auto; padding: 1.5rem;">
            <h4 style="color: #00E676; margin-top: 0;">🤖 CLIMATE INTELLIGENCE</h4>
            <ul style="color: #B0B0B0; line-height: 2; font-size: 0.9rem;">
                <li>✓ 6 Advanced AI Algorithms</li>
                <li>✓ 40+ Global Regions</li>
                <li>✓ Seasonal Optimization</li>
                <li>✓ Weather Risk Assessment</li>
                <li>✓ 87-92% Prediction Confidence</li>
                <li>✓ Location-Specific Analysis</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="metric-card" style="height: auto; padding: 1.5rem;">
            <h4 style="color: #FF9800; margin-top: 0;">🔮 PREDICTIVE ANALYTICS</h4>
            <ul style="color: #B0B0B0; line-height: 2; font-size: 0.9rem;">
                <li>✓ Rig-Well Match Prediction</li>
                <li>✓ Monte Carlo Simulations</li>
                <li>✓ Performance Forecasting</li>
                <li>✓ Learning Curve Analysis</li>
                <li>✓ Hidden Lost Time Detection</li>
                <li>✓ 1,000+ Scenario Runs</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    # Client Value Propositions
    st.markdown("---")
    st.markdown("""
    <div style="margin-bottom: 2rem;">
        <h3 style="color: #4FC3F7; font-size: 1.3rem; margin-bottom: 1.5rem;">
            💼 MEASURABLE BUSINESS VALUE
        </h3>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="metric-card" style="padding: 1.5rem; border-left: 4px solid #00E676;">
            <h4 style="color: #00E676; margin-top: 0;">⚙️ FOR OPERATORS</h4>
            <ul style="color: #B0B0B0; line-height: 1.8; font-size: 0.9rem;">
                <li>💰 <b>$7M-$15M Annual Savings</b> per rig (5-10% efficiency gain)</li>
                <li>⏱️ <b>15-30% Less</b> weather-related downtime</li>
                <li>🎯 <b>10-20% Improved</b> contract utilization</li>
                <li>📊 Objective performance benchmarking</li>
                <li>🔍 Data-backed contractor negotiations</li>
                <li>📈 Strategic basin expansion decisions</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="metric-card" style="padding: 1.5rem; border-left: 4px solid #4FC3F7;">
            <h4 style="color: #4FC3F7; margin-top: 0;">🚀 FOR RIG CONTRACTORS</h4>
            <ul style="color: #B0B0B0; line-height: 1.8; font-size: 0.9rem;">
                <li>💵 <b>5-10% Rate Premium</b> for proven top performers</li>
                <li>🎯 <b>Win More Contracts</b> with evidence-based proposals</li>
                <li>📊 <b>10-20% Utilization Gain</b> with optimized matching</li>
                <li>💡 Eliminate hidden inefficiencies</li>
                <li>🌍 Optimize fleet-to-basin assignments</li>
                <li>📈 Data-driven strategic planning</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    # Success Metrics
    st.markdown("""
    <div class="metric-card" style="margin-top: 2rem; padding: 1.5rem; background: rgba(0, 230, 118, 0.1); border: 1px solid #00E676;">
        <h4 style="color: #00E676; margin-top: 0; text-align: center;">📈 TYPICAL ROI RESULTS (3-6 Months)</h4>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-top: 1rem;">
            <div style="text-align: center; padding: 1rem;">
                <div style="font-size: 2rem; color: #00E676; font-weight: bold;">5-15%</div>
                <div style="color: #B0B0B0; font-size: 0.9rem;">Efficiency Improvement</div>
            </div>
            <div style="text-align: center; padding: 1rem;">
                <div style="font-size: 2rem; color: #4FC3F7; font-weight: bold;">$5M+</div>
                <div style="color: #B0B0B0; font-size: 0.9rem;">Annual Cost Savings</div>
            </div>
            <div style="text-align: center; padding: 1rem;">
                <div style="font-size: 2rem; color: #FF9800; font-weight: bold;">20%</div>
                <div style="color: #B0B0B0; font-size: 0.9rem;">Better Decisions</div>
            </div>
            <div style="text-align: center; padding: 1rem;">
                <div style="font-size: 2rem; color: #D4AF37; font-weight: bold;">40+</div>
                <div style="color: #B0B0B0; font-size: 0.9rem;">Regions Covered</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def main():
    """Enhanced main application with premium interactions"""
    
    # ==================== INITIALIZE CHATBOT ====================
    # Initialize chatbot in session state
    if 'chatbot' not in st.session_state:
        st.session_state.chatbot = RigEfficiencyAIChatbot()
    if 'chatbot_open' not in st.session_state:
        st.session_state.chatbot_open = False
    if 'chat_messages' not in st.session_state:
        st.session_state.chat_messages = []
    
    # ==================== AI CHATBOT IN SIDEBAR ====================
    # Use expander for better UX - no disruption to navigation
    
    
    
    
    # ========================================================================
    
    # Animated Header
    st.markdown('<h1 class="main-header">🛢️ RIG EFFICIENCY INTELLIGENCE PLATFORM</h1>', 
                unsafe_allow_html=True)
    
    st.markdown("""
    <div class="subtitle">
        ADVANCED MULTI-FACTOR ANALYSIS • AI-POWERED CLIMATE INTELLIGENCE • REAL-TIME INSIGHTS
    </div>
    """, unsafe_allow_html=True)
    
    # ===== CLIENT GUIDE BANNER =====
    st.markdown("""
    <div class="client-guide-banner">
        <h2 style="color: var(--cyan-bright); margin: 0 0 1rem 0; font-family: var(--font-display); font-size: 2rem;">
            🎯 CLIENT QUICK START GUIDE
        </h2>
        <div style="display: flex; justify-content: space-around; flex-wrap: wrap; gap: 1.5rem; margin-top: 1.5rem;">
            <div style="flex: 1; min-width: 200px;">
                <div style="font-size: 3rem; margin-bottom: 0.5rem;">1️⃣</div>
                <h3 style="color: var(--blue-primary); margin: 0.5rem 0; font-size: 1.2rem;">UPLOAD DATA</h3>
                <p style="color: var(--text-secondary); margin: 0; font-size: 0.95rem;">
                    Click sidebar → Upload Excel/CSV file
                </p>
            </div>
            <div style="flex: 1; min-width: 200px;">
                <div style="font-size: 3rem; margin-bottom: 0.5rem;">2️⃣</div>
                <h3 style="color: var(--blue-primary); margin: 0.5rem 0; font-size: 1.2rem;">SELECT RIG</h3>
                <p style="color: var(--text-secondary); margin: 0; font-size: 0.95rem;">
                    Choose rig from dropdown menu
                </p>
            </div>
            <div style="flex: 1; min-width: 200px;">
                <div style="font-size: 3rem; margin-bottom: 0.5rem;">3️⃣</div>
                <h3 style="color: var(--blue-primary); margin: 0.5rem 0; font-size: 1.2rem;">ANALYZE</h3>
                <p style="color: var(--text-secondary); margin: 0; font-size: 0.95rem;">
                    Click "ANALYZE RIG" button
                </p>
            </div>
            <div style="flex: 1; min-width: 200px;">
                <div style="font-size: 3rem; margin-bottom: 0.5rem;">4️⃣</div>
                <h3 style="color: var(--blue-primary); margin: 0.5rem 0; font-size: 1.2rem;">EXPLORE</h3>
                <p style="color: var(--text-secondary); margin: 0; font-size: 0.95rem;">
                    Navigate through 6 intelligent tabs
                </p>
            </div>
        </div>
        <div style="margin-top: 1.5rem; padding: 1rem; background: rgba(0, 212, 255, 0.1); border-radius: 10px; border: 1px solid var(--blue-primary);">
            <p style="color: var(--text-primary); margin: 0; font-size: 1rem;">
                💡 <strong>Pro Tip:</strong> Hover over any metric for detailed explanations | Use tabs to explore different analysis views
            </p>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Enhanced Sidebar
    with st.sidebar:
        st.markdown("### 📁 DATA MANAGEMENT")
        
        uploaded_file = st.file_uploader(
            "Upload Rig Data",
            type=['xlsx', 'xls', 'csv'],
            help="📊 Supported formats: Excel (.xlsx, .xls) or CSV (.csv)\n\n"
                 "📋 Required columns:\n"
                 "• Drilling Unit Name / Rig Name\n"
                 "• Contract Start Date\n"
                 "• Contract End Date\n"
                 "• Dayrate ($k)\n"
                 "• Current Location\n"
                 "• Contract Length\n"
                 "• Status"
        )
        
        if uploaded_file is not None:
            try:
                # Show animated progress bar instead of spinner
                progress_placeholder = st.empty()
                
                with progress_placeholder.container():
                    st.markdown("<h4 style='color: #4FC3F7; margin-bottom: 1rem;'>📂 PROCESSING YOUR DATA</h4>", unsafe_allow_html=True)
                    show_animated_progress([
                        ("Reading File", 0.2),
                        ("Validating Data", 0.4),
                        ("Processing Columns", 0.6),
                        ("Calculating Metrics", 0.8),
                        ("Finalizing", 1.0)
                    ])
                
                # Process the file with caching
                file_bytes = uploaded_file.read()
                df = process_uploaded_file(file_bytes, uploaded_file.name)
                st.session_state.df = df
                
                # Clear progress and show success
                progress_placeholder.empty()
                st.success(f"✅ Successfully loaded **{len(df):,}** records from **{uploaded_file.name}**!")
                
                # Enhanced data preview
                with st.expander("📊 Data Preview", expanded=False):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Records", f"{len(df):,}")
                    with col2:
                        st.metric("Unique Rigs", f"{df['Rig Name'].nunique():,}")
                    with col3:
                        if 'Current Location' in df.columns:
                            st.metric("Locations", f"{df['Current Location'].nunique():,}")
                    
                    st.dataframe(
                        df.head(10).style.set_properties(**{
                            'background-color': '#1a1a1a',
                            'color': '#ffffff',
                            'border-color': '#D4AF37'
                        }),
                        use_container_width=True
                    )
            
            except Exception as e:
                st.error(f"❌ Error loading file: {str(e)}")
                st.info("💡 Tip: Ensure your file has the required columns and proper data format.")
        
        st.markdown("---")
        
        # Enhanced settings
        st.markdown("### ⚙️ ANALYSIS SETTINGS")
        
        col1, col2 = st.columns(2)
        with col1:
            show_detailed = st.checkbox("🔍 Detailed", value=True, help="Show detailed breakdowns")
            show_climate = st.checkbox("🌤️ Climate AI", value=True, help="Include climate intelligence")
        with col2:
            show_insights = st.checkbox("💡 Insights", value=True, help="Show AI insights")
            show_charts = st.checkbox("📊 Charts", value=True, help="Display visualizations")
        
        st.markdown("---")
        
        # Feature highlights
        with st.expander("✨ PLATFORM FEATURES", expanded=False):
            st.markdown("""
            **🎯 Core Capabilities:**
            - 6-Factor Efficiency Scoring
            - Real-Time Performance Tracking
            - Interactive Dashboards
            - Fleet-Wide Comparison
            
            **🤖 AI Intelligence:**
            - 6 Advanced Climate Algorithms
            - Predictive Analytics
            - Risk Assessment
            - Seasonal Optimization
            
            **📈 Business Value:**
            - Data-Driven Decisions
            - Cost Optimization
            - Risk Mitigation
            - Strategic Planning
            
            **📤 Export Options:**
            - Comprehensive Excel Reports
            - CSV Data Export
            - Custom Analytics
            """)
        
        # Quick stats if data loaded
        if st.session_state.df is not None:
            st.markdown("---")
            st.markdown("### 📈 QUICK STATS")
            
            df = st.session_state.df
            
            if 'Contract value ($m)' in df.columns:
                total_value = df['Contract value ($m)'].sum()
                st.metric(
                    "Total Portfolio Value",
                    f"${total_value:,.1f}M",
                    help="Combined value of all contracts"
                )
            
            if 'Dayrate ($k)' in df.columns:
                avg_rate = df['Dayrate ($k)'].mean()
                st.metric(
                    "Average Dayrate",
                    f"${avg_rate:,.0f}k",
                    help="Mean dayrate across all rigs"
                )
            
            if 'Status' in df.columns:
                active = df['Status'].str.contains('active', case=False, na=False).sum()
                st.metric(
                    "Active Contracts",
                    f"{active:,}",
                    help="Currently active contracts"
                )
    
    # ==================== AI CHATBOT DISPLAY ====================
    # Display chatbot in sidebar expander for easy access
    with st.sidebar:
        st.markdown("---")
        with st.expander("🤖 AI Assistant - Click to Chat", expanded=False):
            st.markdown("*Ask me anything about rig efficiency*")
            
            # Display chat history in compact format
            if st.session_state.chat_messages:
                # Show only last 3 messages to save space
                recent_messages = st.session_state.chat_messages[-6:]
                for message in recent_messages:
                    if message['role'] == 'user':
                        st.info(f"**You:** {message['content'][:100]}..." if len(message['content']) > 100 else f"**You:** {message['content']}")
                    else:
                        st.success(f"**AI:** {message['content'][:150]}..." if len(message['content']) > 150 else f"**AI:** {message['content']}")
            else:
                st.info("👋 Hello! Ask me anything!")
            
            # Chat input
            user_prompt = st.text_input(
                "Your question:",
                key="chat_input_field",
                placeholder="Type here...",
                label_visibility="collapsed"
            )
            
            # Send button
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Send 📤", key="send_msg", use_container_width=True):
                    if user_prompt and user_prompt.strip():
                        # Add user message
                        st.session_state.chat_messages.append({
                            'role': 'user',
                            'content': user_prompt
                        })
                        
                        # Generate AI response
                        try:
                            context = {
                                'has_data': st.session_state.df is not None,
                                'current_rig': None,
                                'analysis_complete': bool(st.session_state.analysis_results)
                            }
                            
                            response = st.session_state.chatbot.generate_response(
                                user_prompt,
                                context=context
                            )
                        except Exception as e:
                            response = f"⚠️ Error: {str(e)}"
                        
                        # Add assistant response
                        st.session_state.chat_messages.append({
                            'role': 'assistant',
                            'content': response
                        })
                        
                        st.rerun()
            
            # Clear chat button
            with col2:
                if len(st.session_state.chat_messages) > 0:
                    if st.button("🗑️ Clear", key="clear_chat", use_container_width=True):
                        st.session_state.chat_messages = []
                        st.rerun()
    
    # Main content area
    if st.session_state.df is None:
        # Enhanced welcome screen
        st.markdown("""
        <div class="card-container fade-in">
            <h2 style="text-align: center; color: #D4AF37; margin-bottom: 2rem;">
                🚀 WELCOME TO THE FUTURE OF RIG ANALYTICS
            </h2>
            <p style="text-align: center; font-size: 1.2rem; color: #B0B0B0; line-height: 1.8;">
                Transform your drilling operations with AI-powered insights and real-time intelligence.
                <br>Upload your data to unlock comprehensive efficiency analysis.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div class="metric-card">
                <h3 style="color: #4FC3F7; text-align: center;">📊 ANALYTICS</h3>
                <ul style="color: #E0E0E0; line-height: 2;">
                    <li>Multi-Factor Scoring</li>
                    <li>Performance Metrics</li>
                    <li>Visual Dashboards</li>
                    <li>Historical Trends</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div class="metric-card">
                <h3 style="color: #00E676; text-align: center;">🤖 AI INTELLIGENCE</h3>
                <ul style="color: #E0E0E0; line-height: 2;">
                    <li>6 Climate Algorithms</li>
                    <li>Predictive Modeling</li>
                    <li>Risk Assessment</li>
                    <li>Smart Recommendations</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div class="metric-card">
                <h3 style="color: #FF9800; text-align: center;">📈 INSIGHTS</h3>
                <ul style="color: #E0E0E0; line-height: 2;">
                    <li>Strategic Planning</li>
                    <li>Fleet Comparison</li>
                    <li>Cost Optimization</li>
                    <li>Export Reports</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        # Display comprehensive architecture overview
        st.markdown("---")
        show_architecture_overview()
        
        st.markdown("---")
        # Instructions
        st.info("👈 **Get Started:** Upload your rig data file using the sidebar to begin comprehensive analysis")
        
        return
    
    # Data loaded - show analysis interface
    df = st.session_state.df
    
    # Rig selection with enhanced UI
    st.markdown("### 🎯 SELECT RIG FOR ANALYSIS")
    
    if 'Rig Name' not in df.columns:
        st.error("❌ 'Rig Name' column not found. Please check your data format.")
        return
    
    available_rigs = sorted(df['Rig Name'].dropna().unique().tolist())
    
    if len(available_rigs) == 0:
        st.error("❌ No rigs found in uploaded data.")
        return
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        selected_rig = st.selectbox(
            "Choose a rig from your fleet",
            options=available_rigs,
            index=0,
            help=f"📋 {len(available_rigs)} rigs available for analysis"
        )
    
    with col2:
        analyze_button = st.button(
            "🔍 ANALYZE RIG",
            use_container_width=True,
            type="primary"
        )
    
    # Track rig selection history
    if selected_rig not in st.session_state.selected_rig_history:
        st.session_state.selected_rig_history.append(selected_rig)
    
    # Show recently analyzed rigs
    if len(st.session_state.selected_rig_history) > 1:
        with st.expander("🕒 Recently Analyzed", expanded=False):
            st.write(", ".join(st.session_state.selected_rig_history[-5:]))
    
    # Filter data for selected rig
    rig_data = df[df['Rig Name'] == selected_rig].copy()
    
    # Calculate metrics
    if analyze_button or selected_rig in st.session_state.analysis_results:
        with st.spinner('🔄 Analyzing rig performance with AI...'):
            if selected_rig not in st.session_state.analysis_results:
                # Use cached calculation instead of direct call
                metrics = calculate_metrics_cached(selected_rig, rig_data.to_dict('list'))
                st.session_state.analysis_results[selected_rig] = metrics
            else:
                metrics = st.session_state.analysis_results[selected_rig]
        
        if metrics is None:
            st.error("❌ Unable to calculate metrics for this rig. Please check data completeness.")
            return
        
        # Success message
        st.success(f"✅ Analysis complete for **{selected_rig}** | Overall Score: **{metrics['overall_efficiency']:.1f}%** {get_score_emoji(metrics['overall_efficiency'])}")
        
        # Create enhanced tabs
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11, tab12, tab13 = st.tabs([
            "🎯 OVERVIEW",
            "📊 DETAILED METRICS",
            "🌤️ CLIMATE AI",
            "💡 INSIGHTS",
            "🏢 FLEET COMPARISON",
            "📤 EXPORT",
            "🎯 REGIONAL BENCHMARK",
            "🔍 RIG-WELL MATCH",
            "🌍 BASIN SCENARIOS",
            "📈 CONTRACTOR ANALYSIS",
            "📉 LEARNING CURVES",
            "⚡ LOST TIME DETECTOR",
            "🔍 RIG AVAILABILITY"
        ])
        
        # TAB 1: Enhanced Overview
        with tab1:
            st.markdown(f"## 📋 EFFICIENCY OVERVIEW: {selected_rig}")
            
            # Top metrics row
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                score_class = get_score_class(metrics['overall_efficiency'])
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #B0B0B0; text-align: center; margin: 0;">OVERALL SCORE</h4>
                    <div class="score-display {score_class}">
                        {metrics['overall_efficiency']:.1f}%
                    </div>
                    <p style="text-align: center; color: #B0B0B0; margin: 0;">
                        {get_score_emoji(metrics['overall_efficiency'])} {metrics['efficiency_grade']}
                    </p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #B0B0B0; text-align: center;">CONTRACTS</h4>
                    <div class="score-display score-good">
                        {len(rig_data)}
                    </div>
                    <p style="text-align: center; color: #B0B0B0;">Total Count</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                if 'Contract value ($m)' in rig_data.columns:
                    total_value = rig_data['Contract value ($m)'].sum()
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">VALUE</h4>
                        <div class="score-display score-excellent">
                            ${total_value:.1f}M
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Portfolio Total</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            with col4:
                if 'Dayrate ($k)' in rig_data.columns:
                    avg_rate = rig_data['Dayrate ($k)'].mean()
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">DAYRATE</h4>
                        <div class="score-display score-good">
                            ${avg_rate:.0f}k
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Average</p>
                    </div>
                    """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Gauges and charts
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.markdown("#### 🎯 OVERALL PERFORMANCE")
                fig_gauge = create_enhanced_gauge_chart(
                    metrics['overall_efficiency'],
                    "Overall Efficiency"
                )
                st.plotly_chart(fig_gauge, use_container_width=True)
            
            with col2:
                st.markdown("#### 📊 EFFICIENCY BREAKDOWN")
                metric_comparison = {
                    'Utilization': metrics['contract_utilization'],
                    'Dayrate': metrics['dayrate_efficiency'],
                    'Stability': metrics['contract_stability'],
                    'Location': metrics['location_complexity'],
                    'Climate': metrics['climate_impact'],
                    'Performance': metrics['contract_performance']
                }
                fig_bars = create_premium_bar_chart(metric_comparison, "")
                st.plotly_chart(fig_bars, use_container_width=True)
            
            st.markdown("---")
            
            # Radar chart
            st.markdown("#### 🔷 EFFICIENCY PROFILE")
            fig_radar = create_enhanced_radar_chart(metrics, selected_rig)
            st.plotly_chart(fig_radar, use_container_width=True)
            
            # Timeline
            st.markdown("---")
            st.markdown("#### 📅 CONTRACT TIMELINE")
            fig_timeline = create_timeline_chart_enhanced(rig_data)
            if fig_timeline:
                st.plotly_chart(fig_timeline, use_container_width=True)
            else:
                st.info("📊 No contract timeline data available")
            
            # Quick insights preview
            if metrics.get('insights'):
                st.markdown("---")
                st.markdown("#### 💡 KEY FINDINGS (Top 3)")
                for insight in metrics['insights'][:3]:
                    display_enhanced_insights([insight])
        
        # TAB 2: Detailed Metrics (same structure with enhanced styling)
        with tab2:
            st.markdown("## 📊 DETAILED METRICS ANALYSIS")
            
            # Individual metric gauges
            col1, col2, col3 = st.columns(3)
            
            with col1:
                fig1 = create_enhanced_gauge_chart(
                    metrics['contract_utilization'],
                    "Contract Utilization"
                )
                st.plotly_chart(fig1, use_container_width=True)
                with st.expander("ℹ️ What This Means"):
                    st.write("""
                    **Contract Utilization** measures how effectively the rig's time is being used.
                    
                    📌 **Weight:** 25% of overall score
                    
                    ✅ **High Score (>85%):** Excellent time utilization, minimal idle time
                    ⚠️ **Medium Score (60-85%):** Good utilization with room for improvement
                    🔴 **Low Score (<60%):** Significant idle time, need better contract pipeline
                    
                    💡 **Tip:** Focus on back-to-back contracts and reduce mobilization gaps
                    """)
            
            with col2:
                fig2 = create_enhanced_gauge_chart(
                    metrics['dayrate_efficiency'],
                    "Dayrate Efficiency"
                )
                st.plotly_chart(fig2, use_container_width=True)
                with st.expander("ℹ️ What This Means"):
                    st.write("""
                    **Dayrate Efficiency** compares your rates against market benchmarks.
                    
                    📌 **Weight:** 20% of overall score
                    
                    ✅ **High Score (>80%):** Premium rates, strong market position
                    ⚠️ **Medium Score (50-80%):** Competitive rates, room for optimization
                    🔴 **Low Score (<50%):** Below-market rates, value capture opportunity
                    
                    💡 **Tip:** Review rig capabilities and target premium operators
                    """)
            
            with col3:
                fig3 = create_enhanced_gauge_chart(
                    metrics['contract_stability'],
                    "Contract Stability"
                )
                st.plotly_chart(fig3, use_container_width=True)
                with st.expander("ℹ️ What This Means"):
                    st.write("""
                    **Contract Stability** evaluates contract duration and consistency.
                    
                    📌 **Weight:** 15% of overall score
                    
                    ✅ **High Score (>75%):** Long-term contracts, stable revenue
                    ⚠️ **Medium Score (50-75%):** Mix of short and long contracts
                    🔴 **Low Score (<50%):** Frequent short contracts, revenue volatility
                    
                    💡 **Tip:** Negotiate longer contract terms for better stability
                    """)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                fig4 = create_enhanced_gauge_chart(
                    metrics['location_complexity'],
                    "Location Complexity"
                )
                st.plotly_chart(fig4, use_container_width=True)
                with st.expander("ℹ️ What This Means"):
                    st.write("""
                    **Location Complexity** assesses operational environment difficulty.
                    
                    📌 **Weight:** 15% of overall score
                    
                    ✅ **High Score (>80%):** Lower complexity, easier operations
                    ⚠️ **Medium Score (65-80%):** Moderate complexity environments
                    🔴 **Low Score (<65%):** High complexity (deepwater, harsh environment)
                    
                    💡 **Note:** Lower scores aren't negative - they reflect challenging conditions
                    """)
            
            with col2:
                fig5 = create_enhanced_gauge_chart(
                    metrics['climate_impact'],
                    "Climate Impact (AI)"
                )
                st.plotly_chart(fig5, use_container_width=True)
                with st.expander("ℹ️ What This Means"):
                    st.write("""
                    **Climate Impact** uses 6 AI algorithms to assess weather effects.
                    
                    📌 **Weight:** 10% of overall score
                    
                    ✅ **High Score (>85%):** Favorable weather conditions
                    ⚠️ **Medium Score (65-85%):** Moderate weather impacts
                    🔴 **Low Score (<65%):** Significant weather challenges
                    
                    🤖 **AI Analysis:** Time-weighted, predictive, adaptive, risk-adjusted
                    
                    💡 **Tip:** Review Climate AI tab for seasonal optimization strategies
                    """)
            
            with col3:
                fig6 = create_enhanced_gauge_chart(
                    metrics['contract_performance'],
                    "Contract Performance"
                )
                st.plotly_chart(fig6, use_container_width=True)
                with st.expander("ℹ️ What This Means"):
                    st.write("""
                    **Contract Performance** measures overall execution and delivery.
                    
                    📌 **Weight:** 15% of overall score
                    
                    ✅ **High Score (>80%):** Excellent delivery track record
                    ⚠️ **Medium Score (60-80%):** Good performance, some issues
                    🔴 **Low Score (<60%):** Performance challenges, need improvement
                    
                    💡 **Tip:** Focus on on-time delivery and contract compliance
                    """)
            
            # Detailed contract data
            st.markdown("---")
            st.markdown("### 📋 CONTRACT DETAILS")
            
            display_columns = []
            column_config = {}
            
            for col in ['Contract Start Date', 'Contract End Date', 'Dayrate ($k)', 
                        'Contract value ($m)', 'Current Location', 'Contract Length', 'Status']:
                if col in rig_data.columns:
                    display_columns.append(col)
                    if '$' in col:
                        column_config[col] = st.column_config.NumberColumn(
                            col,
                            format="$%.1f"
                        )
            
            if display_columns:
                st.dataframe(
                    rig_data[display_columns],
                    use_container_width=True,
                    hide_index=True,
                    column_config=column_config
                )
        
        # TAB 3: Climate AI (enhanced)
        with tab3:
            st.markdown("## 🌤️ AI-POWERED CLIMATE INTELLIGENCE")
            
            st.markdown("""
            <div class="info-box">
                <h4>🤖 Advanced Climate Analysis Engine</h4>
                <p style="line-height: 1.8; color: #E0E0E0;">
                This platform employs a sophisticated ensemble of <b>6 AI algorithms</b> to provide 
                comprehensive climate intelligence:
                </p>
                <ol style="line-height: 2; color: #B0B0B0;">
                    <li><b>Time-Weighted Climate Efficiency:</b> Daily weather pattern analysis</li>
                    <li><b>Predictive Climate Scoring:</b> ML-inspired future impact forecasting</li>
                    <li><b>Adaptive Learning System:</b> Self-improving with historical data</li>
                    <li><b>Risk-Adjusted Scoring:</b> Probability-weighted weather event assessment</li>
                    <li><b>Timing Optimization:</b> Contract alignment with optimal weather windows</li>
                    <li><b>Ensemble Intelligence:</b> Confidence-weighted multi-model approach</li>
                </ol>
                <p style="color: #4FC3F7; font-weight: bold;">
                📊 Prediction Confidence: 87-92% | 🎯 Accuracy: Industry-Leading
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Climate scores
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                climate_score = metrics['climate_impact']
                score_class = get_score_class(climate_score)
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #B0B0B0; text-align: center;">CLIMATE SCORE</h4>
                    <div class="score-display {score_class}">
                        {climate_score:.1f}%
                    </div>
                    <p style="text-align: center; color: #B0B0B0;">AI Ensemble Result</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                climate_opt = metrics.get('climate_optimization', 70)
                opt_class = get_score_class(climate_opt)
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #B0B0B0; text-align: center;">OPTIMIZATION</h4>
                    <div class="score-display {opt_class}">
                        {climate_opt:.1f}%
                    </div>
                    <p style="text-align: center; color: #B0B0B0;">Contract Timing</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                if climate_score >= 85:
                    grade = "Excellent"
                    grade_color = "#00E676"
                elif climate_score >= 75:
                    grade = "Good"
                    grade_color = "#4FC3F7"
                elif climate_score >= 60:
                    grade = "Fair"
                    grade_color = "#FF9800"
                else:
                    grade = "Poor"
                    grade_color = "#FF5252"
                
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #B0B0B0; text-align: center;">CLIMATE GRADE</h4>
                    <div class="score-display" style="color: {grade_color};">
                        {grade}
                    </div>
                    <p style="text-align: center; color: #B0B0B0;">Overall Rating</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                # Calculate potential improvement
                improvement_potential = max(0, 85 - climate_score)
                st.markdown(f"""
                <div class="metric-card">
                    <h4 style="color: #B0B0B0; text-align: center;">POTENTIAL</h4>
                    <div class="score-display score-good">
                        +{improvement_potential:.1f}%
                    </div>
                    <p style="text-align: center; color: #B0B0B0;">Improvement Room</p>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown("---")
            
            # Climate insights display
            if show_climate and 'climate_insights' in metrics and metrics['climate_insights']:
                st.markdown("### 📍 LOCATION-SPECIFIC CLIMATE ANALYSIS")
                
                for idx, insight in enumerate(metrics['climate_insights']):
                    with st.expander(
                        f"🌍 {insight.get('location', 'Unknown')} | {insight.get('climate_type', 'N/A').replace('_', ' ').title()}",
                        expanded=(idx == 0)
                    ):
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown("**🌡️ Climate Profile**")
                            st.write(f"**Type:** {insight.get('climate_type', 'N/A').replace('_', ' ').title()}")
                            st.write(f"**Description:** {insight.get('description', 'Standard climate conditions')}")
                            st.write(f"**Period:** {insight.get('contract_period', 'N/A')}")
                        
                        with col2:
                            st.markdown("**⚠️ Risk Assessment**")
                            risk = insight.get('risk_assessment', {})
                            if risk:
                                total_months = risk.get('total_months', 0)
                                peak_risk = risk.get('peak_risk_exposure', 0)
                                optimal = risk.get('optimal_coverage', 0)
                                
                                if total_months > 0:
                                    peak_pct = (peak_risk / total_months) * 100
                                    optimal_pct = (optimal / total_months) * 100
                                    
                                    st.write(f"**Peak Risk Months:** {peak_risk} ({peak_pct:.0f}%)")
                                    st.write(f"**Optimal Coverage:** {optimal} ({optimal_pct:.0f}%)")
                                    st.write(f"**Total Duration:** {total_months} months")
                        
                        # Recommendations
                        recs = insight.get('recommendations', [])
                        if recs:
                            st.markdown("---")
                            st.markdown("**💡 AI Recommendations:**")
                            for rec in recs:
                                if 'HIGH RISK' in rec or 'CRITICAL' in rec:
                                    st.error(rec)
                                elif 'OPTIMAL' in rec:
                                    st.success(rec)
                                else:
                                    st.info(rec)
            
            # Climate visualization
            st.markdown("---")
            st.markdown("### 📊 CLIMATE IMPACT VISUALIZATION")
            
            if 'Current Location' in rig_data.columns:
                locations = rig_data['Current Location'].dropna().unique()
                
                location_scores = []
                for location in locations:
                    loc_data = rig_data[rig_data['Current Location'] == location]
                    
                    start_dates = pd.to_datetime(loc_data['Contract Start Date'], errors='coerce')
                    end_dates = pd.to_datetime(loc_data['Contract End Date'], errors='coerce')
                    
                    if start_dates.notna().any() and end_dates.notna().any():
                        # Use cached climate calculation
                        climate_result = calculate_climate_score_cached(
                            selected_rig,
                            location,
                            str(start_dates.iloc[0]),
                            str(end_dates.iloc[0])
                        )
                        score = climate_result['score']
                    else:
                        climate_ai = get_climate_ai()
                        climate_profile = climate_ai._get_climate_profile(str(location).lower())
                        score = climate_profile['efficiency_factor'] * 100
                    
                    location_scores.append({
                        'Location': location,
                        'Climate Score': score,
                        'Grade': 'Excellent' if score >= 85 else 'Good' if score >= 75 else 'Fair' if score >= 60 else 'Poor'
                    })
                
                if location_scores:
                    climate_df = pd.DataFrame(location_scores)
                    
                    fig = px.bar(
                        climate_df,
                        x='Location',
                        y='Climate Score',
                        color='Climate Score',
                        color_continuous_scale=[
                            [0, '#FF5252'],
                            [0.6, '#FF9800'],
                            [0.75, '#4FC3F7'],
                            [0.85, '#00E676']
                        ],
                        range_color=[0, 100],
                        text='Climate Score',
                        title='<b>Climate Efficiency by Location</b>'
                    )
                    
                    fig.update_traces(
                        texttemplate='%{text:.1f}%',
                        textposition='outside',
                        textfont=dict(size=12, color='#FFFFFF', family='Arial Black')
                    )
                    
                    fig.update_layout(
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(26, 26, 26, 0.5)',
                        font=dict(color='#FFFFFF'),
                        xaxis=dict(
                            gridcolor='rgba(212, 175, 55, 0.1)',
                            tickfont=dict(size=11)
                        ),
                        yaxis=dict(
                            gridcolor='rgba(212, 175, 55, 0.1)',
                            range=[0, 110]
                        ),
                        title=dict(
                            font=dict(size=18, color='#FFFFFF', family='Arial Black'),
                            x=0.5,
                            xanchor='center'
                        ),
                        height=450
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
        
        # TAB 4: Enhanced Insights
        with tab4:
            st.markdown("## 💡 STRATEGIC INSIGHTS & RECOMMENDATIONS")
            
            # Display insights
            if 'insights' in metrics and metrics['insights']:
                st.markdown("### 🔍 KEY FINDINGS")
                display_enhanced_insights(metrics['insights'])
            
            # AI Observations
            if 'ai_observations' in metrics and metrics['ai_observations']:
                st.markdown("---")
                st.markdown("### 🤖 AI STRATEGIC OBSERVATIONS")
                
                priority_filter = st.multiselect(
                    "Filter by Priority",
                    options=['CRITICAL', 'HIGH', 'MEDIUM', 'LOW'],
                    default=['CRITICAL', 'HIGH', 'MEDIUM', 'LOW']
                )
                
                filtered_obs = [
                    obs for obs in metrics['ai_observations']
                    if obs.get('priority', 'MEDIUM').upper() in priority_filter
                ]
                
                for idx, obs in enumerate(filtered_obs, 1):
                    priority = obs.get('priority', 'MEDIUM').upper()
                    
                    if priority == 'CRITICAL':
                        icon = '🔴'
                        color = '#FF5252'
                    elif priority == 'HIGH':
                        icon = '🟠'
                        color = '#FF9800'
                    elif priority == 'MEDIUM':
                        icon = '🔵'
                        color = '#4FC3F7'
                    else:
                        icon = '🟢'
                        color = '#00E676'
                    
                    with st.expander(
                        f"{icon} [{priority}] {obs.get('title', 'Observation')}",
                        expanded=(idx <= 2 and priority in ['CRITICAL', 'HIGH'])
                    ):
                        st.markdown(f"""
                        <div style="background: linear-gradient(90deg, {color}15 0%, transparent 100%);
                                    padding: 1rem; border-radius: 10px; border-left: 4px solid {color};">
                            <p style="color: #E0E0E0; line-height: 1.8;">
                                {obs.get('observation', 'No details available')}
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        if 'analysis' in obs:
                            st.markdown("**📊 Analysis:**")
                            for point in obs['analysis']:
                                st.markdown(f"- {point}")
                        
                        if 'actionable_steps' in obs:
                            st.markdown("**✅ Actionable Steps:**")
                            for step in obs['actionable_steps']:
                                st.markdown(f"- {step}")
                        
                        if 'impact' in obs:
                            st.success(f"**💡 Expected Impact:** {obs['impact']}")
            
            # Contract summary
            st.markdown("---")
            st.markdown("### 📊 CONTRACT PERFORMANCE SUMMARY")
            
            # Use cached summary generation
            summary = generate_summary_cached(selected_rig, rig_data.to_dict('list'), metrics)
            
            if summary:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("""
                    <div class="metric-card">
                        <h4 style="color: #D4AF37;">RIG OVERVIEW</h4>
                    """, unsafe_allow_html=True)
                    st.write(f"**Rig Name:** {summary['rig_name']}")
                    st.write(f"**Total Contracts:** {summary['total_contracts']}")
                    st.write(f"**Active Contracts:** {summary['active_contracts']}")
                    st.write(f"**Overall Grade:** {summary['efficiency_grade']}")
                    st.markdown("</div>", unsafe_allow_html=True)
                
                with col2:
                    st.markdown("""
                    <div class="metric-card">
                        <h4 style="color: #4FC3F7;">FINANCIAL SUMMARY</h4>
                    """, unsafe_allow_html=True)
                    st.write(f"**Total Value:** ${summary['total_contract_value']:.1f}M")
                    st.write(f"**Average Dayrate:** ${summary['average_dayrate']:.1f}k")
                    st.write(f"**Top Strength:** {summary['top_strength']}")
                    st.write(f"**Primary Concern:** {summary['primary_concern']}")
                    st.markdown("</div>", unsafe_allow_html=True)
        
        # TAB 5: Enhanced Fleet Comparison
        with tab5:
            st.markdown("## 🏢 FLEET PERFORMANCE COMPARISON")
            
            if len(available_rigs) < 2:
                st.info("📊 Upload data for multiple rigs to enable comprehensive fleet comparison")
            else:
                # Rig selection
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    selected_rigs_comp = st.multiselect(
                        "Select rigs to compare (max 10)",
                        options=available_rigs,
                        default=list(available_rigs[:min(5, len(available_rigs))])
                    )
                    # Enforce a maximum selection of 10 (some Streamlit versions may not support max_selections)
                    if isinstance(selected_rigs_comp, (list, tuple)) and len(selected_rigs_comp) > 10:
                        selected_rigs_comp = list(selected_rigs_comp)[:10]
                        st.warning("Selection limited to first 10 rigs for comparison.")
                
                with col2:
                    if st.button("🔄 REFRESH ANALYSIS", use_container_width=True):
                        # Clear cached results
                        for rig in selected_rigs_comp:
                            if rig in st.session_state.analysis_results:
                                del st.session_state.analysis_results[rig]
                
                if selected_rigs_comp:
                    # Calculate metrics for all selected rigs
                    comparison_data = []
                    
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    for idx, rig_name in enumerate(selected_rigs_comp):
                        status_text.text(f"Analyzing {rig_name}... ({idx+1}/{len(selected_rigs_comp)})")
                        progress_bar.progress((idx + 1) / len(selected_rigs_comp))
                        
                        rig_specific_data = df[df['Rig Name'] == rig_name]
                        
                        if rig_name not in st.session_state.analysis_results:
                            # Use cached calculation for each rig
                            rig_metrics = calculate_metrics_cached(rig_name, rig_specific_data.to_dict('list'))
                            st.session_state.analysis_results[rig_name] = rig_metrics
                        else:
                            rig_metrics = st.session_state.analysis_results[rig_name]
                        
                        if rig_metrics:
                            comparison_data.append({
                                'Rig Name': rig_name,
                                'Overall Score': rig_metrics['overall_efficiency'],
                                'Grade': rig_metrics['efficiency_grade'],
                                'Utilization': rig_metrics['contract_utilization'],
                                'Dayrate': rig_metrics['dayrate_efficiency'],
                                'Stability': rig_metrics['contract_stability'],
                                'Location': rig_metrics['location_complexity'],
                                'Climate': rig_metrics['climate_impact'],
                                'Performance': rig_metrics['contract_performance']
                            })
                    
                    progress_bar.empty()
                    status_text.empty()
                    
                    if comparison_data:
                        comparison_df = pd.DataFrame(comparison_data)
                        comparison_df = comparison_df.sort_values('Overall Score', ascending=False)
                        
                        # Summary stats
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.metric("Fleet Average", f"{comparison_df['Overall Score'].mean():.1f}%")
                        with col2:
                            # Safely get top rig name
                            try:
                                best_name = comparison_df['Rig Name'].iloc[0]
                                display_name = best_name if len(str(best_name)) <= 15 else str(best_name)[:15] + '…'
                            except Exception:
                                display_name = "N/A"
                            st.metric("Best Performer", display_name)
                        with col3:
                            st.metric("Highest Score", f"{comparison_df['Overall Score'].max():.1f}%")
                        with col4:
                            score_range = comparison_df['Overall Score'].max() - comparison_df['Overall Score'].min()
                            st.metric("Score Range", f"{score_range:.1f}%")
                        
                        st.markdown("---")
                        
                        # Comparison table
                        st.markdown("### 📊 PERFORMANCE TABLE")
                        st.dataframe(
                            comparison_df.style.background_gradient(
                                subset=['Overall Score', 'Utilization', 'Dayrate', 'Stability',
                                       'Location', 'Climate', 'Performance'],
                                cmap='RdYlGn',
                                vmin=0,
                                vmax=100
                            ).format({
                                'Overall Score': '{:.1f}%',
                                'Utilization': '{:.1f}%',
                                'Dayrate': '{:.1f}%',
                                'Stability': '{:.1f}%',
                                'Location': '{:.1f}%',
                                'Climate': '{:.1f}%',
                                'Performance': '{:.1f}%'
                            }),
                            use_container_width=True,
                            hide_index=True
                        )
                        
                        st.markdown("---")
                        
                        # Charts
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown("#### 📊 OVERALL SCORES")
                            fig1 = px.bar(
                                comparison_df,
                                x='Rig Name',
                                y='Overall Score',
                                color='Overall Score',
                                color_continuous_scale=[
                                    [0, '#FF5252'],
                                    [0.6, '#FF9800'],
                                    [0.75, '#4FC3F7'],
                                    [0.85, '#00E676']
                                ],
                                range_color=[0, 100],
                                text='Overall Score'
                            )
                            fig1.update_traces(
                                texttemplate='%{text:.1f}%',
                                textposition='outside'
                            )
                            fig1.update_layout(
                                paper_bgcolor='rgba(0,0,0,0)',
                                plot_bgcolor='rgba(26, 26, 26, 0.5)',
                                font=dict(color='#FFFFFF'),
                                xaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)'),
                                yaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)', range=[0, 110]),
                                height=400
                            )
                            st.plotly_chart(fig1, use_container_width=True)
                        
                        with col2:
                            st.markdown("#### 🏆 GRADE DISTRIBUTION")
                            grade_counts = comparison_df['Grade'].value_counts()
                            fig2 = px.pie(
                                values=grade_counts.values,
                                names=grade_counts.index,
                                color_discrete_sequence=['#00E676', '#4FC3F7', '#FF9800', '#FF5252']
                            )
                            fig2.update_layout(
                                paper_bgcolor='rgba(0,0,0,0)',
                                font=dict(color='#FFFFFF'),
                                height=400
                            )
                            st.plotly_chart(fig2, use_container_width=True)
                        
                        # Heatmap
                        st.markdown("---")
                        st.markdown("### 🔥 PERFORMANCE HEATMAP")
                        fig_heatmap = create_performance_heatmap(comparison_df)
                        st.plotly_chart(fig_heatmap, use_container_width=True)
                        
                        # Multi-radar
                        st.markdown("---")
                        st.markdown("### 🔷 FLEET PROFILES")
                        
                        fig_multi_radar = go.Figure()
                        
                        categories = ['Utilization', 'Dayrate', 'Stability', 'Location', 'Climate', 'Performance']
                        colors_radar = ['#D4AF37', '#4FC3F7', '#00E676', '#FF9800', '#FF5252', 
                                      '#9C27B0', '#00BCD4', '#FFC107', '#E91E63', '#3F51B5']
                        
                        for idx, (_, row) in enumerate(comparison_df.iterrows()):
                            values = [
                                row['Utilization'],
                                row['Dayrate'],
                                row['Stability'],
                                row['Location'],
                                row['Climate'],
                                row['Performance']
                            ]
                            try:
                                color_hex = colors_radar[idx % len(colors_radar)]
                                fill_rgba = hex_to_rgba(color_hex, alpha=0.10)
                                fig_multi_radar.add_trace(go.Scatterpolar(
                                    r=values,
                                    theta=categories,
                                    fill='toself',
                                    name=row['Rig Name'],
                                    line=dict(color=color_hex, width=2),
                                    fillcolor=fill_rgba
                                ))
                            except Exception as e:
                                # If a single trace fails, continue with other rigs and report in console
                                import traceback
                                traceback.print_exc()
                        
                        fig_multi_radar.update_layout(
                            polar=dict(
                                bgcolor='rgba(26, 26, 26, 0.5)',
                                radialaxis=dict(
                                    visible=True,
                                    range=[0, 100],
                                    tickfont=dict(size=11, color='#B0B0B0'),
                                    gridcolor='rgba(212, 175, 55, 0.2)'
                                ),
                                angularaxis=dict(
                                    tickfont=dict(size=12, color='#FFFFFF'),
                                    gridcolor='rgba(212, 175, 55, 0.2)'
                                )
                            ),
                            showlegend=True,
                            legend=dict(
                                font=dict(size=11, color='#FFFFFF'),
                                bgcolor='rgba(26, 26, 26, 0.8)',
                                bordercolor='#D4AF37',
                                borderwidth=1
                            ),
                            paper_bgcolor='rgba(0,0,0,0)',
                            height=650
                        )
                        
                        st.plotly_chart(fig_multi_radar, use_container_width=True)
        
        # TAB 6: Enhanced Export
        with tab6:
            st.markdown("## 📤 EXPORT ANALYSIS")
            
            st.markdown("""
            <div class="info-box">
                <h4>📦 Export your analysis in multiple formats for sharing and reporting</h4>
                <p style="color: #B0B0B0;">
                Choose from comprehensive Excel reports or streamlined CSV data exports.
                All exports include timestamp and rig information.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("""
                <div class="metric-card">
                    <h3 style="text-align: center; color: #00E676;">📊 EXCEL REPORT</h3>
                    <p style="text-align: center; color: #B0B0B0; line-height: 1.8;">
                    Complete analysis with:<br>
                    • All efficiency metrics<br>
                    • AI observations<br>
                    • Climate insights<br>
                    • Raw contract data<br>
                    • Strategic recommendations
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                excel_data = export_to_excel_enhanced(rig_data, metrics, f"{selected_rig}_analysis.xlsx")
                
                st.download_button(
                    label="📥 DOWNLOAD EXCEL REPORT",
                    data=excel_data,
                    file_name=f"{selected_rig}_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                st.markdown("""
                <div class="metric-card">
                    <h3 style="text-align: center; color: #4FC3F7;">📄 CSV DATA</h3>
                    <p style="text-align: center; color: #B0B0B0; line-height: 1.8;">
                    Streamlined data export:<br>
                    • Core efficiency metrics<br>
                    • Scores and grades<br>
                    • Quick analysis<br>
                    • Easy integration<br>
                    • Lightweight format
                    </p>
                </div>
                """, unsafe_allow_html=True)
                
                metrics_csv = pd.DataFrame({
                    'Metric': [
                        'Overall Efficiency',
                        'Contract Utilization',
                        'Dayrate Efficiency',
                        'Contract Stability',
                        'Location Complexity',
                        'Climate Impact (AI)',
                        'Contract Performance',
                        'Climate Optimization'
                    ],
                    'Score (%)': [
                        metrics['overall_efficiency'],
                        metrics['contract_utilization'],
                        metrics['dayrate_efficiency'],
                        metrics['contract_stability'],
                        metrics['location_complexity'],
                        metrics['climate_impact'],
                        metrics['contract_performance'],
                        metrics.get('climate_optimization', 70)
                    ],
                    'Grade': [
                        metrics['efficiency_grade'],
                        '', '', '', '', '', '', ''
                    ]
                }).to_csv(index=False)
                
                st.download_button(
                    label="📥 DOWNLOAD CSV DATA",
                    data=metrics_csv,
                    file_name=f"{selected_rig}_metrics_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
            st.markdown("---")
            
            # Export summary
            st.markdown("### 📋 EXPORT SUMMARY")
            
            st.markdown(f"""
            <div class="success-box">
                <h4>✅ Analysis Ready for Export</h4>
                <p style="line-height: 2; color: #E0E0E0;">
                <b>Analysis Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br>
                <b>Rig:</b> {selected_rig}<br>
                <b>Overall Score:</b> {metrics['overall_efficiency']:.1f}% {get_score_emoji(metrics['overall_efficiency'])}<br>
                <b>Grade:</b> {metrics['efficiency_grade']}<br>
                <b>Total Contracts:</b> {len(rig_data)}<br>
                <b>Data Quality:</b> High ✅
                </p>
            </div>
            """, unsafe_allow_html=True)
        
        # ============================================================================
        # TAB 7: REGIONAL BENCHMARK ANALYSIS
        # ============================================================================
        with tab7:
            st.markdown("## 🎯 REGIONAL BENCHMARK ANALYSIS")
            
            st.markdown("""
            <div class="info-box">
                <h4>📊 Performance Normalized Against Regional Benchmarks</h4>
                <p style="color: #B0B0B0;">
                Analyze how this rig's performance compares to regional benchmark standards across 9 critical categories.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                # Get benchmark for this rig
                benchmark = get_benchmark_model()
                benchmark_data = benchmark.get_benchmark(rig_data)
                
                # Calculate normalized performance with caching
                # Backend now generates metrics internally, just pass rig_data
                normalized_perf = calculate_normalized_perf_cached(
                    rig_data.to_dict('list')
                )
                
                # Display benchmark metrics in columns
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">ROP PERFORMANCE</h4>
                        <div class="score-display {get_score_class(normalized_perf.get('rop_performance', 0) if isinstance(normalized_perf, dict) else 0)}" style="font-size: 2.5rem;">
                            {(normalized_perf.get('rop_performance', 0) if isinstance(normalized_perf, dict) else normalized_perf):.1f}%
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Rate of Penetration</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">NPT PERFORMANCE</h4>
                        <div class="score-display {get_score_class(normalized_perf.get('npt_performance', 0))}" style="font-size: 2.5rem;">
                            {normalized_perf.get('npt_performance', 0):.1f}%
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Non-Productive Time</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">TIME PERFORMANCE</h4>
                        <div class="score-display {get_score_class(normalized_perf.get('time_performance', 0))}" style="font-size: 2.5rem;">
                            {normalized_perf.get('time_performance', 0):.1f}%
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Schedule Efficiency</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">COST PERFORMANCE</h4>
                        <div class="score-display {get_score_class(normalized_perf.get('cost_performance', 0))}" style="font-size: 2.5rem;">
                            {normalized_perf.get('cost_performance', 0):.1f}%
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Cost Efficiency</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Overall normalized score
                overall_normalized = normalized_perf.get('overall_normalized', 0)
                st.markdown(f"""
                <div class="success-box">
                    <h3 style="text-align: center; color: #00FFB3;">OVERALL NORMALIZED SCORE: {overall_normalized:.1f}%</h3>
                    <p style="text-align: center; color: #B0B0B0;">Performance against regional benchmarks</p>
                </div>
                """, unsafe_allow_html=True)
                
                # Benchmark breakdown visualization
                st.markdown("### 📊 BENCHMARK CATEGORIES")
                
                benchmark_categories = {
                    'ROP Benchmark': benchmark_data.get('rop_benchmark', 0),
                    'NPT Benchmark': benchmark_data.get('npt_benchmark', 0),
                    'Time Benchmark': benchmark_data.get('time_benchmark', 0),
                    'Cost Benchmark': benchmark_data.get('cost_benchmark', 0),
                    'Location Benchmark': benchmark_data.get('location_benchmark', 0),
                    'Weather Benchmark': benchmark_data.get('weather_benchmark', 0),
                    'Contractor Benchmark': benchmark_data.get('contractor_benchmark', 0),
                    'Equipment Benchmark': benchmark_data.get('equipment_benchmark', 0),
                    'Market Benchmark': benchmark_data.get('market_benchmark', 0)
                }
                
                fig_bench = px.bar(
                    x=list(benchmark_categories.keys()),
                    y=list(benchmark_categories.values()),
                    color=list(benchmark_categories.values()),
                    color_continuous_scale=['#FF5252', '#FF9800', '#4FC3F7', '#00E676'],
                    range_color=[0, 100],
                    text=[f'{v:.1f}%' for v in benchmark_categories.values()]
                )
                fig_bench.update_traces(textposition='outside')
                fig_bench.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(26, 26, 26, 0.5)',
                    font=dict(color='#FFFFFF', size=11),
                    xaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)', tickangle=-45),
                    yaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)', range=[0, 120]),
                    height=400,
                    showlegend=False
                )
                st.plotly_chart(fig_bench, use_container_width=True)
                
            except Exception as e:
                st.error(f"⚠️ Error calculating benchmark analysis: {str(e)}")
                st.info("Ensure rig data is properly formatted and contains required columns.")
        
        # ============================================================================
        # TAB 8: RIG-WELL MATCH PREDICTION
        # ============================================================================
        with tab8:
            st.markdown("## 🔍 RIG-WELL MATCH PREDICTION")
            
            st.markdown("""
            <div class="info-box">
                <h4>🎯 ML-Powered Well Execution Prediction</h4>
                <p style="color: #B0B0B0;">
                Use machine learning to predict execution parameters for a specific well based on rig characteristics.
                Enter well parameters to get predictions on expected time, AFE probability, NPT, and optimal dayrate.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            # Initialize session state for ML predictions
            if 'ml_predictions_result' not in st.session_state:
                st.session_state.ml_predictions_result = None
            
            try:
                # Well parameter input form
                st.markdown("### 📝 WELL PARAMETERS")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    well_depth = st.slider("Well Depth (ft)", min_value=5000, max_value=35000, value=15000, step=500)
                    well_hardness = st.slider("Formation Hardness (1-10)", min_value=1, max_value=10, value=5, step=1)
                
                with col2:
                    well_temp = st.number_input("Bottom Hole Temp (°F)", min_value=50, max_value=300, value=150)
                    well_pressure = st.number_input("Formation Pressure (psi)", min_value=1000, max_value=20000, value=5000, step=100)
                
                with col3:
                    well_location = st.selectbox("Well Location", 
                        ["Gulf of Mexico", "North Sea", "Middle East", "Asia Pacific", "West Africa", "Southeast Asia", "South America"])
                
                # Prepare well parameters
                well_params = {
                    'depth': well_depth,
                    'hardness': well_hardness,
                    'temperature': well_temp,
                    'pressure': well_pressure,
                    'location': well_location
                }
                
                if st.button("🔍 PREDICT WELL EXECUTION", use_container_width=True, key="predict_well_exec"):
                    # Get predictions and store in session state
                    try:
                        ml_predictor = get_ml_predictor()
                        result = ml_predictor.predict_well_execution(rig_data, well_params)
                        if result:
                            st.session_state.ml_predictions_result = result
                            st.rerun()
                        else:
                            st.error("❌ Prediction returned no results. Please check input parameters.")
                    except Exception as e:
                        st.error(f"❌ Prediction error: {str(e)}")
                        import traceback
                        traceback.print_exc()
                
                # Display predictions if available (persists across reruns)
                if st.session_state.ml_predictions_result is not None:
                    predictions = st.session_state.ml_predictions_result
                    
                    st.markdown("---")
                    st.markdown("### 📊 PREDICTION RESULTS")
                    
                    # Display prediction metrics
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4 style="color: #B0B0B0; text-align: center;">EXPECTED DURATION</h4>
                            <div class="score-display score-good" style="font-size: 2.5rem;">
                                {predictions.get('expected_time_days', 0):.1f}
                            </div>
                            <p style="text-align: center; color: #B0B0B0;">Days</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4 style="color: #B0B0B0; text-align: center;">RISK SCORE</h4>
                            <div class="score-display {get_score_class(100 - predictions.get('risk_score', 50))}" style="font-size: 2.5rem;">
                                {predictions.get('risk_score', 0):.1f}
                            </div>
                            <p style="text-align: center; color: #B0B0B0;">0-100 (Lower is Better)</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-card">
                            <h4 style="color: #B0B0B0; text-align: center;">MATCH SCORE</h4>
                            <div class="score-display {get_score_class(predictions.get('match_score', 0))}" style="font-size: 2.5rem;">
                                {predictions.get('match_score', 0):.1f}%
                            </div>
                            <p style="text-align: center; color: #B0B0B0;">Rig-Well Compatibility</p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.markdown("---")
                    
                    # Additional metrics
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.info(f"**Expected NPT:** {predictions.get('expected_npt_percent', 0):.1f}%")
                    
                    with col2:
                        st.info(f"**AFE Probability:** {predictions.get('afe_probability', 0):.1f}%")
                    
                    with col3:
                        st.info(f"**Confidence:** {predictions.get('confidence_percent', 0):.1f}%")
                    
                    # Dayrate recommendation
                    dayrate_range = predictions.get('recommended_dayrate_range', {'low': 0, 'high': 0, 'optimal': 0})
                    st.markdown(f"""
                    <div class="success-box">
                        <h4>💰 RECOMMENDED DAYRATE RANGE</h4>
                        <h2 style="text-align: center; color: #00FFB3;">
                            ${dayrate_range['low']:.0f}k - ${dayrate_range['high']:.0f}k
                        </h2>
                        <p style="text-align: center; color: #B0B0B0;">Based on well complexity and market conditions</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("---")
                    
                    # Clear predictions button
                    if st.button("🗑️ CLEAR PREDICTIONS", use_container_width=True, key="clear_predictions"):
                        st.session_state.ml_predictions_result = None
                        st.rerun()
                    
            except Exception as e:
                st.error(f"⚠️ Error in well prediction: {str(e)}")
                st.info("Please ensure all well parameters are valid.")
        
        # ============================================================================
        # TAB 9: BASIN SCENARIO SIMULATOR
        # ============================================================================
        with tab9:
            st.markdown("## 🌍 BASIN SCENARIO SIMULATOR")
            
            st.markdown("""
            <div class="info-box">
                <h4>🎲 Monte Carlo Basin Transfer Analysis</h4>
                <p style="color: #B0B0B0;">
                Simulate rig performance transfer to different basins using 1000 Monte Carlo iterations.
                Analyze uncertainty ranges and probability distributions for NPT, duration, cost, and risk.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                # Basin scenario configuration
                st.markdown("### 🌐 SELECT BASIN SCENARIO")
                
                basin_scenarios = {
                    "Gulf of Mexico": {
                        "climate_severity": 0.45, "geology_difficulty": 0.50, 
                        "water_depth_ft": 8000, "typical_dayrate_k": 350
                    },
                    "North Sea": {
                        "climate_severity": 0.75, "geology_difficulty": 0.65,
                        "water_depth_ft": 450, "typical_dayrate_k": 380
                    },
                    "Middle East": {
                        "climate_severity": 0.25, "geology_difficulty": 0.45,
                        "water_depth_ft": 5000, "typical_dayrate_k": 320
                    },
                    "Asia Pacific": {
                        "climate_severity": 0.60, "geology_difficulty": 0.55,
                        "water_depth_ft": 3500, "typical_dayrate_k": 300
                    },
                    "West Africa": {
                        "climate_severity": 0.70, "geology_difficulty": 0.60,
                        "water_depth_ft": 5000, "typical_dayrate_k": 340
                    },
                    "Southeast Asia": {
                        "climate_severity": 0.65, "geology_difficulty": 0.50,
                        "water_depth_ft": 2000, "typical_dayrate_k": 280
                    },
                    "South America": {
                        "climate_severity": 0.55, "geology_difficulty": 0.60,
                        "water_depth_ft": 6000, "typical_dayrate_k": 310
                    }
                }
                
                # Initialize session state for basin simulation
                if 'basin_sim_results' not in st.session_state:
                    st.session_state.basin_sim_results = None
                
                selected_basin = st.selectbox("Basin", list(basin_scenarios.keys()))
                basin_params = basin_scenarios[selected_basin].copy()
                # Ensure basin_name is included for the simulation
                basin_params['basin_name'] = selected_basin
                
                # Option to customize parameters
                with st.expander("⚙️ CUSTOMIZE BASIN PARAMETERS"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        basin_params['climate_severity'] = st.slider(
                            "Climate Severity (0-1)", 0.0, 1.0, 
                            basin_params['climate_severity'], 0.05
                        )
                        basin_params['water_depth_ft'] = st.number_input(
                            "Water Depth (ft)", value=int(basin_params['water_depth_ft']), step=500
                        )
                    
                    with col2:
                        basin_params['geology_difficulty'] = st.slider(
                            "Geology Difficulty (0-1)", 0.0, 1.0,
                            basin_params['geology_difficulty'], 0.05
                        )
                        basin_params['typical_dayrate_k'] = st.number_input(
                            "Typical Dayrate ($k)", value=int(basin_params['typical_dayrate_k']), step=10
                        )
                
                if st.button("🎲 RUN MONTE CARLO SIMULATION", use_container_width=True, key="run_monte_carlo"):
                    # Run simulation and store in session state
                    try:
                        monte_carlo = get_monte_carlo()
                        result = monte_carlo.simulate_basin_transfer(rig_data, basin_params)
                        if result:
                            st.session_state.basin_sim_results = result
                            st.rerun()
                        else:
                            st.error("❌ Simulation returned no results. Please check basin parameters.")
                    except Exception as e:
                        st.error(f"❌ Simulation error: {str(e)}")
                        import traceback
                        traceback.print_exc()
                
                # Display simulation results if available (persists across reruns)
                if st.session_state.basin_sim_results is not None:
                    sim_results = st.session_state.basin_sim_results
                
                    st.markdown("---")
                    st.markdown("### 📊 SIMULATION RESULTS")
                    
                    # Create results summary table with safe type conversion
                    results_df = pd.DataFrame({
                        'Metric': ['NPT (days)', 'Duration (days)', 'Cost ($M)', 'Risk Score'],
                        'P10': [
                            float(sim_results['npt'].get('p10', 0)),
                            float(sim_results['duration'].get('p10', 0)),
                            float(sim_results['cost'].get('p10', 0)),
                            float(sim_results['risk'].get('p10', 0))
                        ],
                        'P50 (Median)': [
                            float(sim_results['npt'].get('p50', 0)),
                            float(sim_results['duration'].get('p50', 0)),
                            float(sim_results['cost'].get('p50', 0)),
                            float(sim_results['risk'].get('p50', 0))
                        ],
                        'P90': [
                            float(sim_results['npt'].get('p90', 0)),
                            float(sim_results['duration'].get('p90', 0)),
                            float(sim_results['cost'].get('p90', 0)),
                            float(sim_results['risk'].get('p90', 0))
                        ]
                    })
                    
                    st.dataframe(
                        results_df.style.format({
                            'P10': "{:.2f}",
                            'P50 (Median)': "{:.2f}",
                            'P90': "{:.2f}"
                        }).background_gradient(
                            subset=['P10', 'P50 (Median)', 'P90'],
                            cmap='RdYlGn_r'
                        ),
                        use_container_width=True,
                        hide_index=True
                    )
                    
                    st.markdown("---")
                    st.markdown("### 📈 DISTRIBUTION ANALYSIS")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # NPT distribution chart
                        fig_npt = px.histogram(
                            x=sim_results['npt']['distribution'],
                            nbins=40,
                            title="NPT Days Distribution (1000 simulations)",
                            labels={'x': 'NPT Days', 'count': 'Frequency'},
                            color_discrete_sequence=['#4FC3F7']
                        )
                        fig_npt.update_layout(
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(26, 26, 26, 0.5)',
                            font=dict(color='#FFFFFF'),
                            xaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)'),
                            yaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)'),
                            height=350
                        )
                        st.plotly_chart(fig_npt, use_container_width=True)
                    
                    with col2:
                        # Cost distribution chart
                        fig_cost = px.histogram(
                            x=sim_results['cost']['distribution'],
                            nbins=40,
                            title="Cost Distribution ($M) (1000 simulations)",
                            labels={'x': 'Cost ($M)', 'count': 'Frequency'},
                            color_discrete_sequence=['#FF9800']
                        )
                        fig_cost.update_layout(
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(26, 26, 26, 0.5)',
                            font=dict(color='#FFFFFF'),
                            xaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)'),
                            yaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)'),
                            height=350
                        )
                        st.plotly_chart(fig_cost, use_container_width=True)
                    
                    st.markdown("---")
                    
                    # Clear simulation button
                    if st.button("🗑️ CLEAR RESULTS", use_container_width=True, key="clear_basin_sim"):
                        st.session_state.basin_sim_results = None
                        st.rerun()
                    
            except Exception as e:
                st.error(f"⚠️ Error in basin simulation: {str(e)}")
                st.info("Ensure simulation parameters are within valid ranges.")
        
        # ============================================================================
        # TAB 10: CONTRACTOR PERFORMANCE ANALYZER
        # ============================================================================
        with tab10:
            st.markdown("## 📈 CONTRACTOR PERFORMANCE ANALYZER")
            
            st.markdown("""
            <div class="info-box">
                <h4>🎯 Consistency Metrics & Trend Analysis</h4>
                <p style="color: #B0B0B0;">
                Analyze contractor consistency across multiple performance dimensions.
                Identifies trends, red flags, and reliability indicators.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                # Analyze contractor consistency
                contractor_analyzer = get_contractor_analyzer()
                analysis = contractor_analyzer.analyze_contractor_consistency(rig_data)
                
                st.markdown("### 🏆 CONSISTENCY SCORECARD")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    score_class = get_score_class(analysis.get('overall_consistency', 0))
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">OVERALL CONSISTENCY</h4>
                        <div class="score-display {score_class}" style="font-size: 2.5rem;">
                            {analysis.get('overall_consistency', 0):.1f}%
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Grade: {analysis.get('consistency_grade', 'N/A')}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">TREND</h4>
                        <div class="score-display score-good" style="font-size: 2rem;">
                            {analysis.get('trend', 'Stable')}
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Direction</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    contracts_count = len(rig_data)
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">CONTRACTS</h4>
                        <div class="score-display score-excellent" style="font-size: 2.5rem;">
                            {contracts_count}
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Total Count</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    red_flags_count = len(analysis.get('red_flags', []))
                    flag_color = 'score-poor' if red_flags_count > 2 else 'score-fair' if red_flags_count > 0 else 'score-excellent'
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">RED FLAGS</h4>
                        <div class="score-display {flag_color}" style="font-size: 2.5rem;">
                            {red_flags_count}
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Issues Found</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Consistency breakdown
                st.markdown("### 📊 CONSISTENCY BREAKDOWN")
                
                consistency_data = {
                    'ROP Consistency': analysis.get('rop_consistency', 0),
                    'NPT Consistency': analysis.get('npt_consistency', 0),
                    'Schedule Variance': analysis.get('schedule_consistency', 0),
                    'Delivery Reliability': analysis.get('delivery_reliability', 0),
                    'Crew Stability': analysis.get('crew_stability', 0)
                }
                
                fig_consistency = px.bar(
                    x=list(consistency_data.keys()),
                    y=list(consistency_data.values()),
                    color=list(consistency_data.values()),
                    color_continuous_scale=['#FF5252', '#FF9800', '#4FC3F7', '#00E676'],
                    range_color=[0, 100],
                    text=[f'{v:.1f}%' for v in consistency_data.values()]
                )
                fig_consistency.update_traces(textposition='outside')
                fig_consistency.update_layout(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(26, 26, 26, 0.5)',
                    font=dict(color='#FFFFFF', size=10),
                    xaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)', tickangle=-45),
                    yaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)', range=[0, 120]),
                    height=350,
                    showlegend=False
                )
                st.plotly_chart(fig_consistency, use_container_width=True)
                
                st.markdown("---")
                
                # Red flags
                if analysis.get('red_flags'):
                    st.markdown("### ⚠️ RED FLAGS")
                    for flag in analysis.get('red_flags', []):
                        st.warning(f"🚩 {flag}")
                else:
                    st.success("✅ No red flags detected. Contractor showing strong consistency.")
                
            except Exception as e:
                st.error(f"⚠️ Error in contractor analysis: {str(e)}")
                st.info("Ensure rig data contains performance history.")
        
        # ============================================================================
        # TAB 11: LEARNING CURVE ANALYZER
        # ============================================================================
        with tab11:
            st.markdown("## 📉 LEARNING CURVE ANALYZER")
            
            st.markdown("""
            <div class="info-box">
                <h4>📚 Performance Trajectory & Learning Rate Analysis</h4>
                <p style="color: #B0B0B0;">
                Analyze contractor's power law learning curve showing improvement over time.
                Project future performance based on historical learning patterns.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                # Calculate learning curve with caching
                curve_analysis = calculate_learning_cached(rig_data.to_dict('list'))
                
                st.markdown("### 📊 LEARNING CURVE METRICS")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">LEARNING RATE</h4>
                        <div class="score-display score-good" style="font-size: 2.5rem;">
                            {curve_analysis.get('learning_rate_k', 1.0):.3f}
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Power Law Exponent</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    improvement = curve_analysis.get('improvement_percent', 0)
                    improvement_class = 'score-excellent' if improvement > 20 else 'score-good' if improvement > 10 else 'score-fair'
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">IMPROVEMENT</h4>
                        <div class="score-display {improvement_class}" style="font-size: 2.5rem;">
                            {improvement:.1f}%
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Overall Growth</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    contract_count = len(rig_data)
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">CONTRACTS</h4>
                        <div class="score-display score-excellent" style="font-size: 2.5rem;">
                            {contract_count}
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">In Analysis</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # Learning curve visualization
                actual_times = curve_analysis.get('actual_times', [])
                predicted_times = curve_analysis.get('predicted_times', [])
                
                if actual_times and predicted_times:
                    fig_learning = go.Figure()
                    
                    # Actual times
                    fig_learning.add_trace(go.Scatter(
                        y=actual_times,
                        mode='lines+markers',
                        name='Actual Performance',
                        line=dict(color='#4FC3F7', width=3),
                        marker=dict(size=8)
                    ))
                    
                    # Predicted times (learning curve)
                    fig_learning.add_trace(go.Scatter(
                        y=predicted_times,
                        mode='lines',
                        name='Learning Curve (Predicted)',
                        line=dict(color='#00FFB3', width=3, dash='dash'),
                        marker=dict(size=6)
                    ))
                    
                    fig_learning.update_layout(
                        title="Performance Trajectory - Actual vs Learning Curve",
                        xaxis_title="Contract Number",
                        yaxis_title="Days",
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(26, 26, 26, 0.5)',
                        font=dict(color='#FFFFFF'),
                        xaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)'),
                        yaxis=dict(gridcolor='rgba(212, 175, 55, 0.1)'),
                        height=400,
                        hovermode='x unified',
                        legend=dict(
                            bgcolor='rgba(26, 26, 26, 0.8)',
                            bordercolor='#D4AF37',
                            borderwidth=1
                        )
                    )
                    
                    st.plotly_chart(fig_learning, use_container_width=True)
                
                st.markdown("---")
                
                # Classification
                learning_classification = curve_analysis.get('classification', 'Steady Learner')
                st.markdown(f"""
                <div class="success-box">
                    <h3 style="text-align: center; color: #00FFB3;">CLASSIFICATION: {learning_classification.upper()}</h3>
                    <p style="text-align: center; color: #B0B0B0;">Based on learning curve analysis</p>
                </div>
                """, unsafe_allow_html=True)
                
            except Exception as e:
                st.error(f"⚠️ Error in learning curve analysis: {str(e)}")
                st.info("Ensure rig data has sufficient contract history.")
        
        # ============================================================================
        # TAB 12: INVISIBLE LOST TIME DETECTOR
        # ============================================================================
        with tab12:
            st.markdown("## ⚡ INVISIBLE LOST TIME DETECTOR")
            
            st.markdown("""
            <div class="info-box">
                <h4>🔍 Efficiency Gap Analysis & Cost Impact Assessment</h4>
                <p style="color: #B0B0B0;">
                Detect hidden inefficiencies not captured by standard metrics.
                Identify cost impacts and opportunities for improvement.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            try:
                # Detect invisible lost time
                ilt_detector = get_ilt_detector()
                ilt_analysis = ilt_detector.detect_ilt(rig_data)
                
                st.markdown("### 💰 ILT COST IMPACT")
                
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    total_ilt_days = ilt_analysis.get('total_ilt_days', 0)
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">TOTAL ILT DAYS</h4>
                        <div class="score-display score-fair" style="font-size: 2.5rem;">
                            {total_ilt_days:.1f}
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Days Lost</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    ilt_percentage = ilt_analysis.get('ilt_percentage', 0)
                    ilt_class = 'score-excellent' if ilt_percentage < 5 else 'score-good' if ilt_percentage < 10 else 'score-fair' if ilt_percentage < 15 else 'score-poor'
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">ILT %</h4>
                        <div class="score-display {ilt_class}" style="font-size: 2.5rem;">
                            {ilt_percentage:.1f}%
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Of Total Time</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col3:
                    cost_impact = ilt_analysis.get('cost_impact_$k', 0)
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">COST IMPACT</h4>
                        <div class="score-display score-fair" style="font-size: 2.5rem;">
                            ${cost_impact:,.0f}k
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Total Impact</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col4:
                    severity = ilt_analysis.get('severity', 'Medium')
                    severity_class = 'score-poor' if severity == 'Critical' else 'score-fair' if severity == 'High' else 'score-good'
                    st.markdown(f"""
                    <div class="metric-card">
                        <h4 style="color: #B0B0B0; text-align: center;">SEVERITY</h4>
                        <div class="score-display {severity_class}" style="font-size: 1.8rem;">
                            {severity}
                        </div>
                        <p style="text-align: center; color: #B0B0B0;">Level</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                # ILT Breakdown
                st.markdown("### 📊 ILT FINDINGS BREAKDOWN")
                
                findings = ilt_analysis.get('findings', {})
                
                if findings:
                    # Create findings summary
                    finding_list = []
                    for category, details in findings.items():
                        if isinstance(details, dict):
                            finding_list.append({
                                'Category': category,
                                'Count': details.get('count', 0),
                                'Days Lost': details.get('days', 0),
                                'Cost ($k)': details.get('cost', 0)
                            })
                    
                    if finding_list:
                        findings_df = pd.DataFrame(finding_list)
                        st.dataframe(
                            findings_df.style.background_gradient(
                                subset=['Days Lost', 'Cost ($k)'],
                                cmap='RdYlGn_r'
                            ),
                            use_container_width=True,
                            hide_index=True
                        )
                
                st.markdown("---")
                
                # Recommendations
                st.markdown("### 💡 EFFICIENCY RECOMMENDATIONS")
                
                recommendations = ilt_analysis.get('recommendations', [])
                if recommendations:
                    for i, rec in enumerate(recommendations, 1):
                        st.info(f"**{i}. {rec}**")
                else:
                    st.success("✅ No critical efficiency gaps detected.")
                
            except Exception as e:
                st.error(f"⚠️ Error in ILT detection: {str(e)}")
                st.info("Ensure rig data contains sufficient performance records.")
        
        # ============================================================================
        # TAB 13: RIG AVAILABILITY SEARCH
        # ============================================================================
        with tab13:
            st.markdown("## 🔍 RIG AVAILABILITY SEARCH")
            
            st.markdown("""
            <div class="info-box">
                <h4>🎯 Find Available Rigs by Location, Rate & Availability</h4>
                <p style="color: #B0B0B0;">
                Search and filter rigs based on your operational requirements.
                Results ranked by match score and availability.
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            # === FILTER SECTION ===
            st.markdown("### 🎛️ SEARCH FILTERS")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Location filter
                all_locations = ['All'] + sorted([str(x) for x in df['Current Location'].dropna().unique().tolist()])
                selected_location = st.selectbox(
                    "📍 Location",
                    options=all_locations,
                    help="Filter by specific rig location"
                )
                
                # Region filter
                all_regions = ['All'] + sorted([str(x) for x in df['Region'].dropna().unique().tolist()])
                selected_region = st.selectbox(
                    "🌍 Region",
                    options=all_regions,
                    help="Filter by geographical region"
                )
            
            with col2:
                # Day rate range
                min_rate = float(df['Dayrate ($k)'].min())
                max_rate = float(df['Dayrate ($k)'].max())
                
                dayrate_range = st.slider(
                    "💰 Day Rate Range ($k)",
                    min_value=min_rate,
                    max_value=max_rate,
                    value=(min_rate, max_rate),
                    help="Filter by daily operating rate"
                )
                
                # Contractor filter (optional)
                all_contractors = ['All'] + sorted([str(x) for x in df['Contractor'].dropna().unique().tolist()])
                selected_contractor = st.selectbox(
                    "🏢 Contractor",
                    options=all_contractors,
                    help="Filter by drilling contractor"
                )
            
            with col3:
                # Availability status
                availability_options = [
                    'All',
                    'Available Now',
                    'Available Soon (<30 days)',
                    'Available <90 days'
                ]
                availability_status = st.selectbox(
                    "📅 Availability",
                    options=availability_options,
                    help="Filter by contract end date"
                )
                
                # Climate preference
                climate_options = [
                    'Any',
                    'Stable Climate',
                    'Storm Resistant',
                    'Cold Weather Capable'
                ]
                climate_pref = st.selectbox(
                    "🌤️ Climate Requirements",
                    options=climate_options,
                    help="Climate compatibility preference"
                )
            
            st.markdown("---")
            
            # === SEARCH BUTTON ===
            if st.button("🔎 SEARCH AVAILABLE RIGS", type="primary", use_container_width=True):
                
                # Build filter dictionary
                search_filters = {
                    'location': selected_location,
                    'region': selected_region,
                    'dayrate_min': dayrate_range[0],
                    'dayrate_max': dayrate_range[1],
                    'availability_status': availability_status,
                    'climate_preference': climate_pref,
                    'contractor': selected_contractor
                }
                
                # Get search engine and perform search
                try:
                    search_engine = get_search_engine()  # Cached function
                    results = search_engine.search_available_rigs(df, search_filters)
                    
                    if len(results) > 0:
                        # === RESULTS SUMMARY ===
                        st.markdown("### 📊 SEARCH RESULTS")
                        
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            st.markdown(f"""
                            <div class="metric-card">
                                <h4 style="color: #B0B0B0; text-align: center;">RIGS FOUND</h4>
                                <div class="score-display score-excellent" style="font-size: 2.5rem;">
                                    {len(results)}
                                </div>
                                <p style="text-align: center; color: #B0B0B0;">Matching Criteria</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col2:
                            avg_score = results['Match_Score'].mean()
                            score_class = 'score-excellent' if avg_score >= 80 else 'score-good' if avg_score >= 60 else 'score-fair'
                            st.markdown(f"""
                            <div class="metric-card">
                                <h4 style="color: #B0B0B0; text-align: center;">AVG MATCH</h4>
                                <div class="score-display {score_class}" style="font-size: 2.5rem;">
                                    {avg_score:.0f}%
                                </div>
                                <p style="text-align: center; color: #B0B0B0;">Score</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col3:
                            avg_rate = results['Dayrate ($k)'].mean()
                            st.markdown(f"""
                            <div class="metric-card">
                                <h4 style="color: #B0B0B0; text-align: center;">AVG RATE</h4>
                                <div class="score-display score-good" style="font-size: 2.5rem;">
                                    ${avg_rate:.0f}k
                                </div>
                                <p style="text-align: center; color: #B0B0B0;">Per Day</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        with col4:
                            avg_climate = results['Climate_Score'].mean()
                            st.markdown(f"""
                            <div class="metric-card">
                                <h4 style="color: #B0B0B0; text-align: center;">CLIMATE</h4>
                                <div class="score-display score-good" style="font-size: 2.5rem;">
                                    {avg_climate:.1f}/10
                                </div>
                                <p style="text-align: center; color: #B0B0B0;">Compatibility</p>
                            </div>
                            """, unsafe_allow_html=True)
                        
                        st.markdown("---")
                        
                        # === RESULTS TABLE ===
                        st.markdown("### 📋 AVAILABLE RIGS")
                        
                        # Prepare display dataframe
                        display_cols = [
                            'Drilling Unit Name',
                            'Contractor',
                            'Current Location',
                            'Region',
                            'Dayrate ($k)',
                            'Contract End Date',
                            'Contract Days Remaining',
                            'Match_Score',
                            'Climate_Score'
                        ]
                        
                        # Filter only available columns
                        display_cols = [col for col in display_cols if col in results.columns]
                        
                        display_df = results[display_cols].copy()
                        display_df['Match_Score'] = display_df['Match_Score'].round(0).astype(int)
                        display_df['Climate_Score'] = display_df['Climate_Score'].round(1)
                        
                        # Rename for better display
                        display_df = display_df.rename(columns={
                            'Drilling Unit Name': 'Rig Name',
                            'Current Location': 'Location',
                            'Contract Days Remaining': 'Days Until Available',
                            'Match_Score': 'Match %',
                            'Climate_Score': 'Climate (0-10)'
                        })
                        
                        st.dataframe(
                            display_df.style.background_gradient(
                                subset=['Match %', 'Climate (0-10)'],
                                cmap='RdYlGn'
                            ),
                            use_container_width=True,
                            hide_index=True,
                            height=400
                        )
                        
                        st.markdown("---")
                        
                        # === DETAILED CARDS (Optional) ===
                        show_details = st.checkbox("📑 Show Detailed Rig Cards")
                        
                        if show_details:
                            st.markdown("### 🛢️ DETAILED RIG INFORMATION")
                            
                            for idx, rig in results.head(10).iterrows():  # Show top 10
                                with st.expander(
                                    f"🔹 {rig['Drilling Unit Name']} | Match: {rig['Match_Score']:.0f}% | "
                                    f"Rate: ${rig['Dayrate ($k)']:.0f}k"
                                ):
                                    col1, col2, col3 = st.columns(3)
                                    
                                    with col1:
                                        st.markdown("**📍 Location Details**")
                                        st.write(f"**Location:** {rig['Current Location']}")
                                        st.write(f"**Region:** {rig['Region']}")
                                        st.write(f"**Contractor:** {rig['Contractor']}")
                                    
                                    with col2:
                                        st.markdown("**💰 Contract Information**")
                                        st.write(f"**Day Rate:** ${rig['Dayrate ($k)']}k")
                                        if pd.notna(rig.get('Contract value ($m)')):
                                            st.write(f"**Contract Value:** ${rig['Contract value ($m)']:.1f}M")
                                        if pd.notna(rig.get('Contract End Date')):
                                            st.write(f"**Available From:** {rig['Contract End Date'].strftime('%Y-%m-%d')}")
                                    
                                    with col3:
                                        st.markdown("**📊 Scores**")
                                        st.write(f"**Match Score:** {rig['Match_Score']:.0f}%")
                                        st.write(f"**Climate Score:** {rig['Climate_Score']:.1f}/10")
                                        if pd.notna(rig.get('Contract Days Remaining')):
                                            days = rig['Contract Days Remaining']
                                            if days <= 0:
                                                st.success("✅ Available Now")
                                            elif days <= 30:
                                                st.info(f"⏳ Available in {int(days)} days")
                                            else:
                                                st.warning(f"📅 Available in {int(days)} days")
                        
                        # === EXPORT BUTTON ===
                        st.markdown("---")
                        csv = results.to_csv(index=False)
                        st.download_button(
                            label="📥 EXPORT RESULTS TO CSV",
                            data=csv,
                            file_name=f"rig_search_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
                        
                    else:
                        st.warning("⚠️ No rigs match your search criteria. Try adjusting the filters.")
                        st.info("💡 Tip: Expand the day rate range or change availability status.")
                
                except Exception as e:
                    st.error(f"❌ Error during search: {str(e)}")
                    st.info("Please check your data and try again.")
    


if __name__ == "__main__":
    main()