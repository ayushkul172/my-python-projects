"""
Advanced Rig Efficiency Analysis Tool - Next-Generation AI Assistant
Integrates: ML algorithms, semantic analysis, context awareness, intent prediction,
fuzzy matching, reinforcement learning, multi-turn dialogue, embeddings
"""

import streamlit as st
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any, Tuple, Set
from dataclasses import dataclass, field
from enum import Enum
import json
import re
import logging
import numpy as np
from collections import defaultdict, Counter
import hashlib
from functools import lru_cache
import pickle

# Enhanced logging with more detail
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ============================================================================
# SECTION 1: ENUMS AND DATA STRUCTURES
# ============================================================================

class IntentType(Enum):
    """Enhanced intent classification with more granular types"""
    GREETING = "greeting"
    FAREWELL = "farewell"
    HELP_GENERAL = "help_general"
    HELP_SPECIFIC = "help_specific"
    METRICS_OVERVIEW = "metrics_overview"
    METRICS_CALCULATION = "metrics_calculation"
    METRICS_IMPROVEMENT = "metrics_improvement"
    NAVIGATION = "navigation"
    TROUBLESHOOTING = "troubleshooting"
    TROUBLESHOOTING_UPLOAD = "troubleshooting_upload"
    IMPROVEMENT_SUGGESTIONS = "improvement_suggestions"
    EXPORT = "export"
    UPLOAD = "upload"
    CLIMATE = "climate"
    CLIMATE_RISK = "climate_risk"
    COMPARISON = "comparison"
    BENCHMARK = "benchmark"
    COST_ANALYSIS = "cost_analysis"
    PREDICTION = "prediction"
    RECOMMENDATION = "recommendation"
    CLARIFICATION = "clarification"
    ACKNOWLEDGMENT = "acknowledgment"
    UNKNOWN = "unknown"


class SentimentType(Enum):
    """Enhanced sentiment with intensity"""
    VERY_POSITIVE = "very_positive"
    POSITIVE = "positive"
    NEUTRAL = "neutral"
    NEGATIVE = "negative"
    FRUSTRATED = "frustrated"
    CONFUSED = "confused"
    URGENT = "urgent"


class EntityType(Enum):
    """Entity types for extraction"""
    METRIC = "metric"
    PERCENTAGE = "percentage"
    LOCATION = "location"
    RIG_NAME = "rig_name"
    DATE = "date"
    COST = "cost"
    DURATION = "duration"


# ============================================================================
# SECTION 2: ADVANCED NLP AND ML COMPONENTS
# ============================================================================

class SemanticSimilarity:
    """Compute semantic similarity using TF-IDF and cosine similarity"""
    
    def __init__(self):
        self.vocabulary: Dict[str, int] = {}
        self.idf_scores: Dict[str, float] = {}
        self.document_vectors: List[np.ndarray] = []
        
    def tokenize(self, text: str) -> List[str]:
        """Advanced tokenization with stemming"""
        # Lowercase and remove special characters
        text = re.sub(r'[^a-zA-Z0-9\s]', ' ', text.lower())
        tokens = text.split()
        
        # Simple stemming rules
        stemmed = []
        for token in tokens:
            if token.endswith('ing'):
                token = token[:-3]
            elif token.endswith('ed'):
                token = token[:-2]
            elif token.endswith('s') and len(token) > 3:
                token = token[:-1]
            stemmed.append(token)
        
        return stemmed
    
    def build_vocabulary(self, documents: List[str]):
        """Build vocabulary from document corpus"""
        word_freq = Counter()
        doc_count = defaultdict(int)
        
        for doc in documents:
            tokens = set(self.tokenize(doc))
            for token in tokens:
                doc_count[token] += 1
            word_freq.update(tokens)
        
        # Build vocabulary
        self.vocabulary = {word: idx for idx, word in enumerate(word_freq.keys())}
        
        # Calculate IDF scores
        num_docs = len(documents)
        for word, count in doc_count.items():
            self.idf_scores[word] = np.log((num_docs + 1) / (count + 1)) + 1
    
    def vectorize(self, text: str) -> np.ndarray:
        """Convert text to TF-IDF vector"""
        tokens = self.tokenize(text)
        vector = np.zeros(len(self.vocabulary))
        
        # Calculate term frequency
        token_counts = Counter(tokens)
        total_tokens = len(tokens)
        
        for token, count in token_counts.items():
            if token in self.vocabulary:
                tf = count / total_tokens
                idf = self.idf_scores.get(token, 1.0)
                idx = self.vocabulary[token]
                vector[idx] = tf * idf
        
        # Normalize
        norm = np.linalg.norm(vector)
        if norm > 0:
            vector = vector / norm
        
        return vector
    
    def cosine_similarity(self, vec1: np.ndarray, vec2: np.ndarray) -> float:
        """Calculate cosine similarity between two vectors"""
        if np.linalg.norm(vec1) == 0 or np.linalg.norm(vec2) == 0:
            return 0.0
        return np.dot(vec1, vec2) / (np.linalg.norm(vec1) * np.linalg.norm(vec2))
    
    def find_most_similar(self, query: str, candidates: List[str], top_k: int = 3) -> List[Tuple[str, float]]:
        """Find most similar documents to query"""
        query_vec = self.vectorize(query)
        similarities = []
        
        for candidate in candidates:
            candidate_vec = self.vectorize(candidate)
            similarity = self.cosine_similarity(query_vec, candidate_vec)
            similarities.append((candidate, similarity))
        
        # Sort by similarity and return top k
        similarities.sort(key=lambda x: x[1], reverse=True)
        return similarities[:top_k]


class FuzzyMatcher:
    """Fuzzy string matching using Levenshtein distance"""
    
    @staticmethod
    def levenshtein_distance(s1: str, s2: str) -> int:
        """Calculate edit distance between two strings"""
        if len(s1) < len(s2):
            return FuzzyMatcher.levenshtein_distance(s2, s1)
        
        if len(s2) == 0:
            return len(s1)
        
        previous_row = range(len(s2) + 1)
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        
        return previous_row[-1]
    
    @staticmethod
    def similarity_ratio(s1: str, s2: str) -> float:
        """Calculate similarity ratio (0-1) between two strings"""
        distance = FuzzyMatcher.levenshtein_distance(s1.lower(), s2.lower())
        max_len = max(len(s1), len(s2))
        if max_len == 0:
            return 1.0
        return 1.0 - (distance / max_len)
    
    @staticmethod
    def find_best_match(query: str, candidates: List[str], threshold: float = 0.6) -> Optional[Tuple[str, float]]:
        """Find best matching string from candidates"""
        best_match = None
        best_score = 0.0
        
        for candidate in candidates:
            score = FuzzyMatcher.similarity_ratio(query, candidate)
            if score > best_score and score >= threshold:
                best_score = score
                best_match = candidate
        
        return (best_match, best_score) if best_match else None


class ContextualEmbedding:
    """Simple contextual word embeddings using co-occurrence"""
    
    def __init__(self, embedding_dim: int = 50):
        self.embedding_dim = embedding_dim
        self.word_embeddings: Dict[str, np.ndarray] = {}
        self.co_occurrence: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))
    
    def train(self, documents: List[str], window_size: int = 2):
        """Train embeddings on document corpus"""
        # Build co-occurrence matrix
        for doc in documents:
            words = doc.lower().split()
            for i, word in enumerate(words):
                for j in range(max(0, i - window_size), min(len(words), i + window_size + 1)):
                    if i != j:
                        self.co_occurrence[word][words[j]] += 1
        
        # Generate embeddings using random projection with co-occurrence weights
        vocabulary = list(self.co_occurrence.keys())
        for word in vocabulary:
            # Initialize with random vector
            embedding = np.random.randn(self.embedding_dim)
            
            # Weight by co-occurrence
            for context_word, count in self.co_occurrence[word].items():
                if context_word in vocabulary:
                    embedding += np.random.randn(self.embedding_dim) * np.log1p(count)
            
            # Normalize
            embedding = embedding / (np.linalg.norm(embedding) + 1e-8)
            self.word_embeddings[word] = embedding
    
    def get_embedding(self, word: str) -> np.ndarray:
        """Get embedding for a word"""
        word = word.lower()
        if word in self.word_embeddings:
            return self.word_embeddings[word]
        return np.random.randn(self.embedding_dim) * 0.01
    
    def sentence_embedding(self, sentence: str) -> np.ndarray:
        """Get sentence embedding by averaging word embeddings"""
        words = sentence.lower().split()
        if not words:
            return np.zeros(self.embedding_dim)
        
        embeddings = [self.get_embedding(word) for word in words]
        return np.mean(embeddings, axis=0)
    
    def similarity(self, word1: str, word2: str) -> float:
        """Calculate cosine similarity between two words"""
        emb1 = self.get_embedding(word1)
        emb2 = self.get_embedding(word2)
        
        norm1 = np.linalg.norm(emb1)
        norm2 = np.linalg.norm(emb2)
        
        if norm1 == 0 or norm2 == 0:
            return 0.0
        
        return np.dot(emb1, emb2) / (norm1 * norm2)


class IntentClassifier:
    """Machine learning intent classifier using Naive Bayes"""
    
    def __init__(self):
        self.class_word_counts: Dict[IntentType, Counter] = defaultdict(Counter)
        self.class_totals: Dict[IntentType, int] = defaultdict(int)
        self.vocabulary: Set[str] = set()
        self.prior_probs: Dict[IntentType, float] = {}
        self.is_trained = False
    
    def tokenize(self, text: str) -> List[str]:
        """Tokenize text"""
        text = re.sub(r'[^a-zA-Z0-9\s]', ' ', text.lower())
        return [word for word in text.split() if len(word) > 2]
    
    def train(self, training_data: List[Tuple[str, IntentType]]):
        """Train Naive Bayes classifier"""
        total_docs = len(training_data)
        class_counts = Counter(intent for _, intent in training_data)
        
        # Calculate prior probabilities
        for intent, count in class_counts.items():
            self.prior_probs[intent] = count / total_docs
        
        # Count words per class
        for text, intent in training_data:
            tokens = self.tokenize(text)
            self.vocabulary.update(tokens)
            self.class_word_counts[intent].update(tokens)
            self.class_totals[intent] += len(tokens)
        
        self.is_trained = True
    
    def predict(self, text: str, top_k: int = 3) -> List[Tuple[IntentType, float]]:
        """Predict intent with probability scores"""
        if not self.is_trained:
            return [(IntentType.UNKNOWN, 1.0)]
        
        tokens = self.tokenize(text)
        scores = {}
        
        vocab_size = len(self.vocabulary)
        
        for intent in self.prior_probs.keys():
            # Log probability to avoid underflow
            log_prob = np.log(self.prior_probs[intent])
            
            for token in tokens:
                # Laplace smoothing
                token_count = self.class_word_counts[intent][token]
                token_prob = (token_count + 1) / (self.class_totals[intent] + vocab_size)
                log_prob += np.log(token_prob)
            
            scores[intent] = log_prob
        
        # Convert to probabilities
        max_score = max(scores.values())
        exp_scores = {intent: np.exp(score - max_score) for intent, score in scores.items()}
        total = sum(exp_scores.values())
        probabilities = {intent: score / total for intent, score in exp_scores.items()}
        
        # Sort and return top k
        sorted_intents = sorted(probabilities.items(), key=lambda x: x[1], reverse=True)
        return sorted_intents[:top_k]


class SentimentAnalyzer:
    """Advanced sentiment analysis with intensity scoring"""
    
    def __init__(self):
        self.positive_words = {
            'great', 'good', 'excellent', 'perfect', 'amazing', 'wonderful', 'fantastic',
            'brilliant', 'awesome', 'love', 'appreciate', 'thank', 'thanks', 'helpful',
            'clear', 'easy', 'simple', 'useful', 'nice', 'best', 'better', 'improved'
        }
        
        self.negative_words = {
            'bad', 'terrible', 'awful', 'horrible', 'poor', 'worst', 'hate', 'frustrat',
            'angry', 'upset', 'annoyed', 'difficult', 'hard', 'complex', 'complicated',
            'confus', 'stuck', 'broken', 'error', 'fail', 'wrong', 'problem', 'issue'
        }
        
        self.intensifiers = {
            'very', 'extremely', 'really', 'incredibly', 'absolutely', 'totally',
            'completely', 'utterly', 'so', 'too'
        }
        
        self.negators = {'not', 'no', 'never', 'neither', 'nobody', 'nothing', "n't"}
    
    def analyze(self, text: str) -> Tuple[SentimentType, float]:
        """Analyze sentiment with intensity score"""
        words = text.lower().split()
        
        positive_score = 0.0
        negative_score = 0.0
        intensity_multiplier = 1.0
        negation_active = False
        
        for i, word in enumerate(words):
            # Check for intensifiers
            if word in self.intensifiers:
                intensity_multiplier = 1.5
                continue
            
            # Check for negators
            if word in self.negators:
                negation_active = True
                continue
            
            # Check sentiment
            word_clean = word.strip('.,!?')
            
            if any(pos in word_clean for pos in self.positive_words):
                score = 1.0 * intensity_multiplier
                if negation_active:
                    negative_score += score
                else:
                    positive_score += score
            
            if any(neg in word_clean for neg in self.negative_words):
                score = 1.0 * intensity_multiplier
                if negation_active:
                    positive_score += score
                else:
                    negative_score += score
            
            # Reset modifiers
            intensity_multiplier = 1.0
            negation_active = False
        
        # Calculate net sentiment
        net_score = positive_score - negative_score
        total_score = positive_score + negative_score
        
        if total_score == 0:
            return SentimentType.NEUTRAL, 0.5
        
        sentiment_ratio = (net_score / total_score + 1) / 2  # Normalize to 0-1
        
        # Classify sentiment
        if sentiment_ratio > 0.7:
            return SentimentType.VERY_POSITIVE, sentiment_ratio
        elif sentiment_ratio > 0.55:
            return SentimentType.POSITIVE, sentiment_ratio
        elif sentiment_ratio < 0.3:
            return SentimentType.FRUSTRATED, sentiment_ratio
        elif sentiment_ratio < 0.45:
            return SentimentType.NEGATIVE, sentiment_ratio
        else:
            return SentimentType.NEUTRAL, sentiment_ratio


# ============================================================================
# SECTION 3: CONTEXT MANAGEMENT AND DIALOGUE STATE
# ============================================================================

@dataclass
class ConversationContext:
    """Advanced conversation context with full state tracking"""
    # Data state
    has_data: bool = False
    analysis_complete: bool = False
    current_rig: Optional[str] = None
    efficiency_score: float = 0.0
    
    # Conversation state
    last_intent: Optional[IntentType] = None
    last_entities: Dict[str, List[str]] = field(default_factory=dict)
    sentiment: SentimentType = SentimentType.NEUTRAL
    sentiment_score: float = 0.5
    
    # Topic tracking
    mentioned_topics: List[str] = field(default_factory=list)
    current_topic: Optional[str] = None
    topic_depth: int = 0
    
    # User profile
    expertise_level: str = "beginner"  # beginner, intermediate, expert
    preferred_detail_level: str = "medium"  # brief, medium, detailed
    interaction_count: int = 0
    
    # Multi-turn dialogue
    awaiting_clarification: bool = False
    clarification_context: Optional[str] = None
    last_question: Optional[str] = None
    follow_up_expected: bool = False
    
    # Performance tracking
    successful_resolutions: int = 0
    failed_resolutions: int = 0
    average_confidence: float = 0.0


@dataclass
class ChatMessage:
    """Enhanced chat message with metadata"""
    role: str
    content: str
    timestamp: datetime
    intent: Optional[IntentType] = None
    intent_confidence: float = 0.0
    sentiment: Optional[SentimentType] = None
    sentiment_score: float = 0.5
    entities: Dict[str, List[str]] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)
    response_time: float = 0.0


class DialogueManager:
    """Manages multi-turn dialogue and conversation flow"""
    
    def __init__(self):
        self.conversation_state = {}
        self.topic_transitions = defaultdict(list)
        self.topic_stack = []
        
    def push_topic(self, topic: str):
        """Push new topic to stack"""
        self.topic_stack.append(topic)
    
    def pop_topic(self) -> Optional[str]:
        """Pop topic from stack"""
        return self.topic_stack.pop() if self.topic_stack else None
    
    def current_topic(self) -> Optional[str]:
        """Get current topic"""
        return self.topic_stack[-1] if self.topic_stack else None
    
    def should_clarify(self, intent_confidence: float, entity_count: int) -> bool:
        """Determine if clarification is needed"""
        # Disabled clarification - always provide best answer
        # The advanced knowledge base search will handle ambiguity
        return False  # Never ask for clarification, always provide best response
    
    def generate_clarification_question(self, intent: IntentType, user_input: str) -> str:
        """Generate appropriate clarification question"""
        clarifications = {
            IntentType.METRICS_OVERVIEW: "Which specific metric would you like to know about? (e.g., Contract Utilization, Dayrate Efficiency, Climate Impact)",
            IntentType.IMPROVEMENT_SUGGESTIONS: "Which area would you like to improve? (e.g., efficiency, costs, utilization)",
            IntentType.TROUBLESHOOTING: "What specific issue are you experiencing? (e.g., upload error, missing data, incorrect calculations)",
            IntentType.NAVIGATION: "Which feature or section are you trying to find?",
            IntentType.COMPARISON: "Which rigs would you like to compare, or do you want to compare all rigs?",
        }
        
        return clarifications.get(
            intent,
            "Could you please provide more details about what you're looking for?"
        )


# ============================================================================
# SECTION 4: KNOWLEDGE BASE AND RESPONSE GENERATION
# ============================================================================

class KnowledgeBase:
    """Comprehensive knowledge base with semantic search"""
    
    def __init__(self):
        self.semantic_engine = SemanticSimilarity()
        self.knowledge_entries = []
        self.embeddings = None
        self._init_knowledge()
    
    def _init_knowledge(self):
        """Initialize knowledge base"""
        self.knowledge_entries = [
            {
                "id": "efficiency_overview",
                "topic": "efficiency_metrics",
                "keywords": ["efficiency", "score", "metrics", "overall", "performance"],
                "content": """EFFICIENCY METRICS OVERVIEW

Overall Efficiency = Weighted average of 6 key factors (0-100%)

KEY FACTORS:
1. Contract Utilization (25%) - % of potential contract days worked
2. Dayrate Efficiency (20%) - Operating cost management
3. Contract Stability (15%) - Duration and consistency
4. Location Complexity (15%) - Geographic challenges
5. Climate Impact (10%) - Weather risk management
6. Contract Performance (15%) - Overall execution quality

GRADING SCALE:
- A+ (90-100%): Exceptional performance
- A (80-89%): Excellent performance
- B (70-79%): Good performance
- C (60-69%): Fair performance
- F (<60%): Needs improvement

Higher scores indicate better overall rig performance and profitability."""
            },
            {
                "id": "efficiency_calculation",
                "topic": "efficiency_metrics",
                "keywords": ["calculate", "calculation", "formula", "algorithm", "compute", "how does", "how is"],
                "content": """EFFICIENCY CALCULATION FORMULA

Overall Efficiency = Σ(Factor_i × Weight_i) for i=1 to 6

DETAILED WEIGHTS:
- Contract Utilization: 25% (0.25)
- Dayrate Efficiency: 20% (0.20)
- Contract Stability: 15% (0.15)
- Location Complexity: 15% (0.15)
- Climate Impact: 10% (0.10)
- Contract Performance: 15% (0.15)

EXAMPLE CALCULATION:
If your rig scores:
- Contract Utilization: 85% × 0.25 = 21.25
- Dayrate Efficiency: 75% × 0.20 = 15.00
- Contract Stability: 80% × 0.15 = 12.00
- Location Complexity: 70% × 0.15 = 10.50
- Climate Impact: 65% × 0.10 = 6.50
- Contract Performance: 90% × 0.15 = 13.50

Total = 78.75% (Grade B - Good Performance)

Each factor is calculated based on specific operational data and industry benchmarks."""
            },
            {
                "id": "contract_utilization",
                "topic": "metrics_detail",
                "keywords": ["contract utilization", "utilization", "working days", "active days"],
                "content": """CONTRACT UTILIZATION EXPLAINED

Definition: Percentage of potential contract days actually worked

Formula: (Actual Working Days / Total Contract Days) × 100

TARGET RANGES:
- Excellent: 85-95%
- Good: 75-84%
- Fair: 65-74%
- Poor: <65%

IMPROVEMENT STRATEGIES:
1. Minimize planned downtime
2. Optimize maintenance scheduling
3. Improve weather planning
4. Enhance operational efficiency
5. Better contract negotiation

Weight: 25% (highest impact on overall score)"""
            },
            {
                "id": "dayrate_efficiency",
                "topic": "metrics_detail",
                "keywords": ["dayrate", "day rate", "cost", "operating cost", "expenses"],
                "content": """DAYRATE EFFICIENCY EXPLAINED

Definition: How effectively operating costs are managed relative to dayrate income

Formula: (Dayrate - Operating Costs) / Dayrate × 100

TARGET RANGES:
- Excellent: 75-85%
- Good: 65-74%
- Fair: 55-64%
- Poor: <55%

IMPROVEMENT STRATEGIES:
1. Reduce fuel consumption
2. Optimize crew scheduling
3. Negotiate better supplier rates
4. Preventive maintenance
5. Technology upgrades

Weight: 20% of overall efficiency score"""
            },
            {
                "id": "climate_analysis",
                "topic": "climate",
                "keywords": ["climate", "weather", "risk", "downtime", "hurricane", "storm", "season"],
                "content": """CLIMATE RISK ANALYSIS

WHAT WE ANALYZE:
- Historical weather patterns
- Seasonal risk profiles
- Regional climate characteristics
- Downtime probability
- Cost impact estimates

MAJOR RISK ZONES:
1. Gulf of Mexico: Hurricane season (June-November)
   - Peak risk: August-October
   - Average downtime: 15-25 days/year

2. North Sea: Winter storms (November-March)
   - Peak risk: December-February
   - Average downtime: 10-20 days/year

3. Southeast Asia: Typhoon season (May-December)
   - Peak risk: July-September
   - Average downtime: 12-18 days/year

4. Middle East: Minimal weather risk
   - Year-round operations
   - Average downtime: <5 days/year

MITIGATION STRATEGIES:
- Schedule maintenance during high-risk periods
- Invest in weather forecasting
- Consider weather insurance
- Plan operations around seasonal patterns"""
            },
            {
                "id": "upload_guide",
                "topic": "upload",
                "keywords": ["upload", "file", "data", "import", "excel", "csv", "format"],
                "content": """DATA UPLOAD COMPREHENSIVE GUIDE

REQUIRED FIELDS:
1. Rig Name / Drilling Unit Name (text)
2. Contract Start Date (YYYY-MM-DD)
3. Contract End Date (YYYY-MM-DD)
4. Dayrate in $thousands (numeric)
5. Current Location (text)

OPTIONAL BUT RECOMMENDED:
- Contract Value ($)
- Client Name
- Well Type (Exploration/Development/Appraisal)
- Water Depth (feet/meters)
- Rig Type (Jackup/Semi-sub/Drillship)

SUPPORTED FILE FORMATS:
- Excel (.xlsx, .xls) - Max 10MB
- CSV (comma-separated) - Max 10MB
- TSV (tab-separated) - Max 10MB

FILE PREPARATION TIPS:
1. Remove password protection
2. No merged cells
3. Clear headers in row 1
4. No blank rows
5. Consistent date format
6. Numeric values without $ or commas

COMMON ERRORS:
- "Missing columns" → Check required field names
- "Date format error" → Use YYYY-MM-DD
- "File too large" → Split into multiple files
- "Upload failed" → Close file in Excel first

UPLOAD PROCESS:
1. Click "Upload Data" in sidebar
2. Drag & drop or browse for file
3. Wait for validation (green checkmark)
4. Review data preview
5. Proceed to analysis"""
            },
            {
                "id": "improvement_strategies",
                "topic": "improvement",
                "keywords": ["improve", "increase", "optimize", "enhance", "boost", "better"],
                "content": """PERFORMANCE IMPROVEMENT STRATEGIES

QUICK WINS (0-3 months):
1. Optimize maintenance scheduling
2. Improve crew efficiency
3. Better weather planning
4. Reduce idle time
5. Enhance communication

MEDIUM-TERM (3-12 months):
1. Contract renegotiation
2. Technology upgrades
3. Crew training programs
4. Supplier relationship optimization
5. Process automation

LONG-TERM (12+ months):
1. Fleet modernization
2. Geographic diversification
3. Market positioning
4. Strategic partnerships
5. Innovation initiatives

METRIC-SPECIFIC IMPROVEMENTS:

Contract Utilization:
- Reduce planned downtime
- Improve maintenance efficiency
- Better contract terms
- Enhanced weather planning

Dayrate Efficiency:
- Cost optimization
- Supplier negotiations
- Energy efficiency
- Preventive maintenance

Contract Stability:
- Long-term relationships
- Performance excellence
- Market intelligence
- Competitive pricing

Climate Risk:
- Seasonal planning
- Weather insurance
- Geographic diversity
- Advanced forecasting"""
            },
            {
                "id": "troubleshooting_general",
                "topic": "troubleshooting",
                "keywords": ["problem", "error", "issue", "bug", "not working", "broken", "fix"],
                "content": """GENERAL TROUBLESHOOTING GUIDE

QUICK DIAGNOSTIC CHECKLIST:
□ Is your internet connection stable?
□ Is your browser up to date?
□ Have you tried refreshing (F5)?
□ Is the file format correct?
□ Are all required fields present?

COMMON ISSUES & SOLUTIONS:

1. FILE UPLOAD FAILURES
   - Close file in Excel/Sheets
   - Check file size (<10MB)
   - Verify file format (.xlsx, .csv)
   - Remove password protection
   - Try different browser

2. MISSING OR INCORRECT DATA
   - Verify all required fields
   - Check date formats (YYYY-MM-DD)
   - Ensure numeric fields are numbers
   - Remove special characters
   - Validate data completeness

3. CALCULATION ERRORS
   - Ensure data types are correct
   - Check for negative values
   - Verify date ranges are valid
   - Confirm all fields populated
   - Re-upload if necessary

4. DISPLAY/VISUALIZATION ISSUES
   - Refresh page (Ctrl+F5)
   - Clear browser cache
   - Try different browser
   - Check zoom level (100%)
   - Disable browser extensions

5. EXPORT PROBLEMS
   - Wait for analysis to complete
   - Ensure pop-ups are enabled
   - Check download folder settings
   - Try different export format
   - Verify sufficient disk space

ADVANCED TROUBLESHOOTING:
- Clear cookies and cache
- Try incognito/private mode
- Disable ad blockers
- Check firewall settings
- Update browser to latest version

STILL STUCK?
Provide details about:
- What you were trying to do
- Exact error message
- Steps you've already tried
- Browser and operating system"""
            },
            {
                "id": "navigation_guide",
                "topic": "navigation",
                "keywords": ["where", "find", "locate", "navigate", "menu", "button", "feature", "section"],
                "content": """PLATFORM NAVIGATION GUIDE

MAIN NAVIGATION (Sidebar):
📁 Upload Data - Import rig information
📊 Single Rig Analysis - Analyze one rig in detail
📈 Multi-Rig Comparison - Compare multiple rigs
🎯 Benchmarking - Compare vs industry standards
📥 Export Reports - Download analysis results
⚙️ Settings - Configure preferences

ANALYSIS SECTIONS:
Once you analyze a rig, you'll see these tabs:

1. SUMMARY TAB
   - Overall efficiency score
   - Grade and rating
   - Key metrics overview
   - Quick insights

2. DETAILED METRICS TAB
   - Individual factor scores
   - Component breakdown
   - Performance indicators
   - Historical trends

3. CLIMATE ANALYSIS TAB
   - Weather risk assessment
   - Seasonal patterns
   - Downtime predictions
   - Regional comparisons

4. INSIGHTS & RECOMMENDATIONS TAB
   - AI-generated suggestions
   - Improvement opportunities
   - Best practices
   - Action items

5. COMPARISON TAB
   - Benchmark comparisons
   - Peer analysis
   - Industry standards
   - Competitive positioning

6. WELL-RIG NAVIGATOR TAB
   - Well matching
   - Optimization suggestions
   - Route planning
   - Compatibility analysis

QUICK TIPS:
- Use sidebar for main navigation
- Hover over metrics for tooltips
- Click tabs to switch views
- Export button at bottom of results
- Help icon (?) for context-sensitive help"""
            },
            {
                "id": "export_guide",
                "topic": "export",
                "keywords": ["export", "download", "report", "save", "pdf", "excel", "share"],
                "content": """EXPORT & REPORTING GUIDE

AVAILABLE EXPORT FORMATS:

1. EXCEL (.xlsx)
   - Full data with all metrics
   - Multiple worksheets
   - Formulas included
   - Charts and graphs
   - Customizable format
   - Best for: Detailed analysis

2. PDF REPORT
   - Professional formatted report
   - Executive summary
   - Key visualizations
   - Print-ready format
   - Best for: Presentations, sharing

3. CSV DATA (.csv)
   - Raw data export
   - All calculated metrics
   - Easy import to other tools
   - Best for: Further analysis, databases

4. CHARTS & VISUALIZATIONS
   - Individual chart export
   - PNG or SVG format
   - High resolution
   - Best for: Presentations, documents

EXPORT PROCESS:
1. Complete your analysis
2. Scroll to bottom of results page
3. Click "Export" button
4. Select desired format
5. Choose specific sections (optional)
6. Click "Generate Report"
7. Download automatically starts

CUSTOMIZATION OPTIONS:
- Select which sections to include
- Choose date range
- Include/exclude raw data
- Add custom branding (premium)
- Set report template
- Configure chart styles

REPORT CONTENTS:
- Executive summary
- Overall efficiency scores
- Detailed metric breakdown
- Climate analysis
- Improvement recommendations
- Comparative analysis
- Historical trends
- Visualizations and charts

PRO TIPS:
- Export regularly to track progress
- Use PDF for stakeholder updates
- Use Excel for deep-dive analysis
- Save charts for presentations
- Maintain export archive
- Schedule automated reports (premium)

SHARING REPORTS:
- Email directly from platform
- Generate shareable link
- Set access permissions
- Password protect sensitive data
- Track report views (premium)"""
            },
            {
                "id": "comparison_guide",
                "topic": "comparison",
                "keywords": ["compare", "comparison", "versus", "vs", "multiple rigs", "fleet"],
                "content": """MULTI-RIG COMPARISON GUIDE

WHAT YOU CAN COMPARE:
- Overall efficiency scores
- Individual metrics
- Cost structures
- Operational performance
- Geographic performance
- Time-based trends
- Client satisfaction
- Climate risk profiles

COMPARISON TYPES:

1. SIDE-BY-SIDE COMPARISON
   - 2-5 rigs at once
   - Detailed metric comparison
   - Visual charts
   - Ranking by performance

2. FLEET ANALYSIS
   - All rigs in your fleet
   - Statistical overview
   - Best/worst performers
   - Fleet averages
   - Distribution analysis

3. TEMPORAL COMPARISON
   - Same rig over time
   - Track improvements
   - Identify trends
   - Seasonal patterns

4. BENCHMARK COMPARISON
   - Your rigs vs industry
   - Regional benchmarks
   - Rig type averages
   - Competitive positioning

HOW TO COMPARE:
1. Upload fleet data (or use existing)
2. Go to "Multi-Rig Comparison"
3. Select rigs to compare
4. Choose comparison metrics
5. Generate comparison report
6. Analyze results

INSIGHTS PROVIDED:
- Performance ranking
- Best practice identification
- Improvement priorities
- Resource allocation guidance
- Strategic recommendations
- Outlier detection
- Trend analysis

VISUALIZATION OPTIONS:
- Bar charts (rankings)
- Radar charts (multi-metric)
- Line charts (trends)
- Heat maps (performance matrix)
- Scatter plots (correlations)
- Box plots (distributions)

STRATEGIC VALUE:
- Identify best practices
- Optimize resource allocation
- Recognize underperformers
- Guide investment decisions
- Support strategic planning
- Benchmark against competition
- Track fleet improvements

COMPARISON METRICS:
✓ Overall efficiency
✓ Contract utilization
✓ Dayrate efficiency
✓ Contract stability
✓ Location complexity
✓ Climate impact
✓ Contract performance
✓ Cost per day
✓ Revenue per day
✓ Profitability
✓ Utilization rate
✓ Downtime percentage"""
            },
            {
                "id": "cost_analysis",
                "topic": "cost",
                "keywords": ["cost", "expense", "budget", "spending", "financial", "money", "profit"],
                "content": """COST ANALYSIS & OPTIMIZATION

KEY COST CATEGORIES:

1. OPERATING COSTS
   - Personnel (30-40%)
   - Fuel & energy (20-25%)
   - Maintenance (15-20%)
   - Insurance (5-10%)
   - Administration (5-10%)

2. CAPITAL COSTS
   - Equipment purchases
   - Major upgrades
   - Technology investments
   - Fleet expansion

3. VARIABLE COSTS
   - Weather-related delays
   - Emergency repairs
   - Additional crew
   - Spot market supplies

COST OPTIMIZATION STRATEGIES:

IMMEDIATE ACTIONS:
- Audit current spending
- Identify waste
- Negotiate supplier contracts
- Optimize crew scheduling
- Reduce energy consumption

SHORT-TERM (1-6 months):
- Implement preventive maintenance
- Upgrade to efficient equipment
- Streamline operations
- Consolidate suppliers
- Automate processes

LONG-TERM (6+ months):
- Fleet modernization
- Technology adoption
- Process reengineering
- Strategic partnerships
- Market repositioning

COST METRICS TO TRACK:
- Cost per operating day
- Cost per barrel of oil
- Maintenance cost ratio
- Personnel efficiency
- Energy efficiency
- Administrative overhead

BENCHMARKING:
- Compare to industry averages
- Regional cost differences
- Rig type variations
- Seasonal fluctuations
- Market conditions

ROI CALCULATIONS:
- Cost reduction initiatives
- Technology investments
- Training programs
- Process improvements
- Equipment upgrades"""
            }
        ]
        
        # Build semantic search capability
        documents = [entry["content"] + " " + " ".join(entry["keywords"]) 
                    for entry in self.knowledge_entries]
        self.semantic_engine.build_vocabulary(documents)
    
    def search(self, query: str, top_k: int = 3) -> List[Dict]:
        """Search knowledge base using semantic similarity"""
        candidates = [entry["content"] for entry in self.knowledge_entries]
        similar = self.semantic_engine.find_most_similar(query, candidates, top_k)
        
        results = []
        for content, score in similar:
            # Find corresponding entry
            for entry in self.knowledge_entries:
                if entry["content"] == content:
                    results.append({
                        **entry,
                        "relevance_score": score
                    })
                    break
        
        return results
    
    def get_by_topic(self, topic: str) -> List[Dict]:
        """Get all entries for a specific topic"""
        return [entry for entry in self.knowledge_entries if entry["topic"] == topic]
    
    def get_by_id(self, entry_id: str) -> Optional[Dict]:
        """Get specific entry by ID"""
        for entry in self.knowledge_entries:
            if entry["id"] == entry_id:
                return entry
        return None


class ResponseGenerator:
    """Advanced response generation with templates and personalization"""
    
    def __init__(self):
        self.response_templates = self._init_templates()
        self.fuzzy_matcher = FuzzyMatcher()
    
    def _init_templates(self) -> Dict[str, List[str]]:
        """Initialize response templates"""
        return {
            "greeting": [
                "Hello! I'm your Rig Efficiency AI Assistant. {context}",
                "Hi there! Ready to help with your rig analysis. {context}",
                "Greetings! I'm here to assist with efficiency optimization. {context}"
            ],
            "farewell": [
                "Goodbye! Feel free to return anytime you need help.",
                "Thanks for using the Rig Efficiency Assistant. Have a great day!",
                "See you next time! Don't hesitate to reach out if you need assistance."
            ],
            "clarification": [
                "To better assist you, could you clarify: {question}",
                "I want to make sure I understand correctly. {question}",
                "Just to confirm, {question}"
            ],
            "acknowledgment": [
                "I understand you're asking about {topic}. {response}",
                "Great question about {topic}! {response}",
                "Regarding {topic}: {response}"
            ],
            "error_recovery": [
                "I apologize for the confusion. Let me try a different approach. {response}",
                "Let me rephrase that to be more helpful. {response}",
                "I'll provide a clearer explanation. {response}"
            ]
        }
    
    def generate(self, template_type: str, **kwargs) -> str:
        """Generate response from template"""
        if template_type not in self.response_templates:
            return kwargs.get("response", "I'm here to help!")
        
        templates = self.response_templates[template_type]
        template = np.random.choice(templates)
        
        try:
            return template.format(**kwargs)
        except KeyError:
            return template
    
    def personalize_response(self, response: str, context: ConversationContext) -> str:
        """Personalize response based on user context"""
        # Adjust detail level - DISABLED to prevent truncation
        # Users need full responses regardless of expertise level
        # if context.expertise_level == "beginner" and len(response) > 800:
        #     response = self._simplify_response(response)
        # elif context.expertise_level == "expert" and len(response) < 400:
        #     response = self._enhance_response(response)
        
        # Add sentiment-appropriate tone
        if context.sentiment in [SentimentType.FRUSTRATED, SentimentType.NEGATIVE]:
            response = "I understand this can be challenging. " + response
        elif context.sentiment == SentimentType.CONFUSED:
            response = "Let me clarify that for you. " + response
        elif context.sentiment in [SentimentType.POSITIVE, SentimentType.VERY_POSITIVE]:
            response = "Great! " + response
        
        return response
    
    def _simplify_response(self, response: str) -> str:
        """Simplify response for beginners"""
        # Keep first 500 characters and add summary
        if len(response) > 500:
            summary = response[:500].rsplit('.', 1)[0] + ".\n\n💡 Want more details? Just ask!"
            return summary
        return response
    
    def _enhance_response(self, response: str) -> str:
        """Enhance response for experts"""
        # Add technical details placeholder
        enhancement = "\n\n📊 For technical implementation details, formulas, or API access, please ask specifically."
        return response + enhancement


# ============================================================================
# SECTION 5: MAIN CHATBOT CLASS
# ============================================================================

class AdvancedRigEfficiencyAIChatbot:
    """Next-generation AI chatbot with ML capabilities"""
    
    def __init__(self):
        # Core components
        self.semantic_engine = SemanticSimilarity()
        self.fuzzy_matcher = FuzzyMatcher()
        self.intent_classifier = IntentClassifier()
        self.sentiment_analyzer = SentimentAnalyzer()
        self.knowledge_base = KnowledgeBase()
        self.response_generator = ResponseGenerator()
        self.dialogue_manager = DialogueManager()
        self.embeddings = ContextualEmbedding()
        
        # State management
        self.context = ConversationContext()
        self.chat_history: List[ChatMessage] = []
        self.response_cache: Dict[str, Tuple[str, float]] = {}  # query: (response, timestamp)
        
        # Configuration
        self.max_history = 100
        self.cache_ttl = 3600  # 1 hour
        self.confidence_threshold = 0.6
        
        # Initialize ML components
        self._train_intent_classifier()
        self._train_embeddings()
        
        logger.info("Advanced Rig Efficiency AI Chatbot initialized successfully")
    
    def _train_intent_classifier(self):
        """Train intent classifier with examples"""
        training_data = [
            # Greetings
            ("hello", IntentType.GREETING),
            ("hi there", IntentType.GREETING),
            ("good morning", IntentType.GREETING),
            ("hey", IntentType.GREETING),
            
            # Help requests
            ("I need help", IntentType.HELP_GENERAL),
            ("can you help me", IntentType.HELP_GENERAL),
            ("how do I use this", IntentType.HELP_GENERAL),
            ("what can you do", IntentType.HELP_GENERAL),
            ("guide me about tool", IntentType.HELP_GENERAL),
            ("tell me about this tool", IntentType.HELP_GENERAL),
            ("how does this tool work", IntentType.HELP_GENERAL),
            ("what is this platform", IntentType.HELP_GENERAL),
            ("show me features", IntentType.HELP_GENERAL),
            ("tool guide", IntentType.HELP_GENERAL),
            ("platform guide", IntentType.HELP_GENERAL),
            ("help with metrics", IntentType.HELP_SPECIFIC),
            ("explain efficiency", IntentType.HELP_SPECIFIC),
            
            # Metrics
            ("what is efficiency score", IntentType.METRICS_OVERVIEW),
            ("explain metrics", IntentType.METRICS_OVERVIEW),
            ("what are the factors", IntentType.METRICS_OVERVIEW),
            ("how is efficiency calculated", IntentType.METRICS_CALCULATION),
            ("show me the formula", IntentType.METRICS_CALCULATION),
            ("calculation method", IntentType.METRICS_CALCULATION),
            ("how does it calculate", IntentType.METRICS_CALCULATION),
            ("what is the formula", IntentType.METRICS_CALCULATION),
            ("how do you calculate", IntentType.METRICS_CALCULATION),
            ("explain the calculation", IntentType.METRICS_CALCULATION),
            ("calculation formula", IntentType.METRICS_CALCULATION),
            ("how to improve efficiency", IntentType.METRICS_IMPROVEMENT),
            ("how can I improve", IntentType.METRICS_IMPROVEMENT),
            ("improve my score", IntentType.METRICS_IMPROVEMENT),
            ("increase efficiency", IntentType.METRICS_IMPROVEMENT),
            ("boost performance", IntentType.METRICS_IMPROVEMENT),
            ("increase my score", IntentType.METRICS_IMPROVEMENT),
            ("optimize performance", IntentType.METRICS_IMPROVEMENT),
            ("get better results", IntentType.METRICS_IMPROVEMENT),
            ("enhance efficiency", IntentType.METRICS_IMPROVEMENT),
            
            # Navigation
            ("where is upload", IntentType.NAVIGATION),
            ("find the export button", IntentType.NAVIGATION),
            ("how to navigate", IntentType.NAVIGATION),
            ("where can I find", IntentType.NAVIGATION),
            
            # Upload
            ("how to upload data", IntentType.UPLOAD),
            ("file format", IntentType.UPLOAD),
            ("what file format", IntentType.UPLOAD),
            ("what format do I need", IntentType.UPLOAD),
            ("file format do I need", IntentType.UPLOAD),
            ("supported formats", IntentType.UPLOAD),
            ("upload requirements", IntentType.UPLOAD),
            ("import data", IntentType.UPLOAD),
            ("data format", IntentType.UPLOAD),
            ("excel or csv", IntentType.UPLOAD),
            
            # Climate
            ("weather risk", IntentType.CLIMATE),
            ("climate analysis", IntentType.CLIMATE),
            ("seasonal downtime", IntentType.CLIMATE),
            ("hurricane season", IntentType.CLIMATE_RISK),
            
            # Troubleshooting
            ("upload not working", IntentType.TROUBLESHOOTING_UPLOAD),
            ("error message", IntentType.TROUBLESHOOTING),
            ("something is broken", IntentType.TROUBLESHOOTING),
            ("I'm having trouble", IntentType.TROUBLESHOOTING),
            
            # Comparison
            ("compare rigs", IntentType.COMPARISON),
            ("multi-rig analysis", IntentType.COMPARISON),
            ("which rig is better", IntentType.COMPARISON),
            
            # Export
            ("download report", IntentType.EXPORT),
            ("export data", IntentType.EXPORT),
            ("save results", IntentType.EXPORT),
            
            # Improvements
            ("suggest improvements", IntentType.IMPROVEMENT_SUGGESTIONS),
            ("how to optimize", IntentType.IMPROVEMENT_SUGGESTIONS),
            ("best practices", IntentType.IMPROVEMENT_SUGGESTIONS),
            
            # Cost
            ("reduce costs", IntentType.COST_ANALYSIS),
            ("cost optimization", IntentType.COST_ANALYSIS),
            ("expenses", IntentType.COST_ANALYSIS),
            
            # Farewell
            ("goodbye", IntentType.FAREWELL),
            ("bye", IntentType.FAREWELL),
            ("thanks", IntentType.ACKNOWLEDGMENT),
        ]
        
        self.intent_classifier.train(training_data)
        logger.info(f"Intent classifier trained with {len(training_data)} examples")
    
    def _train_embeddings(self):
        """Train word embeddings on domain corpus"""
        corpus = [entry["content"] for entry in self.knowledge_base.knowledge_entries]
        self.embeddings.train(corpus)
        logger.info("Word embeddings trained on knowledge base")
    
    @lru_cache(maxsize=1000)
    def _get_cached_response(self, query_hash: str) -> Optional[str]:
        """Get cached response if valid"""
        if query_hash in self.response_cache:
            response, timestamp = self.response_cache[query_hash]
            if (datetime.now().timestamp() - timestamp) < self.cache_ttl:
                logger.info(f"Cache hit for query hash: {query_hash[:8]}")
                return response
            else:
                # Remove expired cache entry
                del self.response_cache[query_hash]
        return None
    
    def _cache_response(self, query: str, response: str):
        """Cache response with timestamp"""
        query_hash = hashlib.md5(query.lower().encode()).hexdigest()
        self.response_cache[query_hash] = (response, datetime.now().timestamp())
    
    def generate_response(self, user_input: str, context: Dict = None) -> str:
        """Generate intelligent response with full AI pipeline
        
        Args:
            user_input: User's question or message
            context: Optional context dict from app.py with keys:
                    - has_data: bool
                    - current_rig: str
                    - analysis_complete: bool
        """
        start_time = datetime.now()
        
        try:
            # Input validation
            if not self._validate_input(user_input):
                return "Please provide a valid question (3-500 characters)."
            
            # Update context from app.py if provided
            if context:
                self.context.has_data = context.get('has_data', False)
                self.context.current_rig = context.get('current_rig')
                self.context.analysis_complete = context.get('analysis_complete', False)
            
            # Check cache
            query_hash = hashlib.md5(user_input.lower().encode()).hexdigest()
            cached = self._get_cached_response(query_hash)
            if cached:
                return cached
            
            # NLP Processing Pipeline
            intent_predictions = self.intent_classifier.predict(user_input)
            best_intent, intent_confidence = intent_predictions[0]
            
            # Override for short greetings
            user_lower = user_input.lower().strip()
            if user_lower in ['hey', 'hi', 'hello', 'hey there', 'hi there', 'good morning', 'good afternoon', 'good evening']:
                best_intent = IntentType.GREETING
                intent_confidence = 0.99
            
            sentiment, sentiment_score = self.sentiment_analyzer.analyze(user_input)
            
            # Extract entities (basic pattern matching)
            entities = self._extract_entities(user_input)
            
            # Update conversation context
            self._update_context(best_intent, sentiment, sentiment_score, entities)
            
            # Dialogue management
            if self.dialogue_manager.should_clarify(intent_confidence, len(entities)):
                clarification = self.dialogue_manager.generate_clarification_question(
                    best_intent, user_input
                )
                self.context.awaiting_clarification = True
                self.context.clarification_context = user_input
                return clarification
            
            # Knowledge base search for relevant information
            kb_results = self.knowledge_base.search(user_input, top_k=2)
            
            # Generate response
            response = self._generate_contextual_response(
                user_input, best_intent, intent_confidence, entities, kb_results
            )
            
            # Personalize response
            response = self.response_generator.personalize_response(response, self.context)
            
            # Record interaction
            response_time = (datetime.now() - start_time).total_seconds()
            self._record_interaction(
                user_input, response, best_intent, intent_confidence,
                sentiment, sentiment_score, entities, response_time
            )
            
            # Cache response
            self._cache_response(user_input, response)
            
            # Update user expertise level based on questions
            self._update_expertise_level(user_input, best_intent)
            
            logger.info(f"Response generated in {response_time:.3f}s with intent {best_intent.value} (confidence: {intent_confidence:.2f})")
            
            return response
            
        except Exception as e:
            logger.error(f"Error in generate_response: {str(e)}", exc_info=True)
            return self._generate_error_response()
    
    def _validate_input(self, user_input: str) -> bool:
        """Validate user input"""
        if not user_input or not isinstance(user_input, str):
            return False
        if len(user_input.strip()) < 3:
            return False
        if len(user_input) > 500:
            return False
        return True
    
    def _extract_entities(self, text: str) -> Dict[str, List[str]]:
        """Extract entities from text"""
        entities = {}
        
        # Metrics
        metric_patterns = [
            'efficiency', 'utilization', 'dayrate', 'stability',
            'climate', 'location', 'performance', 'contract'
        ]
        found_metrics = [m for m in metric_patterns if m in text.lower()]
        if found_metrics:
            entities['metrics'] = found_metrics
        
        # Percentages
        percentages = re.findall(r'(\d+(?:\.\d+)?)\s*%', text)
        if percentages:
            entities['percentages'] = percentages
        
        # Locations
        locations = ['gulf', 'north sea', 'southeast asia', 'middle east', 'west africa']
        found_locations = [loc for loc in locations if loc in text.lower()]
        if found_locations:
            entities['locations'] = found_locations
        
        # Dates
        dates = re.findall(r'\d{4}-\d{2}-\d{2}', text)
        if dates:
            entities['dates'] = dates
        
        return entities
    
    def _update_context(self, intent: IntentType, sentiment: SentimentType,
                       sentiment_score: float, entities: Dict):
        """Update conversation context"""
        self.context.last_intent = intent
        self.context.sentiment = sentiment
        self.context.sentiment_score = sentiment_score
        self.context.last_entities = entities
        self.context.interaction_count += 1
        
        # Track topics
        if entities.get('metrics'):
            self.context.mentioned_topics.extend(entities['metrics'])
            self.context.mentioned_topics = list(set(self.context.mentioned_topics))[-10:]
    
    def _generate_contextual_response(self, user_input: str, intent: IntentType,
                                     confidence: float, entities: Dict,
                                     kb_results: List[Dict]) -> str:
        """Generate response based on intent and context"""
        
        # Route to appropriate handler
        if intent == IntentType.GREETING:
            return self._handle_greeting()
        elif intent == IntentType.FAREWELL:
            return self._handle_farewell()
        elif intent == IntentType.HELP_GENERAL:
            return self._handle_help_general()
        elif intent == IntentType.HELP_SPECIFIC:
            return self._handle_help_specific(user_input, entities)
        elif intent in [IntentType.METRICS_OVERVIEW, IntentType.METRICS_CALCULATION, IntentType.METRICS_IMPROVEMENT]:
            return self._handle_metrics(user_input, intent, kb_results)
        elif intent == IntentType.NAVIGATION:
            return self._handle_navigation(user_input, entities)
        elif intent in [IntentType.TROUBLESHOOTING, IntentType.TROUBLESHOOTING_UPLOAD]:
            return self._handle_troubleshooting(user_input, intent)
        elif intent == IntentType.UPLOAD:
            return self._handle_upload()
        elif intent in [IntentType.CLIMATE, IntentType.CLIMATE_RISK]:
            return self._handle_climate(user_input)
        elif intent == IntentType.COMPARISON:
            return self._handle_comparison()
        elif intent == IntentType.EXPORT:
            return self._handle_export()
        elif intent == IntentType.IMPROVEMENT_SUGGESTIONS:
            return self._handle_improvements(user_input, entities)
        elif intent == IntentType.COST_ANALYSIS:
            return self._handle_cost_analysis()
        elif intent == IntentType.ACKNOWLEDGMENT:
            return "You're welcome! Is there anything else you'd like to know?"
        else:
            return self._handle_unknown(user_input, kb_results)
    
    def _handle_greeting(self) -> str:
        """Handle greeting intent"""
        greeting_base = "Hello! I'm your Advanced Rig Efficiency AI Assistant. 🤖\n\n"
        
        if self.context.interaction_count == 0:
            return greeting_base + """I can help you with:
📊 Understanding efficiency metrics and calculations
🗺️ Navigation and platform features
📁 Data upload and file formatting
🔧 Troubleshooting and problem-solving
🌦️ Climate analysis and risk assessment
💡 Performance improvement strategies
📈 Multi-rig comparison and benchmarking
💰 Cost optimization and analysis

What would you like to explore?"""
        else:
            return greeting_base + "Welcome back! What can I help you with today?"
    
    def _handle_farewell(self) -> str:
        """Handle farewell intent"""
        return self.response_generator.generate("farewell")
    
    def _handle_help_general(self) -> str:
        """Handle general help request"""
        return """🤖 RIG EFFICIENCY ANALYSIS TOOL - COMPLETE GUIDE

WHAT THIS TOOL DOES:
This platform analyzes offshore rig performance, calculates efficiency scores, and provides actionable insights to optimize operations and profitability.

📊 KEY FEATURES:

1. EFFICIENCY SCORING
   - Analyzes 6 key performance factors
   - Weighted scoring algorithm (0-100%)
   - Letter grades (A+ to F)
   - Industry benchmarking

2. DATA UPLOAD & ANALYSIS
   - Upload rig contract data (Excel/CSV)
   - Automatic validation & processing
   - Real-time performance calculations
   - Historical trend analysis

3. CLIMATE RISK ASSESSMENT
   - Weather impact analysis
   - Seasonal downtime prediction
   - Geographic risk mapping
   - Hurricane/storm impact

4. MULTI-RIG COMPARISON
   - Compare multiple rigs side-by-side
   - Identify best/worst performers
   - Benchmark against fleet average
   - Spot improvement opportunities

5. IMPROVEMENT RECOMMENDATIONS
   - AI-powered suggestions
   - Quick wins (0-3 months)
   - Medium-term strategies (3-12 months)
   - Long-term initiatives (12+ months)

6. REPORTING & EXPORT
   - PDF reports
   - Excel exports
   - PowerPoint presentations
   - CSV data downloads

🚀 HOW TO GET STARTED:

Step 1: Upload Your Data
   → Click "Upload Data" in sidebar
   → Select Excel/CSV file with rig contracts
   → Wait for validation

Step 2: Review Analysis
   → View efficiency score
   → Check factor breakdown
   → Read AI insights

Step 3: Explore Features
   → Try climate analysis
   → Compare multiple rigs
   → Download reports

Step 4: Take Action
   → Review improvement suggestions
   → Implement recommendations
   → Track progress over time

💡 NEED SPECIFIC HELP?
Ask me about:
- "How is efficiency calculated?"
- "How to upload data?"
- "What file format do I need?"
- "How to improve my score?"
- "Climate risk analysis"
- "Export options"

I'm here to guide you every step of the way! 🎯"""
    
    def _handle_help_specific(self, user_input: str, entities: Dict) -> str:
        """Handle specific help request"""
        # Search knowledge base for specific topic
        results = self.knowledge_base.search(user_input, top_k=1)
        if results and results[0]["relevance_score"] > 0.3:
            return results[0]["content"]
        return self._handle_help_general()
    
    def _handle_metrics(self, user_input: str, intent: IntentType, kb_results: List[Dict]) -> str:
        """Handle metrics-related queries"""
        if intent == IntentType.METRICS_CALCULATION:
            entry = self.knowledge_base.get_by_id("efficiency_calculation")
            return entry["content"] if entry else self._handle_unknown(user_input, kb_results)
        elif intent == IntentType.METRICS_IMPROVEMENT:
            entry = self.knowledge_base.get_by_id("improvement_strategies")
            return entry["content"] if entry else self._handle_unknown(user_input, kb_results)
        else:
            # Check if asking about specific metric
            if 'utilization' in user_input.lower() or 'contract' in user_input.lower():
                entry = self.knowledge_base.get_by_id("contract_utilization")
                if entry:
                    return entry["content"]
            elif 'dayrate' in user_input.lower() or 'cost' in user_input.lower():
                entry = self.knowledge_base.get_by_id("dayrate_efficiency")
                if entry:
                    return entry["content"]
            
            # Default to overview
            entry = self.knowledge_base.get_by_id("efficiency_overview")
            return entry["content"] if entry else self._handle_unknown(user_input, kb_results)
    
    def _handle_navigation(self, user_input: str, entities: Dict) -> str:
        """Handle navigation queries"""
        entry = self.knowledge_base.get_by_id("navigation_guide")
        return entry["content"] if entry else "Navigation help is available in the sidebar menu."
    
    def _handle_troubleshooting(self, user_input: str, intent: IntentType) -> str:
        """Handle troubleshooting queries"""
        if intent == IntentType.TROUBLESHOOTING_UPLOAD or 'upload' in user_input.lower():
            upload_entry = self.knowledge_base.get_by_id("upload_guide")
            troubleshoot_entry = self.knowledge_base.get_by_id("troubleshooting_general")
            
            return f"""{upload_entry['content'] if upload_entry else ''}

TROUBLESHOOTING UPLOAD ISSUES:
{troubleshoot_entry['content'] if troubleshoot_entry else ''}"""
        
        entry = self.knowledge_base.get_by_id("troubleshooting_general")
        return entry["content"] if entry else "Please describe the specific issue you're experiencing."
    
    def _handle_upload(self) -> str:
        """Handle upload queries"""
        entry = self.knowledge_base.get_by_id("upload_guide")
        return entry["content"] if entry else "Upload data via the sidebar Upload Data section."
    
    def _handle_climate(self, user_input: str) -> str:
        """Handle climate queries"""
        entry = self.knowledge_base.get_by_id("climate_analysis")
        return entry["content"] if entry else "Climate analysis assesses weather-related operational risks."
    
    def _handle_comparison(self) -> str:
        """Handle comparison queries"""
        entry = self.knowledge_base.get_by_id("comparison_guide")
        return entry["content"] if entry else "Use Multi-Rig Comparison to analyze multiple rigs simultaneously."
    
    def _handle_export(self) -> str:
        """Handle export queries"""
        entry = self.knowledge_base.get_by_id("export_guide")
        return entry["content"] if entry else "Export options are available at the bottom of analysis results."
    
    def _handle_improvements(self, user_input: str, entities: Dict) -> str:
        """Handle improvement suggestion queries"""
        entry = self.knowledge_base.get_by_id("improvement_strategies")
        
        # Customize based on entities
        if entities.get('metrics'):
            specific_metric = entities['metrics'][0]
            additional = f"\n\n🎯 You mentioned {specific_metric}. Focus on optimizing this metric for quick wins."
            return (entry["content"] if entry else "") + additional
        
        return entry["content"] if entry else "I can suggest improvements once you complete an analysis."
    
    def _handle_cost_analysis(self) -> str:
        """Handle cost analysis queries"""
        entry = self.knowledge_base.get_by_id("cost_analysis")
        return entry["content"] if entry else "Cost analysis helps identify optimization opportunities."
    
    def _handle_unknown(self, user_input: str, kb_results: List[Dict]) -> str:
        """Handle unknown intent"""
        # Try semantic search on knowledge base
        if kb_results and kb_results[0]["relevance_score"] > 0.25:
            return f"""Based on your question, here's what might help:

{kb_results[0]["content"]}

Was this helpful? Feel free to ask for clarification!"""
        
        # Fallback with suggestions
        return """I'm not entirely sure what you're asking about, but I can help with:

🔹 Efficiency Metrics - "What do scores mean?" or "How is it calculated?"
🔹 Improvements - "How to improve performance?" or "Optimization tips?"
🔹 Navigation - "Where is [feature]?" or "How to navigate?"
🔹 Upload - "How to upload data?" or "File requirements?"
🔹 Climate - "Weather risk analysis" or "Climate impact?"
🔹 Comparison - "Compare rigs" or "Benchmarking?"
🔹 Export - "Download reports" or "Export options?"
🔹 Troubleshooting - "Upload not working" or "Error messages?"

What would you like to know about?"""
    
    def _generate_error_response(self) -> str:
        """Generate user-friendly error response"""
        return """I apologize, but I encountered an issue processing your request. 

Please try:
1. Rephrasing your question
2. Being more specific
3. Asking about a particular topic
4. Typing "help" for available topics

I'm here to help! 🤖"""
    
    def _record_interaction(self, user_input: str, response: str, intent: IntentType,
                          intent_confidence: float, sentiment: SentimentType,
                          sentiment_score: float, entities: Dict, response_time: float):
        """Record interaction in chat history"""
        message = ChatMessage(
            role="user",
            content=user_input,
            timestamp=datetime.now(),
            intent=intent,
            intent_confidence=intent_confidence,
            sentiment=sentiment,
            sentiment_score=sentiment_score,
            entities=entities,
            response_time=response_time,
            metadata={
                "context_state": {
                    "expertise_level": self.context.expertise_level,
                    "topics": self.context.mentioned_topics[-5:],
                    "interaction_count": self.context.interaction_count
                }
            }
        )
        
        self.chat_history.append(message)
        
        # Manage history size
        if len(self.chat_history) > self.max_history:
            self.chat_history = self.chat_history[-self.max_history:]
        
        # Track success metrics
        if intent_confidence > 0.7:
            self.context.successful_resolutions += 1
        elif intent_confidence < 0.3:
            self.context.failed_resolutions += 1
        
        # Update average confidence
        total_interactions = self.context.successful_resolutions + self.context.failed_resolutions
        if total_interactions > 0:
            self.context.average_confidence = (
                self.context.successful_resolutions / total_interactions
            )
    
    def _update_expertise_level(self, user_input: str, intent: IntentType):
        """Update user expertise level based on interaction patterns"""
        # Expert indicators
        expert_keywords = ['algorithm', 'formula', 'calculation', 'optimize', 'benchmark', 'api']
        beginner_keywords = ['what is', 'explain', 'how do', 'help', 'don\'t understand']
        
        expert_score = sum(1 for kw in expert_keywords if kw in user_input.lower())
        beginner_score = sum(1 for kw in beginner_keywords if kw in user_input.lower())
        
        # Update expertise level
        if self.context.interaction_count > 5:
            if expert_score > beginner_score * 2:
                self.context.expertise_level = "expert"
            elif beginner_score > expert_score * 2:
                self.context.expertise_level = "beginner"
            else:
                self.context.expertise_level = "intermediate"
    
    def get_conversation_summary(self) -> Dict[str, Any]:
        """Get summary of conversation"""
        return {
            "total_interactions": self.context.interaction_count,
            "successful_resolutions": self.context.successful_resolutions,
            "failed_resolutions": self.context.failed_resolutions,
            "average_confidence": self.context.average_confidence,
            "expertise_level": self.context.expertise_level,
            "main_topics": self.context.mentioned_topics[-10:],
            "current_sentiment": self.context.sentiment.value,
            "chat_history_size": len(self.chat_history)
        }
    
    def reset_conversation(self):
        """Reset conversation state"""
        self.context = ConversationContext()
        self.chat_history = []
        self.dialogue_manager = DialogueManager()
        logger.info("Conversation reset")


# ============================================================================
# SECTION 6: STREAMLIT INTEGRATION (Optional - for testing)
# ============================================================================

def main():
    """Main function for Streamlit app integration"""
    st.set_page_config(
        page_title="Advanced Rig Efficiency AI",
        page_icon="🤖",
        layout="wide"
    )
    
    st.title("🤖 Advanced Rig Efficiency AI Assistant")
    st.markdown("*Next-generation conversational AI with ML capabilities*")
    
    # Initialize chatbot in session state
    if 'chatbot' not in st.session_state:
        st.session_state.chatbot = AdvancedRigEfficiencyAIChatbot()
        st.session_state.messages = []
    
    # Display chat history
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])
    
    # Chat input
    if prompt := st.chat_input("Ask me anything about rig efficiency..."):
        # Display user message
        with st.chat_message("user"):
            st.markdown(prompt)
        st.session_state.messages.append({"role": "user", "content": prompt})
        
        # Generate response
        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                response = st.session_state.chatbot.generate_response(prompt)
                st.markdown(response)
        st.session_state.messages.append({"role": "assistant", "content": response})
    
    # Sidebar with stats
    with st.sidebar:
        st.header("📊 Conversation Stats")
        summary = st.session_state.chatbot.get_conversation_summary()
        
        st.metric("Total Interactions", summary["total_interactions"])
        st.metric("Success Rate", f"{summary['average_confidence']*100:.1f}%")
        st.metric("Expertise Level", summary["expertise_level"].title())
        
        if st.button("🔄 Reset Conversation"):
            st.session_state.chatbot.reset_conversation()
            st.session_state.messages = []
            st.rerun()
        
        st.divider()
        st.caption("Recent Topics")
        for topic in summary["main_topics"][-5:]:
            st.caption(f"• {topic}")


if __name__ == "__main__":
    # For standalone testing
    print("Advanced Rig Efficiency AI Chatbot - Testing Mode")
    print("=" * 60)
    
    chatbot = AdvancedRigEfficiencyAIChatbot()
    
    test_queries = [
        "Hello!",
        "What is efficiency score?",
        "How is efficiency calculated?",
        "How to improve my score?",
        "What's climate risk?",
        "Upload file help",
        "Compare multiple rigs",
        "Cost optimization tips"
    ]
    
    for query in test_queries:
        print(f"\n{'='*60}")
        print(f"USER: {query}")
        print(f"{'='*60}")
        response = chatbot.generate_response(query)
        print(f"AI: {response[:300]}...")
        print()
    
    # Show summary
    print("\n" + "="*60)
    print("CONVERSATION SUMMARY")
    print("="*60)
    summary = chatbot.get_conversation_summary()
    for key, value in summary.items():
        print(f"{key}: {value}")


# ============================================================================
# BACKWARDS COMPATIBILITY ALIAS
# ============================================================================

# Alias for backwards compatibility with app.py
RigEfficiencyAIChatbot = AdvancedRigEfficiencyAIChatbot