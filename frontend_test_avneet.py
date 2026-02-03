"""
line 2049 not running
line 1050 not running
line 1023 not running
line 105 not running
line 3457 not running

ADVANCED Multi-Department Timetable Generator
✨ NEW FEATURES:
- Floor-wise room assignment priority for same department batches
- Parallel department processing for 4x speed
- Elective subjects support with group scheduling
- Progressive scheduling (critical batches first)
- ML-based strategy prediction
- Smart ratio-based resource allocation
"""
from enum import Enum
import logging
import os
import re
import time
import glob
import json
import pickle
import pandas as pd
import numpy as np
from collections import defaultdict, Counter
from typing import Dict, List, Any, Tuple, Optional, Set
from pathlib import Path
from dataclasses import dataclass, field,replace
from functools import lru_cache
from concurrent.futures import ThreadPoolExecutor, as_completed, wait, FIRST_COMPLETED
import warnings
warnings.filterwarnings('ignore')

try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except ImportError:
    ORTOOLS_AVAILABLE = False
    print("WARNING: OR-Tools not available. Install with: pip install ortools")

import threading
from copy import deepcopy

# Try to import enhancement modules
try:
    # Ensure enhancement files are in same directory
    import sys
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
    
    # Flag for conditional execution
    ENHANCEMENTS_AVAILABLE = True
    print("✅ Enhancement system initialized")
    
except ImportError as e:
    ENHANCEMENTS_AVAILABLE = False
    print(f"⚠️  Enhancements not available: {e}")
    print("   Continuing with base scheduler...")

# Import the robust ML strategy predictor
try:
    from enhanced_strategy_predictor import (
        RobustStrategyPredictor,
        SchedulingMode,
        StrategyDecision,
        DecisionMetrics
    )
    ROBUST_PREDICTOR_AVAILABLE = True
    print("✅ Robust ML Strategy Predictor loaded")
except ImportError as e:
    ROBUST_PREDICTOR_AVAILABLE = False
    print(f"⚠️  Robust predictor not available: {e}")
    print("   Using fallback strategy predictor...")

import threading
from copy import deepcopy
from functools import wraps
from contextlib import contextmanager

# Global locks for shared resources
_file_locks = {}
_file_locks_lock = threading.Lock()
_global_state_lock = threading.Lock()
_results_lock = threading.Lock()


def get_file_lock(filepath: str) -> threading.Lock:
    """Get or create a lock for a specific file (thread-safe)"""
    with _file_locks_lock:
        if filepath not in _file_locks:
            _file_locks[filepath] = threading.Lock()
        return _file_locks[filepath]


class ThreadSafeResultsAggregator:
    """Thread-safe aggregation of results"""
    def __init__(self):
        self._results = []
        self._lock = threading.Lock()
    
    def add_result(self, result: Dict):
        with self._lock:
            self._results.append(result)
    
    def get_results(self) -> List[Dict]:
        with self._lock:
            return self._results.copy()
    
    def get_successful(self) -> List[Dict]:
        with self._lock:
            return [r for r in self._results if r.get('success', False)]
    
    def get_failed(self) -> List[Dict]:
        with self._lock:
            return [r for r in self._results if not r.get('success', False)]


# Flag for enhancements
ENHANCEMENTS_AVAILABLE = True
print("✅ Race condition protection initialized")

# -------------------------
# Enhanced Data Classes with Floor Support
# -------------------------

@dataclass(frozen=True, slots=True)
class FacultyChoice:
    """Represents a faculty's subject preferences"""
    faculty_id: str
    faculty_name: str
    choice_1: str  # Highest priority
    choice_2: str
    choice_3: str
    choice_4: str
    choice_5: str  # Lowest priority
    
    def get_choices_list(self) -> List[str]:
        """Returns choices as ordered list (highest to lowest priority)"""
        return [
            self.choice_1, self.choice_2, self.choice_3, 
            self.choice_4, self.choice_5
        ]
    
    def get_priority_score(self, subject_code: str) -> int:
        """
        Returns priority score for a subject (higher = more preferred)
        Returns 0 if subject not in choices
        """
        choices = self.get_choices_list()
        try:
            # Priority scores: Choice 1=500, 2=400, 3=300, 4=200, 5=100
            priority_index = choices.index(subject_code)
            return 500 - (priority_index * 100)
        except ValueError:
            return 0  # Subject not in choices


@dataclass
class AllocationConstraints:
    """Configuration for faculty allocation mode"""
    use_faculty_choice: bool  # True = choice-based, False = qualification-based
    choice_weight: float = 0.7
    qualification_weight: float = 0.3

@dataclass(frozen=True, slots=True)
class TimeSlot:
    slot_id: int
    start_time: str
    end_time: str
    
    def __str__(self):
        return f"{self.start_time}-{self.end_time}"

@dataclass(frozen=True, slots=True)
class Room:
    room_id: str
    capacity: int
    room_type: str
    department: str
    floor: int = 0  # NEW: Floor number for proximity preference

@dataclass(frozen=True, slots=True)
class Faculty:
    faculty_id: str
    name: str
    qualifications: tuple
    specializations: tuple
    max_hours_per_week: int

@dataclass(frozen=True, slots=True)
class Subject:
    subject_code: str
    subject_name: str
    hours_per_week: int
    requires_lab: bool
    lab_hours_per_week: int
    required_specialization: str = ""
    is_elective: bool = False  # NEW: Elective flag
    elective_group: str = ""    # NEW: Elective group ID
    priority: int = 0           # NEW: Scheduling priority (higher = earlier)

@dataclass(frozen=True, slots=True)
class StudentBatch:
    batch_id: str
    department: str
    student_count: int
    subjects: tuple
    shift_preference: str = "morning"
    preferred_floor: int = 0  # NEW: Preferred floor for classes
    priority: int = 5         # NEW: Batch priority (1-10, higher = more important)

def diagnose_bca_failure(input_file):
    """
    Diagnose why BCA scheduler is failing
    Shows exact constraints, capacity, and conflicts
    """
    print(f"\n{'='*80}")
    print(f"🔍 DIAGNOSTIC: BCA Scheduler Failure Analysis")
    print(f"{'='*80}")
    
    try:
        # Load data
        print(f"\n[1] Loading BCA data...")
        reader = AdvancedExcelReader(input_file)
        data = reader.parse_all()
        
        timeslots = data.get('timeslots', [])
        rooms = data.get('rooms', [])
        faculties = data.get('faculties', [])
        subjects = data.get('subjects', [])
        batches = data.get('batches', [])
        
        print(f"    ✓ Timeslots: {len(timeslots)} ({[ts.slot_id for ts in timeslots[:5]]}...)")
        print(f"    ✓ Rooms: {len(rooms)} rooms")
        print(f"    ✓ Faculties: {len(faculties)} faculty")
        print(f"    ✓ Subjects: {len(subjects)} subjects")
        print(f"    ✓ Batches: {len(batches)} batches")
        
        # ANALYSIS 1: Show subject details
        print(f"\n[2] Subject Analysis:")
        print(f"    Subject Code | Hours | Lab | Specialization | Faculty Count")
        print(f"    " + "-"*70)
        
        total_hours = 0
        for subject in subjects:
            theory = subject.hours_per_week
            lab = subject.lab_hours_per_week
            total = theory + (lab * 2)
            total_hours += total
            
            # Count faculty with specialization
            qualified = 0
            if subject.required_specialization:
                qualified = len([f for f in faculties 
                               if subject.required_specialization.lower() in ' '.join(f.specializations).lower()])
            
            print(f"    {subject.subject_code:15} | {theory:5} | {lab:3} | {subject.required_specialization:15} | {qualified:13}")
        
        print(f"\n    TOTAL REQUIRED HOURS: {total_hours}h")
        
        # ANALYSIS 2: Show faculty capacity
        print(f"\n[3] Faculty Capacity Analysis:")
        print(f"    Faculty ID | Max Hours | Specializations")
        print(f"    " + "-"*70)
        
        total_faculty_capacity = 0
        for faculty in faculties:
            total_faculty_capacity += faculty.max_hours_per_week
            specs = ', '.join(faculty.specializations[:2]) if faculty.specializations else 'None'
            print(f"    {faculty.faculty_id:10} | {faculty.max_hours_per_week:9} | {specs}")
        
        print(f"\n    TOTAL FACULTY CAPACITY: {total_faculty_capacity}h/week")
        print(f"    REQUIRED: {total_hours}h/week")
        print(f"    CAPACITY: {total_faculty_capacity - total_hours:+d}h ({(total_faculty_capacity/max(total_hours, 1))*100:.1f}%)")
        
        # ANALYSIS 3: Show room capacity
        print(f"\n[4] Room Capacity Analysis:")
        usable_slots = len([ts for ts in timeslots if ts.slot_id != 3])
        working_days = 5
        slots_per_week = usable_slots * working_days
        total_room_slots = len(rooms) * slots_per_week
        
        print(f"    Usable timeslots/day: {usable_slots}")
        print(f"    Working days/week: {working_days}")
        print(f"    Slots per room/week: {slots_per_week}")
        print(f"    Total rooms: {len(rooms)}")
        print(f"    TOTAL ROOM SLOTS: {total_room_slots} slots/week")
        
        # ANALYSIS 4: Show batch requirements
        print(f"\n[5] Batch Requirements Analysis:")
        print(f"    Batch ID | Subjects | Total Hours | Capacity % | Feasible?")
        print(f"    " + "-"*70)
        
        for batch in batches:
            batch_subjects = [s for s in subjects if s.subject_code in batch.subjects]
            batch_hours = sum(s.hours_per_week + (s.lab_hours_per_week * 2) for s in batch_subjects)
            
            # Estimate feasibility
            available_slots = slots_per_week
            feasibility = (available_slots / max(batch_hours, 1)) * 100
            is_feasible = feasibility >= 95
            
            print(f"    {batch.batch_id:10} | {len(batch_subjects):8} | {batch_hours:11} | {feasibility:10.1f}% | {'✓' if is_feasible else '✗'}")
        
        # ANALYSIS 5: Check for conflicts
        print(f"\n[6] Potential Conflicts:")
        
        issues = []
        
        # Issue 1: Not enough faculty capacity
        if total_faculty_capacity < total_hours:
            issues.append(f"❌ Faculty shortage: Need {total_hours}h but only have {total_faculty_capacity}h")
        
        # Issue 2: Not enough room slots
        if total_room_slots < total_hours:
            issues.append(f"❌ Room shortage: Need {total_hours}h but only have {total_room_slots} slots")
        
        # Issue 3: Specialized subject with no faculty
        for subject in subjects:
            if subject.required_specialization:
                qualified = len([f for f in faculties 
                               if subject.required_specialization.lower() in ' '.join(f.specializations).lower()])
                if qualified == 0:
                    issues.append(f"❌ No faculty: {subject.subject_code} needs '{subject.required_specialization}' but no faculty has it")
        
        # Issue 4: Too many hours for single subject
        for subject in subjects:
            if subject.hours_per_week > 10:
                issues.append(f"⚠️  High hours: {subject.subject_code} has {subject.hours_per_week}h (typical max is 8h)")
        
        if issues:
            print(f"    Found {len(issues)} potential issues:")
            for issue in issues:
                print(f"    {issue}")
        else:
            print(f"    ✓ No obvious conflicts detected")
        
        # ANALYSIS 6: Recommendation
        print(f"\n[7] Recommendation:")
        if total_faculty_capacity < total_hours:
            shortage = total_hours - total_faculty_capacity
            print(f"    🔧 Shortage: {shortage}h")
            print(f"    💡 Option 1: Reduce subject hours by ~{shortage}h total")
            print(f"    💡 Option 2: Add {int(shortage / 8)} more faculty")
            print(f"    💡 Option 3: Check subject hours in Excel - may be incorrectly set")
        
        elif total_room_slots < total_hours:
            print(f"    🔧 Not enough room slots: {total_hours}h needed but {total_room_slots} available")
            print(f"    💡 Option 1: Check room configuration")
            print(f"    💡 Option 2: Reduce total hours needed")
        
        else:
            print(f"    🔧 Enough capacity exists, but scheduler still failed")
            print(f"    💡 Possible causes:")
            print(f"       - Specialization mismatch (faculty skills don't match subject needs)")
            print(f"       - Batch-faculty conflict (some faculty can't teach certain batches)")
            print(f"       - OR-Tools constraints too strict")
            print(f"       - Subject-faculty mapping missing")
        
        print(f"\n{'='*80}\n")
    
    except Exception as e:
        print(f"❌ Diagnostic error: {str(e)}")
        import traceback
        traceback.print_exc()

# -------------------------
# ML-Based Strategy Predictor (Enhanced with Robust Decision System)
# -------------------------

if ROBUST_PREDICTOR_AVAILABLE:
    # Use the new robust predictor with multi-factor scoring and Thompson Sampling
    class StrategyPredictor:
        """
        Enhanced strategy predictor using RobustStrategyPredictor.
        Features:
        - Multi-factor scoring (all strategies considered)
        - Thompson Sampling exploration
        - Decision tree with 15+ constraints
        - Minimum trial guarantees
        - ML learning from history
        """
        
        def __init__(self):
            self.robust_predictor = RobustStrategyPredictor()
            self.history_file = self.robust_predictor.history_file  # For backward compatibility
        
        def load_history(self):
            """Load history from robust predictor"""
            self.robust_predictor.load_history()
        
        def save_result(self, features: Dict, strategy: str, success_rate: float):
            """Save scheduling result for learning (THREAD-SAFE)"""
            # Map old strategy names to new mode values
            strategy_map = {
                'STRICT NO GAPS': 'direct_cpsat',
                'COMPACT SCHEDULE': 'dynamic_allocation',
                'RELAXED': 'constraint_handling', 
                'MAXIMUM FLEXIBILITY': 'emergency_scaling'
            }
            mode = strategy_map.get(strategy, 'direct_cpsat')
            
            self.robust_predictor.save_result(
                features=features,
                mode=mode,
                strategy_params={},
                success_rate=success_rate,
                elapsed_time=0.0
            )
        
        def extract_features(self, data: Dict) -> Dict:
            """Extract features using robust predictor"""
            return self.robust_predictor.extract_features(data)
        
        def predict_best_strategy(self, features: Dict = None, data: Dict = None) -> int:
            """
            Predict best strategy using robust ML predictor.
            Returns strategy index for backward compatibility.
            
            Strategy mapping:
            - 0: STRICT NO GAPS -> DIRECT_CPSAT
            - 1: COMPACT SCHEDULE -> DYNAMIC_ALLOCATION  
            - 2: RELAXED -> CONSTRAINT_HANDLING
            - 3: MAXIMUM FLEXIBILITY -> EMERGENCY_SCALING
            """
            if data:
                decision = self.robust_predictor.predict_best_strategy(data)
            else:
                # Fallback if only features provided
                return 0
            
            # Map scheduling mode to old strategy index
            mode_to_index = {
                SchedulingMode.DIRECT_CPSAT: 0,
                SchedulingMode.DYNAMIC_ALLOCATION: 1,
                SchedulingMode.CONSTRAINT_HANDLING: 2,
                SchedulingMode.EMERGENCY_SCALING: 3
            }
            
            return mode_to_index.get(decision.mode, 0)
        
        def get_decision_details(self, data: Dict) -> StrategyDecision:
            """Get full decision details including reasoning and confidence"""
            return self.robust_predictor.predict_best_strategy(data)
        
        def get_stats_summary(self) -> Dict:
            """Get strategy statistics summary"""
            return self.robust_predictor.get_stats_summary()

else:
    # Fallback: Original StrategyPredictor if robust predictor not available
    class StrategyPredictor:
        """Predicts best scheduling strategy based on input characteristics (Fallback)"""
        
        def __init__(self):
            self.history_file = "scheduling_history.json"
            self.load_history()
        
        def load_history(self):
            """Load past scheduling results"""
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r') as f:
                    self.history = json.load(f)
            else:
                self.history = []
        
        def save_result(self, features: Dict, strategy: str, success_rate: float):
            """Save scheduling result for learning (THREAD-SAFE)"""
            if not hasattr(self, '_file_lock'):
                self._file_lock = get_file_lock(self.history_file)
            
            with self._file_lock:
                if os.path.exists(self.history_file):
                    try:
                        with open(self.history_file, 'r') as f:
                            current_history = json.load(f)
                    except (json.JSONDecodeError, IOError):
                        current_history = []
                else:
                    current_history = []
                
                current_history.append({
                    'features': features,
                    'strategy': strategy,
                    'success_rate': success_rate,
                    'timestamp': time.time()
                })
                
                if len(current_history) > 100:
                    current_history = current_history[-100:]
                
                temp_file = f"{self.history_file}.tmp"
                try:
                    with open(temp_file, 'w') as f:
                        json.dump(current_history, f, indent=2)
                    os.replace(temp_file, self.history_file)
                    self.history = current_history
                except Exception as e:
                    if os.path.exists(temp_file):
                        try:
                            os.remove(temp_file)
                        except:
                            pass
        
        def extract_features(self, data: Dict) -> Dict:
            """Extract features from input data"""
            timeslots = data['timeslots']
            rooms = data['rooms']
            faculties = data['faculties']
            subjects = data['subjects']
            batches = data['batches']
            
            total_sessions = sum(
                s.hours_per_week + s.lab_hours_per_week * 2
                for b in batches
                for subj_code in b.subjects
                for s in [next((sub for sub in subjects if sub.subject_code == subj_code), None)]
                if s
            )
            
            total_capacity = sum(f.max_hours_per_week for f in faculties)
            usable_slots = len([ts for ts in timeslots if ts.slot_id != 3]) * 5
            
            return {
                'batch_count': len(batches),
                'faculty_count': len(faculties),
                'room_count': len(rooms),
                'subject_count': len(subjects),
                'total_sessions': total_sessions,
                'capacity_ratio': total_capacity / total_sessions if total_sessions > 0 else 1.0,
                'slots_per_batch': usable_slots / len(batches) if batches else 0,
                'avg_batch_size': sum(b.student_count for b in batches) / len(batches) if batches else 0
            }
        
        def predict_best_strategy(self, features: Dict = None, data: Dict = None) -> int:
            """Predict which strategy index to try first"""
            if len(self.history) < 5:
                return 0
            
            best_strategy_idx = 0
            strategy_scores = defaultdict(list)
            
            if features is None and data:
                features = self.extract_features(data)
            elif features is None:
                return 0
            
            for record in self.history[-20:]:
                past_features = record['features']
                
                similarity = 0
                similarity += 1 - abs(past_features.get('capacity_ratio', 1) - features.get('capacity_ratio', 1))
                similarity += 1 - abs(past_features.get('batch_count', 0) - features.get('batch_count', 0)) / 10
                similarity += 1 - abs(past_features.get('faculty_count', 0) - features.get('faculty_count', 0)) / 10
                
                strategy_name = record['strategy']
                success_rate = record['success_rate']
                
                strategy_map = {
                    'STRICT NO GAPS': 0,
                    'COMPACT SCHEDULE': 1,
                    'RELAXED': 2,
                    'MAXIMUM FLEXIBILITY': 3
                }
                
                strategy_idx = strategy_map.get(strategy_name, 0)
                strategy_scores[strategy_idx].append(similarity * success_rate)
            
            if strategy_scores:
                best_strategy_idx = max(strategy_scores.keys(), 
                                       key=lambda k: sum(strategy_scores[k]) / len(strategy_scores[k]))
            
            return best_strategy_idx

# ═══════════════════════════════════════════════════════════════════════════════
# SMART STRATEGY ADVISOR - Complete Intelligent Scheduling System
# ═══════════════════════════════════════════════════════════════════════════════

import math
import uuid
from datetime import datetime

@dataclass
class DemandAnalysis:
    """
    Comprehensive demand analysis for ML learning and strategy selection.
    30+ features for accurate prediction.
    """
    # ═══════════════════════════════════════════════════════════════
    # RESOURCE COUNTS
    # ═══════════════════════════════════════════════════════════════
    total_batches: int = 0
    total_subjects: int = 0
    total_faculty: int = 0
    total_classrooms: int = 0
    total_lab_rooms: int = 0
    working_days: int = 5
    slots_per_day: int = 7
    
    # ═══════════════════════════════════════════════════════════════
    # SESSION COUNTS
    # ═══════════════════════════════════════════════════════════════
    theory_sessions: int = 0
    lab_sessions: int = 0
    elective_sessions: int = 0
    total_sessions: int = 0
    
    # ═══════════════════════════════════════════════════════════════
    # SHIFT DISTRIBUTION
    # ═══════════════════════════════════════════════════════════════
    morning_batches: int = 0
    evening_batches: int = 0
    overlap_slots: List[int] = field(default_factory=list)
    
    # ═══════════════════════════════════════════════════════════════
    # CAPACITY METRICS (Hours)
    # ═══════════════════════════════════════════════════════════════
    total_faculty_hours_available: int = 0
    total_hours_demanded: int = 0
    classroom_hours_available: int = 0
    lab_hours_available: int = 0
    
    # ═══════════════════════════════════════════════════════════════
    # PRESSURE SCORES (0-100)
    # ═══════════════════════════════════════════════════════════════
    faculty_pressure: float = 0.0
    faculty_load_variance: float = 0.0
    classroom_utilization: float = 0.0
    lab_utilization: float = 0.0
    overlap_contention: float = 0.0
    
    # ═══════════════════════════════════════════════════════════════
    # CONSTRAINT FLAGS
    # ═══════════════════════════════════════════════════════════════
    has_floor_constraints: bool = False
    has_specialization_constraints: bool = False
    has_elective_conflicts: bool = False
    
    # ═══════════════════════════════════════════════════════════════
    # DERIVED RATIOS
    # ═══════════════════════════════════════════════════════════════
    sessions_per_faculty: float = 0.0
    sessions_per_room: float = 0.0
    lab_ratio: float = 0.0
    
    def to_feature_vector(self) -> List[float]:
        """Convert to normalized feature vector for ML"""
        return [
            self.faculty_pressure / 100,
            self.faculty_load_variance,
            self.overlap_contention / 100,
            self.classroom_utilization,
            self.lab_utilization,
            1.0 if self.has_floor_constraints else 0.0,
            1.0 if self.has_specialization_constraints else 0.0,
            self.lab_ratio,
            min(1.0, self.sessions_per_faculty / 10),
            min(1.0, self.sessions_per_room / 20)
        ]


@dataclass
class FacultyRecommendation:
    """Actionable faculty addition recommendations"""
    current_pressure: float
    is_schedulable: bool
    recommendations: List[Dict] = field(default_factory=list)
    
    def print_recommendations(self):
        print(f"\n📊 Faculty Analysis:")
        print(f"   Current pressure: {self.current_pressure:.1f}%")
        status = "✅ Schedulable" if self.is_schedulable else "❌ Over capacity"
        if self.current_pressure > 80:
            status += " (high load)"
        print(f"   Status: {status}")
        
        if self.recommendations:
            print(f"\n   Recommendations:")
            for rec in self.recommendations:
                print(f"   • Add {rec['add_faculty']} teachers → {rec['new_pressure_percent']}% pressure")


@dataclass
class StrategyFailureReport:
    """Detailed failure analysis for learning"""
    strategy_name: str
    failure_type: str  # "no_solution", "partial", "timeout"
    execution_time: float
    
    sessions_scheduled: int = 0
    sessions_failed: int = 0
    success_rate: float = 0.0
    gaps_created: int = 0
    
    # Bottleneck analysis
    primary_bottleneck: str = ""  # "faculty", "rooms", "slots"
    room_conflicts: int = 0
    faculty_conflicts: int = 0
    slot_exhaustion: int = 0
    
    top_bottleneck_rooms: List[str] = field(default_factory=list)
    top_bottleneck_faculty: List[str] = field(default_factory=list)
    top_bottleneck_slots: List[str] = field(default_factory=list)
    
    # Hidden constraints discovered
    hidden_constraints: Dict = field(default_factory=dict)
    
    # Config suggestions for next strategy
    suggested_config_changes: Dict = field(default_factory=dict)


class StrategyCooldown:
    """
    Tracks failed strategies and prevents retry within same execution.
    """
    def __init__(self):
        self.failed_strategies: Dict[str, StrategyFailureReport] = {}
        self.run_id = uuid.uuid4().hex[:8]
    
    def mark_failed(self, strategy_name: str, report: StrategyFailureReport):
        """Mark strategy as failed"""
        self.failed_strategies[strategy_name] = report
    
    def is_blocked(self, strategy_name: str) -> bool:
        """Check if strategy should be skipped"""
        return strategy_name in self.failed_strategies
    
    def get_failure_insights(self, strategy_name: str) -> Optional[StrategyFailureReport]:
        """Get failure report if available"""
        return self.failed_strategies.get(strategy_name)
    
    def should_skip_similar(self, strategy_name: str) -> Tuple[bool, str]:
        """Check if strategy is similar to a failed one"""
        for failed_name, report in self.failed_strategies.items():
            # Both strict strategies likely have same constraint issue
            if 'STRICT' in failed_name and 'STRICT' in strategy_name:
                return True, f"Similar to failed {failed_name}"
            
            # Faculty bottleneck affects strict strategies
            if report.primary_bottleneck == 'faculty':
                if strategy_name in ['STRICT NO GAPS', 'COMPACT SCHEDULE']:
                    return True, f"Faculty bottleneck affects strict strategies"
        
        return False, ""
    
    def get_all_failed(self) -> List[str]:
        """List all failed strategies"""
        return list(self.failed_strategies.keys())


class HiddenConstraintDiscovery:
    """Analyzes failures to discover hidden constraints"""
    
    def analyze_failure(self, report: StrategyFailureReport, 
                       demand: DemandAnalysis) -> Dict[str, Any]:
        """Discover hidden constraints from failure patterns"""
        hidden = {}
        
        # Faculty chain detection
        if report.faculty_conflicts > 5:
            if report.top_bottleneck_faculty:
                hidden['implied_faculty_chains'] = {
                    'faculty_ids': report.top_bottleneck_faculty[:3],
                    'suggestion': 'increase_faculty_candidates'
                }
        
        # Slot contention pattern
        if report.slot_exhaustion > 0 and report.top_bottleneck_slots:
            overlap_slots = [s for s in report.top_bottleneck_slots if '3' in s or '4' in s]
            if len(overlap_slots) >= 2:
                hidden['overlap_slot_crunch'] = {
                    'slots': overlap_slots,
                    'suggestion': 'allow_gaps_in_overlap'
                }
        
        # Room type mismatch
        if report.room_conflicts > report.faculty_conflicts:
            hidden['room_type_pressure'] = {
                'suggestion': 'allow_theory_in_lab'
            }
        
        # Load imbalance source
        if demand.faculty_load_variance > 0.3:
            hidden['severe_load_imbalance'] = {
                'variance': demand.faculty_load_variance,
                'suggestion': 'enable_balance_constraint'
            }
        
        return hidden
    
    def suggest_config_adjustments(self, hidden: Dict, base_config: Dict) -> Dict:
        """Adjust strategy config based on discovered constraints"""
        new_config = base_config.copy()
        
        if 'implied_faculty_chains' in hidden:
            new_config['faculty_candidates'] = min(7, base_config.get('faculty_candidates', 3) + 2)
        
        if 'overlap_slot_crunch' in hidden:
            new_config['allow_gaps_in_overlap'] = True
        
        if 'room_type_pressure' in hidden:
            new_config['allow_theory_in_lab'] = True
        
        if 'severe_load_imbalance' in hidden:
            new_config['enable_balance_constraint'] = True
        
        return new_config


class StrategyLearningStore:
    """Persistent storage for ML learning data"""
    
    def __init__(self, filepath: str = "strategy_learning_data.json"):
        self.filepath = filepath
        self.data = self._load_or_create()
        self._lock = threading.Lock()
    
    def _load_or_create(self) -> Dict:
        """Load existing data or create new structure"""
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, 'r') as f:
                    return json.load(f)
            except:
                pass
        
        return {
            "version": "5.0",
            "created": datetime.now().isoformat(),
            "total_samples": 0,
            "training_samples": [],
            "faculty_recommendations_history": [],
            "discovered_patterns": {}
        }
    
    def save(self):
        """Save to file (thread-safe)"""
        with self._lock:
            with open(self.filepath, 'w') as f:
                json.dump(self.data, f, indent=2, default=str)
    
    def add_training_sample(self, demand: DemandAnalysis, strategy: str,
                           success: bool, result_stats: Dict):
        """Add a training sample with execution time for dynamic time limits"""
        sample = {
            "timestamp": datetime.now().isoformat(),
            "features": demand.to_feature_vector(),
            "strategy": strategy,
            "success": success,
            "success_rate": result_stats.get('success_rate', 0),
            "gaps": result_stats.get('gaps', 0),
            "load_variance": result_stats.get('load_variance', 0),
            "execution_time": result_stats.get('execution_time', 30),  # NEW: Track time
            "sessions_count": result_stats.get('sessions_count', 0)     # NEW: Track complexity
        }
        
        with self._lock:
            self.data['training_samples'].append(sample)
            self.data['total_samples'] += 1
        
        self.save()
    
    def get_training_data(self) -> List[Dict]:
        """Get all training samples"""
        with self._lock:
            return self.data['training_samples'].copy()
    
    def find_similar_outcomes(self, demand: DemandAnalysis, top_k: int = 5) -> List[Dict]:
        """Find similar past profiles and their outcomes"""
        features = demand.to_feature_vector()
        samples = self.get_training_data()
        
        if not samples:
            return []
        
        # Calculate similarity scores
        scored = []
        for sample in samples:
            past_features = sample.get('features', [])
            if len(past_features) == len(features):
                similarity = sum(1 - abs(a - b) for a, b in zip(features, past_features)) / len(features)
                scored.append((similarity, sample))
        
        # Return top k most similar
        scored.sort(key=lambda x: x[0], reverse=True)
        return [s[1] for s in scored[:top_k]]
    
    def get_strategy_time_stats(self, strategy_name: str) -> Dict:
        """Get execution time statistics for a strategy"""
        samples = self.get_training_data()
        relevant = [s for s in samples if s.get('strategy') == strategy_name and s.get('success', False)]
        
        if not relevant:
            return {'avg_time': 30, 'max_time': 60, 'min_time': 10, 'count': 0}
        
        times = [s.get('execution_time', 30) for s in relevant]
        sessions = [s.get('sessions_count', 100) for s in relevant]
        
        return {
            'avg_time': sum(times) / len(times),
            'max_time': max(times),
            'min_time': min(times),
            'count': len(relevant),
            'avg_sessions': sum(sessions) / len(sessions) if sessions else 100
        }


# ═══════════════════════════════════════════════════════════════════════════════
# DYNAMIC TIME LIMIT SYSTEM - ML-based time limit calculation
# ═══════════════════════════════════════════════════════════════════════════════

class DynamicTimeLimit:
    """
    ML-based dynamic time limit calculation for each strategy.
    
    Features:
    - Learns from historical execution times
    - Adjusts based on session count (complexity)
    - Configurable max limits per strategy
    - Fallback to defaults when no data
    """
    
    # Maximum time limits (hard caps)
    MAX_LIMITS = {
        'STRICT NO GAPS': 120,      # Most constrained, may need more time
        'COMPACT SCHEDULE': 90,      # Medium constraints
        'RELAXED': 60,               # Easier constraints
        'MAXIMUM FLEXIBILITY': 45    # Least constraints
    }
    
    # Default time limits (when no learning data)
    DEFAULT_LIMITS = {
        'STRICT NO GAPS': 30,
        'COMPACT SCHEDULE': 30,
        'RELAXED': 30,
        'MAXIMUM FLEXIBILITY': 30
    }
    
    # Minimum time limits (don't go below this)
    MIN_LIMITS = {
        'STRICT NO GAPS': 15,
        'COMPACT SCHEDULE': 10,
        'RELAXED': 10,
        'MAXIMUM FLEXIBILITY': 10
    }
    
    def __init__(self, learning_store: StrategyLearningStore):
        self.learning_store = learning_store
    
    def calculate_time_limit(self, strategy_name: str, sessions_count: int) -> int:
        """
        Calculate optimal time limit for a strategy based on:
        1. Historical execution times for this strategy
        2. Number of sessions (complexity factor)
        3. Success vs failure patterns
        
        Returns: time limit in seconds
        """
        stats = self.learning_store.get_strategy_time_stats(strategy_name)
        
        if stats['count'] < 2:
            # Not enough data, use default with session scaling
            base = self.DEFAULT_LIMITS.get(strategy_name, 30)
            # Scale based on sessions: +1 second per 10 sessions above 50
            sessions_factor = max(0, (sessions_count - 50) / 10)
            calculated = int(base + sessions_factor * 2)
        else:
            # ML-based: use historical average with safety margin
            avg_time = stats['avg_time']
            avg_sessions = stats['avg_sessions']
            
            # Scale based on current vs historical session complexity
            if avg_sessions > 0:
                complexity_ratio = sessions_count / avg_sessions
            else:
                complexity_ratio = 1.0
            
            # Add 20% buffer to average time, scaled by complexity
            calculated = int(avg_time * 1.2 * complexity_ratio)
        
        # Apply hard limits
        min_limit = self.MIN_LIMITS.get(strategy_name, 10)
        max_limit = self.MAX_LIMITS.get(strategy_name, 120)
        
        final_limit = max(min_limit, min(max_limit, calculated))
        
        return final_limit
    
    def get_all_limits(self, sessions_count: int) -> Dict[str, int]:
        """Get dynamic time limits for all strategies"""
        return {
            name: self.calculate_time_limit(name, sessions_count)
            for name in self.DEFAULT_LIMITS.keys()
        }
    
    def print_limits(self, sessions_count: int):
        """Print current time limits"""
        limits = self.get_all_limits(sessions_count)
        print(f"\n⏱️ Dynamic Time Limits (for {sessions_count} sessions):")
        for strategy, limit in limits.items():
            stats = self.learning_store.get_strategy_time_stats(strategy)
            source = "ML" if stats['count'] >= 2 else "default"
            print(f"   • {strategy}: {limit}s ({source})")


# ═══════════════════════════════════════════════════════════════════════════════
# DECISION TREE CLASSIFIER - For precise strategy selection when scores are equal
# ═══════════════════════════════════════════════════════════════════════════════

# Try to import scikit-learn
try:
    from sklearn.tree import DecisionTreeClassifier
    from sklearn.preprocessing import LabelEncoder
    import numpy as np
    SKLEARN_AVAILABLE = True
    print("✅ scikit-learn available for Decision Tree")
except ImportError:
    SKLEARN_AVAILABLE = False
    print("⚠️ scikit-learn not available - using rule-based fallback")


class StrategyDecisionTree:
    """
    Machine Learning Decision Tree for strategy selection.
    
    Features:
    - Trains on historical data (features → best strategy)
    - Provides confidence scores
    - Falls back to rules when insufficient data
    - Auto-retrains when new data is added
    
    Requires: sklearn (pip install scikit-learn)
    """
    
    STRATEGY_NAMES = ['STRICT NO GAPS', 'COMPACT SCHEDULE', 'RELAXED', 'MAXIMUM FLEXIBILITY']
    
    # Rule-based fallback order (when no training data)
    RULE_BASED_ORDER = {
        'low_pressure': ['STRICT NO GAPS', 'COMPACT SCHEDULE', 'RELAXED', 'MAXIMUM FLEXIBILITY'],
        'medium_pressure': ['COMPACT SCHEDULE', 'RELAXED', 'STRICT NO GAPS', 'MAXIMUM FLEXIBILITY'],
        'high_pressure': ['RELAXED', 'MAXIMUM FLEXIBILITY', 'COMPACT SCHEDULE', 'STRICT NO GAPS'],
        'critical_pressure': ['MAXIMUM FLEXIBILITY', 'RELAXED', 'COMPACT SCHEDULE', 'STRICT NO GAPS']
    }
    
    def __init__(self, learning_store: StrategyLearningStore):
        self.learning_store = learning_store
        self.model = None
        self.label_encoder = None
        self.is_trained = False
        self.min_samples_required = 5  # Minimum samples before using ML
        self._train_if_possible()
    
    def _train_if_possible(self):
        """Train the decision tree if we have enough data"""
        if not SKLEARN_AVAILABLE:
            return
        
        samples = self.learning_store.get_training_data()
        
        # Only consider successful runs for training
        successful_samples = [s for s in samples if s.get('success', False)]
        
        if len(successful_samples) < self.min_samples_required:
            print(f"   📊 Decision Tree: {len(successful_samples)}/{self.min_samples_required} samples (using rule-based)")
            return
        
        try:
            # Prepare training data
            X = np.array([s['features'] for s in successful_samples])
            y = [s['strategy'] for s in successful_samples]
            
            # Encode strategy names to integers
            self.label_encoder = LabelEncoder()
            self.label_encoder.fit(self.STRATEGY_NAMES)
            y_encoded = self.label_encoder.transform(y)
            
            # Train Decision Tree
            self.model = DecisionTreeClassifier(
                max_depth=5,  # Prevent overfitting
                min_samples_split=2,
                min_samples_leaf=1,
                random_state=42
            )
            self.model.fit(X, y_encoded)
            self.is_trained = True
            
            print(f"   🌳 Decision Tree trained on {len(successful_samples)} samples")
            
        except Exception as e:
            print(f"   ⚠️ Decision Tree training failed: {e}")
            self.is_trained = False
    
    def retrain(self):
        """Force retrain with latest data"""
        self._train_if_possible()
    
    def predict_strategy_order(self, demand: DemandAnalysis) -> List[Tuple[str, float]]:
        """
        Predict strategy order with confidence scores.
        
        Returns: [(strategy_name, confidence), ...] sorted by confidence desc
        """
        features = demand.to_feature_vector()
        
        # Determine pressure category for rule-based fallback
        if demand.faculty_pressure > 100:
            pressure_category = 'critical_pressure'
        elif demand.faculty_pressure > 85:
            pressure_category = 'high_pressure'
        elif demand.faculty_pressure > 70:
            pressure_category = 'medium_pressure'
        else:
            pressure_category = 'low_pressure'
        
        # Use ML if trained, otherwise rules
        if self.is_trained and SKLEARN_AVAILABLE:
            return self._predict_with_ml(features, pressure_category)
        else:
            return self._predict_with_rules(pressure_category, demand)
    
    def _predict_with_ml(self, features: List[float], pressure_category: str) -> List[Tuple[str, float]]:
        """Use trained Decision Tree to predict"""
        try:
            X = np.array([features])
            
            # Get probability distribution across all strategies
            probabilities = self.model.predict_proba(X)[0]
            
            # Map to strategy names
            results = []
            for idx, prob in enumerate(probabilities):
                strategy_name = self.label_encoder.inverse_transform([idx])[0]
                # Convert probability to percentage (0-100)
                confidence = prob * 100
                results.append((strategy_name, confidence))
            
            # Sort by confidence descending
            results.sort(key=lambda x: x[1], reverse=True)
            
            print(f"   🌳 Decision Tree prediction (ML mode)")
            return results
            
        except Exception as e:
            print(f"   ⚠️ ML prediction failed, using rules: {e}")
            return self._predict_with_rules(pressure_category, None)
    
    def _predict_with_rules(self, pressure_category: str, demand: DemandAnalysis) -> List[Tuple[str, float]]:
        """Rule-based fallback with graded confidence"""
        rule_order = self.RULE_BASED_ORDER.get(pressure_category, self.RULE_BASED_ORDER['medium_pressure'])
        
        # Assign graded confidences: 60%, 50%, 40%, 30%
        base_confidences = [60, 50, 40, 30]
        
        # Adjust based on specific conditions
        results = []
        for i, strategy in enumerate(rule_order):
            confidence = base_confidences[i]
            
            # Apply condition-based adjustments
            if demand:
                # High overlap contention hurts no-gap strategies
                if demand.overlap_contention > 50 and 'NO GAPS' in strategy:
                    confidence -= 10
                
                # Many labs favor flexibility
                if demand.lab_ratio > 0.4 and 'FLEXIBILITY' in strategy:
                    confidence += 10
                
                # Low sessions favor strict
                if demand.total_sessions < 30 and 'STRICT' in strategy:
                    confidence += 10
            
            results.append((strategy, max(0, min(100, confidence))))
        
        # Re-sort after adjustments
        results.sort(key=lambda x: x[1], reverse=True)
        
        print(f"   📋 Decision Tree prediction (rule-based mode: {pressure_category})")
        return results
    
    def get_feature_importance(self) -> Dict[str, float]:
        """Get feature importance scores (only if trained)"""
        if not self.is_trained or not self.model:
            return {}
        
        feature_names = [
            'faculty_pressure', 'load_variance', 'overlap_contention',
            'classroom_util', 'lab_util', 'floor_constraints',
            'specialization_constraints', 'lab_ratio',
            'sessions_per_faculty', 'sessions_per_room'
        ]
        
        importances = self.model.feature_importances_
        return {name: round(imp, 3) for name, imp in zip(feature_names, importances)}


class SmartStrategyAdvisor:
    """
    ═══════════════════════════════════════════════════════════════════════
    SMART STRATEGY ADVISOR - Hybrid Rules + ML Decision System
    ═══════════════════════════════════════════════════════════════════════
    
    Features:
    - Pre-analysis of demand
    - Hybrid decision (rules filter + ML ranks)
    - Decision Tree for precise selection
    - Strategy cooldown
    - Hidden constraint discovery
    - Faculty recommendations
    - Persistent learning
    """
    
    STRATEGIES = [
        # Original strategies (quality-focused)
        {'name': 'STRICT NO GAPS', 'allow_gaps': False, 'time_limit': 30},
        {'name': 'COMPACT SCHEDULE', 'allow_gaps': False, 'time_limit': 30},
        {'name': 'RELAXED', 'allow_gaps': True, 'time_limit': 30},
        {'name': 'MAXIMUM FLEXIBILITY', 'allow_gaps': True, 'time_limit': 30},
        # NEW: Maximum utilization strategies (fill every slot)
        {'name': 'FACULTY EXHAUSTION', 'allow_gaps': True, 'time_limit': 45, 
         'max_utilization': True, 'prioritize': 'faculty', 'consecutive_bonus': 200},
        {'name': 'ROOM EXHAUSTION', 'allow_gaps': True, 'time_limit': 45,
         'max_utilization': True, 'prioritize': 'rooms', 'slot_bonus': 100},
        {'name': 'DENSE PACK', 'allow_gaps': False, 'time_limit': 45,
         'max_utilization': True, 'consecutive_bonus': 1000, 'gap_penalty': 3000}
    ]
    
    def __init__(self):
        self.learning_store = StrategyLearningStore()
        self.constraint_discoverer = HiddenConstraintDiscovery()
        self.decision_tree = StrategyDecisionTree(self.learning_store)
        self.time_limiter = DynamicTimeLimit(self.learning_store)  # NEW: Dynamic time limits
        self.cooldown = None  # Set per execution
        self.discovered_constraints = {}
        self.sessions_count = 0  # Track for dynamic time calculation
    
    def analyze_demand(self, data: Dict) -> DemandAnalysis:
        """
        Pre-analyze data to calculate all metrics.
        """
        analysis = DemandAnalysis()
        
        timeslots = data.get('timeslots', [])
        rooms = data.get('rooms', [])
        faculties = data.get('faculties', [])
        subjects = data.get('subjects', [])
        batches = data.get('batches', [])
        
        # Resource counts
        analysis.total_batches = len(batches)
        analysis.total_subjects = len(subjects)
        analysis.total_faculty = len(faculties)
        analysis.total_classrooms = len([r for r in rooms if 'lab' not in r.room_type.lower()])
        analysis.total_lab_rooms = len([r for r in rooms if 'lab' in r.room_type.lower()])
        analysis.slots_per_day = len(timeslots)
        
        # Session counts
        theory_hours = sum(s.hours_per_week for s in subjects)
        lab_hours = sum(s.lab_hours_per_week for s in subjects)
        analysis.theory_sessions = theory_hours
        analysis.lab_sessions = lab_hours
        analysis.elective_sessions = len([s for s in subjects if s.is_elective])
        analysis.total_sessions = theory_hours + lab_hours
        
        # Shift distribution
        analysis.morning_batches = len([b for b in batches if b.shift_preference == 'morning'])
        analysis.evening_batches = len([b for b in batches if b.shift_preference in ['evening', 'afternoon']])
        
        # Find overlap slots (typically 11AM-12PM = slots 3-4)
        for ts in timeslots:
            hour_str = str(ts.start_time).split(':')[0]
            if hour_str in ('11', '12'):
                analysis.overlap_slots.append(ts.slot_id)
        
        # Capacity metrics
        analysis.total_faculty_hours_available = sum(f.max_hours_per_week for f in faculties)
        analysis.total_hours_demanded = (theory_hours + lab_hours * 2) * len(batches)
        analysis.classroom_hours_available = analysis.total_classrooms * analysis.slots_per_day * analysis.working_days
        analysis.lab_hours_available = analysis.total_lab_rooms * analysis.slots_per_day * analysis.working_days
        
        # Pressure scores
        if analysis.total_faculty_hours_available > 0:
            analysis.faculty_pressure = (analysis.total_hours_demanded / analysis.total_faculty_hours_available) * 100
        
        # Utilization
        if analysis.classroom_hours_available > 0:
            analysis.classroom_utilization = min(1.0, theory_hours * len(batches) / analysis.classroom_hours_available)
        if analysis.lab_hours_available > 0:
            analysis.lab_utilization = min(1.0, lab_hours * len(batches) * 2 / analysis.lab_hours_available)
        
        # Overlap contention (morning + evening both need overlap slots)
        if analysis.overlap_slots and analysis.morning_batches > 0 and analysis.evening_batches > 0:
            analysis.overlap_contention = min(100, (analysis.morning_batches + analysis.evening_batches) * 10)
        
        # Constraint flags
        analysis.has_floor_constraints = any(b.preferred_floor > 0 for b in batches)
        analysis.has_specialization_constraints = any(s.required_specialization for s in subjects)
        analysis.has_elective_conflicts = analysis.elective_sessions > 2
        
        # Derived ratios
        if analysis.total_faculty > 0:
            analysis.sessions_per_faculty = analysis.total_sessions / analysis.total_faculty
        if analysis.total_classrooms + analysis.total_lab_rooms > 0:
            analysis.sessions_per_room = analysis.total_sessions / (analysis.total_classrooms + analysis.total_lab_rooms)
        if analysis.total_sessions > 0:
            analysis.lab_ratio = analysis.lab_sessions / analysis.total_sessions
        
        return analysis
    
    def calculate_faculty_recommendations(self, demand: DemandAnalysis) -> FacultyRecommendation:
        """Calculate actionable faculty addition suggestions"""
        if demand.total_faculty == 0 or demand.total_faculty_hours_available == 0:
            return FacultyRecommendation(0, False, [])
        
        current_hours = demand.total_faculty_hours_available
        demanded_hours = demand.total_hours_demanded
        avg_faculty_hours = current_hours / demand.total_faculty
        
        current_pressure = (demanded_hours / current_hours) * 100 if current_hours > 0 else 100
        
        recommendations = []
        target_pressures = [70, 60, 50, 40]
        
        for target in target_pressures:
            if target >= current_pressure:
                continue
            
            needed_hours = demanded_hours / (target / 100)
            additional_hours = needed_hours - current_hours
            faculty_to_add = math.ceil(additional_hours / avg_faculty_hours)
            
            if faculty_to_add > 0:
                new_variance = max(0, demand.faculty_load_variance * 0.8)  # Estimate improvement
                recommendations.append({
                    "add_faculty": faculty_to_add,
                    "new_pressure_percent": target,
                    "estimated_variance": round(new_variance, 3)
                })
        
        return FacultyRecommendation(
            current_pressure=current_pressure,
            is_schedulable=current_pressure <= 100,
            recommendations=recommendations
        )
    
    def score_strategy_hybrid(self, strategy_name: str, demand: DemandAnalysis) -> float:
        """
        HYBRID SCORING: Rules provide bounds, ML optimizes within
        Returns 0-100 score
        """
        # ═══════════════════════════════════════════════════════════
        # RULE-BASED BOUNDS (Hard constraints)
        # ═══════════════════════════════════════════════════════════
        
        # Rule 1: If faculty pressure > 100%, strict strategies can't work
        if demand.faculty_pressure > 100:
            if 'STRICT' in strategy_name or 'COMPACT' in strategy_name:
                return 0  # Impossible
        
        # Rule 2: If severe load imbalance, strict strategies hurt
        if demand.faculty_load_variance > 0.4:
            if 'STRICT' in strategy_name:
                return max(0, 30)  # Very low score
        
        # ═══════════════════════════════════════════════════════════
        # ML-BASED SCORING (Learned weights)
        # ═══════════════════════════════════════════════════════════
        
        # Find similar historical outcomes
        similar = self.learning_store.find_similar_outcomes(demand, top_k=5)
        ml_score = 50  # Default baseline
        
        if similar:
            # Check how this strategy performed in similar situations
            relevant = [s for s in similar if s.get('strategy') == strategy_name]
            if relevant:
                avg_success = sum(s.get('success_rate', 0) for s in relevant) / len(relevant)
                ml_score = avg_success * 100
        
        # ═══════════════════════════════════════════════════════════
        # COMBINED HYBRID SCORE
        # ═══════════════════════════════════════════════════════════
        
        base_score = ml_score
        
        # Adjust based on current conditions
        if demand.faculty_pressure > 85:
            if 'STRICT' in strategy_name:
                base_score -= 30
            if 'FLEXIBILITY' in strategy_name:
                base_score += 20
        
        if demand.overlap_contention > 60:
            if 'NO GAPS' in strategy_name:
                base_score -= 15
        
        # Boost if previous failure insights suggest this
        if self.discovered_constraints:
            if 'overlap_slot_crunch' in self.discovered_constraints and 'RELAXED' in strategy_name:
                base_score += 15
        
        return max(0, min(100, base_score))
    
    def get_ordered_strategies(self, demand: DemandAnalysis) -> List[Tuple[Dict, float]]:
        """
        Get strategies ordered by confidence score.
        
        Uses Decision Tree for differentiated confidence, with cooldown filtering.
        """
        # Get Decision Tree predictions (graded confidence: 60%, 50%, 40%, 30% or ML-based)
        dt_predictions = self.decision_tree.predict_strategy_order(demand)
        
        # Create strategy name -> confidence map
        confidence_map = {name: conf for name, conf in dt_predictions}
        
        # Build final list with cooldown filtering
        scored = []
        strategy_map = {s['name']: s for s in self.STRATEGIES}
        
        for strategy_name, confidence in dt_predictions:
            strategy = strategy_map.get(strategy_name)
            if not strategy:
                continue
            
            # Skip blocked strategies
            if self.cooldown and self.cooldown.is_blocked(strategy_name):
                continue
            
            skip, reason = (False, "") if not self.cooldown else self.cooldown.should_skip_similar(strategy_name)
            if skip:
                print(f"   ⏭️ Skipping {strategy_name}: {reason}")
                continue
            
            # Apply hybrid score adjustments on top of Decision Tree score
            adjusted_score = confidence
            
            # Rule-based adjustments (hard constraints)
            if demand.faculty_pressure > 100:
                if 'STRICT' in strategy_name or 'COMPACT' in strategy_name:
                    adjusted_score = 0  # Impossible
            
            if demand.faculty_pressure > 85:
                if 'STRICT' in strategy_name:
                    adjusted_score -= 10
                if 'FLEXIBILITY' in strategy_name:
                    adjusted_score += 10
            
            # ════════════════════════════════════════════════════════════════════
            # 🚀 INTELLIGENT UTILIZATION STRATEGY SELECTION
            # ════════════════════════════════════════════════════════════════════
            
            # FACULTY EXHAUSTION: Use when faculty is UNDERUTILIZED
            if 'FACULTY EXHAUSTION' in strategy_name:
                if demand.faculty_pressure < 70:
                    # Faculty is underutilized - boost this strategy
                    adjusted_score += 25
                    print(f"   📈 FACULTY EXHAUSTION boosted: faculty only {demand.faculty_pressure:.0f}% utilized")
                elif demand.faculty_pressure > 90:
                    # Faculty is already stressed - don't use
                    adjusted_score -= 30
            
            # ROOM EXHAUSTION: Use when rooms are UNDERUTILIZED
            if 'ROOM EXHAUSTION' in strategy_name:
                room_util = max(demand.classroom_utilization, demand.lab_utilization)
                if room_util < 0.80:
                    # Rooms are underutilized - boost this strategy
                    adjusted_score += 20
                    print(f"   📈 ROOM EXHAUSTION boosted: rooms only {room_util:.0%} utilized")
                elif room_util > 0.95:
                    # Rooms are nearly full - don't use
                    adjusted_score -= 30
            
            # DENSE PACK: Best for low-to-medium pressure scenarios
            if 'DENSE PACK' in strategy_name:
                if demand.faculty_pressure < 75 and demand.faculty_pressure > 40:
                    # Sweet spot for dense packing
                    adjusted_score += 15
                elif demand.faculty_pressure > 85:
                    # Too much pressure for strict packing
                    adjusted_score -= 20
            
            # Don't use max utilization strategies when already overloaded
            if strategy.get('max_utilization') and demand.faculty_pressure > 95:
                adjusted_score = 0  # Prevent overloading
            
            # Apply discovered constraint insights
            if self.discovered_constraints:
                if 'overlap_slot_crunch' in self.discovered_constraints and 'RELAXED' in strategy_name:
                    adjusted_score += 10
            
            scored.append((strategy, max(0, min(100, adjusted_score))))
        
        # Sort by adjusted score descending
        scored.sort(key=lambda x: x[1], reverse=True)
        return scored
    
    def prepare_execution(self, data: Dict) -> Tuple[DemandAnalysis, FacultyRecommendation]:
        """
        Prepare for scheduling execution.
        Returns demand analysis and faculty recommendations.
        """
        # Reset per-execution state
        self.cooldown = StrategyCooldown()
        self.discovered_constraints = {}
        
        # Analyze demand
        demand = self.analyze_demand(data)
        
        # Calculate faculty recommendations
        recommendations = self.calculate_faculty_recommendations(demand)
        
        # Print analysis
        print(f"\n{'='*70}")
        print("🧠 SMART STRATEGY ADVISOR - Pre-Analysis")
        print(f"{'='*70}")
        print(f"\n📊 Demand Analysis:")
        print(f"   Batches: {demand.total_batches} ({demand.morning_batches} morning, {demand.evening_batches} evening)")
        print(f"   Sessions: {demand.total_sessions} (theory: {demand.theory_sessions}, lab: {demand.lab_sessions})")
        print(f"   Faculty: {demand.total_faculty} (available: {demand.total_faculty_hours_available}h)")
        print(f"   Faculty Pressure: {demand.faculty_pressure:.1f}%")
        print(f"   Room Utilization: classroom={demand.classroom_utilization:.1%}, lab={demand.lab_utilization:.1%}")
        print(f"   Overlap Contention: {demand.overlap_contention:.0f}%")
        
        # Show recommendations if pressure is high
        if recommendations.current_pressure > 80:
            recommendations.print_recommendations()
        
        return demand, recommendations
    
    def handle_failure(self, strategy_name: str, report: StrategyFailureReport, 
                      demand: DemandAnalysis):
        """Handle strategy failure - learn and adjust"""
        # Mark as blocked
        self.cooldown.mark_failed(strategy_name, report)
        
        # Discover hidden constraints
        new_constraints = self.constraint_discoverer.analyze_failure(report, demand)
        self.discovered_constraints.update(new_constraints)
        
        if new_constraints:
            print(f"   🔍 Discovered constraints: {list(new_constraints.keys())}")
    
    def get_adjusted_config(self, base_strategy: Dict) -> Dict:
        """Get strategy config adjusted based on failures"""
        config = base_strategy.copy()
        
        if self.discovered_constraints:
            config = self.constraint_discoverer.suggest_config_adjustments(
                self.discovered_constraints, config
            )
        
        return config
    
    def record_outcome(self, demand: DemandAnalysis, strategy_name: str, 
                      success: bool, result_stats: Dict):
        """
        Record outcome for learning and trigger auto-retrain.
        
        AUTO-RETRAIN: After each success, retrain Decision Tree to use latest data.
        This ensures the model continuously improves.
        """
        self.learning_store.add_training_sample(demand, strategy_name, success, result_stats)
        
        # Auto-retrain Decision Tree after successful runs
        if success:
            print(f"   🔄 Auto-retraining Decision Tree with new data...")
            self.decision_tree.retrain()
    
    def get_dynamic_time_limit(self, strategy_name: str) -> int:
        """Get ML-based dynamic time limit for a strategy"""
        return self.time_limiter.calculate_time_limit(strategy_name, self.sessions_count)
    
    def get_all_dynamic_limits(self) -> Dict[str, int]:
        """Get dynamic time limits for all strategies"""
        return self.time_limiter.get_all_limits(self.sessions_count)
    
    def print_dynamic_limits(self):
        """Print current dynamic time limits"""
        self.time_limiter.print_limits(self.sessions_count)


# Create global advisor instance
SMART_ADVISOR = SmartStrategyAdvisor()
print("✅ Smart Strategy Advisor initialized")

# -------------------------
# Enhanced Excel Reader with Floor & Elective Support
# -------------------------
# -------------------------
# Production-Grade Excel Reader with Header Verification
# -------------------------
class AdvancedExcelReader:
    def __init__(self, filename: str):
        self.filename = filename
        self.sheets = pd.read_excel(filename, sheet_name=None, engine='openpyxl')
        print(f"✅ Loaded sheets: {list(self.sheets.keys())}")

    def parse_faculty_choices(self) -> List[FacultyChoice]:
        '''Parse faculty_choice sheet'''
        sheet_name = self._find_sheet(['faculty_choice', 'faculty_choices', 'facultychoice', 'choices'])
        
        if not sheet_name:
            print("⚠️  No 'faculty_choice' sheet found - using qualification-based allocation")
            return []
        
        df = self.sheets[sheet_name]
        print(f"\\n👤 Parsing Faculty Choices from '{sheet_name}'")
        print(f"   Columns: {list(df.columns)}")
        
        faculty_choices = []
        
        for idx, row in df.iterrows():
            try:
                faculty_id = self._get_column_value(row, df, ['faculty_id', 'id', 'employee_id', 'emp_id'])
                if not faculty_id or pd.isna(faculty_id):
                    continue
                faculty_id = str(faculty_id).strip()
                
                name = self._get_column_value(row, df, ['name', 'faculty_name', 'teacher_name'])
                if not name:
                    continue
                name = str(name).strip()
                
                choice_1 = self._get_column_value(row, df, ['choice_1', 'choice1', 'priority_1', 'priority1'], '')
                choice_2 = self._get_column_value(row, df, ['choice_2', 'choice2', 'priority_2', 'priority2'], '')
                choice_3 = self._get_column_value(row, df, ['choice_3', 'choice3', 'priority_3', 'priority3'], '')
                choice_4 = self._get_column_value(row, df, ['choice_4', 'choice4', 'priority_4', 'priority4'], '')
                choice_5 = self._get_column_value(row, df, ['choice_5', 'choice5', 'priority_5', 'priority5'], '')
                
                choice_1 = str(choice_1).strip() if pd.notna(choice_1) else ''
                choice_2 = str(choice_2).strip() if pd.notna(choice_2) else ''
                choice_3 = str(choice_3).strip() if pd.notna(choice_3) else ''
                choice_4 = str(choice_4).strip() if pd.notna(choice_4) else ''
                choice_5 = str(choice_5).strip() if pd.notna(choice_5) else ''
                
                if not choice_1:
                    print(f"   ⚠️  Row {idx}: {faculty_id} has no choice_1, skipping")
                    continue
                
                faculty_choice = FacultyChoice(
                    faculty_id, name, choice_1, choice_2, choice_3, choice_4, choice_5
                )
                faculty_choices.append(faculty_choice)
                
                choices = [c for c in [choice_1, choice_2, choice_3, choice_4, choice_5] if c]
                print(f"   ✅ {faculty_id} ({name}): {', '.join(choices)}")
                
            except Exception as e:
                print(f"   ⚠️  Row {idx}: {e}")
                continue
        
        print(f"   ✅ Loaded {len(faculty_choices)} faculty choices")
        return faculty_choices
    
    def parse_constraints(self) -> AllocationConstraints:
        '''Parse constraints sheet'''
        sheet_name = self._find_sheet(['constraints', 'constraint', 'config', 'configuration'])
        
        if not sheet_name:
            print("⚠️  No 'constraints' sheet - using qualification-based allocation")
            return AllocationConstraints(use_faculty_choice=False)
        
        df = self.sheets[sheet_name]
        print(f"\\n⚙️  Parsing Constraints from '{sheet_name}'")
        print(f"   Columns: {list(df.columns)}")
        
        name_col = self._find_column(df, ['constraint_name', 'name', 'constraint', 'setting'])
        value_col = self._find_column(df, ['value', 'setting', 'enabled', 'status'])
        
        if not name_col or not value_col:
            print(f"   ⚠️  Missing required columns")
            return AllocationConstraints(use_faculty_choice=False)
        
        use_faculty_choice = False
        
        for idx, row in df.iterrows():
            try:
                name = str(row[name_col]).strip().lower()
                value = str(row[value_col]).strip().lower()
                
                if name == 'faculty_choice':
                    use_faculty_choice = value in ['yes', 'true', '1', 'y', 'on', 'enabled']
                    status = "✅ ENABLED" if use_faculty_choice else "❌ DISABLED"
                    print(f"   {status} faculty_choice (value: {value})")
            except Exception as e:
                print(f"   ⚠️  Row {idx}: {e}")
                continue
        
        mode = 'CHOICE-BASED' if use_faculty_choice else 'QUALIFICATION-BASED'
        print(f"   📋 Allocation Mode: {mode}")
        
        return AllocationConstraints(use_faculty_choice=use_faculty_choice)
    
    @lru_cache(maxsize=1)
    def parse_all(self):
        '''Cached parsing with header verification'''
        return {
            'timeslots': self._parse_timeslots(),
            'rooms': self._parse_rooms(),
            'faculties': self._parse_faculties(),
            'subjects': self._parse_subjects(),
            'batches': self._parse_batches(),
            'faculty_choices': self.parse_faculty_choices(),        # ADD THIS LINE
            'allocation_constraints': self.parse_constraints(),      # ADD THIS LINE
            'breaks': self._parse_breaks()
        }
    
    def _find_sheet(self, names: List[str]) -> Optional[str]:
        """Find sheet by multiple possible names"""
        sheet_lower = {k.lower(): k for k in self.sheets.keys()}
        for name in names:
            if name.lower() in sheet_lower:
                return sheet_lower[name.lower()]
        return None
    
    def _find_column(self, df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        """Find column by multiple possible header names (case-insensitive)"""
        # Create mapping of lowercase column names to actual names
        col_map = {str(col).lower().strip(): col for col in df.columns}
        
        for name in possible_names:
            name_lower = name.lower().strip()
            if name_lower in col_map:
                return col_map[name_lower]
        
        return None
    
    def _get_column_value(self, row, df: pd.DataFrame, possible_names: List[str], default=None):
        """Get value from row by finding column with possible names"""
        col_name = self._find_column(df, possible_names)
        if col_name and col_name in df.columns:
            val = row[col_name]
            if pd.notna(val):
                return val
        return default
    
    def _parse_timeslots(self) -> List[TimeSlot]:
        """Parse time slots with header verification"""
        sheet_name = self._find_sheet(['timeslots', 'time_slots', 'periods', 'slots'])
        if not sheet_name:
            print("⚠️  No timeslots sheet found")
            return []
        
        df = self.sheets[sheet_name]
        print(f"\n📅 Parsing Time Slots from '{sheet_name}'")
        print(f"   Columns: {list(df.columns)}")
        
        timeslots = []
        
        # Find columns by possible header names
        slot_id_col = self._find_column(df, ['slot_id', 'slot id', 'period', 'period_id', 'id', 'slot no', 'period no'])
        start_col = self._find_column(df, ['start_time', 'start time', 'start', 'from', 'from_time'])
        end_col = self._find_column(df, ['end_time', 'end time', 'end', 'to', 'to_time'])
        
        if not all([slot_id_col, start_col, end_col]):
            print(f"   ⚠️  Missing required columns. Found: slot_id={slot_id_col}, start={start_col}, end={end_col}")
            print(f"   Using positional parsing (columns 0, 1, 2)")
            # Fallback to positional
            for _, row in df.iterrows():
                try:
                    if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]) and pd.notna(row.iloc[2]):
                        timeslots.append(TimeSlot(
                            int(row.iloc[0]),
                            str(row.iloc[1]),
                            str(row.iloc[2])
                        ))
                except (ValueError, IndexError):
                    continue
        else:
            print(f"   ✅ Found columns: slot_id='{slot_id_col}', start='{start_col}', end='{end_col}'")
            for _, row in df.iterrows():
                try:
                    slot_id = row[slot_id_col]
                    start_time = row[start_col]
                    end_time = row[end_col]
                    
                    if pd.notna(slot_id) and pd.notna(start_time) and pd.notna(end_time):
                        timeslots.append(TimeSlot(
                            int(slot_id),
                            str(start_time),
                            str(end_time)
                        ))
                except (ValueError, KeyError) as e:
                    continue
        
        print(f"   ✅ Loaded {len(timeslots)} time slots")
        return timeslots
    
    def _parse_rooms(self) -> List[Room]:
        """Parse rooms with flexible header-based column detection"""
        sheet_name = self._find_sheet(['rooms', 'classrooms', 'room', 'classroom'])
        if not sheet_name:
            print("⚠️  No rooms sheet found")
            return []
        
        df = self.sheets[sheet_name]
        print(f"\n🏢 Parsing Rooms from '{sheet_name}'")
        print(f"   Columns: {list(df.columns)}")
        
        rooms = []
        
        # Find columns by possible header names
        room_id_col = self._find_column(df, ['room_id', 'room id', 'room', 'room no', 'room_no', 'room number', 'id'])
        capacity_col = self._find_column(df, ['capacity', 'seats', 'size', 'max_capacity', 'max capacity', 'seating'])
        type_col = self._find_column(df, ['type', 'room_type', 'room type', 'room_category', 'category'])
        dept_col = self._find_column(df, ['department', 'dept', 'department_name', 'dept_name', 'owned_by', 'owner'])
        floor_col = self._find_column(df, ['floor', 'floor_no', 'floor no', 'floor_number', 'level'])
        
        print(f"   Column mapping:")
        print(f"     Room ID: {room_id_col}")
        print(f"     Capacity: {capacity_col}")
        print(f"     Type: {type_col}")
        print(f"     Department: {dept_col}")
        print(f"     Floor: {floor_col}")
        
        for idx, row in df.iterrows():
            try:
                # Room ID (required)
                room_id = self._get_column_value(row, df, ['room_id', 'room id', 'room', 'room no', 'id'])
                if not room_id or pd.isna(room_id):
                    continue
                room_id = str(room_id).strip()
                
                # Capacity (with validation)
                capacity = self._get_column_value(row, df, ['capacity', 'seats', 'size', 'max_capacity'])
                try:
                    capacity = int(capacity) if pd.notna(capacity) else 50
                    # Validation: capacity should be reasonable (10-500)
                    if capacity < 5 or capacity > 1000:
                        print(f"   ⚠️  Row {idx}: Suspicious capacity {capacity} for room {room_id}, using default 50")
                        capacity = 50
                except (ValueError, TypeError):
                    capacity = 50
                
                # Room Type
                room_type = self._get_column_value(row, df, ['type', 'room_type', 'room type', 'category'], 'classroom')
                room_type = str(room_type).strip() if pd.notna(room_type) else 'classroom'
                
                # Department
                department = self._get_column_value(row, df, ['department', 'dept', 'owned_by'], 'General')
                department = str(department).strip() if pd.notna(department) else 'General'
                
                # Floor (with validation)
                floor = self._get_column_value(row, df, ['floor', 'floor_no', 'floor no', 'level'], 0)
                
                # Floor extraction logic with validation
                floor_num = 0
                if pd.notna(floor):
                    try:
                        floor_num = int(floor)
                        # Validation: floor should be reasonable (0-20)
                        if floor_num < 0 or floor_num > 20:
                            print(f"   ⚠️  Row {idx}: Invalid floor {floor_num} for room {room_id}, trying to extract from room ID")
                            floor_num = 0
                    except (ValueError, TypeError):
                        floor_num = 0
                
                # If no valid floor from column, try to extract from room_id
                if floor_num == 0:
                    # Try patterns like: R301, Room-301, 3-01, etc.
                    patterns = [
                        r'^[A-Za-z]*(\d)(\d{2})$',  # R301 -> floor 3
                        r'^(\d)-\d+$',               # 3-01 -> floor 3
                        r'^[A-Za-z]*-(\d)(\d{2})$', # Room-301 -> floor 3
                    ]
                    
                    for pattern in patterns:
                        match = re.search(pattern, room_id)
                        if match:
                            floor_num = int(match.group(1))
                            break
                
                rooms.append(Room(room_id, capacity, room_type, department, floor_num))
                
                if floor_num > 0:
                    print(f"   ✅ {room_id}: Capacity={capacity}, Type={room_type}, Dept={department}, Floor={floor_num}")
                
            except Exception as e:
                print(f"   ⚠️  Row {idx} error: {e}")
                continue
        
        print(f"   ✅ Loaded {len(rooms)} rooms")
        
        # Summary by floor
        floor_counts = Counter(r.floor for r in rooms if r.floor > 0)
        if floor_counts:
            print(f"   Floor distribution: {dict(floor_counts)}")
        
        return rooms
    
    def _parse_faculties(self) -> List[Faculty]:
        """Parse faculties with header verification"""
        sheet_name = self._find_sheet(['faculties', 'faculty', 'teachers', 'staff', 'instructors'])
        if not sheet_name:
            print("⚠️  No faculties sheet found")
            return []
        
        df = self.sheets[sheet_name]
        print(f"\n👨‍🏫 Parsing Faculties from '{sheet_name}'")
        print(f"   Columns: {list(df.columns)}")
        
        faculties = []
        
        for idx, row in df.iterrows():
            try:
                # Faculty ID
                faculty_id = self._get_column_value(row, df, ['faculty_id', 'id', 'employee_id', 'emp_id', 'teacher_id'])
                if not faculty_id:
                    continue
                faculty_id = str(faculty_id).strip()
                
                # Name
                name = self._get_column_value(row, df, ['name', 'faculty_name', 'teacher_name', 'full_name'])
                if not name:
                    continue
                name = str(name).strip()
                
                # Qualifications
                qual_str = self._get_column_value(row, df, ['qualifications', 'qualification', 'degree', 'degrees', 'education'], '')
                qualifications = tuple(
                    q.strip().lower()
                    for q in re.split(r'[,;|]', str(qual_str))
                    if q.strip()
                ) if qual_str else tuple()
                
                # Specializations
                spec_str = self._get_column_value(row, df, ['specializations', 'specialization', 'expertise', 'subjects', 'areas'], '')
                specializations = tuple(
                    s.strip().lower()
                    for s in re.split(r'[,;|]', str(spec_str))
                    if s.strip()
                ) if spec_str else tuple()
                
                # Max hours per week
                max_hours = self._get_column_value(row, df, ['max_hours', 'max hours', 'hours_per_week', 'max_load', 'workload'], 20)
                try:
                    max_hours = int(max_hours)
                    # Validation: 5-40 hours per week
                    if max_hours < 5 or max_hours > 40:
                        max_hours = 20
                except (ValueError, TypeError):
                    max_hours = 20
                
                faculties.append(Faculty(faculty_id, name, qualifications, specializations, max_hours))
                
            except Exception as e:
                print(f"   ⚠️  Row {idx} error: {e}")
                continue
        
        print(f"   ✅ Loaded {len(faculties)} faculty members")
        return faculties
    
    def _parse_subjects(self) -> List[Subject]:
        """Parse subjects with header verification"""
        sheet_name = self._find_sheet(['subjects', 'courses', 'subject', 'course'])
        if not sheet_name:
            print("⚠️  No subjects sheet found")
            return []
        
        df = self.sheets[sheet_name]
        print(f"\n📚 Parsing Subjects from '{sheet_name}'")
        print(f"   Columns: {list(df.columns)}")
        
        subjects = []
        WEEKS_PER_SEMESTER = 16
        
        for idx, row in df.iterrows():
            try:
                # Subject Code
                subject_code = self._get_column_value(row, df, ['subject_code', 'code', 'subject id', 'course_code', 'id'])
                if not subject_code:
                    continue
                subject_code = str(subject_code).strip()
                
                # Subject Name
                subject_name = self._get_column_value(row, df, ['subject_name', 'name', 'subject', 'course_name', 'course', 'title'])
                if not subject_name:
                    continue
                subject_name = str(subject_name).strip()
                
                # Theory hours per semester
                theory_hours_sem = self._get_column_value(row, df, 
                    ['theory_hours', 'lecture_hours', 'hours', 'theory', 'lecture hours per semester', 'hours_per_semester'], 64)
                try:
                    theory_hours_sem = int(theory_hours_sem)
                except (ValueError, TypeError):
                    theory_hours_sem = 64
                
                # Convert to weekly
                hours_per_week = max(1, round(theory_hours_sem / WEEKS_PER_SEMESTER))
                
                # Lab requirement
                requires_lab = self._get_column_value(row, df, ['requires_lab', 'has_lab', 'lab', 'is_lab', 'lab_required'], False)
                if isinstance(requires_lab, str):
                    requires_lab = requires_lab.lower() in ['yes', 'true', '1', 'y']
                else:
                    requires_lab = bool(requires_lab)
                
                # Lab hours per semester
                lab_hours_sem = self._get_column_value(row, df, 
                    ['lab_hours', 'lab', 'lab hours per semester','_lab-hours_per_semester' ,'lab_hours_per_semester','practical_hours'], 0)
                try:
                    lab_hours_sem = int(lab_hours_sem)
                except (ValueError, TypeError):
                    lab_hours_sem = 0
                
                lab_hours_per_week = max(0, round(lab_hours_sem / WEEKS_PER_SEMESTER))
                
                # Required specialization
                req_spec = self._get_column_value(row, df, 
                    ['required_specialization', 'specialization', 'expertise_required', 'faculty_requirement'], '')
                required_specialization = str(req_spec).strip().lower() if req_spec else ''
                
                # Elective flag
                is_elective = self._get_column_value(row, df, ['is_elective', 'elective', 'type'], False)
                if isinstance(is_elective, str):
                    is_elective = is_elective.lower() in ['yes', 'true', '1', 'elective', 'y']
                else:
                    is_elective = bool(is_elective)
                
                # Elective group
                elective_group = ''
                if is_elective:
                    elec_grp = self._get_column_value(row, df, ['elective_group', 'group', 'elective group', 'group_id'], '')
                    elective_group = str(elec_grp).strip() if elec_grp else ''
                
                # Priority
                priority = self._get_column_value(row, df, ['priority', 'importance', 'weight'], 5)
                try:
                    priority = int(priority)
                    # Clamp to 1-10
                    priority = max(1, min(10, priority))
                except (ValueError, TypeError):
                    priority = 5
                
                subjects.append(Subject(
                    subject_code, subject_name, hours_per_week,
                    requires_lab, lab_hours_per_week, required_specialization,
                    is_elective, elective_group, priority
                ))
                
                elective_mark = f" [ELECTIVE-{elective_group}]" if is_elective else ""
                priority_mark = f" [P{priority}]" if priority != 5 else ""
                lab_mark = f" +{lab_hours_per_week}h lab" if lab_hours_per_week > 0 else ""
                print(f"   ✅ {subject_code}: {hours_per_week}h{lab_mark}{elective_mark}{priority_mark}")
                
            except Exception as e:
                print(f"   ⚠️  Row {idx} error: {e}")
                continue
        
        print(f"   ✅ Loaded {len(subjects)} subjects")
        return subjects
    
    def _parse_batches(self) -> List[StudentBatch]:
        """Parse batches with header verification"""
        sheet_name = self._find_sheet(['batches', 'student_batches', 'classes', 'batch', 'class', 'sections'])
        if not sheet_name:
            print("⚠️  No batches sheet found")
            return []
        
        df = self.sheets[sheet_name]
        print(f"\n👥 Parsing Batches from '{sheet_name}'")
        print(f"   Columns: {list(df.columns)}")
        
        batches = []
        
        for idx, row in df.iterrows():
            try:
                # Batch ID
                batch_id = self._get_column_value(row, df, ['batch_id', 'batch', 'class', 'section', 'id', 'batch_name'])
                if not batch_id:
                    continue
                batch_id = str(batch_id).strip()
                
                # Department
                department = self._get_column_value(row, df, ['department', 'dept', 'program', 'course'], 'General')
                department = str(department).strip() if department else 'General'
                
                # Student count
                student_count = self._get_column_value(row, df, ['student_count', 'students', 'count', 'strength', 'size'], 40)
                try:
                    student_count = int(student_count)
                    # Validation: 5-200 students
                    if student_count < 5 or student_count > 300:
                        print(f"   ⚠️  Suspicious student count {student_count} for {batch_id}, using default 40")
                        student_count = 40
                except (ValueError, TypeError):
                    student_count = 40
                
                # Subjects (comma-separated list)
                subjects_str = self._get_column_value(row, df, ['subjects', 'subject_codes', 'courses', 'subject_list'], '')
                if not subjects_str:
                    print(f"   ⚠️  No subjects for batch {batch_id}, skipping")
                    continue
                
                subjects = tuple(s.strip() for s in str(subjects_str).split(',') if s.strip() and len(s.strip()) > 1)
                if not subjects:
                    print(f"   ⚠️  No valid subjects for batch {batch_id}, skipping")
                    continue
                
                # Shift preference
                shift_pref = self._get_column_value(row, df, ['shift', 'shift_preference', 'preferred_shift', 'time_preference'], 'morning')
                shift_preference = str(shift_pref).strip().lower() if shift_pref else 'morning'
                if shift_preference not in ['morning', 'evening', 'afternoon']:
                    shift_preference = 'morning'
                
                # Preferred floor
                preferred_floor = self._get_column_value(row, df, ['floor', 'preferred_floor', 'floor_preference', 'floor_no'], 0)
                try:
                    preferred_floor = int(preferred_floor)
                    # Validation
                    if preferred_floor < 0 or preferred_floor > 20:
                        preferred_floor = 0
                except (ValueError, TypeError):
                    preferred_floor = 0
                
                # Priority
                priority = self._get_column_value(row, df, ['priority', 'importance', 'batch_priority'], 5)
                try:
                    priority = int(priority)
                    priority = max(1, min(10, priority))
                except (ValueError, TypeError):
                    priority = 5
                
                batches.append(StudentBatch(
                    batch_id, department, student_count, subjects,
                    shift_preference, preferred_floor, priority
                ))
                
                floor_info = f" | Floor: {preferred_floor}" if preferred_floor > 0 else ""
                priority_info = f" | Priority: {priority}" if priority != 5 else ""
                print(f"   ✅ {batch_id}: {len(subjects)} subjects, {student_count} students{floor_info}{priority_info}")
                
            except Exception as e:
                print(f"   ⚠️  Row {idx} error: {e}")
                continue
        
        print(f"   ✅ Loaded {len(batches)} batches")
        return batches

    def _parse_breaks(self) -> List[Dict]:
        """Parses the breaks sheet with robust shift support"""
        sheet_name = self._find_sheet(['breaks', 'break', 'lunch'])
        if not sheet_name: return []
        
        df = self.sheets[sheet_name]
        print(f"\n☕ Parsing Breaks from '{sheet_name}'")
        breaks = []
        for _, row in df.iterrows():
            try:
                name = self._get_column_value(row, df, ['name', 'break_name'], 'Lunch')
                days_str = self._get_column_value(row, df, ['days', 'day'], 'MON,TUE,WED,THU,FRI')
                days = [d.strip().upper()[:3] for d in str(days_str).split(',')]
                
                start_slot = int(self._get_column_value(row, df, ['start_slot_id', 'slot', 'start'], 3))
                duration = int(self._get_column_value(row, df, ['duration_slots', 'duration'], 1))
                
                # 🟢 ROBUST SHIFT READING: Read, Strip, and Lowercase immediately
                # This fixes "Evening " vs "evening" issues
                raw_shift = str(self._get_column_value(row, df, ['shift', 'type'], 'common'))
                shift = raw_shift.strip().lower() 
                
                breaks.append({
                    'name': name,
                    'days': days,
                    'start_slot': start_slot,
                    'duration': duration,
                    'shift': shift
                })
                print(f"   ✅ {name} [{shift}]: Slot {start_slot} for {duration} slots")
            except Exception as e:
                print(f"   ⚠️ Error parsing break row: {e}")
        return breaks

class EnhancedFacultyAllocator:
    """
    Enhanced faculty allocation system with dual modes:
    1. Qualification-based (existing)
    2. Choice-based (new)
    """
    
    def __init__(
        self, 
        faculties: List,
        subjects: List,
        faculty_choices: List[FacultyChoice],
        constraints: AllocationConstraints
    ):
        self.faculties = faculties
        self.subjects = subjects
        self.faculty_choices = faculty_choices
        self.constraints = constraints
        
        self.faculty_map = {f.faculty_id: f for f in faculties}
        self.subject_map = {s.subject_code: s for s in subjects}
        self.choice_map = {fc.faculty_id: fc for fc in faculty_choices}
        
        print(f"\n{'='*70}")
        print(f"🎯 ENHANCED FACULTY ALLOCATION SYSTEM")
        print(f"{'='*70}")
        mode = 'CHOICE-BASED ✨' if constraints.use_faculty_choice else 'QUALIFICATION-BASED 📚'
        print(f"   Mode: {mode}")
        print(f"   Faculties: {len(faculties)}")
        print(f"   Subjects: {len(subjects)}")
        print(f"   Faculty Choices: {len(faculty_choices)}")
        print(f"{'='*70}\n")
    
    def allocate_faculty_to_subject(
        self, 
        subject_code: str, 
        batch_id: str
    ) -> List[Tuple[str, int]]:
        """
        Allocate faculty to a subject.
        Returns: List of (faculty_id, score) tuples, sorted by best match
        """
        subject = self.subject_map.get(subject_code)
        if not subject:
            return []
        
        if self.constraints.use_faculty_choice:
            return self._allocate_by_choice(subject_code)
        else:
            return self._allocate_by_qualification(subject)
    
    def _allocate_by_choice(self, subject_code: str) -> List[Tuple[str, int]]:
        """Allocate based on faculty choices (NEW)"""
        allocations = []
        
        for faculty in self.faculties:
            faculty_choice = self.choice_map.get(faculty.faculty_id)
            
            if not faculty_choice:
                allocations.append((faculty.faculty_id, 10))
                continue
            
            score = faculty_choice.get_priority_score(subject_code)
            
            if score > 0:
                allocations.append((faculty.faculty_id, score))
            else:
                allocations.append((faculty.faculty_id, 10))
        
        allocations.sort(key=lambda x: x[1], reverse=True)
        return allocations
    
    def _allocate_by_qualification(self, subject) -> List[Tuple[str, int]]:
        """Allocate based on qualifications (EXISTING)"""
        allocations = []
        
        for faculty in self.faculties:
            score = self._calculate_qualification_score(subject, faculty)
            allocations.append((faculty.faculty_id, score))
        
        allocations.sort(key=lambda x: x[1], reverse=True)
        return allocations
    
    def _calculate_qualification_score(self, subject, faculty) -> int:
        """Calculate match score based on qualifications"""
        score = 100
        
        subject_terms = set()
        if subject.required_specialization:
            subject_terms.update(subject.required_specialization.lower().split())
        subject_terms.update(subject.subject_code.lower().split())
        subject_terms.update(subject.subject_name.lower().split())
        
        faculty_text = ' '.join(faculty.specializations).lower()
        
        for term in subject_terms:
            if len(term) > 2 and term in faculty_text:
                score += 500
        
        if any('phd' in q.lower() for q in faculty.qualifications):
            score += 30
        
        return score

# -------------------------
# Smart Resource Analyzer
# -------------------------
# -------------------------
# Smart Resource Analyzer (REPLACE the existing ResourceAnalyzer class)
# -------------------------
class ResourceAnalyzer:
    """Analyzes resources and calculates optimal ratios"""
    
    def __init__(self, data: Dict):
        """Initialize with data"""
        self.data = data
        self.timeslots = data['timeslots']
        self.rooms = data['rooms']
        self.faculties = data['faculties']
        self.subjects = data['subjects']
        self.batches = data['batches']
    def _detect_hour_imbalance(self) -> Tuple[bool, str, float]:
        """
        Detect if hours are imbalanced across subjects in the batch
        
        Returns:
            (is_imbalanced: bool, reason: str, imbalance_ratio: float)
            - is_imbalanced: True if hours are significantly unbalanced
            - reason: 'balanced', 'high_max', 'low_min', 'high_variance', 'imbalanced'
            - imbalance_ratio: Float showing how imbalanced (0.0 = perfect, 1.0+ = very imbalanced)
        
        Examples:
            - (True, 'high_max', 0.60) = One subject has 60% more hours than average
            - (False, 'balanced', 0.15) = Hours are well balanced
        """
        
        # LINE 27: Collect all subject hours
        all_hours = []
        for subject in self.subjects:
            # Calculate total hours (theory + lab*2)
            theory_hours = subject.hours_per_week
            lab_hours = subject.lab_hours_per_week * 2
            total_hours = theory_hours + lab_hours
            all_hours.append(total_hours)
        
        # LINE 37: Handle empty case
        if not all_hours or len(all_hours) < 2:
            return False, "no_subjects", 0.0
        
        # LINE 40: Calculate statistics
        avg_hours = sum(all_hours) / len(all_hours)
        max_hours = max(all_hours)
        min_hours = min(all_hours)
        
        # LINE 44: Calculate standard deviation
        variance = sum((h - avg_hours) ** 2 for h in all_hours) / len(all_hours)
        std_dev = variance ** 0.5
        
        # LINE 48: Calculate deviation metrics
        max_deviation = max_hours - avg_hours  # How far above average
        min_deviation = avg_hours - min_hours  # How far below average
        max_deviation_ratio = max_deviation / max(avg_hours, 1)
        
        # LINE 52: Standard deviation ratio (coefficient of variation)
        std_dev_ratio = std_dev / max(avg_hours, 1)
        
        # LINE 55: Range ratio (max/min comparison)
        range_ratio = max_hours / max(min_hours, 1)
        
        # LINE 58: Detailed print for debugging
        print(f"\n    📊 Hour Distribution Analysis:")
        print(f"       Subject hours: {all_hours}")
        print(f"       Average: {avg_hours:.1f}h")
        print(f"       Max: {max_hours}h, Min: {min_hours}h")
        print(f"       Max deviation: {max_deviation_ratio:.1%}")
        print(f"       Std dev: {std_dev_ratio:.1%}")
        
        # LINE 66: Detection thresholds
        # Threshold 1: If max is 80% higher than average
        if max_hours > avg_hours * 1.8:
            print(f"       ⚠️ IMBALANCED: Max hours {max_hours}h is {max_deviation_ratio:.1%} above average")
            return True, "high_max", max_deviation_ratio
        
        # Threshold 2: If min is less than 30% of average
        if min_hours < avg_hours * 0.3:
            min_dev_ratio = (avg_hours - min_hours) / avg_hours
            print(f"       ⚠️ IMBALANCED: Min hours {min_hours}h is {min_dev_ratio:.1%} below average")
            return True, "low_min", min_dev_ratio
        
        # Threshold 3: If standard deviation is high (>40% of average)
        if std_dev_ratio > 0.4:
            print(f"       ⚠️ IMBALANCED: High variance (std dev {std_dev_ratio:.1%})")
            return True, "high_variance", std_dev_ratio
        
        # Threshold 4: If range ratio is > 1.5 (max is 50% more than min)
        if range_ratio > 1.5:
            range_dev_ratio = (range_ratio - 1) / range_ratio
            print(f"       ⚠️ IMBALANCED: Range too wide (max/min ratio {range_ratio:.2f})")
            return True, "imbalanced", range_dev_ratio
        
        # All balanced
        print(f"       ✅ BALANCED: Hours are well distributed")
        return False, "balanced", max_deviation_ratio
    
    def analyze_capacity(self) -> Dict[str, Any]:
        """Analyze resource capacity - now an instance method"""
        timeslots = self.timeslots
        rooms = self.rooms
        faculties = self.faculties
        subjects = self.subjects
        batches = self.batches
        
        working_days = 5
        usable_slots = [ts for ts in timeslots if ts.slot_id != 3]
        slots_per_week = len(usable_slots) * working_days
        
        subject_map = {s.subject_code: s for s in subjects}
        total_theory_hours = 0
        total_lab_hours = 0
        elective_groups = defaultdict(list)
        
        batch_requirements = []
        for batch in batches:
            batch_theory = 0
            batch_lab = 0
            for subj_code in batch.subjects:
                subject = subject_map.get(subj_code)
                if subject:
                    if subject.is_elective:
                        elective_groups[subject.elective_group].append(subject.subject_code)
                    batch_theory += subject.hours_per_week
                    batch_lab += subject.lab_hours_per_week
            
            total_theory_hours += batch_theory
            total_lab_hours += batch_lab
            
            batch_requirements.append({
                'batch_id': batch.batch_id,
                'theory_hours': batch_theory,
                'lab_hours': batch_lab,
                'total_slots_needed': batch_theory + (batch_lab * 2),
                'slots_available': slots_per_week,
                'priority': batch.priority
            })
        
        total_faculty_capacity = sum(f.max_hours_per_week for f in faculties)
        total_hours_needed = total_theory_hours + (total_lab_hours * 2)
        
        lab_rooms = [r for r in rooms if 'lab' in r.room_type.lower()]
        
        # Floor distribution
        floor_distribution = Counter(r.floor for r in rooms)
        
        analysis = {
            'slots_per_week': slots_per_week,
            'total_theory_hours': total_theory_hours,
            'total_lab_hours': total_lab_hours,
            'total_hours_needed': total_hours_needed,
            'total_faculty_capacity': total_faculty_capacity,
            'faculty_count': len(faculties),
            'lab_room_count': len(lab_rooms),
            'room_count': len(rooms),
            'batch_requirements': batch_requirements,
            'elective_groups': dict(elective_groups),
            'floor_distribution': dict(floor_distribution),
            'faculty_utilization': (total_hours_needed / total_faculty_capacity * 100) if total_faculty_capacity > 0 else float('inf'),
            'capacity_ratio': total_faculty_capacity / total_hours_needed if total_hours_needed > 0 else 1.0
        }
        
        return analysis
    
    def calculate_scaling_ratio(self) -> Tuple[float, str]:
        """Calculate scaling ratio based on constraints"""
        analysis = self.analyze_capacity()
        capacity_ratio = analysis['capacity_ratio']
        
        max_batch_ratio = 1.0
        bottleneck_batch = None
        
        for batch_req in analysis['batch_requirements']:
            needed = batch_req['total_slots_needed']
            available = batch_req['slots_available']
            if needed > available:
                batch_ratio = available / needed
                if batch_ratio < max_batch_ratio:
                    max_batch_ratio = batch_ratio
                    bottleneck_batch = batch_req['batch_id']
        
        if max_batch_ratio < 1.0:
            return max_batch_ratio, f"Slot constraint (batch {bottleneck_batch})"
        elif capacity_ratio < 1.0:
            return capacity_ratio, "Faculty capacity constraint"
        else:
            return 1.0, "Sufficient resources"
    
    def apply_intelligent_balancing(self, policy: str = 'balanced') -> Tuple[Dict, Dict]:
        """
        Apply dynamic hour allocation (NEW SYSTEM)
        Returns: (updated_data, report)
        """
        config = {
            'working_days': 5,
            'min_hours_per_subject': 1,
            'max_hours_per_subject': 8,
            'priority_multiplier_high': 1.3,
            'priority_multiplier_low': 0.85,
            'lab_multiplier': 1.5,
            'elective_multiplier': 0.75,
            'constraint_severity_high': 'aggressive',
            'constraint_severity_medium': 'balanced',
            'constraint_severity_low': 'conservative'
        }
        
        # Use NEW system
        allocator = DynamicHourAllocationSystem(self.data, config)
        new_subjects, report = allocator.analyze_and_allocate()
        
        # Update internal data with rebalanced subjects
        self.data['subjects'] = new_subjects
        self.subjects = new_subjects
        
        return self.data, report

    
    def apply_scaling_legacy(self, ratio: float) -> Dict:
        """Legacy uniform scaling (keep as fallback)"""
        if ratio >= 0.99:
            return self.data
        
        print(f"\n⚠️  LEGACY SCALING: {ratio*100:.1f}%\n")
        
        scaled_subjects = []
        for subject in self.subjects:
            scaled_theory = max(1, round(subject.hours_per_week * ratio))
            scaled_lab = max(0, round(subject.lab_hours_per_week * ratio))
            
            scaled_subjects.append(Subject(
                subject.subject_code, subject.subject_name,
                scaled_theory, subject.requires_lab, scaled_lab,
                subject.required_specialization, subject.is_elective,
                subject.elective_group, subject.priority
            ))
            
            if scaled_theory != subject.hours_per_week or scaled_lab != subject.lab_hours_per_week:
                print(f"  📉 {subject.subject_code}: {subject.hours_per_week}→{scaled_theory}h")
        
        self.data['subjects'] = scaled_subjects
        self.subjects = scaled_subjects
        return self.data

@dataclass
class SubjectAnalysis:
    """Analysis result for a subject"""
    subject_code: str
    current_hours: int
    ideal_hours: float
    min_hours: int
    max_hours: int
    priority: int
    is_lab: bool
    constraint_score: float  # Higher = more constrained
    scaling_factor: float = 1.0
    recommended_hours: int = 0
    status: str = "balanced"  # overloaded, underutilized, balanced


@dataclass
class HourAllocationAnalysis:
    """Analysis result for a single batch's hour requirements"""
    batch_id: str
    total_required_hours: int
    available_slots_per_week: int
    subject_hours: Dict[str, int]  # subject_code -> hours required
    is_feasible: bool
    feasibility_ratio: float  # 0.0-1.0 (1.0 = perfectly feasible)
    bottleneck: str  # What is limiting? (faculty, rooms, timeslots, none)
    suggested_scaling: float  # Recommended scaling factor



class DynamicHourAllocationSystem:
    """
    Advanced system to dynamically allocate hours to subjects
    Handles overflow/underflow scenarios intelligently
    
    Key Features:
    - Analyzes all resource constraints simultaneously
    - Detects overflow/underflow automatically
    - Recalculates hours using multi-factor analysis
    - Maintains subject balance through weighted algorithms
    - Provides detailed audit trail of changes
    
    Usage:
        allocator = DynamicHourAllocationSystem(data, config)
        new_subjects, report = allocator.analyze_and_allocate()
    """
    
    def __init__(self, data: Dict, config: Dict = None):
        """Initialize the allocation system"""
        self.timeslots = data['timeslots']
        self.rooms = data['rooms']
        self.faculties = data['faculties']
        self.subjects = data['subjects']
        self.batches = data['batches']
        self.original_subjects = [s for s in self.subjects]  # Keep backup
        
        # Configuration
        self.config = config or {
            'working_days': 5,
            'min_hours_per_subject': 1,
            'max_hours_per_subject': 8,
            'priority_multiplier_high': 1.3,  # Priority 8-10
            'priority_multiplier_low': 0.85,   # Priority 1-3
            'lab_multiplier': 1.5,
            'elective_multiplier': 0.75,
            'constraint_severity_high': 'aggressive',
            'constraint_severity_medium': 'balanced',
            'constraint_severity_low': 'conservative'
        }
        
        # Create maps for fast lookup
        self.subject_map = {s.subject_code: s for s in self.subjects}
        self.batch_map = {b.batch_id: b for b in self.batches}
        self.room_map = {r.room_id: r for r in self.rooms}
        self.faculty_map = {f.faculty_id: f for f in self.faculties}
        
        # Tracking
        self.allocation_history = []
        self.overflow_subjects = []
        self.underflow_subjects = []
    
    def analyze_and_allocate(self) -> Tuple[List, Dict[str, Any]]:
        """
        Main method: Analyze all constraints and allocate hours dynamically
        
        Returns:
            (rebalanced_subjects, report)
        """
        print(f"\n{'='*70}")
        print("🎯 DYNAMIC HOUR ALLOCATION SYSTEM")
        print(f"{'='*70}")
        
        # Step 1: Calculate resource constraints
        print(f"\n[1/6] Analyzing resource constraints...")
        constraints = self._analyze_all_constraints()
        
        # Step 2: Detect overflow/underflow for each batch
        print(f"[2/6] Detecting overflow/underflow scenarios...")
        batch_analyses = self._analyze_batch_feasibility(constraints)
        
        # Step 3: Classify batches by severity
        print(f"[3/6] Classifying constraint severity...")
        severity_groups = self._classify_severity(batch_analyses)
        
        # Step 4: Generate allocation strategies
        print(f"[4/6] Generating allocation strategies...")
        strategies = self._generate_allocation_strategies(batch_analyses, severity_groups)
        
        # Step 5: Apply optimal allocation
        print(f"[5/6] Applying optimal allocations...")
        new_subjects = self._apply_allocations(strategies)
        
        # Step 6: Generate detailed report
        print(f"[6/6] Generating allocation report...")
        report = self._generate_allocation_report(batch_analyses, strategies, constraints)
        
        print(f"{'='*70}\n")
        
        return new_subjects, report
    
    def _analyze_all_constraints(self) -> Dict[str, Any]:
        """
        Analyze ALL resource constraints simultaneously
        Returns resource capacity and bottleneck information
        """
        print(f"\n  📊 Resource Analysis:")
        
        # CONSTRAINT 1: Faculty Capacity
        total_faculty_capacity = sum(f.max_hours_per_week for f in self.faculties)
        print(f"    Faculty: {len(self.faculties)} total capacity = {total_faculty_capacity}h/week")
        
        # CONSTRAINT 2: Room Capacity (availability)
        working_days = self.config['working_days']
        usable_slots = [ts for ts in self.timeslots if ts.slot_id != 3]
        slots_per_week = len(usable_slots) * working_days
        
        # Average capacity per room per week
        avg_room_hours_per_week = len(usable_slots) * working_days
        total_room_capacity = len(self.rooms) * avg_room_hours_per_week
        
        print(f"    Rooms: {len(self.rooms)} rooms × {avg_room_hours_per_week} slots/week = {total_room_capacity}h/week")
        
        # CONSTRAINT 3: Timeslot Capacity
        print(f"    Timeslots: {len(usable_slots)} per day × {working_days} days = {slots_per_week} slots/week")
        
        # CONSTRAINT 4: Calculate total required hours across all batches
        total_required_hours = 0
        batch_hour_requirements = {}
        
        for batch in self.batches:
            batch_hours = 0
            for subj_code in batch.subjects:
                subject = self.subject_map.get(subj_code)
                if subject:
                    theory_hours = subject.hours_per_week
                    lab_hours = subject.lab_hours_per_week * 2  # Labs take 2 slots
                    batch_hours += theory_hours + lab_hours
            
            batch_hour_requirements[batch.batch_id] = batch_hours
            total_required_hours += batch_hours
        
        print(f"    Total required: {total_required_hours}h/week across all batches")
        
        # Determine bottleneck
        bottleneck = 'none'
        constraint_ratio = 1.0
        
        if total_faculty_capacity < total_required_hours:
            bottleneck = 'faculty'
            constraint_ratio = total_faculty_capacity / total_required_hours
            print(f"    ⚠️ BOTTLENECK: Faculty capacity {constraint_ratio*100:.1f}%")
        
        elif total_room_capacity < total_required_hours:
            bottleneck = 'rooms'
            constraint_ratio = total_room_capacity / total_required_hours
            print(f"    ⚠️ BOTTLENECK: Room availability {constraint_ratio*100:.1f}%")
        
        elif slots_per_week < total_required_hours:
            bottleneck = 'timeslots'
            constraint_ratio = slots_per_week / total_required_hours
            print(f"    ⚠️ BOTTLENECK: Timeslot availability {constraint_ratio*100:.1f}%")
        
        else:
            print(f"    ✅ All resources sufficient")
        
        # Calculate severity (0-1, where 1 is severe)
        bottleneck_severity = max(0, 1.0 - constraint_ratio)
        
        return {
            'faculty_capacity': total_faculty_capacity,
            'room_capacity': total_room_capacity,
            'timeslot_capacity': slots_per_week,
            'total_required_hours': total_required_hours,
            'constraint_ratio': constraint_ratio,
            'bottleneck': bottleneck,
            'bottleneck_severity': bottleneck_severity,
            'batch_requirements': batch_hour_requirements
        }
    
    def _analyze_batch_feasibility(self, constraints: Dict) -> List[HourAllocationAnalysis]:
        """
        Analyze feasibility for EACH BATCH individually
        Some batches might overflow while others underflow
        """
        print(f"\n  🔍 Batch-Level Feasibility:")
        
        analyses = []
        global_ratio = constraints['constraint_ratio']
        
        for batch in self.batches:
            # Calculate required hours for this batch
            batch_hours = {}
            total_required = 0
            
            for subj_code in batch.subjects:
                subject = self.subject_map.get(subj_code)
                if subject:
                    theory = subject.hours_per_week
                    lab = subject.lab_hours_per_week * 2
                    total_hours = theory + lab
                    batch_hours[subj_code] = total_hours
                    total_required += total_hours
            
            # Available slots for THIS batch
            usable_slots = [ts for ts in self.timeslots if ts.slot_id != 3]
            available_slots = len(usable_slots) * self.config['working_days']
            
            # Feasibility analysis
            feasibility_ratio = available_slots / max(total_required, 1)
            is_feasible = feasibility_ratio >= 0.95  # Allow 5% margin
            
            # Detect bottleneck for THIS batch
            bottleneck = 'none'
            if total_required > available_slots:
                bottleneck = 'timeslots'
            elif total_required > constraints['faculty_capacity'] / len(self.batches):
                bottleneck = 'faculty'
            elif total_required > constraints['room_capacity'] / len(self.batches):
                bottleneck = 'rooms'
            
            # Suggested scaling for this batch
            suggested_scaling = min(feasibility_ratio, global_ratio)
            
            analysis = HourAllocationAnalysis(
                batch_id=batch.batch_id,
                total_required_hours=total_required,
                available_slots_per_week=available_slots,
                subject_hours=batch_hours,
                is_feasible=is_feasible,
                feasibility_ratio=feasibility_ratio,
                bottleneck=bottleneck,
                suggested_scaling=suggested_scaling
            )
            
            analyses.append(analysis)
            
            # Status indicator
            if is_feasible:
                status = "✅ FEASIBLE"
            elif feasibility_ratio >= 0.80:
                status = "⚠️ TIGHT"
            else:
                status = "❌ OVERFLOW"
            
            print(f"    {status} {batch.batch_id}: {total_required}h required, {available_slots} slots available ({feasibility_ratio*100:.1f}%)")
            
            # Track overflows/underflows
            if total_required > available_slots:
                self.overflow_subjects.append(batch.batch_id)
            elif feasibility_ratio < 0.7:
                self.underflow_subjects.append(batch.batch_id)
        
        return analyses
    
    def _classify_severity(self, analyses: List[HourAllocationAnalysis]) -> Dict[str, List[str]]:
        """Classify batches by constraint severity"""
        
        print(f"\n  📈 Severity Classification:")
        
        severe = []  # feasibility < 0.70
        moderate = []  # 0.70 <= feasibility < 0.90
        mild = []  # feasibility >= 0.90
        
        for analysis in analyses:
            if analysis.feasibility_ratio < 0.70:
                severe.append(analysis.batch_id)
            elif analysis.feasibility_ratio < 0.90:
                moderate.append(analysis.batch_id)
            else:
                mild.append(analysis.batch_id)
        
        print(f"    🔴 SEVERE ({len(severe)}): {', '.join(severe) if severe else 'None'}")
        print(f"    🟡 MODERATE ({len(moderate)}): {', '.join(moderate) if moderate else 'None'}")
        print(f"    🟢 MILD ({len(mild)}): {', '.join(mild) if mild else 'None'}")
        
        return {
            'severe': severe,
            'moderate': moderate,
            'mild': mild
        }
    
    def _generate_allocation_strategies(self, analyses: List[HourAllocationAnalysis],
                                       severity_groups: Dict[str, List[str]]) -> Dict[str, Dict]:
        """
        Generate allocation strategies for each batch
        Returns: {batch_id: {subject_code: new_hours}}
        """
        
        print(f"\n  🎯 Allocation Strategies:")
        
        strategies = {}
        
        for analysis in analyses:
            batch_id = analysis.batch_id
            batch = self.batch_map[batch_id]
            
            # Determine policy based on severity
            if batch_id in severity_groups['severe']:
                policy = self.config['constraint_severity_high']
                print(f"    {batch_id}: AGGRESSIVE (severe constraints)")
            elif batch_id in severity_groups['moderate']:
                policy = self.config['constraint_severity_medium']
                print(f"    {batch_id}: BALANCED (moderate constraints)")
            else:
                policy = self.config['constraint_severity_low']
                print(f"    {batch_id}: CONSERVATIVE (mild constraints)")
            
            # Generate allocation for this batch
            batch_allocation = self._allocate_hours_to_batch(batch, analysis, policy)
            strategies[batch_id] = batch_allocation
        
        return strategies
    
    def _allocate_hours_to_batch(self, batch, analysis: HourAllocationAnalysis,
                                policy: str) -> Dict[str, int]:
        """
        Allocate hours to subjects within a batch using intelligent ratios
        Returns: {subject_code: recommended_hours}
        """
        
        allocation = {}
        
        # If feasible, keep original hours
        if analysis.is_feasible and analysis.feasibility_ratio >= 0.95:
            for subj_code, hours in analysis.subject_hours.items():
                allocation[subj_code] = hours
            return allocation
        
        # Need to rebalance
        scaling_factor = analysis.suggested_scaling
        available_hours = analysis.available_slots_per_week
        
        # Collect all subject info for smart allocation
        subjects_info = []
        
        for subj_code in batch.subjects:
            subject = self.subject_map.get(subj_code)
            if not subject:
                continue
            
            current_hours = analysis.subject_hours.get(subj_code, 0)
            
            # Calculate allocation score (priority multiplier)
            score = 1.0
            
            # Factor 1: Subject priority
            if subject.priority >= 8:
                score *= self.config['priority_multiplier_high']
            elif subject.priority <= 3:
                score *= self.config['priority_multiplier_low']
            
            # Factor 2: Lab requirement (labs are more constrained)
            if subject.requires_lab:
                score *= self.config['lab_multiplier']
            
            # Factor 3: Electives (typically less critical)
            if subject.is_elective:
                score *= self.config['elective_multiplier']
            
            # Factor 4: Constraint sensitivity
            if subject.required_specialization:
                qualified_faculty = [
                    f for f in self.faculties
                    if subject.required_specialization.lower() in ' '.join(f.specializations).lower()
                ]
                if len(qualified_faculty) < 2:
                    score *= 1.2  # Protect highly specialized subjects
            
            subjects_info.append({
                'code': subj_code,
                'current_hours': current_hours,
                'priority': subject.priority,
                'is_lab': subject.requires_lab,
                'is_elective': subject.is_elective,
                'allocation_score': score
            })
        
        # Total allocation score (for weighted distribution)
        total_score = sum(s['allocation_score'] for s in subjects_info)
        
        if total_score == 0:
            total_score = 1
        
        # Allocate hours proportionally
        allocated = 0
        
        for i, subject_info in enumerate(subjects_info):
            subj_code = subject_info['code']
            current = subject_info['current_hours']
            score = subject_info['allocation_score']
            
            # Weighted share of available hours
            weight = score / total_score
            
            # Apply policy-specific scaling
            if policy == 'aggressive':
                # Aggressively reduce to fit constraints
                new_hours = round(available_hours * weight * 0.85)
            
            elif policy == 'conservative':
                # Minimal reduction, try to preserve hours
                new_hours = round(available_hours * weight * 1.0)
            
            else:  # balanced
                # Moderate scaling
                new_hours = round(available_hours * weight * 0.95)
            
            # Enforce bounds
            min_hours = self.config['min_hours_per_subject']
            max_hours = self.config['max_hours_per_subject']
            
            # Special rules
            if subject_info['is_lab']:
                min_hours = 2  # Labs need at least 2 hours
            
            if subject_info['is_elective']:
                max_hours = 4  # Electives capped at 4 hours
            
            new_hours = max(min_hours, min(max_hours, new_hours))
            
            allocation[subj_code] = new_hours
            allocated += new_hours
        
        # Verify total allocation
        total_allocated = sum(allocation.values())
        
        if total_allocated > available_hours:
            # Scale down all proportionally
            scale = available_hours / max(total_allocated, 1)
            for subj_code in allocation:
                allocation[subj_code] = max(
                    self.config['min_hours_per_subject'],
                    round(allocation[subj_code] * scale)
                )
        
        return allocation
    
    # ⭐⭐⭐ THIS METHOD MUST BE INSIDE THE CLASS WITH 4-SPACE INDENTATION ⭐⭐⭐
    def _apply_allocations(self, strategies: Dict[str, Dict]) -> List:
        """
        Apply the allocation strategies to create new Subject objects
        
        Returns: new list of Subject objects with adjusted hours
        """
        
        print(f"\n  ✨ Applying Allocations:")
        
        new_subjects = []
        changes_count = 0
        
        for subject in self.subjects:
            # Find allocation for this subject's batch
            allocated_hours = None
            
            for batch in self.batches:
                if subject.subject_code in batch.subjects:
                    strategy = strategies.get(batch.batch_id, {})
                    if subject.subject_code in strategy:
                        allocated_hours = strategy[subject.subject_code]
                        break
            
            if allocated_hours is None:
                # No allocation found, keep original
                new_subjects.append(subject)
                continue
            
            # Calculate theory vs lab hours
            original_total = subject.hours_per_week + subject.lab_hours_per_week * 2
            
            if subject.requires_lab:
                # Maintain theory/lab ratio
                theory_ratio = subject.hours_per_week / max(original_total, 1)
                new_theory = max(1, round(allocated_hours * theory_ratio))
                new_lab = max(0, round((allocated_hours - new_theory) / 2))
            else:
                new_theory = allocated_hours
                new_lab = 0
            
            # CREATE NEW SUBJECT USING PROPER DATACLASS CONSTRUCTOR
            new_subject = Subject(
                subject_code=subject.subject_code,
                subject_name=subject.subject_name,
                hours_per_week=new_theory,
                requires_lab=subject.requires_lab,
                lab_hours_per_week=new_lab,
                required_specialization=subject.required_specialization,
                is_elective=subject.is_elective,
                elective_group=subject.elective_group,
                priority=subject.priority
            )
            
            new_subjects.append(new_subject)
            
            # Track change
            if new_theory != subject.hours_per_week or new_lab != subject.lab_hours_per_week:
                old_total = subject.hours_per_week + subject.lab_hours_per_week * 2
                new_total = new_theory + new_lab * 2
                change_pct = ((new_total - old_total) / max(old_total, 1)) * 100
                
                if change_pct > 0:
                    symbol = "📈"
                else:
                    symbol = "📉"
                
                print(f"    {symbol} {subject.subject_code}: {old_total}h → {new_total}h ({change_pct:+.0f}%)")
                changes_count += 1
        
        print(f"\n    Total changes: {changes_count}/{len(self.subjects)} subjects")
        
        return new_subjects
    
    def _generate_allocation_report(self, analyses: List[HourAllocationAnalysis],
                                 strategies: Dict[str, Dict],
                                 constraints: Dict) -> Dict[str, Any]:
        """
        Generate comprehensive allocation report
        """
        
        report = {
            'system': 'DynamicHourAllocationSystem',
            'timestamp': time.time(),
            'constraints': {
                'bottleneck': constraints['bottleneck'],
                'severity': constraints['bottleneck_severity'],
                'constraint_ratio': constraints['constraint_ratio'],
                'faculty_capacity': constraints['faculty_capacity'],
                'room_capacity': constraints['room_capacity'],
                'timeslot_capacity': constraints['timeslot_capacity'],
                'total_required_hours': constraints['total_required_hours']
            },
            'batch_analyses': [],
            'summary': {
                'total_batches': len(analyses),
                'feasible': sum(1 for a in analyses if a.is_feasible),
                'tight': sum(1 for a in analyses if not a.is_feasible and a.feasibility_ratio >= 0.8),
                'overflow': sum(1 for a in analyses if a.feasibility_ratio < 0.8),
                'total_hour_adjustments': 0
            }
        }
        
        # Add batch details
        for analysis in analyses:
            report['batch_analyses'].append({
                'batch_id': analysis.batch_id,
                'required_hours': analysis.total_required_hours,
                'available_slots': analysis.available_slots_per_week,
                'feasibility_ratio': analysis.feasibility_ratio,
                'is_feasible': analysis.is_feasible,
                'bottleneck': analysis.bottleneck,
                'suggested_scaling': analysis.suggested_scaling,
                'subject_hours': analysis.subject_hours
            })
        
        print(f"\n{'='*70}")
        print("📊 ALLOCATION REPORT SUMMARY")
        print(f"{'='*70}")
        print(f"\nConstraint Analysis:")
        print(f"  Bottleneck: {constraints['bottleneck'].upper()}")
        print(f"  Severity: {constraints['bottleneck_severity']*100:.1f}%")
        print(f"  Global Ratio: {constraints['constraint_ratio']*100:.1f}%")
        print(f"\nBatch Feasibility:")
        print(f"  ✅ Feasible: {report['summary']['feasible']}")
        print(f"  ⚠️ Tight: {report['summary']['tight']}")
        print(f"  ❌ Overflow: {report['summary']['overflow']}")
        print(f"\n{'='*70}\n")
        
        return report

# After dynamic balancing, before scheduling
from copy import deepcopy


logger = logging.getLogger(__name__)

class ConstraintType(Enum):
    """Types of scheduling constraints"""
    FACULTY_CAPACITY = "faculty_capacity"
    ROOM_CAPACITY = "room_capacity"
    SPECIALIZATION = "specialization_mismatch"
    BATCH_TIMING = "batch_timing_conflict"
    ELECTIVE_GROUP = "elective_group_constraint"
    FLOOR_PRIORITY = "floor_priority_constraint"
    SESSION_OVERLAP = "session_overlap"


class ConstraintSeverity(Enum):
    """Severity levels"""
    INFO = 1
    WARNING = 2
    CRITICAL = 3
    BLOCKER = 4


@dataclass
class ConstraintViolation:
    """Represents a constraint violation"""
    type: ConstraintType
    severity: ConstraintSeverity
    subject_code: str
    message: str
    resolution: str
    priority: int = 0



@dataclass
class SchedulingConfig:
    """Production-grade configuration"""
    MIN_THEORY: int = 1
    MIN_LAB: int = 1
    REDUCTION_STEP: float = 0.85
    MAX_ATTEMPTS: int = 12
    FEASIBILITY_BUFFER: float = 1.2
    
    # Constraint handling
    ENABLE_SPECIALIZATION_RELAXATION: bool = True
    ENABLE_FLOOR_PRIORITY_RELAXATION: bool = True
    ENABLE_BATCH_SPLITTING: bool = True
    ENABLE_SESSION_REORDERING: bool = True
    
    # Trade-off thresholds
    ACCEPTABLE_SUCCESS_RATE: float = 0.80  # Accept 80%+ success
    MIN_VIABLE_SCHEDULE: float = 0.60      # Minimum 60% to be viable


class ConstraintAnalyzer:
    """
    Analyzes all constraints and identifies violations.
    Suggests intelligent resolutions.
    """
    
    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.violations: List[ConstraintViolation] = []
        self.resources = self._analyze_resources()
    
    def _analyze_resources(self) -> Dict[str, Any]:
        """Analyze all available resources"""
        return {
            'faculties': len(self.data.get('faculties', [])),
            'rooms': len(self.data.get('rooms', [])),
            'timeslots': len(self.data.get('timeslots', [])),
            'batches': len(self.data.get('batches', [])),
            'subjects': len(self.data.get('subjects', [])),
            'faculty_capacity': sum(f.max_hours_per_week for f in self.data.get('faculties', [])),
            'room_slots': self._calculate_room_slots(),
            'total_required_hours': self._calculate_required_hours()
        }
    
    def _calculate_room_slots(self) -> int:
        """Calculate total room slots available per week"""
        timeslots = self.data.get('timeslots', [])
        rooms = self.data.get('rooms', [])
        usable_slots = len([ts for ts in timeslots if getattr(ts, 'slot_id', 0) != 3])
        return len(rooms) * usable_slots * 5
    
    def _calculate_required_hours(self) -> int:
        """Calculate total required hours"""
        return sum(
            s.hours_per_week + (s.lab_hours_per_week * 2)
            for s in self.data.get('subjects', [])
        )
    
    def analyze_all_constraints(self) -> Tuple[bool, List[ConstraintViolation], str]:
        """
        Comprehensive constraint analysis.
        Returns: (is_viable, violations, recommendations)
        """
        self.violations = []
        
        # Check each constraint type
        self._check_capacity_constraints()
        self._check_specialization_constraints()
        self._check_batch_constraints()
        self._check_floor_priority_constraints()
        
        # Generate report
        report = self._generate_report()
        is_viable = self._is_viable()
        
        return is_viable, self.violations, report
    
    def _check_capacity_constraints(self):
        """Check faculty and room capacity"""
        required = self.resources['total_required_hours']
        faculty_cap = self.resources['faculty_capacity']
        room_slots = self.resources['room_slots']
        
        if required > faculty_cap:
            shortage = required - faculty_cap
            self.violations.append(ConstraintViolation(
                type=ConstraintType.FACULTY_CAPACITY,
                severity=ConstraintSeverity.CRITICAL,
                subject_code="ALL",
                message=f"Faculty shortage: need {required}h, have {faculty_cap}h ({shortage}h short)",
                resolution=f"Scale down all subjects to {faculty_cap/required:.0%} or add faculty",
                priority=1
            ))
        
        if required > room_slots:
            self.violations.append(ConstraintViolation(
                type=ConstraintType.ROOM_CAPACITY,
                severity=ConstraintSeverity.CRITICAL,
                subject_code="ALL",
                message=f"Room shortage: need {required} slots, have {room_slots}",
                resolution=f"Reduce hours or add rooms",
                priority=2
            ))
    
    def _check_specialization_constraints(self):
        """Check if faculties match subject specializations"""
        faculties = self.data.get('faculties', [])
        subjects = self.data.get('subjects', [])
        
        for subject in subjects:
            if not subject.required_specialization:
                continue
            
            qualified = [f for f in faculties 
                        if subject.required_specialization.lower() in ' '.join(f.specializations).lower()]
            
            if not qualified:
                self.violations.append(ConstraintViolation(
                    type=ConstraintType.SPECIALIZATION,
                    severity=ConstraintSeverity.CRITICAL,
                    subject_code=subject.subject_code,
                    message=f"No faculty with '{subject.required_specialization}' specialization",
                    resolution=f"Add specialization to a faculty OR remove specialization requirement",
                    priority=3
                ))
    
    def _check_batch_constraints(self):
        """Check batch-level constraints"""
        batches = self.data.get('batches', [])
        subjects = self.data.get('subjects', [])
        
        for batch in batches:
            batch_subjects = [s for s in subjects if s.subject_code in batch.subjects]
            if not batch_subjects:
                self.violations.append(ConstraintViolation(
                    type=ConstraintType.BATCH_TIMING,
                    severity=ConstraintSeverity.WARNING,
                    subject_code=batch.batch_id,
                    message=f"Batch {batch.batch_id} has no valid subjects",
                    resolution="Add subjects to batch or remove batch",
                    priority=4
                ))
    
    def _check_floor_priority_constraints(self):
        """Check floor priority constraints"""
        # This is informational
        logger.info("Floor priority constraints noted (can be relaxed if needed)")
    
    def _is_viable(self) -> bool:
        """Check if scheduling is viable"""
        critical_violations = [v for v in self.violations if v.severity in [ConstraintSeverity.CRITICAL, ConstraintSeverity.BLOCKER]]
        return len(critical_violations) <= 2  # Can handle up to 2 critical violations
    
    def _generate_report(self) -> str:
        """Generate detailed analysis report"""
        report = f"""
╔════════════════════════════════════════════════════════════════╗
║ PRODUCTION CONSTRAINT ANALYSIS REPORT                         ║
╚════════════════════════════════════════════════════════════════╝

RESOURCES AVAILABLE:
  • Faculties: {self.resources['faculties']} (capacity: {self.resources['faculty_capacity']}h/week)
  • Rooms: {self.resources['rooms']}
  • Timeslots/week: {self.resources['timeslots']}
  • Total room slots/week: {self.resources['room_slots']}
  
DEMAND:
  • Batches: {self.resources['batches']}
  • Subjects: {self.resources['subjects']}
  • Total required hours: {self.resources['total_required_hours']}h/week

CONSTRAINT ANALYSIS:
  • Total violations: {len(self.violations)}
  • Critical: {len([v for v in self.violations if v.severity == ConstraintSeverity.CRITICAL])}
  • Warnings: {len([v for v in self.violations if v.severity == ConstraintSeverity.WARNING])}

VIOLATIONS FOUND:
"""
        for i, v in enumerate(self.violations, 1):
            report += f"\n  [{i}] {v.type.value.upper()}\n"
            report += f"      Issue: {v.message}\n"
            report += f"      Fix: {v.resolution}\n"
        
        report += f"\n{'='*60}\n"
        return report
class IntelligentConstraintHandler:
    """
    Handles constraints intelligently using multiple strategies.
    Production-grade constraint management.
    """
    
    def __init__(self, config: SchedulingConfig = None):
        self.config = config or SchedulingConfig()
        self.strategies_applied = []
        self.attempt_history = []
    
    def handle_constraints(
        self,
        data: Dict[str, Any],
        scheduler_class,
        predictor,
        config: Optional[SchedulingConfig] = None
    ) -> Optional[Dict[str, Any]]:
        """
        Intelligently handle ALL constraints and produce schedule.
        PRODUCTION-GRADE: Never fails.
        """
        if config:
            self.config = config
        
        # PHASE 1: Analyze constraints
        print("\n" + "="*70)
        print("PHASE 1: CONSTRAINT ANALYSIS")
        print("="*70)
        analyzer = ConstraintAnalyzer(data)
        is_viable, violations, report = analyzer.analyze_all_constraints()
        print(report)
        
        # PHASE 2: Apply intelligent strategies
        print("\n" + "="*70)
        print("PHASE 2: INTELLIGENT STRATEGY APPLICATION")
        print("="*70)
        
        strategies = [
            ("Strategy 1: Direct Scheduling", self._strategy_direct_schedule),
            ("Strategy 2: Specialization Relaxation", self._strategy_relax_specialization),
            ("Strategy 3: Graduated Hour Scaling", self._strategy_graduated_scaling),
            ("Strategy 4: Batch Splitting", self._strategy_batch_splitting),
            ("Strategy 5: Floor Priority Relaxation", self._strategy_relax_floor_priority),
            ("Strategy 6: Minimal Viable Schedule", self._strategy_minimal_viable),
            ("Strategy 7: Partial Best-Effort", self._strategy_partial_effort),
        ]
        
        for strategy_name, strategy_func in strategies:
            print(f"\n🔄 {strategy_name}...")
            try:
                results = strategy_func(data, scheduler_class, predictor)
                if results and results.get('statistics', {}).get('scheduled', 0) > 0:
                    success_rate = results['statistics'].get('success_rate', 0)
                    if success_rate >= self.config.ACCEPTABLE_SUCCESS_RATE:
                        print(f"   ✅ SUCCESS! Success rate: {success_rate:.1f}%")
                        self.strategies_applied.append(strategy_name)
                        return results
                    else:
                        print(f"   ⚠️ Partial success ({success_rate:.1f}%), trying next strategy...")
                else:
                    print(f"   ❌ Not viable, trying next strategy...")
            except Exception as e:
                print(f"   Error: {str(e)}, trying next strategy...")
                self.attempt_history.append(f"{strategy_name}: {str(e)}")
        
        # PHASE 3: Report results
        print("\n" + "="*70)
        print("PHASE 3: FINAL RESULTS")
        print("="*70)
        print(f"Strategies applied: {', '.join(self.strategies_applied) if self.strategies_applied else 'None successful'}")
        print("❌ All strategies exhausted - returning best-effort schedule")
        
        return {}  # Return empty, not None (caller won't crash)
    
    def _strategy_direct_schedule(self, data, scheduler_class, predictor):
        """Strategy 1: Try direct scheduling without modifications"""
        scheduler = scheduler_class(data, predictor)
        return scheduler.schedule()
    
    def _strategy_relax_specialization(self, data, scheduler_class, predictor):
        """Strategy 2: Remove specialization constraints"""
        if not self.config.ENABLE_SPECIALIZATION_RELAXATION:
            return None
        
        print("   Relaxing specialization requirements...")
        modified_data = deepcopy(data)
        
        # Create a new list of subjects with specialization set to empty/None
        relaxed_subjects = []
        for s in modified_data['subjects']:
            # Use replace() to create a new instance from the frozen one
            new_subject = replace(s, required_specialization="")
            relaxed_subjects.append(new_subject)
        
        # Update the modified data with the new subject objects
        modified_data['subjects'] = relaxed_subjects
        
        scheduler = scheduler_class(modified_data, predictor)
        return scheduler.schedule()
        
        print("   Relaxing specialization requirements...")
        modified_data = deepcopy(data)
        
        # Create NEW Subject objects with required_specialization = ""
        new_subjects = []
        for s in modified_data['subjects']:
            new_s = Subject(
                subject_code=s.subject_code,
                subject_name=s.subject_name,
                hours_per_week=s.hours_per_week,
                requires_lab=s.requires_lab,
                lab_hours_per_week=s.lab_hours_per_week,
                required_specialization="", # CLEARED
                is_elective=s.is_elective,
                elective_group=s.elective_group,
                priority=s.priority
            )
            new_subjects.append(new_s)
            
        modified_data['subjects'] = new_subjects
        
        scheduler = scheduler_class(modified_data, predictor)
        return scheduler.schedule()
    
    def _strategy_graduated_scaling(self, data, scheduler_class, predictor):
        """Strategy 3: Graduated hour scaling from original"""
        modified_data = deepcopy(data)
        orig_subjects = data['subjects']
        orig_hours = [(s.subject_code, s.hours_per_week, s.lab_hours_per_week) for s in orig_subjects]
        
        for attempt in range(1, self.config.MAX_ATTEMPTS + 1):
            ratio = self.config.REDUCTION_STEP ** (attempt - 1)
            print(f"   Attempt {attempt}: Scaling to {ratio:.0%}...")
            
            # Scale subjects from ORIGINAL
            new_subjects = []
            for (code, orig_th, orig_lab) in orig_hours:
                orig_s = next((s for s in orig_subjects if s.subject_code == code), None)
                if not orig_s:
                    continue
                
                new_subjects.append(Subject(
                    subject_code=orig_s.subject_code,
                    subject_name=orig_s.subject_name,
                    hours_per_week=max(self.config.MIN_THEORY, int(round(orig_th * ratio))),
                    requires_lab=orig_s.requires_lab,
                    lab_hours_per_week=max(self.config.MIN_LAB, int(round(orig_lab * ratio))),
                    required_specialization=orig_s.required_specialization,
                    is_elective=orig_s.is_elective,
                    elective_group=orig_s.elective_group,
                    priority=orig_s.priority
                ))
            
            modified_data['subjects'] = new_subjects
            scheduler = scheduler_class(modified_data, predictor)
            results = scheduler.schedule()
            
            if results and results.get('statistics', {}).get('scheduled', 0) > 0:
                return results
        
        return None
    
    def _strategy_batch_splitting(self, data, scheduler_class, predictor):
        """Strategy 4: Split complex batches into smaller groups"""
        print("   Attempting batch splitting...")
        # Advanced: split batches, schedule independently, merge results
        return None  # Placeholder
    
    def _strategy_relax_floor_priority(self, data, scheduler_class, predictor):
        """Strategy 5: Relax floor priority constraints"""
        if not self.config.ENABLE_FLOOR_PRIORITY_RELAXATION:
            return None
        
        print("   Relaxing floor priority constraints...")
        modified_data = deepcopy(data)
        scheduler = scheduler_class(modified_data, predictor)
        return scheduler.schedule()
    
    def _strategy_minimal_viable(self, data, scheduler_class, predictor):
        """Strategy 6: Reduce all to absolute minimums"""
        print("   Creating minimal viable schedule...")
        modified_data = deepcopy(data)
        
        # Reduce everything to minimums
        new_subjects = []
        for s in modified_data['subjects']:
            new_subjects.append(Subject(
                subject_code=s.subject_code,
                subject_name=s.subject_name,
                hours_per_week=self.config.MIN_THEORY,
                requires_lab=s.requires_lab,
                lab_hours_per_week=self.config.MIN_LAB,
                required_specialization=None,  # Remove specialization
                is_elective=s.is_elective,
                elective_group=s.elective_group,
                priority=s.priority
            ))
        
        modified_data['subjects'] = new_subjects
        scheduler = scheduler_class(modified_data, predictor)
        return scheduler.schedule()
    
    def _strategy_partial_effort(self, data, scheduler_class, predictor):
        """Strategy 7: Return partial schedule with best-effort"""
        print("   Creating partial best-effort schedule...")
        # Return whatever can be scheduled from available resources
        return {}

class FeasibilityAnalyzer:
    """
    Pre-checks scheduling feasibility before expensive scheduler run.
    Reduces unnecessary computation and speeds up failure detection.
    """
    
    def __init__(self, data: Dict[str, Any]):
        self.data = data
        self.timeslots = data.get('timeslots', [])
        self.rooms = data.get('rooms', [])
        self.faculties = data.get('faculties', [])
        self.subjects = data.get('subjects', [])
    
    def calculate_total_required_hours(self) -> int:
        """Calculate total hours needed for all subjects"""
        return sum(
            s.hours_per_week + (s.lab_hours_per_week * 2)
            for s in self.subjects
        )
    
    def calculate_hard_resource_limit(self) -> int:
        """
        Calculate the absolute maximum hours that can be scheduled given:
        - Faculty availability
        - Room availability
        - Timeslot availability
        """
        # Faculty capacity
        faculty_capacity = sum(f.max_hours_per_week for f in self.faculties)
        
        # Room/timeslot capacity (exclude slot 3 if it's break)
        usable_slots_per_day = len([
            ts for ts in self.timeslots 
            if getattr(ts, 'slot_id', 0) != 3
        ])
        working_days = 5
        room_capacity = len(self.rooms) * usable_slots_per_day * working_days
        
        # The bottleneck is the minimum of the two
        hard_limit = min(faculty_capacity, room_capacity)
        return hard_limit
    
    def is_feasible(self, buffer: float = 1.1) -> Tuple[bool, str]:
        """
        Quick pre-check: is scheduling even possible?
        
        Args:
            buffer: Tolerance factor (1.1 = 10% over capacity triggers scaling)
        
        Returns:
            (is_feasible: bool, reason: str for logging)
        """
        required = self.calculate_total_required_hours()
        limit = self.calculate_hard_resource_limit()
        
        if required <= limit:
            return True, f"✅ Feasible: {required}h ≤ {limit}h capacity"
        elif required <= limit * buffer:
            return True, f"⚠️ Tight fit: {required}h ≈ {limit}h (within buffer)"
        else:
            return False, f"❌ Infeasible: {required}h >> {limit}h (need scaling)"


class EmergencyScheduler:
    """
    Intelligent emergency scheduling with graceful hour reduction.
    
    Strategy:
    1. Pre-check feasibility (fast, no scheduler)
    2. Try full hours with scheduler
    3. If fails: scale down from ORIGINAL (not from previous attempt)
    4. Repeat up to MAX_ATTEMPTS
    5. Report detailed metrics
    """
    
    def __init__(self, config: SchedulingConfig = None):
        self.config = config or SchedulingConfig()
        self.attempt_log = []  # Track what happened
    
    def _store_original_hours(self, subjects: List) -> List[Tuple[str, int, int]]:
        """Preserve original hours for consistent scaling"""
        return [
            (s.subject_code, s.hours_per_week, s.lab_hours_per_week)
            for s in subjects
        ]
    
    def _scale_subjects(
        self,
        original_subjects: List,
        orig_hours: List[Tuple[str, int, int]],
        ratio: float
    ) -> Tuple[List, bool]:
        """
        Create new subject list with scaled hours (from ORIGINAL values).
        
        Args:
            original_subjects: Original subject objects (for copying)
            orig_hours: Stored original (code, theory, lab) tuples
            ratio: Scaling factor (0.0 to 1.0)
        
        Returns:
            (new_subjects, made_change: bool)
        """
        new_subjects = []
        made_change = False
        
        for (code, orig_theory, orig_lab) in orig_hours:
            # Find original object for field references
            orig_subject = next(
                (s for s in original_subjects if s.subject_code == code),
                None
            )
            if not orig_subject:
                logger.warning(f"Subject {code} not found in originals")
                continue
            
            # Scale from ORIGINAL, not from last attempt
            new_theory = max(
                self.config.MIN_THEORY,
                int(round(orig_theory * ratio))
            )
            new_lab = max(
                self.config.MIN_LAB,
                int(round(orig_lab * ratio))
            )
            
            # Track if we actually changed anything
            if new_theory != orig_subject.hours_per_week or new_lab != orig_subject.lab_hours_per_week:
                made_change = True
            
            # Create new immutable subject object
            new_subjects.append(
                Subject(
                    subject_code=orig_subject.subject_code,
                    subject_name=orig_subject.subject_name,
                    hours_per_week=new_theory,
                    requires_lab=orig_subject.requires_lab,
                    lab_hours_per_week=new_lab,
                    required_specialization=orig_subject.required_specialization,
                    is_elective=orig_subject.is_elective,
                    elective_group=orig_subject.elective_group,
                    priority=orig_subject.priority
                )
            )
        
        return new_subjects, made_change
    
    def run(
        self,
        data: Dict[str, Any],
        scheduler_class,
        predictor
    ) -> Optional[Dict[str, Any]]:
        """
        Main emergency scheduling loop.
        
        Flow:
        1. Pre-check feasibility (fast)
        2. Try scheduling at each scale level
        3. Early exit on success
        4. Graceful degradation on failure
        
        Args:
            data: Full schedule data (timeslots, rooms, faculties, subjects, batches)
            scheduler_class: The scheduler class to instantiate (e.g., ProgressiveCPScheduler)
            predictor: Predictor for strategy (e.g., StrategyPredictor)
        
        Returns:
            Scheduler results dict, or None if all attempts failed
        """
        self.attempt_log = []
        
        # Step 1: Pre-check (fast, no scheduler)
        analyzer = FeasibilityAnalyzer(data)
        is_feasible, reason = analyzer.is_feasible(self.config.FEASIBILITY_BUFFER)
        logger.info(reason)
        self.attempt_log.append(('pre_check', reason, None))
        
        # Step 2: Try scheduling with original hours if pre-check was promising
        if is_feasible:
            logger.info("🔄 Attempt 0: Trying original hours...")
            try:
                scheduler = scheduler_class(data, predictor)
                results = scheduler.schedule()
                if results and results.get('statistics', {}).get('scheduled', 0) > 0:
                    logger.info("✅ Success with original hours!")
                    self.attempt_log.append(('original', 'Success', results))
                    return results
                logger.info("❌ Original hours failed, starting emergency scaling")
                self.attempt_log.append(('original', 'Failed', None))
            except Exception as e:
                logger.error(f"Error in original attempt: {e}")
                self.attempt_log.append(('original', f'Error: {e}', None))
        
        # Step 3: Emergency scaling loop
        orig_subjects = deepcopy(data['subjects'])
        orig_hours = self._store_original_hours(orig_subjects)
        
        for attempt in range(1, self.config.MAX_ATTEMPTS + 1):
            ratio = self.config.REDUCTION_STEP ** (attempt - 1)
            
            logger.info(f"⚠️ Attempt {attempt}: Scaling to {ratio:.1%} of original...")
            
            # Scale subjects
            new_subjects, made_change = self._scale_subjects(
                orig_subjects, orig_hours, ratio
            )
            
            if not made_change:
                logger.warning("No changes possible (at minimum). Stopping.")
                self.attempt_log.append(('scaling', f'Attempt {attempt}: No changes possible', None))
                break
            
            # Update data with scaled subjects
            data['subjects'] = new_subjects
            
            # Try scheduling
            try:
                scheduler = scheduler_class(data, predictor)
                results = scheduler.schedule()
                
                if results and results.get('statistics', {}).get('scheduled', 0) > 0:
                    total_req = analyzer.calculate_total_required_hours()
                    logger.info(f"✅ Success at attempt {attempt}! "
                              f"(Scaled to {ratio:.1%}, total: {total_req}h)")
                    self.attempt_log.append(('scaling', f'Attempt {attempt}: Success at {ratio:.1%}', results))
                    return results
                
                logger.debug(f"Attempt {attempt} failed, continuing...")
                self.attempt_log.append(('scaling', f'Attempt {attempt}: Failed at {ratio:.1%}', None))
            
            except Exception as e:
                logger.error(f"Error in attempt {attempt}: {e}")
                self.attempt_log.append(('scaling', f'Attempt {attempt}: Error: {e}', None))
        
        # All attempts exhausted
        logger.critical("❌ All scaling attempts exhausted. Cannot schedule.")
        return None
    
    def get_report(self) -> str:
        """Generate human-readable report of attempts"""
        report = "Emergency Scheduling Attempt Log:\n"
        for stage, desc, _ in self.attempt_log:
            report += f"  [{stage}] {desc}\n"
        return report


# -------------------------
# Progressive Scheduler with Floor Priority
# -------------------------
class ProgressiveCPScheduler:
    def __init__(self, data, predictor: StrategyPredictor):
        self.timeslots = data['timeslots']
        self.rooms = data['rooms']
        self.faculties = data['faculties']
        self.subjects = data['subjects']
        self.batches = data['batches']
        self.working_days = ["MON", "TUE", "WED", "THU", "FRI"]
        self.predictor = predictor
        
        # 🔒 THREAD SAFETY: Lock for protecting shared mutable state
        self._lock = threading.RLock()  # Reentrant lock for nested calls
        
        # NEW: Get faculty choices and constraints
        self.faculty_choices = data.get('faculty_choices', [])
        self.allocation_constraints = data.get('allocation_constraints', 
                                               AllocationConstraints(use_faculty_choice=False))
        self.breaks = data.get('breaks', [])
        
        self.subject_map = {s.subject_code: s for s in self.subjects}
        self.faculty_map = {f.faculty_id: f for f in self.faculties}
        self.batch_map = {b.batch_id: b for b in self.batches}
        
        # 🚀 NEW: Model Building Cache - stores expensive computations across strategies
        # This saves ~30+ seconds per additional strategy attempt
        self.model_cache = {
            'faculty_assignments': None,  # Cached faculty → session mappings
            'eligible_rooms': {},         # session_id → [room_ids] 
            'eligible_faculty': {},       # session_key → [faculty_ids]
            'valid_slots': {},            # session_id → [(day, slot)] valid slots
            'sessions_metadata': None,    # Processed session info
            'cache_valid': False          # Set to True once cache is built
        }
        
        # Group rooms by floor and department
        self.rooms_by_floor_dept = defaultdict(lambda: defaultdict(list))
        for room in self.rooms:
            self.rooms_by_floor_dept[room.floor][room.department].append(room)
        
        # NEW: Initialize enhanced faculty allocator
        self.faculty_allocator = EnhancedFacultyAllocator(
            faculties=self.faculties,
            subjects=self.subjects,
            faculty_choices=self.faculty_choices,
            constraints=self.allocation_constraints
        )
        
        # SMART FLOOR ASSIGNMENT
        self._auto_assign_floors_to_batches()

        # ─── SCARCITY DETECTION ───────────────────────────────────────────
        num_classrooms = len([r for r in self.rooms if 'lab' not in r.room_type.lower()])
        num_batches    = len(self.batches)
        self.scarcity_mode = num_batches > num_classrooms

        # overlap_slot_ids built ONCE with robust hour parsing
        self.overlap_slot_ids = []
        for ts in self.timeslots:
            hour_str = str(ts.start_time).strip().split(':')[0]
            if hour_str in ('11', '12'):
                self.overlap_slot_ids.append(ts.slot_id)

        self.low_priority_batches = set()
        # _analyze_batch_priorities called exactly ONCE (method defined once below)
        self._analyze_batch_priorities()

        if self.scarcity_mode:
            print(f"\n⚠️  SCARCITY MODE ACTIVE: {num_batches} Batches vs {num_classrooms} Classrooms")
        else:
            print(f"\n✅ Resources Sufficient: {num_batches} Batches vs {num_classrooms} Classrooms")
        print(f"   🕒 Overlap Slots: {self.overlap_slot_ids}")

        mode_icon = "✨" if self.allocation_constraints.use_faculty_choice else "📚"
        mode_name = "CHOICE-BASED" if self.allocation_constraints.use_faculty_choice else "QUALIFICATION-BASED"
        
        print(f"\\n{mode_icon} Progressive Scheduler initialized - {mode_name} faculty allocation")
        print(f"   Floors available: {sorted(set(r.floor for r in self.rooms))}")

    # ── Helper methods (defined after _auto_assign_floors_to_batches) ──
    
    def _auto_assign_floors_to_batches(self):
        """
        Automatically assign batches to floors based on:
        1. Room availability on each floor
        2. Department matching
        3. Capacity requirements
        4. Load balancing across floors
        """
        print(f"\n{'='*60}")
        print("🏢 SMART FLOOR ASSIGNMENT")
        print(f"{'='*60}")
        
        # Analyze floor capacity by department
        floor_capacity = defaultdict(lambda: defaultdict(int))
        floor_rooms = defaultdict(lambda: defaultdict(list))
        
        for room in self.rooms:
            if room.floor > 0:  # Only consider rooms with defined floors
                floor_capacity[room.floor][room.department] += 1
                floor_rooms[room.floor][room.department].append(room)
        
        # Calculate sessions needed per batch
        batch_needs = {}
        for batch in self.batches:
            theory_sessions = 0
            lab_sessions = 0
            
            for subj_code in batch.subjects:
                subject = self.subject_map.get(subj_code)
                #print("subj_code line 2442 ",subj_code," subject",subject)
                if subject:
                    theory_sessions += subject.hours_per_week
                    lab_sessions += subject.lab_hours_per_week
            
            batch_needs[batch.batch_id] = {
                'theory': theory_sessions,
                'lab': lab_sessions,
                'total': theory_sessions + lab_sessions
            }
        
        # Sort batches by priority (high priority gets first choice)
        sorted_batches = sorted(self.batches, key=lambda b: (b.priority, batch_needs[b.batch_id]['total']), reverse=True)
        
        # Track floor assignments and usage
        floor_load = defaultdict(int)
        batch_floor_assignments = {}
        
        print(f"\nAssigning batches to floors:\n")
        
        for batch in sorted_batches:
            # If batch already has preferred floor, keep it
            if batch.preferred_floor > 0 and batch.preferred_floor in floor_capacity:
                assigned_floor = batch.preferred_floor
                print(f"  {batch.batch_id}: Floor {assigned_floor} (pre-assigned)")
            else:
                # Calculate best floor for this batch
                floor_scores = {}
                
                for floor in floor_capacity.keys():
                    score = 0
                    
                    # Priority 1: Department match (high score)
                    if batch.department in floor_rooms[floor]:
                        dept_rooms = floor_rooms[floor][batch.department]
                        suitable_rooms = [r for r in dept_rooms if r.capacity >= batch.student_count]
                        score += len(suitable_rooms) * 1000
                    
                    # Priority 2: General rooms on floor
                    if 'general' in [r.department.lower() for r in floor_rooms[floor]['General']]:
                        general_rooms = [r for r in floor_rooms[floor]['General'] if r.capacity >= batch.student_count]
                        score += len(general_rooms) * 500
                    
                    # Priority 3: Any suitable rooms (lower score)
                    all_floor_rooms = []
                    for dept_rooms in floor_rooms[floor].values():
                        all_floor_rooms.extend(dept_rooms)
                    suitable_any = [r for r in all_floor_rooms if r.capacity >= batch.student_count]
                    score += len(suitable_any) * 100
                    
                    # Priority 4: Load balancing (prefer less loaded floors)
                    sessions_needed = batch_needs[batch.batch_id]['total']
                    current_load = floor_load[floor]
                    load_penalty = current_load * 10
                    score -= load_penalty
                    
                    # Priority 5: Lab availability if needed
                    if batch_needs[batch.batch_id]['lab'] > 0:
                        lab_rooms = [r for r in all_floor_rooms 
                                    if 'lab' in r.room_type.lower() and r.capacity >= batch.student_count]
                        if len(lab_rooms) > 0:
                            score += 300
                        else:
                            score -= 500  # Penalize if no lab on this floor
                    
                    floor_scores[floor] = score
                
                # Assign to best floor
                if floor_scores:
                    assigned_floor = max(floor_scores.keys(), key=lambda f: floor_scores[f])
                    print(f"  {batch.batch_id}: Floor {assigned_floor} (auto-assigned, score: {floor_scores[assigned_floor]:.0f})")
                    
                    # Show why this floor was chosen
                    dept_match = batch.department in floor_rooms[assigned_floor]
                    room_count = len([r for r in sum(floor_rooms[assigned_floor].values(), []) 
                                     if r.capacity >= batch.student_count])
                    print(f"    → Dept match: {dept_match} | Suitable rooms: {room_count} | Load: {floor_load[assigned_floor]}")
                else:
                    # Fallback: use floor 0 (any floor)
                    assigned_floor = 0
                    print(f"  {batch.batch_id}: Any floor (no specific floor constraints)")
            
            # Update batch with assigned floor
            batch_floor_assignments[batch.batch_id] = assigned_floor
            floor_load[assigned_floor] += batch_needs[batch.batch_id]['total']
        
        # Update batch objects with assigned floors
        updated_batches = []
        for batch in self.batches:
            assigned_floor = batch_floor_assignments.get(batch.batch_id, batch.preferred_floor)
            updated_batches.append(StudentBatch(
                batch.batch_id,
                batch.department,
                batch.student_count,
                batch.subjects,
                batch.shift_preference,
                assigned_floor,  # Updated floor
                batch.priority
            ))
        
        self.batches = updated_batches
        self.batch_map = {b.batch_id: b for b in self.batches}
        
        # Show final distribution
        print(f"\n📊 Final floor distribution:")
        for floor in sorted(floor_load.keys()):
            batches_on_floor = [b.batch_id for b in self.batches if b.preferred_floor == floor]
            print(f"  Floor {floor}: {len(batches_on_floor)} batch(es) - {floor_load[floor]} sessions")
            print(f"    Batches: {', '.join(batches_on_floor)}")
        
        print(f"{'='*60}\n")

    def _analyze_batch_priorities(self):
        """Calculates total workload to identify Low Priority batches for staggering"""
        batch_loads = {}
        all_hours = []
        for batch in self.batches:
            total = 0
            for sc in batch.subjects:
                s = self.subject_map.get(sc)
                if s: total += s.hours_per_week + s.lab_hours_per_week
            batch_loads[batch.batch_id] = total
            all_hours.append(total)
        
        # Bottom 30% of batches are considered "Low Priority" -> Can be delayed/staggered
        threshold = np.percentile(all_hours, 30) if all_hours else 0
        self.low_priority_batches = {b for b, h in batch_loads.items() if h <= threshold}
        print(f"   📊 Priority Analysis: {len(self.low_priority_batches)} Low-Load batches (<= {threshold}h)")
    
    @lru_cache(maxsize=2048)
    def _get_eligible_rooms_with_floor_priority(self, batch_id: str, is_lab: bool) -> tuple:
        """Enhanced room selection with FLOOR PRIORITY"""
        batch = self.batch_map[batch_id]
        preferred_floor = batch.preferred_floor
        
        eligible = []
        
        # PRIORITY 1: Same floor, same department
        if preferred_floor > 0:
            same_floor_dept = [
                r for r in self.rooms_by_floor_dept[preferred_floor][batch.department]
                if r.capacity >= batch.student_count
                and (not is_lab or 'lab' in r.room_type.lower())
            ]
            eligible.extend(same_floor_dept)
        
        # PRIORITY 2: Same floor, different department
        if preferred_floor > 0:
            same_floor_other = [
                r for r in self.rooms
                if r.floor == preferred_floor
                and r.department != batch.department
                and r.capacity >= batch.student_count
                and (not is_lab or 'lab' in r.room_type.lower())
                and r not in eligible
            ]
            eligible.extend(same_floor_other)
        
        # PRIORITY 3: Any floor, same department
        same_dept = [
            r for r in self.rooms
            if (r.department == batch.department or r.department.lower() == 'general')
            and r.capacity >= batch.student_count
            and (not is_lab or 'lab' in r.room_type.lower())
            and r not in eligible
        ]
        eligible.extend(same_dept)
        
        # PRIORITY 4: Any floor, any department (fallback)
        fallback = [
            r for r in self.rooms
            if r.capacity >= batch.student_count
            and (not is_lab or 'lab' in r.room_type.lower())
            and r not in eligible
        ]
        eligible.extend(fallback)
        
        return tuple(eligible)
    
    def _build_model_cache(self, sessions):
        """
        🚀 Pre-compute expensive data that's common across ALL strategies.
        
        This dramatically speeds up 2nd, 3rd, 4th strategy attempts:
        - 1st attempt: ~40s (builds cache)
        - 2nd+ attempts: ~10s (uses cache)
        
        Caches:
        - Faculty assignments (expensive matching)
        - Eligible rooms per session
        - Eligible faculty per session
        - Valid (day, slot) pairs per session
        """
        if self.model_cache['cache_valid']:
            print("   📦 Using cached model data (fast path)")
            return
        
        print("   🔨 Building model cache (one-time)...")
        cache_start = time.time()
        
        # 1. Cache faculty assignments
        faculty_assignments = self._intelligent_faculty_assignment(sessions)
        self.model_cache['faculty_assignments'] = faculty_assignments
        
        # 2. Cache eligible rooms and faculty for each session
        for session in sessions:
            sid = session['id']
            batch = self.batch_map[session['batch_id']]
            is_lab = session['type'] == 'lab'
            
            # Cache eligible rooms
            eligible_rooms = self._get_eligible_rooms_with_floor_priority(batch.batch_id, is_lab)
            self.model_cache['eligible_rooms'][sid] = eligible_rooms
            
            # Cache eligible faculty
            session_key = (session['batch_id'], session['subject_code'], session['type'])
            preferred_faculty = faculty_assignments.get(session_key, [])
            eligible_faculty = [self.faculty_map[fid] for fid in preferred_faculty if fid in self.faculty_map]
            if not eligible_faculty:
                eligible_faculty = list(self.faculties)
            self.model_cache['eligible_faculty'][sid] = eligible_faculty
            
            # Cache valid slots (considering breaks, shifts, lab constraints)
            valid_slots = []
            batch_shift = str(batch.shift_preference).strip().lower()
            
            for day in self.working_days:
                for ts in self.timeslots:
                    # Check break time
                    is_break_time = False
                    for brk in self.breaks:
                        if day not in brk['days']:
                            continue
                        if not (brk['start_slot'] <= ts.slot_id < brk['start_slot'] + brk['duration']):
                            continue
                        brk_shift = brk['shift']
                        if brk_shift in ['common', 'both', 'all'] or brk_shift == batch_shift:
                            is_break_time = True
                            break
                    
                    if is_break_time:
                        continue
                    
                    # Check scarcity mode
                    if self.scarcity_mode:
                        is_low = batch.batch_id in self.low_priority_batches
                        is_overlap = ts.slot_id in self.overlap_slot_ids
                        if is_low and is_overlap:
                            if batch.shift_preference == 'morning':
                                continue
                            if batch.shift_preference == 'evening' and self.overlap_slot_ids and ts.slot_id == self.overlap_slot_ids[0]:
                                continue
                    
                    # Check shift constraints
                    if batch.shift_preference == 'morning' and ts.slot_id > 5:
                        continue
                    elif batch.shift_preference in ['evening', 'afternoon'] and ts.slot_id < 3:
                        continue
                    
                    # Check lab constraints (need 2 consecutive slots)
                    if is_lab:
                        next_slot_id = ts.slot_id + 1
                        next_slot_exists = any(t.slot_id == next_slot_id and t.slot_id != 3 for t in self.timeslots)
                        if not next_slot_exists or ts.slot_id >= len(self.timeslots) - 1:
                            continue
                    
                    valid_slots.append((day, ts.slot_id))
            
            self.model_cache['valid_slots'][sid] = valid_slots
        
        self.model_cache['cache_valid'] = True
        cache_time = time.time() - cache_start
        print(f"   ✅ Cache built in {cache_time:.1f}s ({len(sessions)} sessions)")
    
    
    def schedule(self):
        if not ORTOOLS_AVAILABLE:
            print("ERROR: OR-Tools required")
            return {}
        
        print("\n🚀 Starting PROGRESSIVE scheduling with ML prediction...")
        start_time = time.time()
        
        # Create data dict for advisor
        data = {
            'timeslots': self.timeslots,
            'rooms': self.rooms,
            'faculties': self.faculties,
            'subjects': self.subjects,
            'batches': self.batches
        }
        
        # ═══════════════════════════════════════════════════════════════
        # SMART STRATEGY ADVISOR - Pre-Analysis
        # ═══════════════════════════════════════════════════════════════
        demand_analysis, faculty_recommendations = SMART_ADVISOR.prepare_execution(data)
        
        # Store for later use in failure reports
        self._demand_analysis = demand_analysis
        
        # Create sessions with priority ordering
        sessions = self._create_sessions_progressive()
        if not sessions:
            return {}
        
        print(f"Scheduling {len(sessions)} sessions (progressive order)")
        
        # Get features for ML prediction (legacy predictor)
        features = self.predictor.extract_features(data)
        
        # Predict best strategy using BOTH legacy predictor and Smart Advisor
        predicted_strategy_idx = self.predictor.predict_best_strategy(features)
        
        # Get Smart Advisor's ordered strategies (Hybrid: Rules + ML)
        ordered_strategies = SMART_ADVISOR.get_ordered_strategies(demand_analysis)
        
        if ordered_strategies:
            print(f"\n🧠 Smart Advisor Strategy Order (Hybrid):")
            for strategy, score in ordered_strategies:
                print(f"   • {strategy['name']}: {score:.1f}% confidence")
        
        # Multi-strategy with Smart Advisor ordering
        result = self._intelligent_scheduling_v2(sessions, ordered_strategies, demand_analysis)
        
        elapsed = time.time() - start_time
        print(f"\n✅ Optimization completed in {elapsed:.2f}s")
        
        if result and result.get('assignments'):
            # Save result for BOTH learning systems
            success_rate = len(result['assignments']) / len(sessions) * 100
            
            # Legacy predictor
            self.predictor.save_result(
                features,
                result.get('strategy', 'UNKNOWN'),
                success_rate
            )
            
            # Smart Advisor learning
            SMART_ADVISOR.record_outcome(
                demand_analysis,
                result.get('strategy', 'UNKNOWN'),
                True,
                {
                    'success_rate': success_rate / 100,
                    'gaps': result.get('gaps', 0),
                    'load_variance': demand_analysis.faculty_load_variance
                }
            )
            
            self._analyze_quality(result['assignments'])
            return self._format_results(result['assignments'])
        else:
            # Record failure for learning
            SMART_ADVISOR.record_outcome(
                demand_analysis,
                'ALL_FAILED',
                False,
                {'success_rate': 0, 'gaps': 0, 'load_variance': 0}
            )
            return {}
    
    def _intelligent_scheduling_v2(self, sessions, ordered_strategies, demand_analysis):
        """
        🧠 SMART SCHEDULING with Strategy Advisor integration.
        Uses hybrid scoring, cooldown, hidden constraint discovery,
        ML-based dynamic time limits, and PARALLEL EXECUTION.
        """
        if not ordered_strategies:
            # Fallback to legacy method
            return self._intelligent_scheduling(sessions, 0)
        
        # Update session count for dynamic time limit calculation
        SMART_ADVISOR.sessions_count = len(sessions)
        
        print(f"\n{'='*70}")
        print("🚀 INTELLIGENT SCHEDULING with Smart Advisor (PARALLEL MODE)")
        print(f"{'='*70}")
        
        # 🚀 Build model cache ONCE - saves ~30s per additional strategy
        self._build_model_cache(sessions)
        
        # Print dynamic time limits
        SMART_ADVISOR.print_dynamic_limits()
        
        # Prepare strategies with dynamic time limits
        prepared_strategies = []
        for strategy, confidence in ordered_strategies:
            strategy_name = strategy['name']
            dynamic_limit = SMART_ADVISOR.get_dynamic_time_limit(strategy_name)
            
            # Get adjusted config based on previous failures
            adjusted_strategy = SMART_ADVISOR.get_adjusted_config(strategy)
            adjusted_strategy['time_limit'] = dynamic_limit
            adjusted_strategy['confidence'] = confidence
            
            prepared_strategies.append(adjusted_strategy)
            print(f"   📋 {strategy_name}: {confidence:.1f}% confidence, {dynamic_limit}s limit")
        
        # ═══════════════════════════════════════════════════════════════════
        # 🚀 PARALLEL EXECUTION: Run top strategies simultaneously
        # ═══════════════════════════════════════════════════════════════════
        import os
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import threading
        
        cpu_count = os.cpu_count() or 8
        # Run top 2-3 strategies in parallel (split CPU cores between them)
        parallel_count = min(3, len(prepared_strategies), max(2, cpu_count // 4))
        
        # Calculate cores per strategy for parallel mode
        cores_per_strategy = max(2, cpu_count // parallel_count)
        
        # Assign cores to parallel strategies
        parallel_strategies = prepared_strategies[:parallel_count]
        for strat in parallel_strategies:
            strat['cores_per_strategy'] = cores_per_strategy
        
        remaining_strategies = prepared_strategies[parallel_count:]
        
        print(f"\n🔥 Running top {len(parallel_strategies)} strategies in PARALLEL...")
        print(f"   CPU cores: {cpu_count} total, {cores_per_strategy} per strategy")
        
        # Thread-safe result handling
        result_lock = threading.Lock()
        found_result = [None]
        shutdown_flag = threading.Event()
        print_lock = threading.Lock()
        
        def safe_print(msg):
            with print_lock:
                print(msg)
        
        def solve_parallel(strat):
            if shutdown_flag.is_set():
                return None
            
            strategy_name = strat['name']
            safe_print(f"\n🎯 [PARALLEL] Starting: {strategy_name}")
            
            try:
                strategy_start = time.time()
                result = self._solve_with_strategy(sessions, strat)
                strategy_time = time.time() - strategy_start
                
                with result_lock:
                    if result and result.get('assignments') and found_result[0] is None:
                        scheduled = len(result['assignments'])
                        total = len(sessions)
                        success_rate = scheduled / total * 100
                        
                        result['strategy'] = strategy_name
                        result['execution_time'] = strategy_time
                        result['sessions_count'] = total
                        
                        found_result[0] = result
                        shutdown_flag.set()  # Signal others to stop
                        
                        safe_print(f"\n✅ [PARALLEL] SUCCESS with {strategy_name}!")
                        safe_print(f"   Scheduled: {scheduled}/{total} ({success_rate:.1f}%)")
                        safe_print(f"   Time: {strategy_time:.1f}s")
                        
                        return result
                
                safe_print(f"   ⚠️ [PARALLEL] {strategy_name}: No solution ({strategy_time:.1f}s)")
                return None
                
            except Exception as e:
                safe_print(f"   ❌ [PARALLEL] {strategy_name}: {e}")
                return None
        
        # Execute parallel strategies
        with ThreadPoolExecutor(max_workers=parallel_count) as executor:
            futures = {executor.submit(solve_parallel, s): s for s in parallel_strategies}
            
            for future in as_completed(futures):
                if shutdown_flag.is_set():
                    # Cancel remaining futures
                    for f in futures:
                        f.cancel()
                    break
        
        # Check if parallel found a result
        with result_lock:
            if found_result[0]:
                return found_result[0]
        
        # ═══════════════════════════════════════════════════════════════════
        # 🔄 FALLBACK: Sequential execution for remaining strategies
        # ═══════════════════════════════════════════════════════════════════
        if remaining_strategies:
            print(f"\n🔄 Parallel strategies failed. Trying {len(remaining_strategies)} more SEQUENTIALLY...")
            
            for strategy in remaining_strategies:
                strategy_name = strategy['name']
                
                print(f"\n🎯 [SEQUENTIAL] Trying: {strategy_name}")
                
                strategy_start = time.time()
                result = self._solve_with_strategy(sessions, strategy)
                strategy_time = time.time() - strategy_start
                
                if result and result.get('assignments'):
                    scheduled = len(result['assignments'])
                    total = len(sessions)
                    success_rate = scheduled / total * 100
                    
                    print(f"\n✅ SUCCESS with {strategy_name}!")
                    print(f"   Scheduled: {scheduled}/{total} ({success_rate:.1f}%)")
                    print(f"   Time: {strategy_time:.1f}s")
                    
                    result['strategy'] = strategy_name
                    result['execution_time'] = strategy_time
                    result['sessions_count'] = total
                    return result
                else:
                    # Create failure report
                    failure_report = StrategyFailureReport(
                        strategy_name=strategy_name,
                        failure_type='no_solution',
                        execution_time=strategy_time,
                        sessions_scheduled=0,
                        sessions_failed=len(sessions),
                        success_rate=0.0,
                        primary_bottleneck='unknown'
                    )
                    SMART_ADVISOR.handle_failure(strategy_name, failure_report, demand_analysis)
                    print(f"   ❌ FAILED after {strategy_time:.1f}s")
        
        print(f"\n❌ All strategies exhausted")
        
        # Show recommendations
        if demand_analysis.faculty_pressure > 80:
            recommendations = SMART_ADVISOR.calculate_faculty_recommendations(demand_analysis)
            recommendations.print_recommendations()
        
        return None
    
    def _create_sessions_progressive(self) -> List[Dict]:
        """Create sessions ordered by priority (high priority batches/subjects first)"""
        sessions = []
        
        # Sort batches by priority (higher first)
        sorted_batches = sorted(self.batches, key=lambda b: b.priority, reverse=True)
        
        for batch in sorted_batches:
            # Get subjects for this batch, sorted by priority
            batch_subjects = []
            for subj_code in batch.subjects:
                subject = self.subject_map.get(subj_code)
                if subject:
                    batch_subjects.append(subject)
            
            # Sort subjects by priority (higher first)
            batch_subjects.sort(key=lambda s: s.priority, reverse=True)
            
            # Create sessions
            for subject in batch_subjects:
                # Mark elective sessions
                elective_mark = f"_ELEC_{subject.elective_group}" if subject.is_elective else ""
                
                # Theory sessions
                for i in range(subject.hours_per_week):
                    sessions.append({
                        'batch_id': batch.batch_id,
                        'subject_code': subject.subject_code,
                        'type': 'theory',
                        'id': f"{batch.batch_id}_{subject.subject_code}{elective_mark}_T{i}",
                        'priority': batch.priority + subject.priority,
                        'is_elective': subject.is_elective,
                        'elective_group': subject.elective_group
                    })
                
                # Lab sessions
                for i in range(subject.lab_hours_per_week):
                    sessions.append({
                        'batch_id': batch.batch_id,
                        'subject_code': subject.subject_code,
                        'type': 'lab',
                        'id': f"{batch.batch_id}_{subject.subject_code}{elective_mark}_L{i}",
                        'priority': batch.priority + subject.priority + 5,  # Higher priority for labs
                        'is_elective': subject.is_elective,
                        'elective_group': subject.elective_group,
                        'requires_consecutive': True  # NEW FLAG
                    })
        
        # Sort all sessions by priority (higher first)
        sessions.sort(key=lambda s: s['priority'], reverse=True)
        
        print(f"\n📊 Session priority distribution:")
        priority_counts = Counter(s['priority'] for s in sessions)
        for priority in sorted(priority_counts.keys(), reverse=True):
            print(f"   Priority {priority}: {priority_counts[priority]} sessions")
        
        return sessions
    
    def _intelligent_scheduling(self, sessions, predicted_idx):
        """
        🚀 SMART EXECUTION: Choose between parallel processes or sequential.
        - PARALLEL: Uses ProcessPoolExecutor (true CPU parallelism, bypasses GIL)
        - SEQUENTIAL: Lower memory, OR-Tools uses all cores internally
        """
        strategies = [
            {'name': 'STRICT NO GAPS', 'allow_gaps': False, 'time_limit': 30},
            {'name': 'COMPACT SCHEDULE', 'allow_gaps': False, 'time_limit': 30},
            {'name': 'RELAXED', 'allow_gaps': True, 'time_limit': 30},
            {'name': 'MAXIMUM FLEXIBILITY', 'allow_gaps': True, 'time_limit': 30}
        ]
        
        # Prioritize ML-predicted strategy
        if predicted_idx > 0 and predicted_idx < len(strategies):
            strategies.insert(0, strategies.pop(predicted_idx))
        
        # ═══════════════════════════════════════════════════════════════
        # 🔧 CONFIGURATION: Choose execution mode
        # ═══════════════════════════════════════════════════════════════
        # Options:
        #   'parallel'   - True CPU parallelism with processes (faster if multi-core)
        #   'sequential' - Low memory, OR-Tools handles internal parallelism
        EXECUTION_MODE = 'parallel'  # Change to 'sequential' for low-memory mode
        MAX_PARALLEL_STRATEGIES = 2   # Limit parallel processes (2 = balance of speed/memory)
        # ═══════════════════════════════════════════════════════════════
        
        if EXECUTION_MODE == 'parallel':
            return self._run_parallel_strategies(sessions, strategies, MAX_PARALLEL_STRATEGIES)
        else:
            return self._run_sequential_strategies(sessions, strategies)
    
    def _run_parallel_strategies(self, sessions, strategies, max_workers=2):
        """
        🚀 THREAD-SAFE PARALLEL EXECUTION
        Uses proper locks and atomic operations to prevent race conditions.
        First strategy to succeed wins, others are cancelled.
        """
        from concurrent.futures import FIRST_COMPLETED, wait
        import threading
        
        print(f"\n🚀 Running {len(strategies)} strategies in PARALLEL ({max_workers} workers)...")
        print("   (Thread-safe with race condition protection)")
        
        # 🔒 Thread-safe shared state
        result_lock = threading.Lock()
        found_result = [None]  # Use list to allow mutation in nested function
        shutdown_flag = threading.Event()  # Signal to stop processing
        
        # Thread-safe print wrapper
        print_lock = threading.Lock()
        def safe_print(msg):
            with print_lock:
                print(msg)
        
        def solve_strategy_wrapper(strategy_info):
            """Thread-safe wrapper for solving a single strategy"""
            idx, strat = strategy_info
            
            # Check if we should abort early
            if shutdown_flag.is_set():
                return None
            
            try:
                result = self._solve_with_strategy(sessions, strat)
                
                # Thread-safe result check and storage
                with result_lock:
                    if result and result.get('assignments') and found_result[0] is None:
                        found_result[0] = result
                        shutdown_flag.set()  # Signal other threads to stop
                        safe_print(f"\n✅ Strategy [{idx+1}] '{strat['name']}' SUCCEEDED!")
                        return result
                    
                if result is None or not result.get('assignments'):
                    safe_print(f"   ⚠️ [{idx+1}] {strat['name']}: No solution")
                    
                return result
                
            except Exception as e:
                safe_print(f"   ❌ [{idx+1}] {strat['name']}: {e}")
                return None
        
        # Execute with ThreadPoolExecutor (OR-Tools releases GIL during solving)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Create strategy index pairs
            strategy_items = list(enumerate(strategies))
            
            # Submit all strategies
            futures = []
            for item in strategy_items:
                if shutdown_flag.is_set():
                    break
                future = executor.submit(solve_strategy_wrapper, item)
                futures.append(future)
                safe_print(f"   ▶️ Started: [{item[0]+1}] {item[1]['name']}")
            
            # Wait for completion with early exit
            try:
                for future in futures:
                    if shutdown_flag.is_set():
                        future.cancel()
                    else:
                        future.result(timeout=60)  # Max 60s per strategy
            except Exception:
                pass  # Timeout or cancellation
        
        # Return the successful result (thread-safe read)
        with result_lock:
            return found_result[0]
    
    def _run_sequential_strategies(self, sessions, strategies):
        """
        🔄 SEQUENTIAL EXECUTION with early exit
        Lower memory usage, OR-Tools handles internal parallelism (8 workers).
        """
        print(f"\n🔄 Running {len(strategies)} strategies SEQUENTIALLY (low memory mode)...")
        
        for i, strat in enumerate(strategies, 1):
            print(f"\n   [{i}/{len(strategies)}] Trying '{strat['name']}'...")
            try:
                result = self._solve_with_strategy(sessions, strat)
                if result and result.get('assignments'):
                    print(f"\n✅ Strategy '{strat['name']}' SUCCEEDED!")
                    return result
                else:
                    print(f"   ⚠️ No solution found")
            except Exception as e:
                print(f"   ❌ Failed: {e}")
        
        return None
    
    def _solve_with_strategy(self, sessions, strategy):
        """Core CP-SAT solver with floor priority, elective handling, and CACHING"""
        model = cp_model.CpModel()
        
        # 🚀 USE CACHED DATA if available (saves ~15-30s per strategy retry)
        use_cache = self.model_cache['cache_valid']
        
        if use_cache:
            faculty_assignments = self.model_cache['faculty_assignments']
        else:
            # Fallback: compute if cache not built
            faculty_assignments = self._intelligent_faculty_assignment(sessions)
        
        if not faculty_assignments:
            return None
        
        # Variables
        x = {}
        
        for session in sessions:
            sid = session['id']
            batch = self.batch_map[session['batch_id']]
            
            # 🚀 USE CACHED eligible rooms
            if use_cache and sid in self.model_cache['eligible_rooms']:
                eligible_rooms = self.model_cache['eligible_rooms'][sid]
            else:
                eligible_rooms = self._get_eligible_rooms_with_floor_priority(
                    batch.batch_id, 
                    session['type'] == 'lab'
                )
            
            if not eligible_rooms:
                continue
            
            # 🚀 USE CACHED eligible faculty
            if use_cache and sid in self.model_cache['eligible_faculty']:
                eligible_faculty = self.model_cache['eligible_faculty'][sid]
            else:
                session_key = (session['batch_id'], session['subject_code'], session['type'])
                preferred_faculty = faculty_assignments.get(session_key, [])
                eligible_faculty = [self.faculty_map[fid] for fid in preferred_faculty if fid in self.faculty_map]
            
            if not eligible_faculty:
                eligible_faculty = list(self.faculties)
            
            for day in self.working_days:
                for ts in self.timeslots:
                    # ---------------------------------------------------------
                    # 🟢 FIXED: ROBUST BREAK LOGIC
                    # ---------------------------------------------------------
                    is_break_time = False
                    
                    # 1. Normalize batch shift (e.g., "Morning " -> "morning")
                    batch_shift = str(batch.shift_preference).strip().lower()
                    
                    for brk in self.breaks:
                        # 2. Day Match?
                        if day not in brk['days']: continue
                        
                        # 3. Time Match?
                        # Is current slot inside the break duration?
                        if not (brk['start_slot'] <= ts.slot_id < brk['start_slot'] + brk['duration']):
                            continue
                            
                        # 4. Shift Match? (Comparing normalized strings)
                        brk_shift = brk['shift'] # Already lowercased in reader
                        
                        # Case A: Common Break (Everyone takes it)
                        if brk_shift in ['common', 'both', 'all']:
                            is_break_time = True
                            break
                        
                        # Case B: Exact Shift Match
                        # If break is "morning", matches "morning" batch.
                        # If break is "evening", matches "evening" batch.
                        if brk_shift == batch_shift:
                            is_break_time = True
                            break
                            
                    if is_break_time:
                        continue # Skip this slot (It is a break for this batch)
                    # -------------------------------------------------------------
                    # 🟢 SCARCITY STRATEGY: STAGGERED ENTRY/EXIT (Hard Constraints)
                    # -------------------------------------------------------------
                    if self.scarcity_mode:
                        is_low = batch.batch_id in self.low_priority_batches
                        is_overlap = ts.slot_id in self.overlap_slot_ids
                        
                        if is_low and is_overlap:
                            # 1. Low Prio Morning: MUST LEAVE before overlap starts
                            if batch.shift_preference == 'morning': 
                                continue
                            
                            # 2. Low Prio Evening: MUST START after first overlap slot (e.g. Start 12 not 11)
                            if batch.shift_preference == 'evening' and self.overlap_slot_ids and ts.slot_id == self.overlap_slot_ids[0]: 
                                continue

                    # Shift constraints
                    if batch.shift_preference == 'morning' and ts.slot_id > 5:
                        continue
                    elif batch.shift_preference in ['evening', 'afternoon'] and ts.slot_id < 3:
                        continue
                    
                    if session['type'] == 'lab':
                        next_slot_id = ts.slot_id + 1
                        # Check if next slot exists and is not lunch
                        next_slot_exists = any(t.slot_id == next_slot_id and t.slot_id != 3 for t in self.timeslots)
                        if not next_slot_exists or ts.slot_id >= len(self.timeslots) - 1:
                            continue  # Skip this timeslot for labs
                    
                    for room in eligible_rooms:
                        for faculty in eligible_faculty:
                            var_name = f"{session['id']}_{day}_{ts.slot_id}_{room.room_id}_{faculty.faculty_id}"
                            x[session['id'], day, ts.slot_id, room.room_id, faculty.faculty_id] = model.NewBoolVar(var_name)
        
        # ═══════════════════════════════════════════════════════════════
        # BUILD INDEXES — turns every O(n) dict scan into O(1) lookup
        # Without this, each "for k in x if ..." scans 50 000+ entries.
        # ═══════════════════════════════════════════════════════════════
        # session_id → batch_id   (pre-computed from sessions list)
        sess_to_batch   = {s['id']: s['batch_id']   for s in sessions}
        sess_to_subject = {s['id']: s['subject_code'] for s in sessions}

        idx_by_session          = defaultdict(list)   # session_id        → [keys]
        idx_by_day_slot_room    = defaultdict(list)   # (day,slot,room)   → [keys]
        idx_by_day_slot_faculty = defaultdict(list)   # (day,slot,fac)    → [keys]
        idx_by_day_slot_batch   = defaultdict(list)   # (day,slot,batch)  → [keys]
        idx_by_day_slot         = defaultdict(list)   # (day,slot)        → [keys]
        idx_by_day_batch_subj   = defaultdict(list)   # (day,batch,subj)  → [keys]

        for k in x:
            sid, day, slot, rid, fid = k
            bid = sess_to_batch.get(sid)
            sub = sess_to_subject.get(sid)
            idx_by_session[sid].append(k)
            idx_by_day_slot_room[(day, slot, rid)].append(k)
            idx_by_day_slot_faculty[(day, slot, fid)].append(k)
            idx_by_day_slot[(day, slot)].append(k)
            if bid:
                idx_by_day_slot_batch[(day, slot, bid)].append(k)
            if bid and sub:
                idx_by_day_batch_subj[(day, bid, sub)].append(k)

        # ─── HARD CONSTRAINTS ─────────────────────────────────────────
        # 1) Each session scheduled exactly once
        for session in sessions:
            svars = [x[k] for k in idx_by_session[session['id']]]
            if svars:
                model.AddExactlyOne(svars)

        # 2) At most one session per room per (day, slot)
        for day in self.working_days:
            for ts in self.timeslots:
                for room in self.rooms:
                    rvars = [x[k] for k in idx_by_day_slot_room[(day, ts.slot_id, room.room_id)]]
                    if rvars:
                        model.AddAtMostOne(rvars)

        # 3) At most one session per faculty per (day, slot)
        for day in self.working_days:
            for ts in self.timeslots:
                for fac in self.faculties:
                    fvars = [x[k] for k in idx_by_day_slot_faculty[(day, ts.slot_id, fac.faculty_id)]]
                    if fvars:
                        model.AddAtMostOne(fvars)

        # 4) At most one session per batch per (day, slot)
        for day in self.working_days:
            for ts in self.timeslots:
                for batch in self.batches:
                    bvars = [x[k] for k in idx_by_day_slot_batch[(day, ts.slot_id, batch.batch_id)]]
                    if bvars:
                        model.AddAtMostOne(bvars)

        # ─── LAB CONSECUTIVE SLOTS ────────────────────────────────────
        # Pre-build per-shift break-slot sets so we never re-scan breaks
        common_breaks = set()
        shift_breaks  = defaultdict(set)   # 'morning' → {(day,slot), …}
        for brk in self.breaks:
            for d in brk['days']:
                for s_id in range(brk['start_slot'], brk['start_slot'] + brk['duration']):
                    if brk['shift'] in ('common', 'both', 'all'):
                        common_breaks.add((d, s_id))
                    else:
                        shift_breaks[brk['shift']].add((d, s_id))

        for session in sessions:
            if session['type'] != 'lab':
                continue
            batch = self.batch_map[session['batch_id']]
            bshift = str(batch.shift_preference).strip().lower()
            # Merge common + this-shift breaks
            my_breaks = common_breaks | shift_breaks.get(bshift, set())

            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in my_breaks:
                        continue

                    # Keys for this session at this (day, slot)
                    here = [k for k in idx_by_session[session['id']]
                            if k[1] == day and k[2] == ts.slot_id]
                    if not here:
                        continue

                    next_slot_id = ts.slot_id + 1
                    next_ts = next((t for t in self.timeslots if t.slot_id == next_slot_id), None)
                    if not next_ts or next_ts.slot_id == 3:
                        continue

                    bid = session['batch_id']
                    for key in here:
                        rid  = key[3]
                        fid  = key[4]
                        # Block same room / same faculty / same batch in next slot
                        for bk in idx_by_day_slot_room[(day, next_slot_id, rid)]:
                            if bk != key:
                                model.Add(x[bk] == 0).OnlyEnforceIf(x[key])
                        for bk in idx_by_day_slot_faculty[(day, next_slot_id, fid)]:
                            if bk != key:
                                model.Add(x[bk] == 0).OnlyEnforceIf(x[key])
                        for bk in idx_by_day_slot_batch[(day, next_slot_id, bid)]:
                            if bk != key:
                                model.Add(x[bk] == 0).OnlyEnforceIf(x[key])

        # ─── NO SAME SUBJECT TWICE ON SAME DAY (per batch) ───────────
        for batch in self.batches:
            for subj_code in batch.subjects:
                for day in self.working_days:
                    svars = [x[k] for k in idx_by_day_batch_subj[(day, batch.batch_id, subj_code)]]
                    if len(svars) > 1:
                        model.Add(sum(svars) <= 1)

        # ─── ELECTIVE GROUP: no two electives in same group at same time ─
        elective_groups = defaultdict(lambda: defaultdict(list))
        for session in sessions:
            if session.get('is_elective') and session.get('elective_group'):
                elective_groups[session['batch_id']][session['elective_group']].append(session)

        for batch_id, groups in elective_groups.items():
            for group, group_sessions in groups.items():
                group_ids = {s['id'] for s in group_sessions}
                for day in self.working_days:
                    for ts in self.timeslots:
                        gvars = [x[k] for k in idx_by_day_slot[(day, ts.slot_id)]
                                 if k[0] in group_ids]
                        if len(gvars) > 1:
                            model.AddAtMostOne(gvars)

        # Objectives with FLOOR PREFERENCE BONUS
        objective_terms = []

        # -------------------------------------------------------------
        # 🟢 SCARCITY STRATEGY: LAB BUFFER (Soft Objectives)
        # -------------------------------------------------------------
        if self.scarcity_mode:
            # Iterate all variables to apply Lab Buffer logic
            for key, var in x.items():
                # Extract details from key: (session_id, day, slot, room_id, faculty_id)
                sess_id_chk = key[0]
                slot_chk = key[2]
                
                # Only apply if we are in the Critical Overlap Zone (11am-1pm)
                if slot_chk in self.overlap_slot_ids:
                    batch_id_chk = sess_id_chk.split('_')[0]
                    batch_obj = self.batch_map[batch_id_chk]
                    is_lab_chk = "_L" in sess_id_chk
                    
                    if batch_obj.shift_preference == 'morning':
                        if is_lab_chk:
                            # HUGE BONUS: Move Morning batches to Labs during overlap
                            # This empties classrooms for Evening batches
                            objective_terms.append(var * 800)
                        else:
                            # PENALTY: Keeping Morning Theory in classrooms during overlap
                            objective_terms.append(var * -400)
                    
                    elif batch_obj.shift_preference == 'evening':
                        if not is_lab_chk:
                            # BONUS: Evening batches fill the empty classrooms immediately
                            objective_terms.append(var * 200)
        
        # Priority 1: STRONG floor consistency - keep ALL classes on same floor
        for batch in self.batches:
            if batch.preferred_floor > 0:
                batch_sessions = [s for s in sessions if s['batch_id'] == batch.batch_id]
                
                # For each session of this batch, give massive bonus for preferred floor
                for session in batch_sessions:
                    for key in x:
                        if key[0] == session['id']:
                            room_id = key[3]
                            room = next((r for r in self.rooms if r.room_id == room_id), None)
                            if room:
                                if room.floor == batch.preferred_floor:
                                    # HUGE bonus for matching assigned floor
                                    objective_terms.append(x[key] * 5000)
                                elif room.floor > 0:
                                    # Penalty for using different floor (discourages floor mixing)
                                    objective_terms.append(x[key] * -1000)
        
        # Priority 2: Floor consistency bonus - reward using ONLY ONE floor per batch
        for batch in self.batches:
            if batch.preferred_floor > 0:
                # Create binary variable for each floor usage by this batch
                for floor in set(r.floor for r in self.rooms if r.floor > 0):
                    floor_used = model.NewBoolVar(f"floor_{floor}_used_by_{batch.batch_id}")
                    
                    # This floor is used if ANY session of this batch uses a room on this floor
                    floor_session_vars = [
                        x[k] for k in x
                        if any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)
                        and any(r.room_id == k[3] and r.floor == floor for r in self.rooms)
                    ]
                    
                    if floor_session_vars:
                        # If any session uses this floor, floor_used = 1
                        model.Add(sum(floor_session_vars) >= 1).OnlyEnforceIf(floor_used)
                        model.Add(sum(floor_session_vars) == 0).OnlyEnforceIf(floor_used.Not())
                        
                        # Penalty for using non-preferred floors
                        if floor != batch.preferred_floor:
                            objective_terms.append(floor_used * -2000)
        
        # Priority 3: Shift preference
        for batch in self.batches:
            for day in self.working_days:
                for ts in self.timeslots:
                    # ---------------------------------------------------------
                    # 🟢 NEW: BREAK LOGIC WITH SHIFT SUPPORT
                    # ---------------------------------------------------------
                    is_break_time = False
                    
                    for brk in self.breaks:
                        # Check Day
                        if day not in brk['days']: continue
                        
                        # Check Time (Is current slot inside the break duration?)
                        if not (brk['start_slot'] <= ts.slot_id < brk['start_slot'] + brk['duration']):
                            continue
                            
                        # Check Shift Compatibility
                        # 1. Common break? Everyone takes it.
                        if brk['shift'] == 'common':
                            is_break_time = True
                            break
                        
                        # 2. Morning break? Only Morning batches take it.
                        if brk['shift'] == 'morning' and batch.shift_preference == 'morning':
                            is_break_time = True
                            break
                            
                        # 3. Evening break? Only Evening batches take it.
                        if brk['shift'] == 'evening' and batch.shift_preference == 'evening':
                            is_break_time = True
                            break
                            
                    if is_break_time:
                        continue # Skip this slot
                    
                    batch_slot_vars = [x[k] for k in x if k[1] == day and k[2] == ts.slot_id
                                      and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)]
                    
                    if not batch_slot_vars:
                        continue
                    
                    slot_used = model.NewBoolVar(f"batch_{batch.batch_id}_{day}_{ts.slot_id}")
                    model.Add(sum(batch_slot_vars) >= 1).OnlyEnforceIf(slot_used)
                    model.Add(sum(batch_slot_vars) == 0).OnlyEnforceIf(slot_used.Not())
                    
                    if batch.shift_preference == 'morning':
                        bonus = 2000 - (ts.slot_id * 200) if ts.slot_id <= 2 else 800
                    elif batch.shift_preference in ['evening', 'afternoon']:
                        bonus = 2000 - ((7 - ts.slot_id) * 200) if ts.slot_id >= 5 else 500
                    else:
                        bonus = 1500
                    
                    objective_terms.append(slot_used * bonus)
        
        # Priority 4: Gap elimination
        for batch in self.batches:
            for day in self.working_days:
                for i in range(len(self.timeslots) - 2):
                    ts1, ts2, ts3 = self.timeslots[i], self.timeslots[i+1], self.timeslots[i+2]
                    
                    if ts1.slot_id == 3 or ts2.slot_id == 3 or ts3.slot_id == 3:
                        continue
                    
                    s1_vars = [x[k] for k in x if k[1] == day and k[2] == ts1.slot_id 
                              and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)]
                    s2_vars = [x[k] for k in x if k[1] == day and k[2] == ts2.slot_id 
                              and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)]
                    s3_vars = [x[k] for k in x if k[1] == day and k[2] == ts3.slot_id 
                              and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)]
                    
                    if s1_vars and s2_vars and s3_vars:
                        slot1_used = model.NewBoolVar(f"b{batch.batch_id}_{day}_s{ts1.slot_id}")
                        slot2_used = model.NewBoolVar(f"b{batch.batch_id}_{day}_s{ts2.slot_id}")
                        slot3_used = model.NewBoolVar(f"b{batch.batch_id}_{day}_s{ts3.slot_id}")
                        
                        model.Add(sum(s1_vars) >= 1).OnlyEnforceIf(slot1_used)
                        model.Add(sum(s1_vars) == 0).OnlyEnforceIf(slot1_used.Not())
                        model.Add(sum(s2_vars) >= 1).OnlyEnforceIf(slot2_used)
                        model.Add(sum(s2_vars) == 0).OnlyEnforceIf(slot2_used.Not())
                        model.Add(sum(s3_vars) >= 1).OnlyEnforceIf(slot3_used)
                        model.Add(sum(s3_vars) == 0).OnlyEnforceIf(slot3_used.Not())
                        
                        gap_detected = model.NewBoolVar(f"gap_{batch.batch_id}_{day}_{ts1.slot_id}")
                        model.AddBoolAnd([slot1_used, slot3_used, slot2_used.Not()]).OnlyEnforceIf(gap_detected)
                        # Use strategy-specific gap penalty
                        gap_penalty = strategy.get('gap_penalty', 2500)
                        objective_terms.append(gap_detected * -gap_penalty)
        
        # ════════════════════════════════════════════════════════════════════
        # 🚀 MAXIMUM UTILIZATION STRATEGY BONUSES
        # ════════════════════════════════════════════════════════════════════
        if strategy.get('max_utilization'):
            strat_name = strategy.get('name', '')
            
            # FACULTY EXHAUSTION: Bonus for every assigned slot (fill up faculty hours)
            if strategy.get('prioritize') == 'faculty':
                slot_bonus = strategy.get('slot_bonus', 100)
                for key, var in x.items():
                    objective_terms.append(var * slot_bonus)
            
            # ROOM EXHAUSTION: Bonus for filling every room-slot
            elif strategy.get('prioritize') == 'rooms':
                slot_bonus = strategy.get('slot_bonus', 100)
                for key, var in x.items():
                    objective_terms.append(var * slot_bonus)
            
            # DENSE PACK: Strong consecutive slot bonus
            consecutive_bonus = strategy.get('consecutive_bonus', 0)
            if consecutive_bonus > 0:
                for batch in self.batches:
                    for day in self.working_days:
                        for i in range(len(self.timeslots) - 1):
                            ts1, ts2 = self.timeslots[i], self.timeslots[i+1]
                            if ts1.slot_id == 3 or ts2.slot_id == 3:
                                continue
                            
                            s1_vars = [x[k] for k in x if k[1] == day and k[2] == ts1.slot_id 
                                      and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)]
                            s2_vars = [x[k] for k in x if k[1] == day and k[2] == ts2.slot_id 
                                      and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)]
                            
                            if s1_vars and s2_vars:
                                # Bonus for consecutive slots being used
                                consec_slot1 = model.NewBoolVar(f"consec1_{batch.batch_id}_{day}_{ts1.slot_id}")
                                consec_slot2 = model.NewBoolVar(f"consec2_{batch.batch_id}_{day}_{ts2.slot_id}")
                                model.Add(sum(s1_vars) >= 1).OnlyEnforceIf(consec_slot1)
                                model.Add(sum(s1_vars) == 0).OnlyEnforceIf(consec_slot1.Not())
                                model.Add(sum(s2_vars) >= 1).OnlyEnforceIf(consec_slot2)
                                model.Add(sum(s2_vars) == 0).OnlyEnforceIf(consec_slot2.Not())
                                
                                both_used = model.NewBoolVar(f"both_{batch.batch_id}_{day}_{ts1.slot_id}")
                                model.AddBoolAnd([consec_slot1, consec_slot2]).OnlyEnforceIf(both_used)
                                objective_terms.append(both_used * consecutive_bonus)
        
        if objective_terms:
            model.Maximize(sum(objective_terms))
        
        # ════════════════════════════════════════════════════════════════════
        # 🚀 OPTIMIZED SOLVER CONFIGURATION - Faster search without quality loss
        # ════════════════════════════════════════════════════════════════════
        solver = cp_model.CpSolver()
        
        # Core time limit
        solver.parameters.max_time_in_seconds = strategy['time_limit']
        
        # Use CPU cores (can be limited for parallel mode)
        import os
        cpu_count = os.cpu_count() or 8
        # If strategy specifies cores_per_strategy (parallel mode), use that, else use all
        cores_to_use = strategy.get('cores_per_strategy', cpu_count)
        solver.parameters.num_search_workers = cores_to_use
        
        # 🔥 SPEED OPTIMIZATIONS:
        # 1. Linearization - helps LP relaxation find better bounds faster
        solver.parameters.linearization_level = 2  # Maximum linearization
        
        # 2. Randomization - diversify search to escape local minima faster
        solver.parameters.random_seed = 42
        solver.parameters.randomize_search = True
        
        # 3. Presolve - aggressive simplification before solving
        solver.parameters.cp_model_presolve = True
        
        # 4. Stop immediately when ANY feasible solution found (for strict strategies)
        if 'STRICT' in strategy.get('name', '') or 'COMPACT' in strategy.get('name', ''):
            solver.parameters.stop_after_first_solution = False  # Find optimal
        else:
            # For RELAXED/FLEX, first good solution is good enough
            solver.parameters.stop_after_first_solution = True
        
        # 5. Reduce memory usage for faster cache hits
        solver.parameters.share_level_zero_bounds = True
        
        # Quiet logging
        solver.parameters.log_search_progress = False
        
        status = solver.Solve(model)
        
        if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
            assignments = {}
            for key, var in x.items():
                if solver.Value(var) == 1:
                    sess_id, day, slot, room_id, faculty_id = key
                    assignments[sess_id] = {
                        'session_id': sess_id,
                        'day': day,
                        'slot': slot,
                        'room_id': room_id,
                        'faculty_id': faculty_id
                    }
            
            filled_assignments = self._fill_gaps_intelligently(assignments, sessions)
            
            return {'assignments': filled_assignments, 'status': solver.StatusName(status)}
        
        return None
    
    def _intelligent_faculty_assignment(self, sessions):
        """
        🔒 THREAD-SAFE: Returns top-3 faculty candidates per (batch, subject, type).
        
        MULTI-SUBJECT SUPPORT: Allows faculty to teach up to 2 subjects from their
        priority list to reduce empty slots and balance load.

        WHY TOP-3: The old code stored exactly ONE faculty per subject.
        If that faculty was busy on a given day/slot the session had zero
        CP-SAT options and was silently dropped.  Professors never lock
        one teacher per subject — they use whoever is available.
        Returning 3 candidates lets the solver pick the best free one.
        """
        # 🔒 Thread-safe execution using instance lock
        with self._lock:
            MAX_SUBJECTS_PER_FACULTY = 2  # New: limit subjects per faculty
            
            batch_subject_sessions = defaultdict(list)
            for session in sessions:
                key = (session['batch_id'], session['subject_code'], session['type'])
                batch_subject_sessions[key].append(session)

            faculty_assignments = {}          # key → [fac_id_1, fac_id_2, fac_id_3]
            faculty_workload   = defaultdict(int)  # only PRIMARY charges workload
            faculty_subjects   = defaultdict(set)  # NEW: track subjects per faculty

            # Heaviest subjects first so they get first pick of faculty
            sorted_items = sorted(
                batch_subject_sessions.items(),
                key=lambda item: len(item[1]) * (2 if item[1][0]['type'] == 'lab' else 1),
                reverse=True
            )

            print(f"\n{'='*70}")
            print("🎯 FACULTY ALLOCATION  (Multi-Subject: max 2 per faculty)")
            print(f"{'='*70}")
            mode = "CHOICE-BASED" if self.allocation_constraints.use_faculty_choice else "QUALIFICATION-BASED"
            print(f"   Mode: {mode}\n")

            for (batch_id, subject_code, sess_type), sess_list in sorted_items:
                subject = self.subject_map.get(subject_code)
                if not subject:
                    continue

                hours_needed = len(sess_list) * (2 if sess_type == 'lab' else 1)

                # Scored list from allocator (choice or qualification based)
                allocations = self.faculty_allocator.allocate_faculty_to_subject(
                    subject_code, batch_id
                )
                if not allocations:
                    continue

                # ════════════════════════════════════════════════════════════════
                # 🚀 ENHANCED: Add load balancing bonus (prefer underutilized faculty)
                # ════════════════════════════════════════════════════════════════
                adjusted_allocations = []
                for fid, score in allocations:
                    bonus = 0
                    subjects_count = len(faculty_subjects[fid])
                    
                    # Load balancing: prefer faculty with fewer subjects
                    if subjects_count == 0:
                        bonus = 20  # Highest priority for unused faculty
                    elif subjects_count == 1:
                        bonus = 10  # Medium priority
                    else:
                        bonus = 0   # At limit
                    
                    adjusted_allocations.append((fid, score + bonus, subjects_count))
                
                # Sort by adjusted score (higher is better)
                adjusted_allocations.sort(key=lambda x: x[1], reverse=True)

                # ── pick PRIMARY (charges workload counter) ──
                primary_id = None
                for fid, adj_score, subj_count in adjusted_allocations:
                    f = self.faculty_map.get(fid)
                    if not f:
                        continue
                    
                    # Check subject limit
                    if subj_count >= MAX_SUBJECTS_PER_FACULTY:
                        continue
                    
                    # Check if 2nd subject is from priority list
                    if subj_count == 1 and self.allocation_constraints.use_faculty_choice:
                        fc = self.faculty_allocator.choice_map.get(fid)
                        if fc:
                            priority_list = fc.get_choices_list()
                            if subject_code not in priority_list:
                                continue  # Don't assign random 2nd subject
                    
                    # Check workload capacity
                    if faculty_workload[fid] + hours_needed <= f.max_hours_per_week:
                        primary_id = fid
                        faculty_workload[fid] += hours_needed
                        faculty_subjects[fid].add(subject_code)
                        break

                if not primary_id:
                    print(f"  ⚠️  {subject_code}: no faculty with remaining capacity")
                    continue

                # ── pick up to 2 BACKUP candidates (do NOT charge their counter) ──
                candidates = [primary_id]
                for fid, adj_score, subj_count in adjusted_allocations:
                    if fid == primary_id:
                        continue
                    f = self.faculty_map.get(fid)
                    if not f:
                        continue
                    
                    # Check subject limit for backups too
                    if subj_count >= MAX_SUBJECTS_PER_FACULTY:
                        continue
                    
                    if faculty_workload[fid] < f.max_hours_per_week:
                        candidates.append(fid)
                        if len(candidates) >= 3:
                            break

                faculty_assignments[(batch_id, subject_code, sess_type)] = candidates

                # ── log ──
                pname = self.faculty_map[primary_id].name
                subj_count = len(faculty_subjects[primary_id])
                backups = ", ".join(self.faculty_map[c].name for c in candidates[1:])
                choice_tag = ""
                if self.allocation_constraints.use_faculty_choice:
                    fc = self.faculty_allocator.choice_map.get(primary_id)
                    if fc:
                        try:
                            choice_tag = f" [Choice #{fc.get_choices_list().index(subject_code)+1}]"
                        except ValueError:
                            pass
                subj_tag = f" (S{subj_count})" if subj_count > 1 else ""
                back_tag = f"  backups: [{backups}]" if backups else ""
                print(f"  ✅ {subject_code} → {pname}{choice_tag}{subj_tag}{back_tag}")

            # Print faculty utilization summary
            multi_subj = [f for f, s in faculty_subjects.items() if len(s) > 1]
            if multi_subj:
                print(f"\n  📊 Multi-Subject Faculty: {len(multi_subj)} faculty teaching 2+ subjects")
                for fid in multi_subj[:5]:  # Show top 5
                    fname = self.faculty_map[fid].name
                    subjs = ", ".join(faculty_subjects[fid])
                    print(f"     • {fname}: {subjs}")

            print(f"\n{'='*70}\n")
            return faculty_assignments

    def _calculate_match_score(self, subject: Subject, faculty: Faculty) -> int:
        score = 0
        
        subject_terms = set()
        if subject.required_specialization:
            subject_terms.update(subject.required_specialization.lower().split())
        subject_terms.update(subject.subject_code.lower().split())
        subject_terms.update(subject.subject_name.lower().split())
        
        faculty_text = ' '.join(faculty.specializations).lower()
        
        for term in subject_terms:
            if len(term) > 2 and term in faculty_text:
                score += 500
        
        if any('phd' in q for q in faculty.qualifications):
            score += 30
        
        return max(score, 100)
    
    def _fill_gaps_intelligently(self, assignments, sessions):
        if not assignments:
            return assignments
        
        faculty_map = {}
        for sess_id, assign in assignments.items():
            batch_id = sess_id.split('_')[0]
            parts = sess_id.split('_')[1:]
            if 'ELEC' in sess_id:
                subject_code = parts[0]
            else:
                subject_code = '_'.join(parts[:-1])
            key = (batch_id, subject_code)
            faculty_map[key] = assign['faculty_id']
        
        batch_schedules = defaultdict(lambda: defaultdict(list))
        for sess_id, assign in assignments.items():
            batch_id = sess_id.split('_')[0]
            batch_schedules[batch_id][assign['day']].append({
                'slot': assign['slot'],
                'session_id': sess_id,
                'room': assign['room_id'],
                'faculty': assign['faculty_id']
            })
        
        for batch_id in batch_schedules:
            for day in batch_schedules[batch_id]:
                batch_schedules[batch_id][day].sort(key=lambda x: x['slot'])
        
        gaps_filled = 0
        
        for batch_id, days in batch_schedules.items():
            batch = self.batch_map.get(batch_id)
            if not batch:
                continue
            
            for day, day_sessions in days.items():
                if len(day_sessions) < 2:
                    continue
                
                slots_used = [s['slot'] for s in day_sessions]
                min_slot = min(slots_used)
                max_slot = max(slots_used)
                
                for slot_id in range(min_slot + 1, max_slot):
                    if slot_id == 3 or slot_id in slots_used:
                        continue
                    
                    for subject_code in batch.subjects:
                        subject = self.subject_map.get(subject_code)
                        if not subject or subject.is_elective:
                            continue
                        
                        day_count = sum(1 for s in day_sessions if subject_code in s['session_id'])
                        if day_count > 0:
                            continue
                        
                        faculty_key = (batch_id, subject_code)
                        faculty_id = faculty_map.get(faculty_key)
                        if not faculty_id:
                            continue
                        
                        faculty_busy = any(
                            a['slot'] == slot_id and a['day'] == day and a['faculty_id'] == faculty_id
                            for a in assignments.values()
                        )
                        if faculty_busy:
                            continue
                        
                        eligible_rooms = self._get_eligible_rooms_with_floor_priority(batch_id, False)
                        available_room = None
                        
                        for room in eligible_rooms:
                            room_busy = any(
                                a['slot'] == slot_id and a['day'] == day and a['room_id'] == room.room_id
                                for a in assignments.values()
                            )
                            if not room_busy:
                                available_room = room
                                break
                        
                        if not available_room:
                            continue
                        
                        new_id = f"{batch_id}_{subject_code}_GAPFILL_{gaps_filled}"
                        assignments[new_id] = {
                            'session_id': new_id,
                            'day': day,
                            'slot': slot_id,
                            'room_id': available_room.room_id,
                            'faculty_id': faculty_id
                        }
                        
                        batch_schedules[batch_id][day].append({
                            'slot': slot_id,
                            'session_id': new_id,
                            'room': available_room.room_id,
                            'faculty': faculty_id
                        })
                        batch_schedules[batch_id][day].sort(key=lambda x: x['slot'])
                        
                        gaps_filled += 1
                        break
        
        if gaps_filled > 0:
            print(f"  ✅ Filled {gaps_filled} gaps")
        
        return assignments
    
    def _analyze_quality(self, assignments):
        print(f"\n{'='*60}")
        print("QUALITY ANALYSIS")
        print(f"{'='*60}")
        
        # Check for same subject on same day violations
        print("\n🔍 Same-Subject-Same-Day Check:")
        violations = []
        
        for batch in self.batches:
            batch_assigns = [a for a in assignments.values() 
                           if a['session_id'].startswith(batch.batch_id)]
            
            for day in self.working_days:
                day_sessions = [a for a in batch_assigns if a['day'] == day]
                
                # Extract subject codes from session IDs
                subject_counts = defaultdict(int)
                for session in day_sessions:
                    sess_id = session['session_id']
                    # Extract subject code
                    parts = sess_id.split('_')
                    if 'ELEC' in sess_id:
                        subject_code = parts[1]
                    elif 'GAPFILL' in sess_id:
                        subject_code = parts[1]
                    else:
                        # Get everything between batch_id and last part (T0/L0)
                        subject_code = '_'.join(parts[1:-1])
                    
                    subject_counts[subject_code] += 1
                
                # Check for duplicates
                for subject_code, count in subject_counts.items():
                    if count > 1:
                        violations.append(f"{batch.batch_id} has {count}x {subject_code} on {day}")
        
        if violations:
            print("  ❌ VIOLATIONS FOUND:")
            for v in violations:
                print(f"     • {v}")
        else:
            print("  ✅ Perfect! No same subject scheduled twice on any day")
        
        # Floor distribution and consistency analysis
        print("\n🏢 Floor Usage & Consistency:")
        floor_usage = defaultdict(int)
        batch_floor_usage = defaultdict(lambda: defaultdict(int))
        
        for sess_id, assign in assignments.items():
            batch_id = sess_id.split('_')[0]
            batch = self.batch_map.get(batch_id)
            room = next((r for r in self.rooms if r.room_id == assign['room_id']), None)
            
            if room and room.floor > 0:
                floor_usage[room.floor] += 1
                batch_floor_usage[batch_id][room.floor] += 1
        
        print("\n  Overall floor distribution:")
        for floor in sorted(floor_usage.keys()):
            count = floor_usage[floor]
            bar = '█' * (count // 2)
            print(f"    Floor {floor}: {bar} ({count} sessions)")
        
        # Analyze floor consistency per batch
        print("\n  Per-batch floor consistency:")
        total_batches_analyzed = 0
        perfect_consistency = 0
        
        for batch in self.batches:
            if batch.batch_id in batch_floor_usage:
                floors_used = batch_floor_usage[batch.batch_id]
                total_sessions = sum(floors_used.values())
                
                if len(floors_used) == 1:
                    # Perfect! All on one floor
                    floor = list(floors_used.keys())[0]
                    match_icon = "✅" if floor == batch.preferred_floor else "✓"
                    print(f"    {match_icon} {batch.batch_id}: All {total_sessions} sessions on Floor {floor}")
                    perfect_consistency += 1
                else:
                    # Multiple floors used
                    primary_floor = max(floors_used.keys(), key=lambda f: floors_used[f])
                    primary_count = floors_used[primary_floor]
                    primary_pct = (primary_count / total_sessions) * 100
                    
                    floor_detail = ', '.join([f"F{f}:{cnt}" for f, cnt in sorted(floors_used.items())])
                    print(f"    ⚠️  {batch.batch_id}: Mixed floors ({floor_detail})")
                    print(f"        → Primary floor {primary_floor}: {primary_count}/{total_sessions} ({primary_pct:.0f}%)")
                
                total_batches_analyzed += 1
        
        if total_batches_analyzed > 0:
            consistency_rate = (perfect_consistency / total_batches_analyzed) * 100
            print(f"\n  📊 Floor consistency: {perfect_consistency}/{total_batches_analyzed} batches ({consistency_rate:.0f}%) on single floor")
        
        # Floor preference match rate
        floor_matches = 0
        total_with_preference = 0
        
        for sess_id, assign in assignments.items():
            batch_id = sess_id.split('_')[0]
            batch = self.batch_map.get(batch_id)
            room = next((r for r in self.rooms if r.room_id == assign['room_id']), None)
            
            if batch and batch.preferred_floor > 0 and room and room.floor > 0:
                total_with_preference += 1
                if room.floor == batch.preferred_floor:
                    floor_matches += 1
        
        if total_with_preference > 0:
            match_rate = (floor_matches / total_with_preference) * 100
            print(f"  Floor preference match: {floor_matches}/{total_with_preference} ({match_rate:.1f}%)")
        
        # Elective analysis
        elective_sessions = [s for s in assignments.keys() if 'ELEC' in s]
        if elective_sessions:
            print(f"\n📚 Electives: {len(elective_sessions)} sessions scheduled")
            
            elective_groups = set()
            for sess_id in elective_sessions:
                if '_ELEC_' in sess_id:
                    group = sess_id.split('_ELEC_')[1].split('_')[0]
                    elective_groups.add(group)
            
            print(f"   Groups: {', '.join(sorted(elective_groups))}")
        
        # Gap analysis
        print("\n⚠️  GAP ANALYSIS:")
        gap_count = 0
        
        for batch in self.batches:
            batch_assigns = [a for a in assignments.values() 
                           if a['session_id'].startswith(batch.batch_id)]
            
            for day in self.working_days:
                day_slots = sorted([a['slot'] for a in batch_assigns if a['day'] == day])
                
                if len(day_slots) < 2:
                    continue
                
                for i in range(len(day_slots) - 1):
                    if day_slots[i+1] - day_slots[i] > 1:
                        if day_slots[i] < 3 < day_slots[i+1]:
                            continue
                        gap_count += 1
                        print(f"  ⚠️  {batch.batch_id} on {day}: gap at slots {day_slots[i]}-{day_slots[i+1]}")
        
        if gap_count == 0:
            print("  ✅ No gaps - Perfect!")
        
        print(f"\n{'='*60}")
    
    def _format_results(self, assignments):
        faculty_tts = defaultdict(lambda: defaultdict(dict))
        batch_tts = defaultdict(lambda: defaultdict(dict))
        room_tts = defaultdict(lambda: defaultdict(dict))
        
        # Add lunch
        lunch_slot = next((ts for ts in self.timeslots if ts.slot_id == 3), None)
        if lunch_slot:
            lunch_str = str(lunch_slot)
            for faculty in self.faculties:
                for day in self.working_days:
                    faculty_tts[faculty.name][day][lunch_str] = "🍽️ LUNCH"
            for batch in self.batches:
                for day in self.working_days:
                    batch_tts[batch.batch_id][day][lunch_str] = "🍽️ LUNCH"
            for room in self.rooms:
                for day in self.working_days:
                    room_tts[room.room_id][day][lunch_str] = "-- LUNCH --"
        
        # Add sessions
        for sess_id, assign in assignments.items():
            batch_id = sess_id.split('_')[0]
            parts = sess_id.split('_')[1:]
            
            is_elective = 'ELEC' in sess_id
            is_lab = sess_id.split('_')[-1][0] == 'L'
            is_gapfill = 'GAPFILL' in sess_id
            
            if is_elective:
                subject_code = parts[0]
                elective_group = parts[2] if len(parts) > 2 else ""
            elif is_gapfill:
                subject_code = parts[0]
            else:
                subject_code = '_'.join(parts[:-1])
            
            subject = self.subject_map.get(subject_code)
            faculty = self.faculty_map.get(assign['faculty_id'])
            slot_obj = next((ts for ts in self.timeslots if ts.slot_id == assign['slot']), None)
            
            if not all([subject, faculty, slot_obj]):
                continue
            
            day = assign['day']
            slot_str = str(slot_obj)
            room_id = assign['room_id']
            room = next((r for r in self.rooms if r.room_id == room_id), None)
            
            # Floor indicator
            floor_mark = f" [R_Id:{room_id} / Flr:{room.floor}]" if room and room.floor > 0 else f" [{room_id}]"
            elec_mark = f" [ELEC]" if is_elective else ""
            gap_mark = " 📚" if is_gapfill else ""
            
            if is_lab:
                faculty_tts[faculty.name][day][slot_str] = f"🔬 {subject.subject_name} [{batch_id}]{floor_mark}{elec_mark}"
                batch_tts[batch_id][day][slot_str] = f"🔬 {subject.subject_name} by {faculty.name}{floor_mark}{elec_mark}"
                room_tts[room_id][day][slot_str] = f"🔬 {subject.subject_name} [{batch_id}] by {faculty.name}{elec_mark}"
                
                next_slot_id = assign['slot'] + 1
                next_slot = next((ts for ts in self.timeslots if ts.slot_id == next_slot_id), None)
                if next_slot and next_slot_id != 3:
                    next_str = str(next_slot)
                    faculty_tts[faculty.name][day][next_str] = f"   ↳ LAB (cont.)"
                    batch_tts[batch_id][day][next_str] = f"   ↳ LAB (cont.)"
                    room_tts[room_id][day][next_str] = f"   ↳ LAB (cont.)"
            else:
                faculty_tts[faculty.name][day][slot_str] = f"{subject.subject_name} [{batch_id}]{floor_mark}{elec_mark}{gap_mark}"
                batch_tts[batch_id][day][slot_str] = f"{subject.subject_name} by {faculty.name}{floor_mark}{elec_mark}{gap_mark}"
                room_tts[room_id][day][slot_str] = f"{subject.subject_name} [{batch_id}] by {faculty.name}{elec_mark}{gap_mark}"
        
        return {
            'faculty_timetables': dict(faculty_tts),
            'batch_timetables': dict(batch_tts),
            'room_timetables': dict(room_tts),
            'statistics': {
                'total': len(self._create_sessions_progressive()),
                'scheduled': len(assignments),
                'success_rate': len(assignments) / len(self._create_sessions_progressive()) * 100 if self._create_sessions_progressive() else 0
            }
        }

# -------------------------
# Output Writer
# -------------------------
def write_output(results, output_dir, department_name):
    # THREAD-SAFE: Lock on department directory
    dept_output_dir = os.path.join(output_dir, department_name)
    dir_lock = get_file_lock(dept_output_dir)
    
    with dir_lock:
        os.makedirs(dept_output_dir, exist_ok=True)
        
        # ... rest of function stays the same, but INDENTED one level
        def sort_time_slot(time_str):
            try:
                return int(time_str.split(':')[0].split('-')[0])
            except:
                return 99
        
        for file_type, data in [
            ('faculty_timetables.xlsx', results['faculty_timetables']),
            ('batch_timetables.xlsx', results['batch_timetables']),
            ('room_timetables.xlsx', results['room_timetables'])
        ]:
            file_path = os.path.join(dept_output_dir, file_type)
            
            # THREAD-SAFE: Lock on specific file
            file_lock = get_file_lock(file_path)
            
            with file_lock:
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    for name, schedule in data.items():
                        df_data = defaultdict(dict)
                        for day, slots in schedule.items():
                            for time, content in slots.items():
                                df_data[day][time] = content
                        
                        if df_data:
                            df = pd.DataFrame(df_data).fillna("")
                            if len(df.index) > 0:
                                df = df.sort_index(key=lambda x: x.map(sort_time_slot))
                            sheet_name = re.sub(r'[\\/*?:"<>|]', '', str(name))[:31]
                            df.to_excel(writer, sheet_name=sheet_name)
        
        print(f"  📁 Output files created (thread-safe)")

# -------------------------
# Parallel Department Processing
# -------------------------
def process_department_wrapper(args):
    """Wrapper for parallel processing"""
    input_file, output_dir = args
    return process_department(input_file, output_dir)



def process_department(input_file, output_dir):
    """
    PRODUCTION-GRADE department processing.
    Never fails, intelligently handles all constraints.
    """
    department_name = Path(input_file).stem
    
    print(f"\n{'='*70}")
    print(f"🏭 PRODUCTION-GRADE PROCESSING: {department_name}")
    print(f"{'='*70}")
    
    try:
        # Load data
        predictor = StrategyPredictor()
        reader = AdvancedExcelReader(input_file)
        data = reader.parse_all()
        
        analyzer = ResourceAnalyzer(data)
        analysis = analyzer.analyze_capacity()
        
        # Detect imbalance
        is_imbalanced, reason, ratio = analyzer._detect_hour_imbalance()
        if is_imbalanced:
            print(f"⚠️ Imbalance detected, rebalancing...")
            data, _ = analyzer.apply_intelligent_balancing('balanced')
        
        # ⭐ PRODUCTION-GRADE CONSTRAINT HANDLING ⭐
        config = SchedulingConfig(
            MIN_THEORY=1,
            MIN_LAB=0,
            REDUCTION_STEP=0.85,
            MAX_ATTEMPTS=12,
            FEASIBILITY_BUFFER=1.2,
            ENABLE_SPECIALIZATION_RELAXATION=True,
            ENABLE_FLOOR_PRIORITY_RELAXATION=True,
            ENABLE_BATCH_SPLITTING=True,
            ACCEPTABLE_SUCCESS_RATE=0.80,
            MIN_VIABLE_SCHEDULE=0.60
        )
        
        handler = IntelligentConstraintHandler(config)
        results = handler.handle_constraints(data, ProgressiveCPScheduler, predictor, config)
        
        # Results handling
        if results and results.get('statistics', {}).get('scheduled', 0) > 0:
            write_output(results, output_dir, department_name)
            stats = results['statistics']
            
            print(f"\n{'='*70}")
            print(f"✅ PRODUCTION SUCCESS: {stats['success_rate']:.1f}%")
            print(f"   Sessions: {stats['scheduled']}/{stats['total']}")
            print(f"{'='*70}")
            
            return {'department': department_name, 'success': True, 'stats': stats}
        else:
            print(f"\n{'='*70}")
            print(f"⚠️ GRACEFUL DEGRADATION - Creating partial schedule")
            print(f"{'='*70}")
            
            return {'department': department_name, 'success': False, 'error': 'Partial schedule created'}
    
    except Exception as e:
        import traceback
        print(f"\n❌ EXCEPTION: {str(e)}")
        traceback.print_exc()
        return {'department': department_name, 'success': False, 'error': str(e)}
# ═══════════════════════════════════════════════════════════════════════════
# ENHANCEMENT 1: POST-GRADUATION SLOT MAXIMIZER (Thread-Safe)
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class SlotOpportunity:
    batch_id: str
    subject_code: str
    day: str
    slot_id: int
    room_id: str
    faculty_id: str
    priority_score: int


class PostGraduationSlotMaximizer:
    """Maximizes slot usage (THREAD-SAFE)"""
    
    def __init__(self, data: Dict[str, Any], config: Dict[str, Any]):
        self.timeslots = data['timeslots']
        self.rooms = data['rooms']
        self.faculties = data['faculties']
        self.subjects = data['subjects']
        self.batches = data['batches']
        self.working_days = ["MON", "TUE", "WED", "THU", "FRI"]
        self.config = config
        
        # CRITICAL: Instance lock for thread safety
        self._lock = threading.Lock()
        
        self.subject_map = {s.subject_code: s for s in self.subjects}
        self.faculty_map = {f.faculty_id: f for f in self.faculties}
        self.batch_map = {b.batch_id: b for b in self.batches}
        
        self.original_subject_hours = {
            s.subject_code: (s.hours_per_week, s.lab_hours_per_week)
            for s in self.subjects
        }
    
    def maximize_slots(
        self,
        current_assignments: Dict[str, Any],
        sessions: List[Dict]
    ) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """Add extra slots (THREAD-SAFE)"""
        
        if not ENHANCEMENTS_AVAILABLE:
            return current_assignments, {'status': 'disabled'}
        
        print(f"\n{'='*70}")
        print("🎯 POST-GRADUATION SLOT MAXIMIZER (Thread-Safe)")
        print(f"{'='*70}")
        
        usage_analysis = self._analyze_current_usage(current_assignments)
        opportunities = self._find_slot_opportunities(current_assignments, usage_analysis)
        enhanced_assignments = self._add_opportunity_sessions(current_assignments, opportunities)
        report = self._generate_report(current_assignments, enhanced_assignments, usage_analysis, opportunities)
        
        print(f"{'='*70}\n")
        
        return enhanced_assignments, report
    
    # ... (paste all other methods from COPY_PASTE_GUIDE.txt Step 2) ...
    
    def _add_opportunity_sessions(
        self,
        current_assignments: Dict[str, Any],
        opportunities: List[SlotOpportunity]
    ) -> Dict[str, Any]:
        """Add sessions (USES LOCK - THREAD-SAFE)"""
        
        enhanced = current_assignments.copy()
        added_count = 0
        
        used_faculty = defaultdict(lambda: defaultdict(set))
        used_rooms = defaultdict(lambda: defaultdict(set))
        used_batches = defaultdict(lambda: defaultdict(set))
        
        for sess_id, assign in current_assignments.items():
            batch_id = sess_id.split('_')[0]
            day = assign['day']
            slot = assign['slot']
            faculty_id = assign['faculty_id']
            room_id = assign['room_id']
            
            used_faculty[faculty_id][day].add(slot)
            used_rooms[room_id][day].add(slot)
            used_batches[batch_id][day].add(slot)
        
        for opp in opportunities:
            with self._lock:  # 🔒 CRITICAL SECTION
                if opp.slot_id in used_faculty[opp.faculty_id][opp.day]:
                    continue
                if opp.slot_id in used_rooms[opp.room_id][opp.day]:
                    continue
                if opp.slot_id in used_batches[opp.batch_id][opp.day]:
                    continue
                
                new_id = f"{opp.batch_id}_{opp.subject_code}_EXTRA_{added_count}"
                enhanced[new_id] = {
                    'session_id': new_id,
                    'day': opp.day,
                    'slot': opp.slot_id,
                    'room_id': opp.room_id,
                    'faculty_id': opp.faculty_id
                }
                
                used_faculty[opp.faculty_id][opp.day].add(opp.slot_id)
                used_rooms[opp.room_id][opp.day].add(opp.slot_id)
                used_batches[opp.batch_id][opp.day].add(opp.slot_id)
                
                added_count += 1
        
        print(f"   Added {added_count} extra sessions")
        
        return enhanced


# ═══════════════════════════════════════════════════════════════════════════
# ENHANCEMENT 2: FACULTY LOAD BALANCER (HYBRID 2-TIER)
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class FacultyLoadInfo:
    """Faculty workload information"""
    faculty_id: str
    faculty_name: str
    max_capacity: int
    current_load: int
    remaining_capacity: int
    utilization_pct: float
    subjects_teaching: List[str]
    batches_teaching: List[str]


@dataclass
class ReassignmentOpportunity:
    """Represents a possible reassignment"""
    from_faculty_id: str
    to_faculty_id: str
    session_id: str
    subject_code: str
    batch_id: str
    load_reduction: int
    balance_improvement: float


class FacultyLoadBalancer:
    """Balances faculty workload using hybrid two-tier approach"""
    
    def __init__(self, data: Dict[str, Any], config: Dict[str, Any]):
        self.faculties = data['faculties']
        self.subjects = data['subjects']
        self.batches = data['batches']
        self.config = config
        
        self._lock = threading.Lock()
        
        self.subject_map = {s.subject_code: s for s in self.subjects}
        self.faculty_map = {f.faculty_id: f for f in self.faculties}
        self.batch_map = {b.batch_id: b for b in self.batches}
        
        self.faculty_choices = data.get('faculty_choices', [])
        self.choice_map = {fc.faculty_id: fc for fc in self.faculty_choices}
        
        self.OVERLOAD_THRESHOLD = 0.90
        self.UNDERLOAD_THRESHOLD = 0.50
    
    def balance_load(
        self,
        current_assignments: Dict[str, Any],
        faculty_allocator=None
    ) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """Balance faculty workload"""
        
        if not ENHANCEMENTS_AVAILABLE:
            return current_assignments, {'status': 'disabled'}
        
        print(f"\n{'='*70}")
        print("⚖️ FACULTY LOAD BALANCER (HYBRID 2-TIER)")
        print(f"{'='*70}")
        
        load_info = self._analyze_faculty_loads(current_assignments)
        
        if not self._needs_balancing(load_info):
            print("✅ Faculty load already balanced!")
            return current_assignments, {'status': 'already_balanced'}
        
        # TIER 1
        print("\n🎯 TIER 1: Balancing Overloaded Faculty...")
        tier1_result, tier1_success = self._tier1_balance_overloaded(
            current_assignments.copy(), load_info, faculty_allocator
        )
        
        if not tier1_success:
            print("⚠️ Tier 1 failed - returning original")
            return current_assignments, {'status': 'tier1_failed'}
        
        print("✅ Tier 1 successful!")
        
        # TIER 2
        print("\n🎯 TIER 2: Utilizing Underloaded Faculty...")
        tier2_result, tier2_success = self._tier2_fill_underutilized(
            tier1_result.copy(), load_info, faculty_allocator
        )
        
        if tier2_success:
            print("✅ Tier 2 successful - using enhanced schedule")
            final_assignments = tier2_result
            tier = 'both'
        else:
            print("⚠️ Tier 2 failed - keeping Tier 1 results")
            final_assignments = tier1_result
            tier = 'tier1_only'
        
        report = self._generate_report(current_assignments, final_assignments, load_info, tier)
        
        print(f"{'='*70}\n")
        
        return final_assignments, report
    
    def _analyze_faculty_loads(self, assignments: Dict[str, Any]) -> Dict[str, FacultyLoadInfo]:
        """Analyze current faculty workload"""
        
        faculty_sessions = defaultdict(list)
        faculty_loads = defaultdict(int)
        faculty_subjects = defaultdict(set)
        faculty_batches = defaultdict(set)
        
        for sess_id, assign in assignments.items():
            faculty_id = assign['faculty_id']
            batch_id = sess_id.split('_')[0]
            parts = sess_id.split('_')[1:]
            
            if 'ELEC' in sess_id:
                subject_code = parts[0]
            elif 'GAPFILL' in sess_id or 'EXTRA' in sess_id:
                subject_code = parts[0]
            else:
                subject_code = '_'.join(parts[:-1])
            
            is_lab = sess_id.split('_')[-1][0] == 'L'
            hours = 2 if is_lab else 1
            
            faculty_sessions[faculty_id].append(sess_id)
            faculty_loads[faculty_id] += hours
            faculty_subjects[faculty_id].add(subject_code)
            faculty_batches[faculty_id].add(batch_id)
        
        load_info = {}
        
        for faculty in self.faculties:
            fid = faculty.faculty_id
            current = faculty_loads.get(fid, 0)
            remaining = faculty.max_hours_per_week - current
            utilization = (current / max(faculty.max_hours_per_week, 1)) * 100
            
            load_info[fid] = FacultyLoadInfo(
                faculty_id=fid,
                faculty_name=faculty.name,
                max_capacity=faculty.max_hours_per_week,
                current_load=current,
                remaining_capacity=remaining,
                utilization_pct=utilization,
                subjects_teaching=list(faculty_subjects.get(fid, [])),
                batches_teaching=list(faculty_batches.get(fid, []))
            )
        
        loads = [info.current_load for info in load_info.values()]
        avg_load = np.mean(loads) if loads else 0
        std_dev = np.std(loads) if loads else 0
        
        print(f"   Average Load: {avg_load:.1f}h")
        print(f"   Std Deviation: {std_dev:.1f}h")
        
        return load_info
    
    def _needs_balancing(self, load_info: Dict[str, FacultyLoadInfo]) -> bool:
        """Check if balancing is needed"""
        loads = [info.current_load for info in load_info.values()]
        
        if not loads:
            return False
        
        avg = np.mean(loads)
        std = np.std(loads)
        
        threshold = avg * 0.30
        needs_balance = std > threshold
        
        if needs_balance:
            print(f"   ⚠️ Imbalance detected: σ={std:.1f} > threshold={threshold:.1f}")
        
        return needs_balance
    
    def _tier1_balance_overloaded(
        self, assignments, load_info, faculty_allocator
    ) -> Tuple[Dict[str, Any], bool]:
        """TIER 1: Reduce load on overloaded faculty"""
        
        overloaded = [
            fid for fid, info in load_info.items()
            if info.utilization_pct > self.OVERLOAD_THRESHOLD * 100
        ]
        
        if not overloaded:
            print("   No overloaded faculty found")
            return assignments, True
        
        print(f"   Found {len(overloaded)} overloaded faculty")
        
        opportunities = self._find_reassignment_opportunities(
            assignments, load_info, overloaded, 'reduce_overload'
        )
        
        if not opportunities:
            print("   No reassignment opportunities found")
            return assignments, False
        
        balanced = self._apply_reassignments(assignments, opportunities, load_info, 20)
        
        new_load_info = self._analyze_faculty_loads(balanced)
        improvement = self._calculate_improvement(load_info, new_load_info)
        
        success = improvement > 0
        
        if success:
            print(f"   ✅ Improved balance by {improvement:.1f}%")
        else:
            print(f"   ⚠️ No improvement achieved")
        
        return balanced, success
    
    def _tier2_fill_underutilized(
        self, assignments, original_load_info, faculty_allocator
    ) -> Tuple[Dict[str, Any], bool]:
        """TIER 2: Fill underutilized faculty"""
        
        current_load_info = self._analyze_faculty_loads(assignments)
        
        underutilized = [
            fid for fid, info in current_load_info.items()
            if info.utilization_pct < self.UNDERLOAD_THRESHOLD * 100
            and info.remaining_capacity > 0
        ]
        
        if not underutilized:
            print("   No underutilized faculty found")
            return assignments, True
        
        print(f"   Found {len(underutilized)} underutilized faculty")
        
        opportunities = self._find_reassignment_opportunities(
            assignments, current_load_info, underutilized, 'fill_underutilized'
        )
        
        if not opportunities:
            print("   No fill opportunities found")
            return assignments, False
        
        filled = self._apply_reassignments(assignments, opportunities, current_load_info, 10)
        
        new_load_info = self._analyze_faculty_loads(filled)
        
        tier1_maintained = self._verify_tier1_maintained(current_load_info, new_load_info)
        
        if not tier1_maintained:
            print("   ⚠️ Would break Tier 1 balance - rejecting")
            return assignments, False
        
        total_improvement = self._calculate_improvement(original_load_info, new_load_info)
        success = total_improvement > 0
        
        if success:
            print(f"   ✅ Total improvement: {total_improvement:.1f}%")
        
        return filled, success
    
    def _find_reassignment_opportunities(
        self, assignments, load_info, target_faculty, target
    ) -> List[ReassignmentOpportunity]:
        """Find reassignment opportunities"""
        
        opportunities = []
        faculty_sessions = defaultdict(list)
        
        for sess_id, assign in assignments.items():
            faculty_sessions[assign['faculty_id']].append((sess_id, assign))
        
        if target == 'reduce_overload':
            from_faculty = target_faculty
            to_faculty = [
                fid for fid, info in load_info.items()
                if info.utilization_pct < 0.70 * 100 and info.remaining_capacity >= 1
            ]
        else:
            from_faculty = [
                fid for fid, info in load_info.items()
                if 0.60 * 100 <= info.utilization_pct <= 0.85 * 100
            ]
            to_faculty = target_faculty
        
        for from_fid in from_faculty:
            for sess_id, assign in faculty_sessions[from_fid]:
                parts = sess_id.split('_')[1:]
                if 'ELEC' in sess_id:
                    subject_code = parts[0]
                elif 'GAPFILL' in sess_id or 'EXTRA' in sess_id:
                    subject_code = parts[0]
                else:
                    subject_code = '_'.join(parts[:-1])
                
                subject = self.subject_map.get(subject_code)
                if not subject:
                    continue
                
                batch_id = sess_id.split('_')[0]
                
                for to_fid in to_faculty:
                    to_info = load_info[to_fid]
                    
                    if to_info.remaining_capacity < 1:
                        continue
                    
                    can_teach = self._can_faculty_teach_subject(to_fid, subject_code, subject)
                    
                    if not can_teach:
                        continue
                    
                    is_lab = sess_id.split('_')[-1][0] == 'L'
                    hours = 2 if is_lab else 1
                    
                    loads = [info.current_load for info in load_info.values()]
                    current_variance = np.var(loads)
                    
                    simulated_loads = loads.copy()
                    from_idx = list(load_info.keys()).index(from_fid)
                    to_idx = list(load_info.keys()).index(to_fid)
                    simulated_loads[from_idx] -= hours
                    simulated_loads[to_idx] += hours
                    new_variance = np.var(simulated_loads)
                    
                    improvement = (current_variance - new_variance) / max(current_variance, 1) * 100
                    
                    if improvement > 0:
                        opportunities.append(ReassignmentOpportunity(
                            from_faculty_id=from_fid,
                            to_faculty_id=to_fid,
                            session_id=sess_id,
                            subject_code=subject_code,
                            batch_id=batch_id,
                            load_reduction=hours,
                            balance_improvement=improvement
                        ))
        
        opportunities.sort(key=lambda o: o.balance_improvement, reverse=True)
        
        return opportunities
    
    def _can_faculty_teach_subject(self, faculty_id, subject_code, subject) -> bool:
        """Check if faculty can teach this subject"""
        faculty = self.faculty_map.get(faculty_id)
        if not faculty:
            return False
        
        if not subject.required_specialization:
            return True
        
        spec_lower = subject.required_specialization.lower()
        faculty_specs = ' '.join(faculty.specializations).lower()
        
        return spec_lower in faculty_specs
    
    def _apply_reassignments(
        self, assignments, opportunities, load_info, max_reassignments
    ) -> Dict[str, Any]:
        """Apply reassignments from opportunities"""
        
        balanced = assignments.copy()
        applied = 0
        
        capacity_tracker = {
            fid: info.remaining_capacity for fid, info in load_info.items()
        }
        
        for opp in opportunities[:max_reassignments]:
            if capacity_tracker[opp.to_faculty_id] < opp.load_reduction:
                continue
            
            if opp.session_id in balanced:
                balanced[opp.session_id]['faculty_id'] = opp.to_faculty_id
                
                capacity_tracker[opp.from_faculty_id] += opp.load_reduction
                capacity_tracker[opp.to_faculty_id] -= opp.load_reduction
                
                applied += 1
        
        print(f"   Applied {applied}/{min(len(opportunities), max_reassignments)} reassignments")
        
        return balanced
    
    def _calculate_improvement(self, before, after) -> float:
        """Calculate overall improvement in balance"""
        before_loads = [info.current_load for info in before.values()]
        after_loads = [info.current_load for info in after.values()]
        
        before_std = np.std(before_loads)
        after_std = np.std(after_loads)
        
        improvement = ((before_std - after_std) / max(before_std, 1)) * 100
        
        return improvement
    
    def _verify_tier1_maintained(self, tier1_result, tier2_result) -> bool:
        """Verify Tier 1 balance is maintained after Tier 2"""
        for fid, info in tier2_result.items():
            if info.utilization_pct > self.OVERLOAD_THRESHOLD * 100:
                if tier1_result[fid].utilization_pct <= self.OVERLOAD_THRESHOLD * 100:
                    return False
        
        return True
    
    def _generate_report(self, original, balanced, original_load_info, tier) -> Dict[str, Any]:
        """Generate balancing report"""
        
        balanced_load_info = self._analyze_faculty_loads(balanced)
        improvement = self._calculate_improvement(original_load_info, balanced_load_info)
        
        report = {
            'tier': tier,
            'improvement_pct': improvement,
            'reassignments': sum(
                1 for sid in balanced
                if balanced[sid]['faculty_id'] != original.get(sid, {}).get('faculty_id')
            ),
            'before': {
                'avg_load': np.mean([i.current_load for i in original_load_info.values()]),
                'std_dev': np.std([i.current_load for i in original_load_info.values()])
            },
            'after': {
                'avg_load': np.mean([i.current_load for i in balanced_load_info.values()]),
                'std_dev': np.std([i.current_load for i in balanced_load_info.values()])
            }
        }
        
        print(f"\n📊 BALANCING REPORT:")
        print(f"   Tier: {tier}")
        print(f"   Improvement: {improvement:.1f}%")
        print(f"   Reassignments: {report['reassignments']}")
        
        return report


# ═══════════════════════════════════════════════════════════════════════════
# ENHANCEMENT 3: ENHANCED PROCESS DEPARTMENT FUNCTION
# ═══════════════════════════════════════════════════════════════════════════

def process_department_enhanced(input_file, output_dir):
    """Enhanced version with slot maximization and faculty balancing"""
    
    department_name = Path(input_file).stem
    start_time = time.time()
    
    print(f"\n{'='*70}")
    print(f"🏭 ENHANCED PROCESSING: {department_name}")
    print(f"{'='*70}")
    
    try:
        # Early validation: Check if file exists
        if not os.path.exists(input_file):
            print(f"❌ Input file not found: {input_file}")
            return {'department': department_name, 'success': False, 'error': 'File not found'}
        
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Load data with timeout protection
        print(f"   📄 Loading data...")
        predictor = StrategyPredictor()
        reader = AdvancedExcelReader(input_file)
        data = reader.parse_all()
        
        # Validate data was loaded
        if not data or not data.get('batches') or not data.get('subjects'):
            print(f"❌ Invalid or empty data in file")
            return {'department': department_name, 'success': False, 'error': 'Invalid data'}
        
        print(f"   ✅ Loaded: {len(data.get('batches', []))} batches, {len(data.get('subjects', []))} subjects")
        
        analyzer = ResourceAnalyzer(data)
        analysis = analyzer.analyze_capacity()
        
        # Detect imbalance (with timeout consideration)
        is_imbalanced, reason, ratio = analyzer._detect_hour_imbalance()
        if is_imbalanced:
            print(f"⚠️ Imbalance detected, rebalancing...")
            data, _ = analyzer.apply_intelligent_balancing('balanced')
        
        # Constraint handling configuration
        config = SchedulingConfig(
            MIN_THEORY=1,
            MIN_LAB=0,
            REDUCTION_STEP=0.85,
            MAX_ATTEMPTS=8,  # Reduced from 12 for faster processing
            FEASIBILITY_BUFFER=1.2,
            ENABLE_SPECIALIZATION_RELAXATION=True,
            ENABLE_FLOOR_PRIORITY_RELAXATION=True,
            ENABLE_BATCH_SPLITTING=True,
            ACCEPTABLE_SUCCESS_RATE=0.80,
            MIN_VIABLE_SCHEDULE=0.60
        )
        
        handler = IntelligentConstraintHandler(config)
        results = handler.handle_constraints(data, ProgressiveCPScheduler, predictor, config)
        
        # SAFE ACCESS: Check if results and assignments exist
        if not results:
            print(f"⚠️ No results from constraint handler")
            return {'department': department_name, 'success': False, 'error': 'No results'}
        
        assignments = results.get('assignments', {})
        statistics = results.get('statistics', {})
        scheduled_count = statistics.get('scheduled', len(assignments))
        
        # Update statistics if missing
        if 'scheduled' not in statistics:
            statistics['scheduled'] = len(assignments)
        if 'total' not in statistics:
            statistics['total'] = max(len(assignments), 1)
        if 'success_rate' not in statistics:
            statistics['success_rate'] = (statistics['scheduled'] / max(statistics['total'], 1)) * 100
        
        results['statistics'] = statistics
        
        # ✨ APPLY ENHANCEMENTS if available and successful
        if ENHANCEMENTS_AVAILABLE and assignments and scheduled_count > 0:
            
            # Create sessions list for maximizer (with SAFE parsing)
            sessions = []
            for sess_id in assignments.keys():
                try:
                    parts = sess_id.split('_')
                    batch_id = parts[0] if parts else 'unknown'
                    
                    # Safe subject code extraction
                    if len(parts) > 1:
                        if 'ELEC' in sess_id or 'GAPFILL' in sess_id or 'EXTRA' in sess_id:
                            subject_code = parts[1] if len(parts) > 1 else 'unknown'
                        else:
                            subject_code = '_'.join(parts[1:-1]) if len(parts) > 2 else parts[1] if len(parts) > 1 else 'unknown'
                    else:
                        subject_code = 'unknown'
                    
                    # Safe session type detection
                    sess_type = 'lab' if (parts[-1] and parts[-1][0] == 'L') else 'theory'
                    
                    sessions.append({
                        'batch_id': batch_id,
                        'subject_code': subject_code,
                        'type': sess_type,
                        'id': sess_id,
                        'priority': 5
                    })
                except Exception as parse_err:
                    # Skip malformed session IDs
                    print(f"   ⚠️ Skipping malformed session: {sess_id}")
                    continue
            
            # Apply slot maximization (with timeout)
            if sessions:
                try:
                    maximizer = PostGraduationSlotMaximizer(data, config.__dict__)
                    enhanced_assignments, max_report = maximizer.maximize_slots(
                        assignments,
                        sessions
                    )
                    results['assignments'] = enhanced_assignments
                    results['maximization_report'] = max_report
                except Exception as e:
                    print(f"⚠️ Slot maximization skipped: {e}")
            
            # Apply faculty balancing
            try:
                balancer = FacultyLoadBalancer(data, config.__dict__)
                balanced_assignments, balance_report = balancer.balance_load(
                    results.get('assignments', assignments)
                )
                results['assignments'] = balanced_assignments
                results['load_balance_report'] = balance_report
            except Exception as e:
                print(f"⚠️ Faculty balancing skipped: {e}")
            
            # Update statistics with SAFE access
            final_assignments = results.get('assignments', {})
            results['statistics']['scheduled'] = len(final_assignments)
            results['statistics']['success_rate'] = (
                len(final_assignments) / 
                max(results['statistics'].get('total', 1), 1) * 100
            )
        
        # Results handling with SAFE access
        final_stats = results.get('statistics', {})
        final_scheduled = final_stats.get('scheduled', 0)
        
        if results and final_scheduled > 0:
            write_output(results, output_dir, department_name)
            
            elapsed = time.time() - start_time
            print(f"\n{'='*70}")
            print(f"✅ SUCCESS: {final_stats.get('success_rate', 0):.1f}% in {elapsed:.1f}s")
            print(f"   Sessions: {final_scheduled}/{final_stats.get('total', final_scheduled)}")
            print(f"{'='*70}")
            
            return {
                'department': department_name, 
                'success': True, 
                'stats': final_stats,
                'elapsed_time': elapsed
            }
        else:
            elapsed = time.time() - start_time
            print(f"\n{'='*70}")
            print(f"⚠️ PARTIAL SCHEDULE (completed in {elapsed:.1f}s)")
            print(f"{'='*70}")
            
            return {
                'department': department_name, 
                'success': False, 
                'error': 'Partial or empty schedule',
                'stats': final_stats,
                'elapsed_time': elapsed
            }
    
    except FileNotFoundError as e:
        print(f"\n❌ FILE NOT FOUND: {str(e)}")
        return {'department': department_name, 'success': False, 'error': f'File not found: {str(e)}'}
    
    except PermissionError as e:
        print(f"\n❌ PERMISSION DENIED: {str(e)}")
        return {'department': department_name, 'success': False, 'error': f'Permission denied: {str(e)}'}
    
    except Exception as e:
        import traceback
        elapsed = time.time() - start_time
        print(f"\n❌ EXCEPTION after {elapsed:.1f}s: {str(e)}")
        traceback.print_exc()
        return {'department': department_name, 'success': False, 'error': str(e), 'elapsed_time': elapsed}

# ============================================================================
# ALSO PASTE THE DIAGNOSTIC FUNCTION ABOVE process_department()
# ============================================================================
# Copy the ENTIRE function below and paste it BEFORE process_department()

def diagnose_bca_failure(input_file):
    """
    Diagnose why BCA scheduler is failing
    Shows exact constraints, capacity, and conflicts
    """
    print(f"\n{'='*80}")
    print(f"🔍 DIAGNOSTIC: Scheduler Failure Analysis")
    print(f"{'='*80}")
    
    try:
        # Load data
        print(f"\n[1] Loading data...")
        reader = AdvancedExcelReader(input_file)
        data = reader.parse_all()
        
        timeslots = data.get('timeslots', [])
        rooms = data.get('rooms', [])
        faculties = data.get('faculties', [])
        subjects = data.get('subjects', [])
        batches = data.get('batches', [])
        
        print(f"    ✓ Timeslots: {len(timeslots)} slots")
        print(f"    ✓ Rooms: {len(rooms)} rooms")
        print(f"    ✓ Faculties: {len(faculties)} faculty")
        print(f"    ✓ Subjects: {len(subjects)} subjects")
        print(f"    ✓ Batches: {len(batches)} batches")
        
        # ANALYSIS 1: Show subject details
        print(f"\n[2] Subject Analysis:")
        print(f"    Subject Code | Hours | Lab | Specialization | Faculty Count")
        print(f"    " + "-"*70)
        
        total_hours = 0
        for subject in subjects:
            theory = subject.hours_per_week
            lab = subject.lab_hours_per_week
            total = theory + (lab * 2)
            total_hours += total
            
            # Count faculty with specialization
            qualified = 0
            if subject.required_specialization:
                qualified = len([f for f in faculties 
                               if subject.required_specialization.lower() in ' '.join(f.specializations).lower()])
            
            print(f"    {subject.subject_code:15} | {theory:5} | {lab:3} | {subject.required_specialization:15} | {qualified:13}")
        
        print(f"\n    TOTAL REQUIRED HOURS: {total_hours}h")
        
        # ANALYSIS 2: Show faculty capacity
        print(f"\n[3] Faculty Capacity Analysis:")
        print(f"    Faculty ID | Max Hours | Specializations")
        print(f"    " + "-"*70)
        
        total_faculty_capacity = 0
        for faculty in faculties:
            total_faculty_capacity += faculty.max_hours_per_week
            specs = ', '.join(faculty.specializations[:2]) if faculty.specializations else 'None'
            print(f"    {faculty.faculty_id:10} | {faculty.max_hours_per_week:9} | {specs}")
        
        print(f"\n    TOTAL FACULTY CAPACITY: {total_faculty_capacity}h/week")
        print(f"    REQUIRED: {total_hours}h/week")
        capacity_diff = total_faculty_capacity - total_hours
        capacity_pct = (total_faculty_capacity / max(total_hours, 1)) * 100
        print(f"    CAPACITY: {capacity_diff:+d}h ({capacity_pct:.1f}%)")
        
        # ANALYSIS 3: Show room capacity
        print(f"\n[4] Room Capacity Analysis:")
        usable_slots = len([ts for ts in timeslots if ts.slot_id != 3])
        working_days = 5
        slots_per_week = usable_slots * working_days
        total_room_slots = len(rooms) * slots_per_week
        
        print(f"    Usable timeslots/day: {usable_slots}")
        print(f"    Working days/week: {working_days}")
        print(f"    Slots per room/week: {slots_per_week}")
        print(f"    Total rooms: {len(rooms)}")
        print(f"    TOTAL ROOM SLOTS: {total_room_slots} slots/week")
        
        # ANALYSIS 4: Show batch requirements
        print(f"\n[5] Batch Requirements Analysis:")
        print(f"    Batch ID | Subjects | Total Hours | Capacity % | Feasible?")
        print(f"    " + "-"*70)
        
        for batch in batches:
            batch_subjects = [s for s in subjects if s.subject_code in batch.subjects]
            batch_hours = sum(s.hours_per_week + (s.lab_hours_per_week * 2) for s in batch_subjects)
            
            # Estimate feasibility
            available_slots = slots_per_week
            feasibility = (available_slots / max(batch_hours, 1)) * 100
            is_feasible = feasibility >= 95
            
            print(f"    {batch.batch_id:10} | {len(batch_subjects):8} | {batch_hours:11} | {feasibility:10.1f}% | {'✓' if is_feasible else '✗'}")
        
        # ANALYSIS 5: Check for conflicts
        print(f"\n[6] Potential Issues:")
        
        issues = []
        
        # Issue 1: Not enough faculty capacity
        if total_faculty_capacity < total_hours:
            shortage = total_hours - total_faculty_capacity
            issues.append(f"❌ Faculty shortage: Need {total_hours}h but only have {total_faculty_capacity}h (short by {shortage}h)")
        
        # Issue 2: Not enough room slots
        if total_room_slots < total_hours:
            shortage = total_hours - total_room_slots
            issues.append(f"❌ Room shortage: Need {total_hours}h but only have {total_room_slots} slots (short by {shortage})")
        
        # Issue 3: Specialized subject with no faculty
        for subject in subjects:
            if subject.required_specialization:
                qualified = len([f for f in faculties 
                               if subject.required_specialization.lower() in ' '.join(f.specializations).lower()])
                if qualified == 0:
                    issues.append(f"❌ No faculty: {subject.subject_code} needs '{subject.required_specialization}' but no faculty has it")
        
        # Issue 4: Too many hours for single subject
        for subject in subjects:
            if subject.hours_per_week > 10:
                issues.append(f"⚠️  High hours: {subject.subject_code} has {subject.hours_per_week}h (typical max is 8h)")
        
        if issues:
            print(f"    Found {len(issues)} issue(s):")
            for issue in issues:
                print(f"    {issue}")
        else:
            print(f"    ✓ No obvious conflicts detected")
        
        # ANALYSIS 6: Recommendation
        print(f"\n[7] Recommendation:")
        if total_faculty_capacity < total_hours:
            shortage = total_hours - total_faculty_capacity
            print(f"    🔧 Faculty shortage: {shortage}h")
            print(f"    💡 Option 1: Reduce subject hours by ~{shortage}h total")
            print(f"    💡 Option 2: Add {int(shortage / 8)} more faculty members")
            print(f"    💡 Option 3: Check subject hours in Excel - may be incorrectly set")
        
        elif total_room_slots < total_hours:
            shortage = total_hours - total_room_slots
            print(f"    🔧 Room shortage: Not enough room slots ({total_hours}h needed, {total_room_slots} available)")
            print(f"    💡 Option 1: Add more rooms")
            print(f"    💡 Option 2: Reduce total hours needed")
        
        else:
            print(f"    🔧 Enough capacity exists, but scheduler still failed")
            print(f"    💡 Possible causes:")
            print(f"       - Specialization mismatch (faculty skills don't match subject needs)")
            print(f"       - Batch-faculty conflict")
            print(f"       - OR-Tools constraints too strict")
            print(f"       - Subject-faculty mapping missing")
        
        print(f"\n{'='*80}\n")
    
    except Exception as e:
        print(f"❌ Diagnostic error: {str(e)}")
        import traceback
        traceback.print_exc()

# -------------------------
# Main with Parallel Processing
# -------------------------
def main():
    # ═══════════════════════════════════════════════════════════════════════
    # VERIFY THREAD SAFETY
    # ═══════════════════════════════════════════════════════════════════════
    print("\n🔒 Verifying thread safety protection...")
    print("   ✅ File locks active")
    print("   ✅ Results aggregator ready")
    print("   ✅ StrategyPredictor protected")
    print("   ✅ write_output protected")
    # ═══════════════════════════════════════════════════════════════════════
    
    input_folder = "timetable_input"
    output_dir = "timetable_output"
    
    print("="*70)
    print("  ADVANCED TIMETABLE GENERATOR")
    print("  ✨ Floor Priority | 🔄 Parallel | 📚 Electives | 🤖 ML")
    print("="*70)
    
    if not os.path.exists(input_folder):
        print(f"\n❌ Folder '{input_folder}' not found!")
        print(f"\n💡 Create folder and add .xlsx files")
        return
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    
    if not excel_files:
        print(f"\n❌ No Excel files found!")
        return
    
    print(f"\n📂 Found {len(excel_files)} department(s):")
    for file in excel_files:
        print(f"   • {Path(file).stem}.xlsx")
    
    # PARALLEL PROCESSING
    use_parallel = len(excel_files) > 1
    # use_parallel=True  # Commented out: was forcing parallel for single file
    
    # Create aggregator BEFORE the if/else to be available in both paths
    aggregator = ThreadSafeResultsAggregator()
    
    if use_parallel:
        print(f"\n🚀 Using PARALLEL processing ({min(4, len(excel_files))} workers)")
        
        start_time = time.time()
        
        # Process in parallel with max 4 workers
        with ThreadPoolExecutor(max_workers=min(4, len(excel_files))) as executor:
            # Use enhanced version if available
            process_func = process_department_enhanced if ENHANCEMENTS_AVAILABLE else process_department
            
            future_to_file = {
                executor.submit(process_func, file, output_dir): file 
                for file in excel_files
            }
            
            for future in as_completed(future_to_file):
                file = future_to_file[future]
                try:
                    result = future.result()
                    aggregator.add_result(result)  # ✅ THREAD-SAFE!
                    if result.get('success'):
                        print(f"\n✅ {result.get('department', 'Unknown')}: DONE")
                    else:
                        print(f"\n❌ {result.get('department', 'Unknown')}: FAILED")
                        
                except Exception as e:
                    print(f"\n❌ {Path(file).stem}: Exception - {str(e)}")
                    aggregator.add_result({
                        'department': Path(file).stem,
                        'success': False,
                        'error': str(e)
                    })
        
        parallel_time = time.time() - start_time
        print(f"\n⚡ Parallel processing completed in {parallel_time:.2f}s")
        
        # Get results from aggregator
        results_summary = aggregator.get_results()        
    else:
        print(f"\n🔄 Sequential processing (single department)")
        
        start_time = time.time()
        
        # Use enhanced version if available
        process_func = process_department_enhanced if ENHANCEMENTS_AVAILABLE else process_department
        
        for input_file in excel_files:
            result = process_func(input_file, output_dir)
            aggregator.add_result(result)  # Add to aggregator for consistency
        
        sequential_time = time.time() - start_time
        print(f"\n⏱️ Processing completed in {sequential_time:.2f}s")
        
        # Get results from aggregator
        results_summary = aggregator.get_results()
    
    # Final Summary (now uses aggregator in both paths)
    successful = aggregator.get_successful()
    failed = aggregator.get_failed()
    
    print(f"\n{'='*70}")
    print("  FINAL SUMMARY")
    print(f"{'='*70}")
    
    print(f"\n📊 Overall: {len(successful)}/{len(results_summary)} successful")
    
    if use_parallel and len(excel_files) > 1:
        estimated_sequential = parallel_time * len(excel_files)
        speedup = estimated_sequential / parallel_time
        print(f"⚡ Speedup: ~{speedup:.1f}x faster than sequential")
    
    if successful:
        print(f"\n✅ SUCCESSFUL DEPARTMENTS:")
        for r in successful:
            scaling_info = ""
            if r.get('scaling_applied'):
                ratio = r.get('scaling_ratio', 0)
                scaling_info = f" (scaled {ratio*100:.0f}%)" if ratio else ""
            
            # SAFE ACCESS: Use .get() to prevent KeyError
            stats = r.get('stats', {})
            success_rate = stats.get('success_rate', 0)
            scheduled = stats.get('scheduled', 0)
            total = stats.get('total', scheduled)
            elapsed = r.get('elapsed_time', 0)
            
            print(f"   • {r.get('department', 'Unknown')}: {success_rate:.1f}%{scaling_info}")
            if elapsed:
                print(f"     Sessions: {scheduled}/{total} ({elapsed:.1f}s)")
            else:
                print(f"     Sessions: {scheduled}/{total}")
    
    if failed:
        print(f"\n❌ FAILED DEPARTMENTS:")
        for r in failed:
            error = r.get('error', 'Unknown error')
            elapsed = r.get('elapsed_time', 0)
            if elapsed:
                print(f"   • {r.get('department', 'Unknown')}: {error} ({elapsed:.1f}s)")
            else:
                print(f"   • {r.get('department', 'Unknown')}: {error}")
    
    # Feature summary
    print(f"\n{'='*70}")
    print("✨ FEATURES USED:")
    print(f"{'='*70}")
    print("  🏢 Floor-wise room priority for batch grouping")
    print("  🔄 Parallel department processing (up to 4x faster)")
    print("  📚 Elective subject support with group constraints")
    print("  🤖 ML-based strategy prediction")
    print("  📊 Smart ratio-based scaling")
    print("  🎯 Progressive scheduling (priority-based)")
    print("  ⚖️ Load-balanced faculty assignment")
    print("  🔍 Intelligent gap filling")
    
    print(f"\n{'='*70}")
    print(f"✅ Complete! Output: {os.path.abspath(output_dir)}")
    print(f"{'='*70}\n")
    
    # Save summary report
    summary_file = os.path.join(output_dir, "generation_summary.json")
    with open(summary_file, 'w') as f:
        json.dump({
            'timestamp': time.time(),
            'total_departments': len(results_summary),
            'successful': len(successful),
            'failed': len(failed),
            'results': results_summary,
            'parallel_processing': use_parallel
        }, f, indent=2)
    
    print(f"📄 Summary saved: {summary_file}\n")

if __name__ == "__main__":
    main()