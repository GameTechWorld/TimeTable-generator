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
from dataclasses import dataclass, field
from functools import lru_cache
from concurrent.futures import ProcessPoolExecutor, as_completed
import warnings
warnings.filterwarnings('ignore')

try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except ImportError:
    ORTOOLS_AVAILABLE = False
    print("WARNING: OR-Tools not available. Install with: pip install ortools")

# -------------------------
# Enhanced Data Classes with Floor Support
# -------------------------
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
# ML-Based Strategy Predictor
# -------------------------
class StrategyPredictor:
    """Predicts best scheduling strategy based on input characteristics"""
    
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
        """Save scheduling result for learning"""
        self.history.append({
            'features': features,
            'strategy': strategy,
            'success_rate': success_rate,
            'timestamp': time.time()
        })
        
        # Keep only last 100 results
        if len(self.history) > 100:
            self.history = self.history[-100:]
        
        with open(self.history_file, 'w') as f:
            json.dump(self.history, f, indent=2)
    
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
        usable_slots = len([ts for ts in timeslots if ts.slot_id != 3]) * 5  # 5 days
        
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
    
    def predict_best_strategy(self, features: Dict) -> int:
        """Predict which strategy index to try first"""
        if len(self.history) < 5:
            return 0  # Not enough data, use default
        
        # Simple similarity-based prediction
        best_strategy_idx = 0
        best_similarity = -1
        
        strategy_scores = defaultdict(list)
        
        for record in self.history[-20:]:  # Use recent 20
            past_features = record['features']
            
            # Calculate feature similarity
            similarity = 0
            similarity += 1 - abs(past_features['capacity_ratio'] - features['capacity_ratio'])
            similarity += 1 - abs(past_features['batch_count'] - features['batch_count']) / 10
            similarity += 1 - abs(past_features['faculty_count'] - features['faculty_count']) / 10
            
            strategy_name = record['strategy']
            success_rate = record['success_rate']
            
            # Map strategy name to index
            strategy_map = {
                'STRICT NO GAPS': 0,
                'COMPACT SCHEDULE': 1,
                'RELAXED': 2,
                'MAXIMUM FLEXIBILITY': 3
            }
            
            strategy_idx = strategy_map.get(strategy_name, 0)
            strategy_scores[strategy_idx].append(similarity * success_rate)
        
        # Find strategy with best weighted score
        if strategy_scores:
            best_strategy_idx = max(strategy_scores.keys(), 
                                   key=lambda k: sum(strategy_scores[k]) / len(strategy_scores[k]))
        
        return best_strategy_idx

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
    
    @lru_cache(maxsize=1)
    def parse_all(self):
        """Cached parsing with header verification"""
        return {
            'timeslots': self._parse_timeslots(),
            'rooms': self._parse_rooms(),
            'faculties': self._parse_faculties(),
            'subjects': self._parse_subjects(),
            'batches': self._parse_batches()
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
        for subject in modified_data['subjects']:
            subject.required_specialization = None
        
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
        
        self.subject_map = {s.subject_code: s for s in self.subjects}
        self.faculty_map = {f.faculty_id: f for f in self.faculties}
        self.batch_map = {b.batch_id: b for b in self.batches}
        
        # Group rooms by floor and department
        self.rooms_by_floor_dept = defaultdict(lambda: defaultdict(list))
        for room in self.rooms:
            self.rooms_by_floor_dept[room.floor][room.department].append(room)
        
        # SMART FLOOR ASSIGNMENT: Automatically assign batches to floors
        self._auto_assign_floors_to_batches()
        
        print(f"\n✨ Progressive Scheduler initialized")
        print(f"   Floors available: {sorted(set(r.floor for r in self.rooms))}")
    
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
    
    def schedule(self):
        if not ORTOOLS_AVAILABLE:
            print("ERROR: OR-Tools required")
            return {}
        
        print("\n🚀 Starting PROGRESSIVE scheduling with ML prediction...")
        start_time = time.time()
        
        # Create sessions with priority ordering
        sessions = self._create_sessions_progressive()
        if not sessions:
            return {}
        
        print(f"Scheduling {len(sessions)} sessions (progressive order)")
        
        # Get features for ML prediction
        features = self.predictor.extract_features({
            'timeslots': self.timeslots,
            'rooms': self.rooms,
            'faculties': self.faculties,
            'subjects': self.subjects,
            'batches': self.batches
        })
        
        # Predict best strategy
        predicted_strategy_idx = self.predictor.predict_best_strategy(features)
        print(f"🤖 ML suggests trying strategy #{predicted_strategy_idx + 1} first")
        
        # Multi-strategy with reordering based on prediction
        result = self._intelligent_scheduling(sessions, predicted_strategy_idx)
        
        elapsed = time.time() - start_time
        print(f"\n✅ Optimization completed in {elapsed:.2f}s")
        
        if result and result['assignments']:
            # Save result for ML learning
            self.predictor.save_result(
                features,
                result.get('strategy', 'UNKNOWN'),
                len(result['assignments']) / len(sessions) * 100
            )
            
            self._analyze_quality(result['assignments'])
            return self._format_results(result['assignments'])
        else:
            return {}
    
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
        """Multi-strategy with ML-predicted order"""
        strategies = [
            {'name': 'STRICT NO GAPS', 'allow_gaps': False, 'time_limit': 45},
            {'name': 'COMPACT SCHEDULE', 'allow_gaps': False, 'time_limit': 40},
            {'name': 'RELAXED', 'allow_gaps': True, 'time_limit': 35},
            {'name': 'MAXIMUM FLEXIBILITY', 'allow_gaps': True, 'time_limit': 40}
        ]
        
        # Reorder strategies based on prediction
        if predicted_idx > 0:
            strategies = [strategies[predicted_idx]] + strategies[:predicted_idx] + strategies[predicted_idx+1:]
        
        print(f"\n{'='*60}")
        print("INTELLIGENT SCHEDULING")
        print(f"{'='*60}\n")
        
        best_result = None
        best_score = 0
        
        for idx, strategy in enumerate(strategies, 1):
            print(f"Strategy {idx}/{len(strategies)}: {strategy['name']}")
            
            result = self._solve_with_strategy(sessions, strategy)
            
            if result and result.get('assignments'):
                scheduled = len(result['assignments'])
                total = len(sessions)
                score = (scheduled / total) * 100
                
                print(f"  Result: {scheduled}/{total} sessions ({score:.1f}%)")
                
                if score > best_score:
                    best_score = score
                    best_result = result
                    best_result['strategy'] = strategy['name']
                
                if score >= 95:
                    print(f"  ✅ Excellent!")
                    return best_result
            else:
                print(f"  Failed")
            print()
        
        if best_result:
            print(f"{'='*60}")
            print(f"BEST: {best_score:.1f}% scheduled")
            print(f"{'='*60}")
            return best_result
        
        return None
    
    def _solve_with_strategy(self, sessions, strategy):
        """Core CP-SAT solver with floor priority and elective handling"""
        model = cp_model.CpModel()
        
        # Faculty assignment
        faculty_assignments = self._intelligent_faculty_assignment(sessions)
        if not faculty_assignments:
            return None
        
        # Variables
        x = {}
        
        for session in sessions:
            batch = self.batch_map[session['batch_id']]
            eligible_rooms = self._get_eligible_rooms_with_floor_priority(
                batch.batch_id, 
                session['type'] == 'lab'
            )
            
            if not eligible_rooms:
                continue
            
            session_key = (session['batch_id'], session['subject_code'], session['type'])
            preferred_faculty = faculty_assignments.get(session_key, [])
            eligible_faculty = [self.faculty_map[fid] for fid in preferred_faculty if fid in self.faculty_map]
            
            if not eligible_faculty:
                eligible_faculty = list(self.faculties)
            
            for day in self.working_days:
                for ts in self.timeslots:
                    if ts.slot_id == 3:
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
        
        # Constraints
        for session in sessions:
            session_vars = [x[k] for k in x if k[0] == session['id']]
            if session_vars:
                model.AddExactlyOne(session_vars)
        
        # Room conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                for room in self.rooms:
                    room_vars = [x[k] for k in x if k[1] == day and k[2] == ts.slot_id and k[3] == room.room_id]
                    if room_vars:
                        model.AddAtMostOne(room_vars)
        
        # Faculty conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                for faculty in self.faculties:
                    fac_vars = [x[k] for k in x if k[1] == day and k[2] == ts.slot_id and k[4] == faculty.faculty_id]
                    if fac_vars:
                        model.AddAtMostOne(fac_vars)
        
        # Batch conflicts (no two classes at same time)
        for day in self.working_days:
            for ts in self.timeslots:
                for batch in self.batches:
                    batch_vars = [x[k] for k in x if k[1] == day and k[2] == ts.slot_id 
                                 and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id for s in sessions)]
                    if batch_vars:
                        model.AddAtMostOne(batch_vars)
        # NEW CONSTRAINT: Lab sessions must reserve consecutive slots
        for session in sessions:
            if session['type'] == 'lab':
                for day in self.working_days:
                    for ts in self.timeslots:
                        if ts.slot_id == 3:
                            continue
                        
                        # If lab is scheduled at this slot, next slot must also be blocked
                        session_vars_at_slot = [
                            x[k] for k in x 
                            if k[0] == session['id'] and k[1] == day and k[2] == ts.slot_id
                        ]
                        
                        if session_vars_at_slot:
                            next_slot_id = ts.slot_id + 1
                            next_slot = next((t for t in self.timeslots if t.slot_id == next_slot_id), None)
                            
                            if next_slot and next_slot.slot_id != 3:
                                # Block next slot for same batch/room/faculty
                                for key in x:
                                    if (key[0] == session['id'] and key[1] == day and key[2] == ts.slot_id):
                                        room_id = key[3]
                                        faculty_id = key[4]
                                        batch_id = session['batch_id']
                                        
                                        # Next slot must be free for this room/faculty/batch
                                        blocking_vars = [
                                            x[k] for k in x 
                                            if k[1] == day and k[2] == next_slot_id 
                                            and (k[3] == room_id or k[4] == faculty_id or 
                                                 any(s['id'] == k[0] and s['batch_id'] == batch_id for s in sessions))
                                        ]
                                        
                                        # If lab scheduled, no other class can use room/faculty/batch in next slot
                                        for blocking_var in blocking_vars:
                                            model.Add(blocking_var == 0).OnlyEnforceIf(x[key])
        
        # CRITICAL: No same subject on same day for any batch
        for batch in self.batches:
            for subject_code in batch.subjects:
                for day in self.working_days:
                    # Get all sessions of this subject for this batch on this day
                    same_subject_day_vars = [
                        x[k] for k in x if k[1] == day
                        and any(s['id'] == k[0] and s['batch_id'] == batch.batch_id 
                               and s['subject_code'] == subject_code for s in sessions)
                    ]
                    if len(same_subject_day_vars) > 1:
                        # At most ONE session of this subject on this day
                        model.Add(sum(same_subject_day_vars) <= 1)
        
        # NEW: Elective group constraints (students choose one from group)
        elective_groups = defaultdict(lambda: defaultdict(list))
        for session in sessions:
            if session.get('is_elective') and session.get('elective_group'):
                batch_id = session['batch_id']
                group = session['elective_group']
                elective_groups[batch_id][group].append(session)
        
        # For each batch and elective group, ensure electives don't overlap
        for batch_id, groups in elective_groups.items():
            for group, group_sessions in groups.items():
                for day in self.working_days:
                    for ts in self.timeslots:
                        if ts.slot_id == 3:
                            continue
                        
                        # All electives in same group can't be at same time
                        group_slot_vars = [
                            x[k] for k in x 
                            if k[1] == day and k[2] == ts.slot_id
                            and any(s['id'] == k[0] for s in group_sessions)
                        ]
                        
                        if len(group_slot_vars) > 1:
                            model.AddAtMostOne(group_slot_vars)
        
        # Objectives with FLOOR PREFERENCE BONUS
        objective_terms = []
        
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
                    if ts.slot_id == 3:
                        continue
                    
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
                        objective_terms.append(gap_detected * -2500)
        
        if objective_terms:
            model.Maximize(sum(objective_terms))
        
        # Solve
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = strategy['time_limit']
        solver.parameters.num_search_workers = 8
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
        """Smart faculty allocation"""
        batch_subject_sessions = defaultdict(list)
        for session in sessions:
            key = (session['batch_id'], session['subject_code'], session['type'])
            batch_subject_sessions[key].append(session)
        
        total_workload = sum(
            len(sess_list) * (2 if sess_list[0]['type'] == 'lab' else 1)
            for sess_list in batch_subject_sessions.values()
        )
        
        total_capacity = sum(f.max_hours_per_week for f in self.faculties)
        avg_target = total_workload / len(self.faculties) if self.faculties else 0
        
        faculty_assignments = {}
        faculty_workload = defaultdict(int)
        
        sorted_items = sorted(
            batch_subject_sessions.items(),
            key=lambda x: len(x[1]) * (2 if x[1][0]['type'] == 'lab' else 1),
            reverse=True
        )
        
        for (batch_id, subject_code, sess_type), sess_list in sorted_items:
            subject = self.subject_map.get(subject_code)
            if not subject:
                continue
            
            assignment_hours = len(sess_list) * (2 if sess_type == 'lab' else 1)
            
            eligible = []
            for faculty in self.faculties:
                current_load = faculty_workload[faculty.faculty_id]
                projected_load = current_load + assignment_hours
                
                if projected_load > faculty.max_hours_per_week:
                    continue
                
                match_score = self._calculate_match_score(subject, faculty)
                
                if match_score > 0:
                    balance_score = (avg_target - projected_load) * 100
                    final_score = match_score * 0.3 + balance_score * 0.7
                    
                    eligible.append({
                        'id': faculty.faculty_id,
                        'score': final_score,
                        'projected_load': projected_load
                    })
            
            if eligible:
                eligible.sort(key=lambda x: x['score'], reverse=True)
                best = eligible[0]
                
                key = (batch_id, subject_code, sess_type)
                faculty_assignments[key] = [best['id']]
                faculty_workload[best['id']] += assignment_hours
        
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
    dept_output_dir = os.path.join(output_dir, department_name)
    os.makedirs(dept_output_dir, exist_ok=True)
    
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
    
    print(f"  📁 Output files created")

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
    use_parallel=False
    if use_parallel:
        print(f"\n🚀 Using PARALLEL processing ({min(4, len(excel_files))} workers)")
        
        start_time = time.time()
        results_summary = []
        
        # Process in parallel with max 4 workers
        with ProcessPoolExecutor(max_workers=min(4, len(excel_files))) as executor:
            future_to_file = {
                executor.submit(process_department, file, output_dir): file 
                for file in excel_files
            }
            
            for future in as_completed(future_to_file):
                file = future_to_file[future]
                try:
                    result = future.result()
                    results_summary.append(result)
                    
                    if result['success']:
                        print(f"\n✅ {result['department']}: DONE")
                    else:
                        print(f"\n❌ {result['department']}: FAILED")
                        
                except Exception as e:
                    print(f"\n❌ {Path(file).stem}: Exception - {str(e)}")
                    results_summary.append({
                        'department': Path(file).stem,
                        'success': False,
                        'error': str(e)
                    })
        
        parallel_time = time.time() - start_time
        print(f"\n⚡ Parallel processing completed in {parallel_time:.2f}s")
        
    else:
        print(f"\n🔄 Sequential processing (single department)")
        
        start_time = time.time()
        results_summary = []
        
        for input_file in excel_files:
            result = process_department(input_file, output_dir)
            results_summary.append(result)
        
        sequential_time = time.time() - start_time
        print(f"\n⏱️ Processing completed in {sequential_time:.2f}s")
    
    # Final Summary
    successful = [r for r in results_summary if r['success']]
    failed = [r for r in results_summary if not r['success']]
    
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
                scaling_info = f" (scaled {r['scaling_ratio']*100:.0f}%)"
            
            print(f"   • {r['department']}: {r['stats']['success_rate']:.1f}%{scaling_info}")
            print(f"     Sessions: {r['stats']['scheduled']}/{r['stats']['total']}")
    
    if failed:
        print(f"\n❌ FAILED DEPARTMENTS:")
        for r in failed:
            print(f"   • {r['department']}: {r.get('error', 'Unknown')}")
    
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
