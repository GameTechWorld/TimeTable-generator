"""
Simplified Timetable Generator - No Credit Validation

This version focuses purely on scheduling without NEP 2020 credit compliance.
It generates conflict-free timetables based on basic constraints only.

Features:
- Basic constraint satisfaction (no conflicts)
- Faculty-subject matching (relaxed)
- Room allocation
- Time slot assignment
- Excel input/output

Removed:
- Credit structure validation
- NEP 2020 compliance checks
- Complex prerequisite validation
- Course category requirements
"""

import os
import copy
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple, Any
from collections import defaultdict
import re

# External Libraries
import pandas as pd

# Optional fuzzy matching
try:
    import Levenshtein
    def similarity(a, b):
        return Levenshtein.ratio(a, b)
except ImportError:
    import difflib
    def similarity(a, b):
        return difflib.SequenceMatcher(None, a, b).ratio()

# OR-Tools for constraint solving
try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except ImportError:
    ORTOOLS_AVAILABLE = False

# -------------------------
# Simplified Data Models
# -------------------------
@dataclass
class TimeSlot:
    slot_id: int
    start_time: str
    end_time: str
    
    def __str__(self): 
        return f"{self.start_time}-{self.end_time}"

@dataclass
class Room:
    room_id: str
    capacity: int
    room_type: str
    department: str

@dataclass
class Faculty:
    faculty_id: str
    name: str
    qualifications: List[str] = field(default_factory=list)
    specializations: List[str] = field(default_factory=list)
    max_hours_per_week: int = 20
    availability: List[Tuple[str, int]] = field(default_factory=list)

@dataclass
class Subject:
    subject_code: str
    subject_name: str
    hours_per_week: int = 4
    requires_lab: bool = False
    lab_hours_per_week: int = 0
    required_room_type: str = "classroom"

@dataclass
class StudentBatch:
    batch_id: str
    department: str
    student_count: int
    subjects: List[str] = field(default_factory=list)

@dataclass
class ClassSession:
    session_id: str
    batch_id: str
    subject_code: str
    session_type: str  # "theory" or "lab"
    assigned_day: Optional[str] = None
    assigned_slot: Optional[int] = None
    assigned_room: Optional[str] = None
    assigned_faculty: Optional[str] = None

# -------------------------
# Utilities
# -------------------------
def normalize(s: str) -> str:
    return str(s).strip().lower() if s is not None else ""

def parse_room_type(s: str) -> str:
    s2 = normalize(s)
    if any(k in s2 for k in ["computer", "lab", "it"]): return "computer_lab"
    if "physics" in s2: return "physics_lab"
    if "chemistry" in s2 or "chem" in s2: return "chemistry_lab"
    if "seminar" in s2: return "seminar_hall"
    return "classroom"

def fuzzy_find(cols: List[str], target: str, cutoff: float = 0.6) -> Optional[str]:
    t = normalize(target)
    scores = sorted([(c, similarity(normalize(c), t)) for c in cols], key=lambda x: x[1], reverse=True)
    return scores[0][0] if scores and scores[0][1] >= cutoff else None

# -------------------------
# Simplified Excel Adapter
# -------------------------
class SimplifiedExcelAdapter:
    def __init__(self, path: str):
        if not os.path.exists(path):
            raise FileNotFoundError(f"Input file not found at: {path}")
        self.wb = pd.read_excel(path, sheet_name=None)
        self.sheets = {normalize(k): v for k, v in self.wb.items()}

    def get_sheet_df(self, candidates: List[str]) -> Optional[pd.DataFrame]:
        for c in candidates:
            df = self.sheets.get(normalize(c))
            if df is not None and not df.empty:
                return df
        return None
    
    def _parse_list_from_cell(self, cell_content):
        if pd.isna(cell_content):
            return []
        return [item.strip() for item in str(cell_content).split(',') if item.strip()]

    def _parse_day_slots(self, cell_content):
        if pd.isna(cell_content):
            return []
        day_slots = []
        for token in str(cell_content).split(','):
            parts = token.strip().upper().replace("-", "_").split('_')
            if len(parts) >= 2 and parts[-1].isdigit():
                day_slots.append(("_".join(parts[:-1]), int(parts[-1])))
        return day_slots

    def _parse_timeslots(self) -> List[TimeSlot]:
        df = self.get_sheet_df(["timeslots", "time_slots", "periods"])
        if df is None:
            # Create default time slots
            return [
                TimeSlot(1, "08:00", "09:00"), TimeSlot(2, "09:00", "10:00"),
                TimeSlot(3, "10:00", "11:00"), TimeSlot(4, "11:30", "12:30"),
                TimeSlot(5, "12:30", "13:30"), TimeSlot(6, "14:30", "15:30"),
                TimeSlot(7, "15:30", "16:30"), TimeSlot(8, "16:30", "17:30")
            ]
        
        cols = list(df.columns)
        id_col = fuzzy_find(cols, "slot_id") or fuzzy_find(cols, "period_id")
        s_col = fuzzy_find(cols, "start_time")
        e_col = fuzzy_find(cols, "end_time")
        
        if not all([id_col, s_col, e_col]):
            print("WARNING: Timeslots sheet incomplete, using defaults")
            return self._parse_timeslots()  # Use defaults
        
        return [TimeSlot(int(r[id_col]), str(r[s_col]), str(r[e_col])) 
                for _, r in df.iterrows() if pd.notna(r[id_col])]

    def _parse_rooms(self) -> List[Room]:
        df = self.get_sheet_df(["rooms", "classrooms"])
        if df is None:
            print("WARNING: No rooms sheet found, creating default rooms")
            return [
                Room("R101", 60, "classroom", "General"),
                Room("R102", 60, "classroom", "General"), 
                Room("LAB01", 30, "computer_lab", "CSE"),
                Room("LAB02", 30, "computer_lab", "ECE")
            ]
        
        cols = list(df.columns)
        id_col = fuzzy_find(cols, "room_id")
        cap_col = fuzzy_find(cols, "capacity")
        type_col = fuzzy_find(cols, "room_type")
        dept_col = fuzzy_find(cols, "department")
        
        if not id_col:
            print("ERROR: No room_id column found")
            return []
        
        rooms = []
        for _, row in df.iterrows():
            if pd.notna(row[id_col]):
                rooms.append(Room(
                    room_id=str(row[id_col]),
                    capacity=int(row.get(cap_col, 60)) if cap_col and pd.notna(row.get(cap_col)) else 60,
                    room_type=parse_room_type(str(row.get(type_col, "classroom"))) if type_col else "classroom",
                    department=str(row.get(dept_col, "General")) if dept_col else "General"
                ))
        return rooms

    def _parse_faculties(self) -> List[Faculty]:
        df = self.get_sheet_df(["faculties", "teachers", "staff", "faculty"])
        if df is None:
            print("WARNING: No faculty sheet found")
            return []
        
        cols = list(df.columns)
        id_col = fuzzy_find(cols, "faculty_id")
        name_col = fuzzy_find(cols, "name")
        qual_col = fuzzy_find(cols, "qualifications")
        spec_col = fuzzy_find(cols, "specializations")
        max_h_col = fuzzy_find(cols, "max_hours_per_week")
        avail_col = fuzzy_find(cols, "availability")
        
        if not all([id_col, name_col]):
            print("ERROR: Faculty sheet missing required columns")
            return []
        
        faculties = []
        for _, row in df.iterrows():
            if pd.notna(row[id_col]):
                faculties.append(Faculty(
                    faculty_id=str(row[id_col]),
                    name=str(row[name_col]),
                    qualifications=self._parse_list_from_cell(row.get(qual_col)) if qual_col else [],
                    specializations=self._parse_list_from_cell(row.get(spec_col)) if spec_col else [],
                    max_hours_per_week=int(row.get(max_h_col, 20)) if max_h_col and pd.notna(row.get(max_h_col)) else 20,
                    availability=self._parse_day_slots(row.get(avail_col)) if avail_col else []
                ))
        return faculties

    def _parse_subjects(self) -> List[Subject]:
        df = self.get_sheet_df(["subjects", "courses"])
        if df is None:
            print("WARNING: No subjects sheet found, creating minimal default subjects")
            return [
                Subject("CS101", "Programming", 4, True, 2, "computer_lab"),
                Subject("MATH101", "Mathematics", 4, False, 0, "classroom"),
                Subject("PHY101", "Physics", 4, False, 0, "classroom")
            ]
        
        cols = list(df.columns)
        print(f"DEBUG: Subjects sheet columns: {cols}")
        
        code_col = fuzzy_find(cols, "subject_code")
        name_col = fuzzy_find(cols, "subject_name") 
        hours_col = fuzzy_find(cols, "hours_per_week")
        lab_col = fuzzy_find(cols, "requires_lab")
        lab_hours_col = fuzzy_find(cols, "lab_hours_per_week")
        room_type_col = fuzzy_find(cols, "required_room_type")
        
        print(f"DEBUG: Found columns - code: {code_col}, name: {name_col}, hours: {hours_col}")
        
        if not code_col:
            print("ERROR: No subject_code column found, trying alternatives...")
            # Try alternatives
            for col in cols:
                if 'code' in col.lower() or 'id' in col.lower():
                    code_col = col
                    print(f"DEBUG: Using {col} as subject_code")
                    break
        
        if not name_col:
            print("ERROR: No subject_name column found, trying alternatives...")
            for col in cols:
                if 'name' in col.lower() or 'title' in col.lower():
                    name_col = col
                    print(f"DEBUG: Using {col} as subject_name")
                    break
        
        if not all([code_col, name_col]):
            print("ERROR: Cannot find basic subject columns, using defaults")
            return [
                Subject("CS101", "Programming", 4, True, 2, "computer_lab"),
                Subject("MATH101", "Mathematics", 4, False, 0, "classroom")
            ]
        
        subjects = []
        for idx, row in df.iterrows():
            try:
                if pd.notna(row[code_col]) and str(row[code_col]).strip():
                    
                    # Extract values with defaults
                    subject_code = str(row[code_col]).strip()
                    subject_name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else subject_code
                    hours_per_week = 4  # Default
                    requires_lab = False
                    lab_hours_per_week = 0
                    required_room_type = "classroom"
                    
                    # Try to get hours_per_week
                    if hours_col and pd.notna(row[hours_col]):
                        try:
                            hours_per_week = int(float(row[hours_col]))
                        except (ValueError, TypeError):
                            print(f"DEBUG: Invalid hours value for {subject_code}, using default 4")
                    
                    # Try to get lab information
                    if lab_col and pd.notna(row[lab_col]):
                        lab_value = str(row[lab_col]).lower()
                        requires_lab = lab_value in ['true', '1', 'yes', 'y']
                    
                    if lab_hours_col and pd.notna(row[lab_hours_col]):
                        try:
                            lab_hours_per_week = int(float(row[lab_hours_col]))
                            if lab_hours_per_week > 0:
                                requires_lab = True
                        except (ValueError, TypeError):
                            pass
                    
                    # Try to get room type
                    if room_type_col and pd.notna(row[room_type_col]):
                        required_room_type = str(row[room_type_col]).strip()
                    elif requires_lab:
                        required_room_type = "computer_lab"  # Default for labs
                    
                    subject = Subject(
                        subject_code=subject_code,
                        subject_name=subject_name,
                        hours_per_week=hours_per_week,
                        requires_lab=requires_lab,
                        lab_hours_per_week=lab_hours_per_week,
                        required_room_type=required_room_type
                    )
                    subjects.append(subject)
                    print(f"DEBUG: Added subject {subject_code} - {subject_name} ({hours_per_week}h, lab={requires_lab})")
                    
            except Exception as e:
                print(f"WARNING: Error parsing subject row {idx}: {e}")
                continue
        
        if not subjects:
            print("WARNING: No valid subjects found, creating defaults")
            return [
                Subject("DEFAULT_SUBJ", "Default Subject", 4, False, 0, "classroom")
            ]
        
        print(f"INFO: Successfully parsed {len(subjects)} subjects")
        return subjects

    def _parse_batches(self) -> List[StudentBatch]:
        df = self.get_sheet_df(["batches", "student_batches", "classes"])
        if df is None:
            print("WARNING: No batches sheet found")
            return []
        
        cols = list(df.columns)
        id_col = fuzzy_find(cols, "batch_id")
        dept_col = fuzzy_find(cols, "department")
        count_col = fuzzy_find(cols, "student_count")
        subjects_col = fuzzy_find(cols, "subjects")
        
        if not all([id_col, count_col]):
            print("ERROR: Batches sheet missing required columns")
            return []
        
        batches = []
        for _, row in df.iterrows():
            if pd.notna(row[id_col]):
                batches.append(StudentBatch(
                    batch_id=str(row[id_col]),
                    department=str(row.get(dept_col, "General")) if dept_col else "General",
                    student_count=int(row[count_col]),
                    subjects=self._parse_list_from_cell(row.get(subjects_col)) if subjects_col else []
                ))
        return batches

    def parse(self) -> Dict[str, Any]:
        """Parse all data with detailed debugging"""
        print("INFO: Starting to parse Excel data...")
        
        parsed_data = {
            "timeslots": self._parse_timeslots(),
            "rooms": self._parse_rooms(),
            "faculties": self._parse_faculties(),
            "subjects": self._parse_subjects(),
            "batches": self._parse_batches()
        }
        
        # Debug output
        print(f"INFO: Parsed data summary:")
        print(f"  - Timeslots: {len(parsed_data['timeslots'])}")
        print(f"  - Rooms: {len(parsed_data['rooms'])}")
        print(f"  - Faculty: {len(parsed_data['faculties'])}")
        print(f"  - Subjects: {len(parsed_data['subjects'])}")
        print(f"  - Batches: {len(parsed_data['batches'])}")
        
        # Detailed debug for key data
        if parsed_data['rooms']:
            print(f"\nRoom details:")
            for room in parsed_data['rooms'][:3]:  # Show first 3
                print(f"  - {room.room_id}: capacity={room.capacity}, type={room.room_type}")
        
        if parsed_data['faculties']:
            print(f"\nFaculty details:")
            for faculty in parsed_data['faculties'][:3]:  # Show first 3
                print(f"  - {faculty.faculty_id} ({faculty.name}): specializations={faculty.specializations}")
        
        if parsed_data['subjects']:
            print(f"\nSubject details:")
            for subject in parsed_data['subjects'][:3]:  # Show first 3
                print(f"  - {subject.subject_code} ({subject.subject_name}): {subject.hours_per_week}h/week, lab={subject.requires_lab}")
        
        if parsed_data['batches']:
            print(f"\nBatch details:")
            for batch in parsed_data['batches']:
                print(f"  - {batch.batch_id} ({batch.department}): {batch.student_count} students, subjects={batch.subjects}")
        
        return parsed_data

# -------------------------
# Simplified Scheduler
# -------------------------
class SimplifiedScheduler:
    def __init__(self, timeslots: List[TimeSlot], rooms: List[Room], faculties: List[Faculty],
                 subjects: List[Subject], batches: List[StudentBatch], 
                 working_days: List[str] = None):
        
        self.timeslots = sorted(timeslots, key=lambda t: t.slot_id)
        self.rooms = rooms
        self.faculties = faculties
        self.subjects = subjects
        self.batches = batches
        self.working_days = working_days or ["MON", "TUE", "WED", "THU", "FRI"]

        # Create lookup maps
        self.subject_map = {s.subject_code: s for s in subjects}
        self.faculty_map = {f.faculty_id: f for f in faculties}
        self.room_map = {r.room_id: r for r in rooms}
        self.batch_map = {b.batch_id: b for b in batches}
        
        # Break slots (lunch break)
        self.break_mask = {("MON", 4), ("TUE", 4), ("WED", 4), ("THU", 4), ("FRI", 4)}  # Slot 4 lunch break

        self.sessions: List[ClassSession] = self._create_sessions()
        self._compute_eligibles()

    def _create_sessions(self) -> List[ClassSession]:
        sessions = []
        sid_counter = 0
        
        for batch in self.batches:
            for subject_code in batch.subjects:
                subject = self.subject_map.get(subject_code)
                if not subject:
                    print(f"WARNING: Subject {subject_code} not found for batch {batch.batch_id}")
                    continue
                
                # Create theory sessions
                for _ in range(subject.hours_per_week):
                    sessions.append(ClassSession(
                        session_id=f"S{sid_counter}",
                        batch_id=batch.batch_id,
                        subject_code=subject_code,
                        session_type="theory"
                    ))
                    sid_counter += 1
                
                # Create lab sessions
                for _ in range(subject.lab_hours_per_week):
                    sessions.append(ClassSession(
                        session_id=f"S{sid_counter}",
                        batch_id=batch.batch_id,
                        subject_code=subject_code,
                        session_type="lab"
                    ))
                    sid_counter += 1
        
        return sessions

    def _compute_eligibles(self):
        """Compute eligible faculty and rooms - MAXIMALLY FLEXIBLE VERSION"""
        self.eligible_faculties: Dict[str, List[str]] = defaultdict(list)
        self.eligible_rooms: Dict[str, List[str]] = defaultdict(list)
        
        print(f"DEBUG: Computing eligibilities for {len(self.sessions)} sessions...")
        print(f"DEBUG: Available resources: {len(self.rooms)} rooms, {len(self.faculties)} faculty")
        
        for session in self.sessions:
            subject = self.subject_map.get(session.subject_code)
            if not subject:
                print(f"DEBUG: Subject {session.subject_code} not found")
                continue
                
            batch = self.batch_map[session.batch_id]
            print(f"DEBUG: Processing {subject.subject_name} for {batch.batch_id} (students: {batch.student_count})")

            # MAXIMALLY FLEXIBLE ROOM ASSIGNMENT
            # Any room that can fit the students is eligible
            for room in self.rooms:
                room_suitable = False
                
                # Basic capacity check
                if room.capacity >= batch.student_count:
                    room_suitable = True
                    print(f"DEBUG: Room {room.room_id} suitable for {subject.subject_name} (capacity: {room.capacity} >= {batch.student_count})")
                else:
                    # Even if capacity is slightly less, allow it (maybe some students absent)
                    if room.capacity >= batch.student_count * 0.8:  # Allow 20% flexibility
                        room_suitable = True
                        print(f"DEBUG: Room {room.room_id} suitable for {subject.subject_name} with flexibility (capacity: {room.capacity} vs {batch.student_count})")
                
                if room_suitable:
                    self.eligible_rooms[session.session_id].append(room.room_id)
            
            # MAXIMALLY FLEXIBLE FACULTY ASSIGNMENT  
            # ANY faculty can teach ANY subject (maximum flexibility)
            for faculty in self.faculties:
                can_teach = True  # Start with assumption anyone can teach anything
                
                # Optional: Very loose subject-faculty matching (but still allow if no match)
                subject_keywords = session.subject_code.lower()
                faculty_keywords = ' '.join(faculty.specializations).lower()
                
                # Just for scoring preference, not for exclusion
                preference_match = False
                if any(keyword in faculty_keywords for keyword in ['cs', 'computer', 'programming', 'software']) and 'cs' in subject_keywords:
                    preference_match = True
                elif any(keyword in faculty_keywords for keyword in ['ec', 'electronic', 'circuit', 'vlsi']) and 'ec' in subject_keywords:
                    preference_match = True
                elif any(keyword in faculty_keywords for keyword in ['ph', 'physic']) and 'ph' in subject_keywords:
                    preference_match = True
                elif any(keyword in faculty_keywords for keyword in ['ma', 'math']) and 'ma' in subject_keywords:
                    preference_match = True
                
                # Add faculty regardless of specialization match
                self.eligible_faculties[session.session_id].append(faculty.faculty_id)
                
                if preference_match:
                    print(f"DEBUG: Faculty {faculty.faculty_id} preferred for {subject.subject_name}")
                else:
                    print(f"DEBUG: Faculty {faculty.faculty_id} available for {subject.subject_name} (flexible assignment)")
            
            print(f"DEBUG: Session {session.session_id} - Eligible rooms: {len(self.eligible_rooms[session.session_id])}, Eligible faculty: {len(self.eligible_faculties[session.session_id])}")
        
        # Summary debug info
        sessions_without_rooms = sum(1 for sid in self.eligible_rooms if not self.eligible_rooms[sid])
        sessions_without_faculty = sum(1 for sid in self.eligible_faculties if not self.eligible_faculties[sid])
        
        print(f"DEBUG: Sessions without eligible rooms: {sessions_without_rooms}")
        print(f"DEBUG: Sessions without eligible faculty: {sessions_without_faculty}")
        
        if sessions_without_rooms > 0:
            print("DEBUG: Adding emergency room assignments...")
            # Emergency: assign largest room to sessions without rooms
            largest_room = max(self.rooms, key=lambda r: r.capacity) if self.rooms else None
            if largest_room:
                for session in self.sessions:
                    if not self.eligible_rooms[session.session_id]:
                        self.eligible_rooms[session.session_id].append(largest_room.room_id)
                        print(f"DEBUG: Emergency assignment - {largest_room.room_id} to session {session.session_id}")
        
        if sessions_without_faculty > 0:
            print("DEBUG: Adding emergency faculty assignments...")
            # Emergency: assign first available faculty to sessions without faculty
            if self.faculties:
                emergency_faculty = self.faculties[0]
                for session in self.sessions:
                    if not self.eligible_faculties[session.session_id]:
                        self.eligible_faculties[session.session_id].append(emergency_faculty.faculty_id)
                        print(f"DEBUG: Emergency assignment - {emergency_faculty.faculty_id} to session {session.session_id}")

    def _basic_validation(self) -> bool:
        """Enhanced validation with debugging and flexible acceptance"""
        print("\nINFO: Performing enhanced validation...")
        
        issues = []
        warnings = []
        
        # Check data completeness
        if not self.sessions:
            issues.append("No sessions found - check that batches have subjects assigned")
        
        if not self.rooms:
            issues.append("No rooms found - check rooms sheet")
            
        if not self.faculties:
            issues.append("No faculty found - check faculties sheet")
        
        # Check individual sessions
        for session in self.sessions:
            subject = self.subject_map.get(session.subject_code)
            if not subject:
                issues.append(f"Subject {session.subject_code} not found in subjects sheet")
                continue
                
            batch = self.batch_map[session.batch_id]
            
            # Room availability
            eligible_rooms = self.eligible_rooms[session.session_id]
            if not eligible_rooms:
                issues.append(f"No suitable rooms for {subject.subject_name} ({batch.batch_id})")
                print(f"  DEBUG: Room details for {subject.subject_name}:")
                print(f"    - Required capacity: {batch.student_count}")
                print(f"    - Session type: {session.session_type}")
                print(f"    - Required room type: {subject.required_room_type}")
                for room in self.rooms:
                    print(f"    - Available: {room.room_id} (capacity: {room.capacity}, type: {room.room_type})")
            else:
                print(f"  OK: {subject.subject_name} has {len(eligible_rooms)} eligible rooms")
            
            # Faculty availability  
            eligible_faculty = self.eligible_faculties[session.session_id]
            if not eligible_faculty:
                issues.append(f"No faculty available for {subject.subject_name} ({batch.batch_id})")
                print(f"  DEBUG: Faculty details for {subject.subject_name}:")
                for faculty in self.faculties:
                    print(f"    - Available: {faculty.faculty_id} ({faculty.name}) - specializations: {faculty.specializations}")
            else:
                print(f"  OK: {subject.subject_name} has {len(eligible_faculty)} eligible faculty")
        
        # Print summary
        print(f"\nValidation Summary:")
        print(f"  - Total sessions: {len(self.sessions)}")
        print(f"  - Critical issues: {len(issues)}")
        print(f"  - Warnings: {len(warnings)}")
        
        if issues:
            print("\nCritical Issues Found:")
            for issue in issues[:10]:  # Limit output
                print(f"  - {issue}")
            if len(issues) > 10:
                print(f"  ... and {len(issues) - 10} more issues")
        
        # Be more permissive - only fail if there are no sessions at all
        if not self.sessions:
            print("ERROR: Cannot proceed without any sessions")
            return False
        
        # If we have some eligible assignments, proceed even with warnings
        total_eligible_rooms = sum(len(rooms) for rooms in self.eligible_rooms.values())
        total_eligible_faculty = sum(len(faculty) for faculty in self.eligible_faculties.values())
        
        if total_eligible_rooms > 0 and total_eligible_faculty > 0:
            print("INFO: Some eligible assignments found - proceeding with flexible scheduling")
            return True
        else:
            print("ERROR: No valid assignments possible - check your data")
            return False

    def schedule(self, time_limit: int = 60) -> Dict[str, Any]:
        if not ORTOOLS_AVAILABLE:
            return {"error": "OR-Tools library required. Install with: pip install ortools"}
        
        if not self._basic_validation():
            return {
                "error": "Validation failed. Check faculty and room assignments.",
                "statistics": {"total": len(self.sessions), "scheduled": 0, "unscheduled": len(self.sessions)}
            }

        print("INFO: Starting simplified scheduling...")
        result = self._solve_schedule(time_limit)
        return self._format_result(result) if result else {"error": "Failed to generate schedule"}

    def _solve_schedule_with_consecutive_classes(self, time_limit: int = 60) -> Optional[Dict[str, Any]]:
        """Enhanced solver with stronger consecutive constraints"""
        model = cp_model.CpModel()
        x = {}
        
        # Group sessions by batch for easier consecutive scheduling
        batch_sessions = defaultdict(list)
        for session in self.sessions:
            batch_sessions[session.batch_id].append(session)
        
        print(f"INFO: Scheduling {len(self.sessions)} sessions for {len(batch_sessions)} batches with STRICT consecutive constraints")
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            # Check faculty availability
                            faculty = self.faculty_map[faculty_id]
                            if faculty.availability and (day, ts.slot_id) not in faculty.availability:
                                continue
                            
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        # Basic constraints
        # 1. Each session assigned exactly once
        for session in self.sessions:
            session_vars = [var for (sid, d, slot, rid, fid), var in x.items() if sid == session.session_id]
            if session_vars:
                model.AddExactlyOne(session_vars)

        # 2. No room conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for room in self.rooms:
                    room_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                               if d == day and slot == ts.slot_id and rid == room.room_id]
                    if room_vars:
                        model.AddAtMostOne(room_vars)

        # 3. No faculty conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for faculty in self.faculties:
                    faculty_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                  if d == day and slot == ts.slot_id and fid == faculty.faculty_id]
                    if faculty_vars:
                        model.AddAtMostOne(faculty_vars)

        # 4. No batch conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for batch_id in batch_sessions.keys():
                    batch_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                if d == day and slot == ts.slot_id and 
                                any(s.session_id == sid and s.batch_id == batch_id for s in self.sessions)]
                    if batch_vars:
                        model.AddAtMostOne(batch_vars)

        # 5. Faculty workload limits
        for faculty in self.faculties:
            weekly_sessions = [var for (sid, d, slot, rid, fid), var in x.items() if fid == faculty.faculty_id]
            if weekly_sessions:
                model.Add(sum(weekly_sessions) <= faculty.max_hours_per_week)

        # 6. STRICT SUBJECT AND FACULTY CONSTRAINTS
        print("INFO: Adding strict subject and faculty assignment constraints...")
        
        for batch_id, sessions in batch_sessions.items():
            print(f"INFO: Processing constraints for batch {batch_id}")
            
            # Group sessions by subject for this batch
            subject_sessions = defaultdict(list)
            for session in sessions:
                subject_sessions[session.subject_code].append(session)
            
            # CONSTRAINT 1: EXACTLY ONE SESSION PER SUBJECT PER DAY (STRICT)
            for subject_code, subj_sessions in subject_sessions.items():
                print(f"INFO: Adding constraint - max 1 session per day for {subject_code} in {batch_id}")
                for day in self.working_days:
                    day_subject_vars = []
                    for session in subj_sessions:
                        session_vars_this_day = [
                            var for (sid, d, slot, rid, fid), var in x.items()
                            if sid == session.session_id and d == day
                        ]
                        day_subject_vars.extend(session_vars_this_day)
                    
                    if day_subject_vars:
                        # HARD CONSTRAINT: At most 1 session of this subject on this day
                        model.Add(sum(day_subject_vars) <= 1)
            
            # CONSTRAINT 2: SAME FACULTY FOR ALL SESSIONS OF SAME SUBJECT (CONSISTENCY)
            for subject_code, subj_sessions in subject_sessions.items():
                if len(subj_sessions) > 1:  # Only needed if subject has multiple sessions
                    print(f"INFO: Adding faculty consistency constraint for {subject_code} in {batch_id}")
                    
                    # Create faculty assignment variables for this subject
                    faculty_assigned = {}
                    for faculty in self.faculties:
                        if faculty.faculty_id in self.eligible_faculties[subj_sessions[0].session_id]:
                            faculty_assigned[faculty.faculty_id] = model.NewBoolVar(f'faculty_{faculty.faculty_id}_assigned_to_{subject_code}_{batch_id}')
                    
                    if faculty_assigned:
                        # Exactly one faculty must be assigned to this subject for this batch
                        model.AddExactlyOne(list(faculty_assigned.values()))
                        
                        # If a faculty is assigned to this subject, they must teach ALL sessions of this subject
                        for faculty_id, faculty_var in faculty_assigned.items():
                            for session in subj_sessions:
                                session_by_this_faculty = [
                                    var for (sid, d, slot, rid, fid), var in x.items()
                                    if sid == session.session_id and fid == faculty_id
                                ]
                                
                                session_by_any_faculty = [
                                    var for (sid, d, slot, rid, fid), var in x.items()
                                    if sid == session.session_id
                                ]
                                
                                if session_by_this_faculty and session_by_any_faculty:
                                    # If faculty is assigned to subject AND session is scheduled, then this faculty must teach it
                                    session_scheduled = model.NewBoolVar(f'session_{session.session_id}_scheduled')
                                    model.Add(sum(session_by_any_faculty) >= 1).OnlyEnforceIf(session_scheduled)
                                    model.Add(sum(session_by_any_faculty) == 0).OnlyEnforceIf(session_scheduled.Not())
                                    
                                    # If both faculty assigned and session scheduled, then faculty must teach this session
                                    both_conditions = model.NewBoolVar(f'both_{faculty_id}_{session.session_id}')
                                    model.AddBoolAnd([faculty_var, session_scheduled]).OnlyEnforceIf(both_conditions)
                                    model.Add(sum(session_by_this_faculty) >= 1).OnlyEnforceIf(both_conditions)
            
            # CONSTRAINT 3: CONSECUTIVE SCHEDULING WITH GAP FILLING
            for day in self.working_days:
                # Get non-break slots for this day
                available_slots = [ts.slot_id for ts in self.timeslots if (day, ts.slot_id) not in self.break_mask]
                available_slots.sort()
                
                # Track which slots are used by this batch
                slot_used = {}
                for slot_id in available_slots:
                    slot_used[slot_id] = model.NewBoolVar(f'batch_{batch_id}_day_{day}_slot_{slot_id}')
                    
                    # Connect to actual session assignments
                    assignments_in_slot = [
                        var for (sid, d, slot, rid, fid), var in x.items()
                        if d == day and slot == slot_id and 
                        any(s.session_id == sid and s.batch_id == batch_id for s in sessions)
                    ]
                    
                    if assignments_in_slot:
                        model.Add(sum(assignments_in_slot) >= 1).OnlyEnforceIf(slot_used[slot_id])
                        model.Add(sum(assignments_in_slot) == 0).OnlyEnforceIf(slot_used[slot_id].Not())
                    else:
                        model.Add(slot_used[slot_id] == 0)
                
                # No gaps between used slots
                for i in range(len(available_slots) - 2):
                    slot_i = available_slots[i]
                    slot_mid = available_slots[i + 1]  
                    slot_j = available_slots[i + 2]
                    
                    if slot_i in slot_used and slot_mid in slot_used and slot_j in slot_used:
                        # If slots i and j are both used, then slot_mid must also be used
                        both_ends = model.NewBoolVar(f'both_ends_{batch_id}_{day}_{i}')
                        model.AddBoolAnd([slot_used[slot_i], slot_used[slot_j]]).OnlyEnforceIf(both_ends)
                        model.Add(slot_used[slot_mid] == 1).OnlyEnforceIf(both_ends)

        # 7. SUBJECT COMPLETION AND DISTRIBUTION
        objective_terms = []
        
        # Priority 1: Minimize use of later time slots
        for (sid, d, slot, rid, fid), var in x.items():
            if slot > 5:  # Very late slots
                objective_terms.append(var * 50)
            elif slot > 3:  # Afternoon slots  
                objective_terms.append(var * 20)
            elif slot > 1:  # Mid-morning slots
                objective_terms.append(var * 5)
        
        # Priority 2: Prefer fewer days with classes per batch
        for batch_id, sessions in batch_sessions.items():
            for day in self.working_days:
                day_used = model.NewBoolVar(f'day_used_{batch_id}_{day}')
                day_sessions = [
                    var for (sid, d, slot, rid, fid), var in x.items()
                    if d == day and any(s.session_id == sid and s.batch_id == batch_id for s in sessions)
                ]
                if day_sessions:
                    model.Add(sum(day_sessions) >= 1).OnlyEnforceIf(day_used)
                    model.Add(sum(day_sessions) == 0).OnlyEnforceIf(day_used.Not())
                    objective_terms.append(day_used * 30)  # Penalty for each day used
        
        # Priority 3: Balance daily workload
        for batch_id, sessions in batch_sessions.items():
            daily_counts = []
            for day in self.working_days:
                day_count = model.NewIntVar(0, len(sessions), f'daily_count_{batch_id}_{day}')
                day_assignments = [
                    var for (sid, d, slot, rid, fid), var in x.items()
                    if d == day and any(s.session_id == sid and s.batch_id == batch_id for s in sessions)
                ]
                if day_assignments:
                    model.Add(day_count == sum(day_assignments))
                else:
                    model.Add(day_count == 0)
                daily_counts.append(day_count)
            
            # Minimize variance in daily session counts
            if len(daily_counts) > 1:
                avg_sessions = len(sessions) // len(self.working_days) + 1
                for day_count in daily_counts:
                    deviation = model.NewIntVar(0, len(sessions), f'deviation_{batch_id}')
                    model.AddAbsEquality(deviation, day_count - avg_sessions)
                    objective_terms.append(deviation * 10)

        # Objective function: Minimize gaps and spread
        objective_terms = []
        
        # Heavy penalty for using later slots (encourage morning scheduling)
        for (sid, d, slot, rid, fid), var in x.items():
            if slot > 4:
                objective_terms.append(var * 20)
            elif slot > 2:
                objective_terms.append(var * 10)
        
        # Penalty for spreading batch sessions across multiple days
        for batch_id in batch_sessions.keys():
            for day in self.working_days:
                day_used = model.NewBoolVar(f'day_used_{batch_id}_{day}')
                day_sessions = [
                    var for (sid, d, slot, rid, fid), var in x.items()
                    if d == day and any(s.session_id == sid and s.batch_id == batch_id for s in batch_sessions[batch_id])
                ]
                if day_sessions:
                    model.Add(sum(day_sessions) >= 1).OnlyEnforceIf(day_used)
                    model.Add(sum(day_sessions) == 0).OnlyEnforceIf(day_used.Not())
                    objective_terms.append(day_used * 50)  # Penalty for each day used

        if objective_terms:
            model.Minimize(sum(objective_terms))

        # Solve with more time for consecutive constraints
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit * 2  # More time for complex constraints
        solver.parameters.num_search_workers = max(1, os.cpu_count() - 1) if os.cpu_count() else 1
        
        print("INFO: Solving with enhanced consecutive constraints...")
        status = solver.Solve(model)
        print(f"INFO: Enhanced solver status: {solver.StatusName(status)}")
        
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            # Extract and verify solution
            assigned = {}
            for (sid, d, slot, rid, fid), var in x.items():
                if solver.Value(var) == 1:
                    session = next(s for s in self.sessions if s.session_id == sid)
                    result_session = copy.deepcopy(session)
                    result_session.assigned_day = d
                    result_session.assigned_slot = slot
                    result_session.assigned_room = rid
                    result_session.assigned_faculty = fid
                    assigned[sid] = result_session
            
            # Verify no gaps exist
            gaps_found = self._verify_and_report_consecutiveness(assigned, detailed=True)
            
            if gaps_found == 0:
                print("SUCCESS: All batches have perfectly consecutive schedules!")
            else:
                print(f"WARNING: Found {gaps_found} gaps in batch schedules")
            
            return {
                "assigned": assigned,
                "unscheduled": len(self.sessions) - len(assigned),
                "solver_status": solver.StatusName(status),
                "gaps_found": gaps_found
            }
        else:
            print("WARNING: Enhanced consecutive scheduling failed. Trying basic approach...")
            return self._solve_basic_schedule(time_limit)
        """Alternative solver that ensures consecutive classes for each batch"""
        model = cp_model.CpModel()
        x = {}
        
        # Group sessions by batch for easier consecutive scheduling
        batch_sessions = defaultdict(list)
        for session in self.sessions:
            batch_sessions[session.batch_id].append(session)
        
        print(f"INFO: Scheduling {len(self.sessions)} sessions for {len(batch_sessions)} batches")
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            # Check faculty availability
                            faculty = self.faculty_map[faculty_id]
                            if faculty.availability and (day, ts.slot_id) not in faculty.availability:
                                continue
                            
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        # Basic constraints (same as before)
        # 1. Each session assigned exactly once
        for session in self.sessions:
            session_vars = [var for (sid, d, slot, rid, fid), var in x.items() if sid == session.session_id]
            if session_vars:
                model.AddExactlyOne(session_vars)

        # 2. No room conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for room in self.rooms:
                    room_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                               if d == day and slot == ts.slot_id and rid == room.room_id]
                    if room_vars:
                        model.AddAtMostOne(room_vars)

        # 3. No faculty conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for faculty in self.faculties:
                    faculty_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                  if d == day and slot == ts.slot_id and fid == faculty.faculty_id]
                    if faculty_vars:
                        model.AddAtMostOne(faculty_vars)

        # 4. No batch conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for batch_id in batch_sessions.keys():
                    batch_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                if d == day and slot == ts.slot_id and 
                                any(s.session_id == sid and s.batch_id == batch_id for s in self.sessions)]
                    if batch_vars:
                        model.AddAtMostOne(batch_vars)

        # 6. SIMPLIFIED CONSECUTIVE CLASSES CONSTRAINT
        for batch_id, sessions in batch_sessions.items():
            for day in self.working_days:
                # Get available slots for this day (excluding breaks)
                available_slots = [ts.slot_id for ts in self.timeslots if (day, ts.slot_id) not in self.break_mask]
                available_slots.sort()
                
                if len(available_slots) < 2:  # Need at least 2 slots to enforce consecutiveness
                    continue
                
                # Create boolean variables for each slot to indicate if batch is scheduled there
                slot_used = {}
                for slot_id in available_slots:
                    slot_used[slot_id] = model.NewBoolVar(f'slot_used_{batch_id}_{day}_{slot_id}')
                    
                    # Connect slot_used to actual assignments
                    slot_assignments = [var for (sid, d, slot, rid, fid), var in x.items()
                                      if d == day and slot == slot_id and 
                                      any(s.session_id == sid and s.batch_id == batch_id for s in sessions)]
                    
                    if slot_assignments:
                        # slot_used is true if any session is assigned to this slot
                        model.AddMaxEquality(slot_used[slot_id], slot_assignments)
                    else:
                        model.Add(slot_used[slot_id] == 0)
                
                # Simple consecutive constraint: no gaps between used slots
                # If slot i and slot i+2 are used, then slot i+1 must also be used
                for i in range(len(available_slots) - 2):
                    slot1, slot2, slot3 = available_slots[i], available_slots[i+1], available_slots[i+2]
                    
                    if slot1 in slot_used and slot2 in slot_used and slot3 in slot_used:
                        # Create implication: if slot1 and slot3 are used, then slot2 must be used
                        gap_exists = model.NewBoolVar(f'gap_{batch_id}_{day}_{i}')
                        
                        # gap_exists is true if slot1 used AND slot3 used AND slot2 not used
                        model.AddBoolAnd([slot_used[slot1], slot_used[slot3], slot_used[slot2].Not()]).OnlyEnforceIf(gap_exists)
                        
                        # Don't allow gaps
                        model.Add(gap_exists == 0)

        # Faculty workload limits
        for faculty in self.faculties:
            weekly_sessions = [var for (sid, d, slot, rid, fid), var in x.items() if fid == faculty.faculty_id]
            if weekly_sessions:
                model.Add(sum(weekly_sessions) <= faculty.max_hours_per_week)

        # Objective: Minimize spread across days (encourage concentration)
        objective_terms = []
        
        for batch_id in batch_sessions.keys():
            # Count number of days this batch has classes
            days_used = []
            for day in self.working_days:
                day_used = model.NewBoolVar(f'day_used_{batch_id}_{day}')
                day_sessions = [var for (sid, d, slot, rid, fid), var in x.items()
                              if d == day and any(s.session_id == sid and s.batch_id == batch_id 
                                                for s in batch_sessions[batch_id])]
                if day_sessions:
                    model.Add(sum(day_sessions) >= 1).OnlyEnforceIf(day_used)
                    model.Add(sum(day_sessions) == 0).OnlyEnforceIf(day_used.Not())
                    days_used.append(day_used)
            
            # Minimize number of days used (prefer concentrated schedules)
            if days_used:
                objective_terms.extend([var * 10 for var in days_used])

        if objective_terms:
            model.Minimize(sum(objective_terms))

        # Solve
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit
        solver.parameters.num_search_workers = max(1, os.cpu_count() - 1) if os.cpu_count() else 1
        
        print("INFO: Solving with consecutive class constraints...")
        status = solver.Solve(model)
        print(f"INFO: Solver status: {solver.StatusName(status)}")
        
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            # Extract solution
            assigned = {}
            for (sid, d, slot, rid, fid), var in x.items():
                if solver.Value(var) == 1:
                    session = next(s for s in self.sessions if s.session_id == sid)
                    result_session = copy.deepcopy(session)
                    result_session.assigned_day = d
                    result_session.assigned_slot = slot
                    result_session.assigned_room = rid
                    result_session.assigned_faculty = fid
                    assigned[sid] = result_session
            
            # Verify consecutive scheduling
            self._verify_consecutive_scheduling(assigned)
            
            return {
                "assigned": assigned,
                "unscheduled": len(self.sessions) - len(assigned),
                "solver_status": solver.StatusName(status)
            }
        else:
            print("WARNING: Could not find solution with consecutive constraints. Trying relaxed approach...")
            return self._solve_basic_schedule(time_limit)  # Fallback to basic method

    def _solve_basic_schedule(self, time_limit: int = 60) -> Optional[Dict[str, Any]]:
        """Basic scheduling without consecutive constraints - fallback method"""
        model = cp_model.CpModel()
        x = {}
        
        print("INFO: Using basic scheduling (no consecutive constraints)...")
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            # Check faculty availability
                            faculty = self.faculty_map[faculty_id]
                            if faculty.availability and (day, ts.slot_id) not in faculty.availability:
                                continue
                            
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        # Basic constraints only
        # 1. Each session assigned exactly once
        for session in self.sessions:
            session_vars = [var for (sid, d, slot, rid, fid), var in x.items() if sid == session.session_id]
            if session_vars:
                model.AddExactlyOne(session_vars)

        # 2. No room conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for room in self.rooms:
                    room_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                               if d == day and slot == ts.slot_id and rid == room.room_id]
                    if room_vars:
                        model.AddAtMostOne(room_vars)

        # 3. No faculty conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for faculty in self.faculties:
                    faculty_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                  if d == day and slot == ts.slot_id and fid == faculty.faculty_id]
                    if faculty_vars:
                        model.AddAtMostOne(faculty_vars)

        # 4. No batch conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for batch in self.batches:
                    batch_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                if d == day and slot == ts.slot_id and 
                                any(s.session_id == sid and s.batch_id == batch.batch_id for s in self.sessions)]
                    if batch_vars:
                        model.AddAtMostOne(batch_vars)

        # 5. Faculty workload limits
        for faculty in self.faculties:
            weekly_sessions = [var for (sid, d, slot, rid, fid), var in x.items() if fid == faculty.faculty_id]
            if weekly_sessions:
                model.Add(sum(weekly_sessions) <= faculty.max_hours_per_week)

        # Simple objective: minimize spread
        objective_terms = []
        for (sid, d, slot, rid, fid), var in x.items():
            if slot > 4:  # Later slots get penalty
                objective_terms.append(var * 2)

        if objective_terms:
            model.Minimize(sum(objective_terms))

        # Solve
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit
        solver.parameters.num_search_workers = max(1, os.cpu_count() - 1) if os.cpu_count() else 1
        
        status = solver.Solve(model)
        print(f"INFO: Basic solver status: {solver.StatusName(status)}")
        
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            # Extract solution
            assigned = {}
            for (sid, d, slot, rid, fid), var in x.items():
                if solver.Value(var) == 1:
                    session = next(s for s in self.sessions if s.session_id == sid)
                    result_session = copy.deepcopy(session)
                    result_session.assigned_day = d
                    result_session.assigned_slot = slot
                    result_session.assigned_room = rid
                    result_session.assigned_faculty = fid
                    assigned[sid] = result_session
            
            return {
                "assigned": assigned,
                "unscheduled": len(self.sessions) - len(assigned),
                "solver_status": solver.StatusName(status)
            }
        else:
            return {
                "assigned": {},
                "unscheduled": len(self.sessions),
                "solver_status": solver.StatusName(status)
            }

    def _verify_and_report_consecutiveness(self, assigned: Dict[str, ClassSession], detailed: bool = False) -> int:
        """Enhanced verification with subject and faculty consistency analysis"""
        print("\n" + "="*80)
        print("         COMPREHENSIVE SCHEDULING VERIFICATION")
        print("="*80)
        
        batch_schedules = defaultdict(lambda: defaultdict(list))
        batch_subjects = defaultdict(lambda: defaultdict(list))
        batch_faculty_assignments = defaultdict(lambda: defaultdict(list))
        
        total_violations = 0
        
        # Collect scheduling data
        for session in assigned.values():
            batch_schedules[session.batch_id][session.assigned_day].append(session.assigned_slot)
            batch_subjects[session.batch_id][session.assigned_day].append(session.subject_code)
            batch_faculty_assignments[session.batch_id][session.subject_code].append(session.assigned_faculty)
        
        # Check each batch
        for batch_id in batch_schedules.keys():
            print(f"\nBatch: {batch_id}")
            print("-" * 50)
            
            # 1. CHECK SUBJECT REPETITION ON SAME DAY
            subject_violations = 0
            for day, subjects in batch_subjects[batch_id].items():
                subject_counts = {}
                for subject in subjects:
                    subject_counts[subject] = subject_counts.get(subject, 0) + 1
                
                day_violations = [subj for subj, count in subject_counts.items() if count > 1]
                if day_violations:
                    print(f"  {day}: SUBJECT REPETITION VIOLATION - {day_violations} appear multiple times")
                    subject_violations += len(day_violations)
                    total_violations += len(day_violations)
            
            # 2. CHECK FACULTY CONSISTENCY PER SUBJECT
            faculty_violations = 0
            for subject, faculty_list in batch_faculty_assignments[batch_id].items():
                unique_faculty = set(faculty_list)
                if len(unique_faculty) > 1:
                    print(f"  FACULTY CONSISTENCY VIOLATION - {subject}: taught by {list(unique_faculty)}")
                    faculty_violations += 1
                    total_violations += 1
                else:
                    print(f"  Faculty OK - {subject}: consistently taught by {list(unique_faculty)[0]}")
            
            # 3. CHECK CONSECUTIVE SCHEDULING
            gaps_found = 0
            for day, slots in batch_schedules[batch_id].items():
                if len(slots) <= 1:
                    subjects_today = batch_subjects[batch_id][day]
                    print(f"  {day}: {slots} [{', '.join(subjects_today)}] - OK")
                    continue
                
                slots.sort()
                subjects_today = batch_subjects[batch_id][day]
                day_gaps = []
                
                # Find gaps
                for i in range(len(slots) - 1):
                    current_slot = slots[i]
                    next_expected = current_slot + 1
                    actual_next = slots[i + 1]
                    
                    if actual_next > next_expected:
                        gap_slots = list(range(next_expected, actual_next))
                        non_break_gaps = [s for s in gap_slots if (day, s) not in self.break_mask]
                        day_gaps.extend(non_break_gaps)
                
                # Format display with subject details
                slot_details = []
                for i, slot in enumerate(slots):
                    slot_obj = next((ts for ts in self.timeslots if ts.slot_id == slot), None)
                    time_str = slot_obj.start_time if slot_obj else str(slot)
                    subject = subjects_today[i] if i < len(subjects_today) else "?"
                    slot_details.append(f"{slot}({time_str},{subject})")
                
                if day_gaps:
                    print(f"  {day}: {slot_details} - GAPS at slots {day_gaps} ❌")
                    gaps_found += len(day_gaps)
                    total_violations += len(day_gaps)
                else:
                    print(f"  {day}: {slot_details} - CONSECUTIVE ✓")
            
            # Batch Summary
            print(f"\n  BATCH {batch_id} SUMMARY:")
            if subject_violations == 0:
                print(f"    ✓ Subject Distribution: No repetitions")
            else:
                print(f"    ❌ Subject Distribution: {subject_violations} violations")
            
            if faculty_violations == 0:
                print(f"    ✓ Faculty Consistency: All subjects have consistent teachers")
            else:
                print(f"    ❌ Faculty Consistency: {faculty_violations} subjects have multiple teachers")
            
            if gaps_found == 0:
                print(f"    ✓ Consecutive Scheduling: No gaps found")
            else:
                print(f"    ❌ Consecutive Scheduling: {gaps_found} gaps found")
        
        # Overall Summary
        print("\n" + "="*80)
        if total_violations == 0:
            print("RESULT: PERFECT SCHEDULING - All constraints satisfied!")
        else:
            print(f"RESULT: {total_violations} TOTAL VIOLATIONS FOUND")
            print("Issues to fix:")
            print("- Subject repetitions on same day")
            print("- Multiple teachers for same subject in same batch")  
            print("- Gaps in consecutive scheduling")
        print("="*80)
        
        return total_violations

    def _verify_consecutive_scheduling(self, assigned: Dict[str, ClassSession]):
        """Simple verification method (kept for backward compatibility)"""
        return self._verify_and_report_consecutiveness(assigned, detailed=False)

    def _compact_schedule_to_start_early(self, assigned: Dict[str, ClassSession]) -> int:
        """Ensure all batch schedules start from first available slot with no early gaps"""
        print("Step 4: Compacting schedules to start from first time slot...")
        corrections = 0
        
        # Group sessions by batch and day
        batch_day_sessions = defaultdict(lambda: defaultdict(list))
        for session in assigned.values():
            batch_day_sessions[session.batch_id][session.assigned_day].append(session)
        
        for batch_id, days in batch_day_sessions.items():
            for day, sessions in days.items():
                if len(sessions) == 0:
                    continue
                
                # Sort sessions by current slot assignment
                sessions.sort(key=lambda s: s.assigned_slot)
                current_slots = [s.assigned_slot for s in sessions]
                
                # Get available slots (excluding breaks)
                available_slots = [ts.slot_id for ts in self.timeslots if (day, ts.slot_id) not in self.break_mask]
                available_slots.sort()
                
                # Check if we can compact to start from first available slot
                if len(sessions) <= len(available_slots):
                    # Check if other batches are using the early slots we need
                    required_slots = available_slots[:len(sessions)]
                    conflicting_sessions = []
                    
                    for other_session in assigned.values():
                        if (other_session.batch_id != batch_id and 
                            other_session.assigned_day == day and 
                            other_session.assigned_slot in required_slots):
                            conflicting_sessions.append(other_session)
                    
                    if not conflicting_sessions:
                        # No conflicts, can compact to start from first slot
                        old_slots = current_slots.copy()
                        for i, session in enumerate(sessions):
                            session.assigned_slot = required_slots[i]
                        
                        if current_slots != [s.assigned_slot for s in sessions]:
                            corrections += 1
                            print(f"    Compacted {batch_id} on {day}: {old_slots} → {[s.assigned_slot for s in sessions]}")
                    else:
                        # Try to find the earliest possible consecutive block
                        corrections += self._find_earliest_consecutive_block(batch_id, day, sessions, available_slots, assigned)
        
        return corrections

    def _find_earliest_consecutive_block(self, batch_id: str, day: str, sessions: List[ClassSession], 
                                       available_slots: List[int], assigned: Dict[str, ClassSession]) -> int:
        """Find the earliest possible consecutive block for sessions"""
        num_sessions = len(sessions)
        
        # Try each possible starting position
        for start_idx in range(len(available_slots) - num_sessions + 1):
            candidate_slots = available_slots[start_idx:start_idx + num_sessions]
            
            # Check if these slots are free (no other batches using them)
            slots_free = True
            for other_session in assigned.values():
                if (other_session.batch_id != batch_id and 
                    other_session.assigned_day == day and 
                    other_session.assigned_slot in candidate_slots):
                    slots_free = False
                    break
            
            if slots_free:
                # Found earliest available consecutive block
                old_slots = [s.assigned_slot for s in sessions]
                for i, session in enumerate(sessions):
                    session.assigned_slot = candidate_slots[i]
                
                print(f"    Moved {batch_id} on {day} to earliest block: {old_slots} → {candidate_slots}")
                return 1
        
        return 0

    def _post_process_schedule_validation(self, assigned: Dict[str, ClassSession]) -> Dict[str, ClassSession]:
        """Post-processing validation and correction system to fix rule violations"""
        print("\n" + "="*60)
        print("         POST-PROCESSING VALIDATION & CORRECTION")
        print("="*60)
        
        corrected_assigned = dict(assigned)  # Work on a copy
        corrections_made = 0
        
        # Step 1: Fix faculty consistency violations
        corrections_made += self._fix_faculty_consistency(corrected_assigned)
        
        # Step 2: Fix subject repetition violations
        corrections_made += self._fix_subject_repetition(corrected_assigned)
        
        # Step 3: Fix gaps by shifting or filling
        corrections_made += self._fix_gaps_by_shifting(corrected_assigned)
        
        # Step 4: Compact schedules to start from first available slot
        corrections_made += self._compact_schedule_to_start_early(corrected_assigned)
        
        print(f"\nPOST-PROCESSING SUMMARY: {corrections_made} corrections made")
        print("="*60)
        
        return corrected_assigned

    def _fix_faculty_consistency(self, assigned: Dict[str, ClassSession]) -> int:
        """Ensure same faculty teaches all sessions of same subject for same batch"""
        print("Step 1: Fixing faculty consistency violations...")
        corrections = 0
        
        # Group sessions by batch and subject
        batch_subject_sessions = defaultdict(lambda: defaultdict(list))
        for session in assigned.values():
            batch_subject_sessions[session.batch_id][session.subject_code].append(session)
        
        for batch_id, subjects in batch_subject_sessions.items():
            for subject_code, sessions in subjects.items():
                if len(sessions) > 1:
                    # Find the most qualified faculty for this subject
                    faculty_sessions = defaultdict(list)
                    for session in sessions:
                        faculty_sessions[session.assigned_faculty].append(session)
                    
                    # Choose the faculty who teaches the most sessions of this subject
                    best_faculty = max(faculty_sessions.keys(), 
                                     key=lambda f: len(faculty_sessions[f]))
                    
                    # Reassign all sessions to the best faculty
                    for session in sessions:
                        if session.assigned_faculty != best_faculty:
                            old_faculty = session.assigned_faculty
                            session.assigned_faculty = best_faculty
                            corrections += 1
                            print(f"  Fixed: {batch_id} {subject_code} reassigned from {old_faculty} to {best_faculty}")
        
        return corrections

    def _fix_subject_repetition(self, assigned: Dict[str, ClassSession]) -> int:
        """Fix same subject appearing multiple times on same day"""
        print("Step 2: Fixing subject repetition violations...")
        corrections = 0
        
        # Group sessions by batch and day
        batch_day_sessions = defaultdict(lambda: defaultdict(list))
        for session in assigned.values():
            batch_day_sessions[session.batch_id][session.assigned_day].append(session)
        
        for batch_id, days in batch_day_sessions.items():
            for day, sessions in days.items():
                # Find subject repetitions
                subject_sessions = defaultdict(list)
                for session in sessions:
                    subject_sessions[session.subject_code].append(session)
                
                # For each subject that appears more than once
                for subject_code, subj_sessions in subject_sessions.items():
                    if len(subj_sessions) > 1:
                        print(f"  Found repetition: {batch_id} has {len(subj_sessions)} {subject_code} sessions on {day}")
                        
                        # Keep the first session, move others to different days
                        sessions_to_move = subj_sessions[1:]
                        corrections += self._move_sessions_to_different_days(sessions_to_move, batch_id, assigned)
        
        return corrections

    def _move_sessions_to_different_days(self, sessions_to_move: List[ClassSession], 
                                       batch_id: str, assigned: Dict[str, ClassSession]) -> int:
        """Move sessions to different days to avoid repetition"""
        corrections = 0
        
        for session in sessions_to_move:
            original_day = session.assigned_day
            
            # Find alternative days where this subject doesn't already exist
            for alt_day in self.working_days:
                if alt_day == original_day:
                    continue
                
                # Check if this subject already exists on this day for this batch
                day_subjects = set()
                for other_session in assigned.values():
                    if (other_session.batch_id == batch_id and 
                        other_session.assigned_day == alt_day):
                        day_subjects.add(other_session.subject_code)
                
                if session.subject_code not in day_subjects:
                    # Check if we can find a suitable slot
                    available_slots = self._find_available_slots(alt_day, batch_id, assigned)
                    if available_slots:
                        # Move to first available slot
                        session.assigned_day = alt_day
                        session.assigned_slot = available_slots[0]
                        corrections += 1
                        print(f"    Moved {session.subject_code} from {original_day} to {alt_day}")
                        break
        
        return corrections

    def _find_available_slots(self, day: str, batch_id: str, assigned: Dict[str, ClassSession]) -> List[int]:
        """Find available time slots for a batch on a given day"""
        # Get all time slots except breaks
        available_slots = [ts.slot_id for ts in self.timeslots if (day, ts.slot_id) not in self.break_mask]
        
        # Remove slots already occupied by this batch
        occupied_slots = set()
        for session in assigned.values():
            if session.batch_id == batch_id and session.assigned_day == day:
                occupied_slots.add(session.assigned_slot)
        
        return [slot for slot in available_slots if slot not in occupied_slots]

    def _fix_gaps_by_shifting(self, assigned: Dict[str, ClassSession]) -> int:
        """Fix gaps by shifting classes earlier or filling with other subjects"""
        print("Step 3: Fixing gaps by shifting and filling...")
        corrections = 0
        
        # Group sessions by batch and day
        batch_day_sessions = defaultdict(lambda: defaultdict(list))
        for session in assigned.values():
            batch_day_sessions[session.batch_id][session.assigned_day].append(session)
        
        for batch_id, days in batch_day_sessions.items():
            for day, sessions in days.items():
                if len(sessions) <= 1:
                    continue
                
                # Sort sessions by time slot
                sessions.sort(key=lambda s: s.assigned_slot)
                slots = [s.assigned_slot for s in sessions]
                
                # Find gaps
                gaps = self._find_gaps_in_schedule(slots, day)
                
                if gaps:
                    print(f"  Found {len(gaps)} gaps in {batch_id} on {day}")
                    corrections += self._fill_or_shift_gaps(batch_id, day, sessions, gaps, assigned)
        
        return corrections

    def _find_gaps_in_schedule(self, slots: List[int], day: str) -> List[int]:
        """Find gaps in a schedule (excluding breaks)"""
        if len(slots) <= 1:
            return []
        
        gaps = []
        for i in range(len(slots) - 1):
            current_slot = slots[i]
            next_slot = slots[i + 1]
            
            # Find missing slots between current and next
            for gap_slot in range(current_slot + 1, next_slot):
                # Only count as gap if it's not a break
                if (day, gap_slot) not in self.break_mask:
                    gaps.append(gap_slot)
        
        return gaps

    def _fill_or_shift_gaps(self, batch_id: str, day: str, sessions: List[ClassSession], 
                          gaps: List[int], assigned: Dict[str, ClassSession]) -> int:
        """Fill gaps by shifting sessions or finding subjects to fill gaps"""
        corrections = 0
        
        # Strategy 1: Try to shift all sessions to start earlier (eliminate gaps)
        if self._try_shift_sessions_earlier(batch_id, day, sessions, assigned):
            corrections += len(gaps)
            print(f"    Shifted all sessions earlier to eliminate gaps")
        else:
            # Strategy 2: Try to fill gaps with other subjects from this batch
            corrections += self._try_fill_gaps_with_subjects(batch_id, day, gaps, assigned)
        
        return corrections

    def _try_shift_sessions_earlier(self, batch_id: str, day: str, sessions: List[ClassSession], 
                                  assigned: Dict[str, ClassSession]) -> bool:
        """Try to shift all sessions earlier to eliminate gaps"""
        # Get available early slots
        early_slots = [ts.slot_id for ts in self.timeslots[:4] if (day, ts.slot_id) not in self.break_mask]
        early_slots.sort()
        
        if len(early_slots) >= len(sessions):
            # Check if early slots are available (not used by other batches)
            occupied_by_others = set()
            for other_session in assigned.values():
                if (other_session.batch_id != batch_id and 
                    other_session.assigned_day == day and 
                    other_session.assigned_slot in early_slots):
                    occupied_by_others.add(other_session.assigned_slot)
            
            available_early_slots = [slot for slot in early_slots if slot not in occupied_by_others]
            
            if len(available_early_slots) >= len(sessions):
                # Shift sessions to early consecutive slots
                for i, session in enumerate(sessions):
                    session.assigned_slot = available_early_slots[i]
                return True
        
        return False

    def _try_fill_gaps_with_subjects(self, batch_id: str, day: str, gaps: List[int], 
                                   assigned: Dict[str, ClassSession]) -> int:
        """Try to fill gaps with other subjects scheduled on different days"""
        corrections = 0
        
        # Get subjects that this batch has on other days
        batch_subjects_other_days = set()
        day_subjects = set()
        
        for session in assigned.values():
            if session.batch_id == batch_id:
                if session.assigned_day == day:
                    day_subjects.add(session.subject_code)
                else:
                    batch_subjects_other_days.add(session.subject_code)
        
        # Find subjects that could be moved to fill gaps
        fillable_subjects = batch_subjects_other_days - day_subjects
        
        for gap_slot in gaps:
            if not fillable_subjects:
                break
                
            # Find a session of a fillable subject to move here
            for session in list(assigned.values()):
                if (session.batch_id == batch_id and 
                    session.subject_code in fillable_subjects and 
                    session.assigned_day != day):
                    
                    # Move this session to fill the gap
                    old_day = session.assigned_day
                    session.assigned_day = day
                    session.assigned_slot = gap_slot
                    
                    fillable_subjects.remove(session.subject_code)
                    corrections += 1
                    print(f"    Filled gap at slot {gap_slot} with {session.subject_code} (moved from {old_day})")
                    break
        
        return corrections

    def schedule(self, time_limit: int = 60) -> Dict[str, Any]:
        if not ORTOOLS_AVAILABLE:
            return {"error": "OR-Tools library required. Install with: pip install ortools"}
        
        if not self._basic_validation():
            return {
                "error": "Validation failed. Check faculty and room assignments.",
                "statistics": {"total": len(self.sessions), "scheduled": 0, "unscheduled": len(self.sessions)}
            }

        print("INFO: Starting STRICT scheduling with multiple solution attempts...")
        
        # Try multiple solution strategies and pick the best one
        best_result = self._find_best_violation_free_solution(time_limit)
        
        return self._format_result(best_result) if best_result else {"error": "Failed to generate violation-free schedule"}

    def _find_best_violation_free_solution(self, time_limit: int) -> Optional[Dict[str, Any]]:
        """Try multiple solution strategies and return the best violation-free result"""
        print("="*70)
        print("         EXHAUSTIVE SOLUTION SEARCH FOR VIOLATION-FREE SCHEDULE")
        print("="*70)
        
        solution_attempts = [
            ("Strict Constraints with Early Start Priority", self._solve_with_early_start_priority),
            ("Balanced Daily Distribution", self._solve_with_balanced_distribution),
            ("Subject Completion Priority", self._solve_with_subject_completion),
            ("Faculty Minimization Priority", self._solve_with_faculty_minimization),
            ("Compact Block Scheduling", self._solve_with_compact_blocks)
        ]
        
        best_solution = None
        best_violations = float('inf')
        all_results = []
        
        for strategy_name, solve_method in solution_attempts:
            print(f"\n--- Attempting Strategy: {strategy_name} ---")
            
            try:
                result = solve_method(time_limit // len(solution_attempts))
                
                if result and result.get("assigned"):
                    # Post-process to fix violations
                    corrected_assigned = self._post_process_schedule_validation(result["assigned"])
                    
                    # Count violations in corrected solution
                    violations = self._count_total_violations(corrected_assigned)
                    
                    result["assigned"] = corrected_assigned
                    result["violations"] = violations
                    all_results.append((strategy_name, result, violations))
                    
                    print(f"Strategy '{strategy_name}': {violations} violations found")
                    
                    if violations < best_violations:
                        best_violations = violations
                        best_solution = result
                        print(f"NEW BEST SOLUTION: {violations} violations")
                    
                    # If we found a perfect solution, stop here
                    if violations == 0:
                        print(f"PERFECT SOLUTION FOUND with '{strategy_name}'!")
                        break
                        
            except Exception as e:
                print(f"Strategy '{strategy_name}' failed: {e}")
                continue
        
        # Report all attempts
        print(f"\n" + "="*70)
        print("         SOLUTION ATTEMPT SUMMARY")
        print("="*70)
        
        for strategy_name, result, violations in all_results:
            status = "PERFECT" if violations == 0 else f"{violations} violations"
            scheduled = len(result.get("assigned", {}))
            print(f"{strategy_name}: {scheduled} sessions scheduled, {status}")
        
        if best_solution and best_violations == 0:
            print(f"\nSUCCESS: Found violation-free solution!")
        elif best_solution:
            print(f"\nBEST RESULT: {best_violations} violations (could not achieve violation-free schedule)")
        else:
            print(f"\nFAILURE: No valid solutions found")
            
        print("="*70)
        
        return best_solution

    def _solve_with_early_start_priority(self, time_limit: int) -> Optional[Dict[str, Any]]:
        """Solve with priority on early time slots and strict constraints"""
        model = cp_model.CpModel()
        x = {}
        
        batch_sessions = defaultdict(list)
        for session in self.sessions:
            batch_sessions[session.batch_id].append(session)
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        # Add all strict constraints
        self._add_basic_constraints(model, x, batch_sessions)
        self._add_strict_subject_constraints(model, x, batch_sessions)
        self._add_strict_faculty_constraints(model, x, batch_sessions)
        self._add_consecutive_constraints(model, x, batch_sessions)
        
        # Objective: Minimize late slots (prioritize early start)
        objective_terms = []
        for (sid, d, slot, rid, fid), var in x.items():
            objective_terms.append(var * slot * 100)  # Heavy penalty for late slots
        
        model.Minimize(sum(objective_terms))
        return self._solve_model(model, x, time_limit, "Early Start Priority")

    def _solve_with_balanced_distribution(self, time_limit: int) -> Optional[Dict[str, Any]]:
        """Solve with priority on balanced daily distribution"""
        model = cp_model.CpModel()
        x = {}
        
        batch_sessions = defaultdict(list)
        for session in self.sessions:
            batch_sessions[session.batch_id].append(session)
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        self._add_basic_constraints(model, x, batch_sessions)
        self._add_strict_subject_constraints(model, x, batch_sessions)
        self._add_strict_faculty_constraints(model, x, batch_sessions)
        self._add_consecutive_constraints(model, x, batch_sessions)
        
        # Objective: Balance daily distribution
        objective_terms = []
        for batch_id, sessions in batch_sessions.items():
            daily_counts = []
            for day in self.working_days:
                day_count = model.NewIntVar(0, len(sessions), f'daily_count_{batch_id}_{day}')
                day_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                           if d == day and any(s.session_id == sid and s.batch_id == batch_id for s in sessions)]
                if day_vars:
                    model.Add(day_count == sum(day_vars))
                daily_counts.append(day_count)
            
            # Minimize variance between days
            avg_per_day = len(sessions) // len(self.working_days)
            for day_count in daily_counts:
                deviation = model.NewIntVar(0, len(sessions), f'dev_{batch_id}')
                model.AddAbsEquality(deviation, day_count - avg_per_day)
                objective_terms.append(deviation * 10)
        
        model.Minimize(sum(objective_terms))
        return self._solve_model(model, x, time_limit, "Balanced Distribution")

    def _solve_with_subject_completion(self, time_limit: int) -> Optional[Dict[str, Any]]:
        """Solve with priority on completing subjects efficiently"""
        model = cp_model.CpModel()
        x = {}
        
        batch_sessions = defaultdict(list)
        for session in self.sessions:
            batch_sessions[session.batch_id].append(session)
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        self._add_basic_constraints(model, x, batch_sessions)
        self._add_strict_subject_constraints(model, x, batch_sessions)
        self._add_strict_faculty_constraints(model, x, batch_sessions)
        self._add_consecutive_constraints(model, x, batch_sessions)
        
        # Objective: Complete subjects quickly (fewer days per subject)
        objective_terms = []
        for batch_id, sessions in batch_sessions.items():
            subject_sessions = defaultdict(list)
            for session in sessions:
                subject_sessions[session.subject_code].append(session)
            
            for subject_code, subj_sessions in subject_sessions.items():
                for day in self.working_days:
                    day_used = model.NewBoolVar(f'subj_{subject_code}_{batch_id}_{day}')
                    day_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                               if d == day and any(s.session_id == sid and s.subject_code == subject_code 
                                                 and s.batch_id == batch_id for s in subj_sessions)]
                    if day_vars:
                        model.Add(sum(day_vars) >= 1).OnlyEnforceIf(day_used)
                        model.Add(sum(day_vars) == 0).OnlyEnforceIf(day_used.Not())
                        objective_terms.append(day_used * 5)  # Penalty for each day a subject is used
        
        model.Minimize(sum(objective_terms))
        return self._solve_model(model, x, time_limit, "Subject Completion")

    def _solve_with_faculty_minimization(self, time_limit: int) -> Optional[Dict[str, Any]]:
        """Solve with priority on minimizing faculty conflicts"""
        model = cp_model.CpModel()
        x = {}
        
        batch_sessions = defaultdict(list)
        for session in self.sessions:
            batch_sessions[session.batch_id].append(session)
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        self._add_basic_constraints(model, x, batch_sessions)
        self._add_strict_subject_constraints(model, x, batch_sessions)
        self._add_strict_faculty_constraints(model, x, batch_sessions)
        self._add_consecutive_constraints(model, x, batch_sessions)
        
        # Objective: Prefer best-qualified faculty
        objective_terms = []
        for (sid, d, slot, rid, fid), var in x.items():
            session = next(s for s in self.sessions if s.session_id == sid)
            score = self.faculty_subject_scores.get((fid, session.subject_code), 0)
            objective_terms.append(var * (-score))  # Negative because we want to maximize good matches
        
        model.Minimize(sum(objective_terms))
        return self._solve_model(model, x, time_limit, "Faculty Minimization")

    def _solve_with_compact_blocks(self, time_limit: int) -> Optional[Dict[str, Any]]:
        """Solve with priority on creating compact schedule blocks"""
        model = cp_model.CpModel()
        x = {}
        
        batch_sessions = defaultdict(list)
        for session in self.sessions:
            batch_sessions[session.batch_id].append(session)
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        self._add_basic_constraints(model, x, batch_sessions)
        self._add_strict_subject_constraints(model, x, batch_sessions)
        self._add_strict_faculty_constraints(model, x, batch_sessions)
        self._add_consecutive_constraints(model, x, batch_sessions)
        
        # Objective: Minimize number of days used per batch
        objective_terms = []
        for batch_id, sessions in batch_sessions.items():
            for day in self.working_days:
                day_used = model.NewBoolVar(f'day_used_{batch_id}_{day}')
                day_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                           if d == day and any(s.session_id == sid and s.batch_id == batch_id for s in sessions)]
                if day_vars:
                    model.Add(sum(day_vars) >= 1).OnlyEnforceIf(day_used)
                    model.Add(sum(day_vars) == 0).OnlyEnforceIf(day_used.Not())
                    objective_terms.append(day_used * 20)  # Penalty for each day used
        
    def _add_basic_constraints(self, model, x, batch_sessions):
        """Add basic scheduling constraints"""
        # 1. Each session assigned exactly once
        for session in self.sessions:
            session_vars = [var for (sid, d, slot, rid, fid), var in x.items() if sid == session.session_id]
            if session_vars:
                model.AddExactlyOne(session_vars)

        # 2. No room conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for room in self.rooms:
                    room_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                               if d == day and slot == ts.slot_id and rid == room.room_id]
                    if room_vars:
                        model.AddAtMostOne(room_vars)

        # 3. No faculty conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for faculty in self.faculties:
                    faculty_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                  if d == day and slot == ts.slot_id and fid == faculty.faculty_id]
                    if faculty_vars:
                        model.AddAtMostOne(faculty_vars)

        # 4. No batch conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for batch_id in batch_sessions.keys():
                    batch_vars = [var for (sid, d, slot, rid, fid), var in x.items()
                                if d == day and slot == ts.slot_id and 
                                any(s.session_id == sid and s.batch_id == batch_id for s in self.sessions)]
                    if batch_vars:
                        model.AddAtMostOne(batch_vars)

        # 5. Faculty workload limits
        for faculty in self.faculties:
            weekly_sessions = [var for (sid, d, slot, rid, fid), var in x.items() if fid == faculty.faculty_id]
            if weekly_sessions:
                model.Add(sum(weekly_sessions) <= faculty.max_hours_per_week)

    def _add_strict_subject_constraints(self, model, x, batch_sessions):
        """Add STRICT constraints to prevent subject repetition on same day"""
        for batch_id, sessions in batch_sessions.items():
            subject_sessions = defaultdict(list)
            for session in sessions:
                subject_sessions[session.subject_code].append(session)
            
            # HARD CONSTRAINT: Exactly one session per subject per day
            for subject_code, subj_sessions in subject_sessions.items():
                for day in self.working_days:
                    day_subject_vars = []
                    for session in subj_sessions:
                        session_vars_this_day = [
                            var for (sid, d, slot, rid, fid), var in x.items()
                            if sid == session.session_id and d == day
                        ]
                        day_subject_vars.extend(session_vars_this_day)
                    
                    if day_subject_vars:
                        model.Add(sum(day_subject_vars) <= 1)

    def _add_strict_faculty_constraints(self, model, x, batch_sessions):
        """Add STRICT constraints for faculty consistency per subject per batch"""
        for batch_id, sessions in batch_sessions.items():
            subject_sessions = defaultdict(list)
            for session in sessions:
                subject_sessions[session.subject_code].append(session)
            
            # HARD CONSTRAINT: Same faculty for all sessions of same subject
            for subject_code, subj_sessions in subject_sessions.items():
                if len(subj_sessions) > 1:
                    # For each pair of sessions of the same subject, they must have the same faculty
                    for i in range(len(subj_sessions)):
                        for j in range(i + 1, len(subj_sessions)):
                            session_i = subj_sessions[i]
                            session_j = subj_sessions[j]
                            
                            # Get all possible faculty for these sessions
                            faculty_i_vars = defaultdict(list)
                            faculty_j_vars = defaultdict(list)
                            
                            for (sid, d, slot, rid, fid), var in x.items():
                                if sid == session_i.session_id:
                                    faculty_i_vars[fid].append(var)
                                elif sid == session_j.session_id:
                                    faculty_j_vars[fid].append(var)
                            
                            # If session i is taught by faculty f, then session j must also be taught by faculty f
                            for faculty_id in faculty_i_vars.keys():
                                if faculty_id in faculty_j_vars:
                                    i_by_faculty = model.NewBoolVar(f'session_{session_i.session_id}_by_{faculty_id}')
                                    j_by_faculty = model.NewBoolVar(f'session_{session_j.session_id}_by_{faculty_id}')
                                    
                                    model.Add(sum(faculty_i_vars[faculty_id]) >= 1).OnlyEnforceIf(i_by_faculty)
                                    model.Add(sum(faculty_i_vars[faculty_id]) == 0).OnlyEnforceIf(i_by_faculty.Not())
                                    model.Add(sum(faculty_j_vars[faculty_id]) >= 1).OnlyEnforceIf(j_by_faculty)
                                    model.Add(sum(faculty_j_vars[faculty_id]) == 0).OnlyEnforceIf(j_by_faculty.Not())
                                    
                                    # If i is taught by this faculty, j must also be taught by this faculty
                                    model.AddImplication(i_by_faculty, j_by_faculty)

    def _add_consecutive_constraints(self, model, x, batch_sessions):
        """Add STRICT constraints for consecutive scheduling"""
        for batch_id, sessions in batch_sessions.items():
            for day in self.working_days:
                available_slots = [ts.slot_id for ts in self.timeslots if (day, ts.slot_id) not in self.break_mask]
                available_slots.sort()
                
                # Track which slots are used by this batch
                slot_used = {}
                for slot_id in available_slots:
                    slot_used[slot_id] = model.NewBoolVar(f'batch_{batch_id}_day_{day}_slot_{slot_id}')
                    
                    assignments_in_slot = [
                        var for (sid, d, slot, rid, fid), var in x.items()
                        if d == day and slot == slot_id and 
                        any(s.session_id == sid and s.batch_id == batch_id for s in sessions)
                    ]
                    
                    if assignments_in_slot:
                        model.Add(sum(assignments_in_slot) >= 1).OnlyEnforceIf(slot_used[slot_id])
                        model.Add(sum(assignments_in_slot) == 0).OnlyEnforceIf(slot_used[slot_id].Not())
                    else:
                        model.Add(slot_used[slot_id] == 0)
                
                # HARD CONSTRAINT: No gaps between used slots
                for i in range(len(available_slots) - 2):
                    slot_i = available_slots[i]
                    slot_mid = available_slots[i + 1]  
                    slot_j = available_slots[i + 2]
                    
                    if slot_i in slot_used and slot_mid in slot_used and slot_j in slot_used:
                        both_ends = model.NewBoolVar(f'both_ends_{batch_id}_{day}_{i}')
                        model.AddBoolAnd([slot_used[slot_i], slot_used[slot_j]]).OnlyEnforceIf(both_ends)
                        model.Add(slot_used[slot_mid] == 1).OnlyEnforceIf(both_ends)

    def _solve_model(self, model, x, time_limit: int, strategy_name: str) -> Optional[Dict[str, Any]]:
        """Solve the model with given constraints and return result"""
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit
        solver.parameters.num_search_workers = max(1, os.cpu_count() - 1) if os.cpu_count() else 1
        
        print(f"INFO: Solving with '{strategy_name}' strategy...")
        status = solver.Solve(model)
        
        print(f"INFO: {strategy_name} solver status: {solver.StatusName(status)}")
        
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            assigned = {}
            for (sid, d, slot, rid, fid), var in x.items():
                if solver.Value(var) == 1:
                    session = next(s for s in self.sessions if s.session_id == sid)
                    result_session = copy.deepcopy(session)
                    result_session.assigned_day = d
                    result_session.assigned_slot = slot
                    result_session.assigned_room = rid
                    result_session.assigned_faculty = fid
                    assigned[sid] = result_session
            
            return {
                "assigned": assigned,
                "unscheduled": len(self.sessions) - len(assigned),
                "solver_status": solver.StatusName(status),
                "strategy": strategy_name
            }
        else:
            return None

    def _count_total_violations(self, assigned: Dict[str, ClassSession]) -> int:
        """Count all three types of violations strictly"""
        violations = 0
        
        # Group sessions for analysis
        batch_schedules = defaultdict(lambda: defaultdict(list))
        batch_subjects = defaultdict(lambda: defaultdict(list))
        batch_faculty_assignments = defaultdict(lambda: defaultdict(set))
        
        for session in assigned.values():
            batch_schedules[session.batch_id][session.assigned_day].append(session.assigned_slot)
            batch_subjects[session.batch_id][session.assigned_day].append(session.subject_code)
            batch_faculty_assignments[session.batch_id][session.subject_code].add(session.assigned_faculty)
        
        # COUNT VIOLATION 1: Subject repetitions on same day
        for batch_id, days in batch_subjects.items():
            for day, subjects in days.items():
                subject_counts = {}
                for subject in subjects:
                    subject_counts[subject] = subject_counts.get(subject, 0) + 1
                
                # Count excess occurrences
                for subject, count in subject_counts.items():
                    if count > 1:
                        violations += (count - 1)  # Each extra occurrence is a violation
        
        # COUNT VIOLATION 2: Multiple teachers for same subject in same batch
        for batch_id, subjects in batch_faculty_assignments.items():
            for subject, faculty_set in subjects.items():
                if len(faculty_set) > 1:
                    violations += (len(faculty_set) - 1)  # Each extra faculty is a violation
        
        # COUNT VIOLATION 3: Gaps in consecutive scheduling
        for batch_id, days in batch_schedules.items():
            for day, slots in days.items():
                if len(slots) > 1:
                    slots.sort()
                    for i in range(len(slots) - 1):
                        current_slot = slots[i]
                        next_slot = slots[i + 1]
                        
                        # Count non-break gaps
                        gap_slots = list(range(current_slot + 1, next_slot))
                        non_break_gaps = [s for s in gap_slots if (day, s) not in self.break_mask]
                        violations += len(non_break_gaps)
        
        return violations
        model = cp_model.CpModel()
        x = {}
        
        # Create variables
        for session in self.sessions:
            for day in self.working_days:
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    for room_id in self.eligible_rooms[session.session_id]:
                        for faculty_id in self.eligible_faculties[session.session_id]:
                            # Check faculty availability
                            faculty = self.faculty_map[faculty_id]
                            if faculty.availability and (day, ts.slot_id) not in faculty.availability:
                                continue
                            
                            x[session.session_id, day, ts.slot_id, room_id, faculty_id] = model.NewBoolVar(
                                f"x_{session.session_id}_{day}_{ts.slot_id}_{room_id}_{faculty_id}"
                            )

        # Constraints
        # 1. Each session assigned exactly once
        for session in self.sessions:
            model.AddExactlyOne([
                var for (sid, d, slot, rid, fid), var in x.items() if sid == session.session_id
            ])

        # 2. No room conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for room in self.rooms:
                    model.AddAtMostOne([
                        var for (sid, d, slot, rid, fid), var in x.items()
                        if d == day and slot == ts.slot_id and rid == room.room_id
                    ])

        # 3. No faculty conflicts
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for faculty in self.faculties:
                    model.AddAtMostOne([
                        var for (sid, d, slot, rid, fid), var in x.items()
                        if d == day and slot == ts.slot_id and fid == faculty.faculty_id
                    ])

        # 4. No batch conflicts (students can't be in two places)
        for day in self.working_days:
            for ts in self.timeslots:
                if (day, ts.slot_id) in self.break_mask:
                    continue
                for batch in self.batches:
                    batch_sessions = [
                        var for (sid, d, slot, rid, fid), var in x.items()
                        if d == day and slot == ts.slot_id and 
                        next((s for s in self.sessions if s.session_id == sid), None) and
                        next((s for s in self.sessions if s.session_id == sid), None).batch_id == batch.batch_id
                    ]
                    if batch_sessions:
                        model.AddAtMostOne(batch_sessions)

        # 5. Faculty workload limits
        for faculty in self.faculties:
            weekly_sessions = [
                var for (sid, d, slot, rid, fid), var in x.items() if fid == faculty.faculty_id
            ]
            if weekly_sessions:
                model.Add(sum(weekly_sessions) <= faculty.max_hours_per_week)

        # 6. CONSECUTIVE CLASSES CONSTRAINT - Key addition for student batches
        for batch in self.batches:
            for day in self.working_days:
                # Get all possible sessions for this batch on this day
                day_sessions = []
                for session in self.sessions:
                    if session.batch_id == batch.batch_id:
                        session_vars = [
                            var for (sid, d, slot, rid, fid), var in x.items()
                            if sid == session.session_id and d == day
                        ]
                        if session_vars:
                            # Create a boolean variable indicating if this session is scheduled on this day
                            session_scheduled = model.NewBoolVar(f'batch_{batch.batch_id}_{day}_{session.session_id}')
                            model.Add(sum(session_vars) == 1).OnlyEnforceIf(session_scheduled)
                            model.Add(sum(session_vars) == 0).OnlyEnforceIf(session_scheduled.Not())
                            day_sessions.append((session.session_id, session_scheduled, session_vars))

                if not day_sessions:
                    continue

                # If there are sessions on this day, they should be consecutive
                # First, determine which time slots are used
                slot_used = {}
                for ts in self.timeslots:
                    if (day, ts.slot_id) in self.break_mask:
                        continue
                    slot_used[ts.slot_id] = model.NewBoolVar(f'slot_used_{batch.batch_id}_{day}_{ts.slot_id}')
                    
                    # A slot is used if any session of this batch is scheduled in it
                    slot_sessions = [
                        var for (sid, d, slot, rid, fid), var in x.items()
                        if d == day and slot == ts.slot_id and 
                        any(s.session_id == sid and s.batch_id == batch.batch_id for s in self.sessions)
                    ]
                    
                    if slot_sessions:
                        model.AddMaxEquality(slot_used[ts.slot_id], slot_sessions)

                # Consecutive constraint: if slots i and i+2 are used, then slot i+1 must also be used
                # (unless i+1 is a break)
                available_slots = [ts.slot_id for ts in self.timeslots if (day, ts.slot_id) not in self.break_mask]
                available_slots.sort()
                
                for i in range(len(available_slots) - 2):
                    slot1, slot2, slot3 = available_slots[i], available_slots[i+1], available_slots[i+2]
                    
                    if slot1 in slot_used and slot2 in slot_used and slot3 in slot_used:
                        # If both slot1 and slot3 are used, then slot2 must be used (no gaps)
                        gap_violation = model.NewBoolVar(f'gap_{batch.batch_id}_{day}_{slot1}_{slot3}')
                        
                        # gap_violation = 1 if (slot1 used AND slot3 used AND slot2 not used)
                        model.AddBoolAnd([slot_used[slot1], slot_used[slot3], slot_used[slot2].Not()]).OnlyEnforceIf(gap_violation)
                        
                        # Prevent gap violations (make them impossible)
                        model.Add(gap_violation == 0)

        # Objective: Basic preferences
        objective_terms = []
        
        # Minimize total sessions per day per batch (encourage concentration)
        for batch in self.batches:
            for day in self.working_days:
                day_session_count = model.NewIntVar(0, len(self.sessions), f'day_count_{batch.batch_id}_{day}')
                day_sessions = [
                    var for (sid, d, slot, rid, fid), var in x.items()
                    if d == day and any(s.session_id == sid and s.batch_id == batch.batch_id for s in self.sessions)
                ]
                if day_sessions:
                    model.Add(day_session_count == sum(day_sessions))
                    # Penalty for having sessions spread across too many days
                    # Prefer fewer days with more sessions each
                    busy_day = model.NewBoolVar(f'busy_day_{batch.batch_id}_{day}')
                    model.Add(day_session_count >= 1).OnlyEnforceIf(busy_day)
                    model.Add(day_session_count == 0).OnlyEnforceIf(busy_day.Not())
                    objective_terms.append(busy_day * 5)  # Small penalty for each busy day

        # Prefer morning slots for better attendance
        for (sid, d, slot, rid, fid), var in x.items():
            if slot > 5:  # Later time slots
                objective_terms.append(var * 2)

        if objective_terms:
            model.Minimize(sum(objective_terms))

        # Solve
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = time_limit
        solver.parameters.num_search_workers = max(1, os.cpu_count() - 1) if os.cpu_count() else 1
        
        status = solver.Solve(model)
        print(f"INFO: Solver status: {solver.StatusName(status)}")
        
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            # Extract solution
            assigned = {}
            for (sid, d, slot, rid, fid), var in x.items():
                if solver.Value(var) == 1:
                    session = next(s for s in self.sessions if s.session_id == sid)
                    result_session = copy.deepcopy(session)
                    result_session.assigned_day = d
                    result_session.assigned_slot = slot
                    result_session.assigned_room = rid
                    result_session.assigned_faculty = fid
                    assigned[sid] = result_session
            
            return {
                "assigned": assigned,
                "unscheduled": len(self.sessions) - len(assigned),
                "solver_status": solver.StatusName(status)
            }
        else:
            return {
                "assigned": {},
                "unscheduled": len(self.sessions),
                "solver_status": solver.StatusName(status)
            }

    def _format_result(self, result: Dict[str, Any]) -> Dict[str, Any]:
        assigned = result.get("assigned", {})
        
        # Create timetables
        faculty_tts = {f.name: {d: {str(ts): "" for ts in self.timeslots} for d in self.working_days} 
                      for f in self.faculties}
        batch_tts = {b.batch_id: {d: {str(ts): "" for ts in self.timeslots} for d in self.working_days} 
                    for b in self.batches}
        room_tts = {r.room_id: {d: {str(ts): "" for ts in self.timeslots} for d in self.working_days} 
                   for r in self.rooms}

        # Add breaks
        for tt_group in [faculty_tts, batch_tts, room_tts]:
            for grid in tt_group.values():
                for day, slot_id in self.break_mask:
                    if day in grid and slot_id <= len(self.timeslots):
                        slot_str = str(self.timeslots[slot_id - 1]) if slot_id > 0 else str(self.timeslots[3])
                        if slot_str in grid[day]:
                            grid[day][slot_str] = "BREAK"

        # Populate with sessions
        for session in assigned.values():
            day = session.assigned_day
            slot_obj = next((ts for ts in self.timeslots if ts.slot_id == session.assigned_slot), None)
            if not slot_obj:
                continue
                
            slot_str = str(slot_obj)
            subject = self.subject_map.get(session.subject_code)
            faculty = self.faculty_map.get(session.assigned_faculty)
            
            if subject and faculty:
                # Faculty timetable
                faculty_tts[faculty.name][day][slot_str] = f"{subject.subject_name} [{session.batch_id}] @ {session.assigned_room}"
                
                # Batch timetable  
                batch_tts[session.batch_id][day][slot_str] = f"{subject.subject_name} by {faculty.name} @ {session.assigned_room}"
                
                # Room timetable
                room_tts[session.assigned_room][day][slot_str] = f"{subject.subject_name} [{session.batch_id}] by {faculty.name}"

        return {
            "faculty_timetables": faculty_tts,
            "batch_timetables": batch_tts,
            "room_timetables": room_tts,
            "statistics": {
                "total_sessions": len(self.sessions),
                "scheduled_sessions": len(assigned),
                "unscheduled_sessions": result.get('unscheduled', 0),
                "success_rate": (len(assigned) / len(self.sessions) * 100) if self.sessions else 0
            },
            "solver_info": {"status": result.get("solver_status", "Unknown")}
        }

# -------------------------
# Template and Output Functions
# -------------------------
def create_simple_template(filename: str):
    template_sheets = {
        "timeslots": ["slot_id", "start_time", "end_time"],
        "rooms": ["room_id", "capacity", "room_type", "department"],
        "faculties": ["faculty_id", "name", "qualifications", "specializations", "max_hours_per_week", "availability"],
        "subjects": ["subject_code", "subject_name", "hours_per_week", "requires_lab", "lab_hours_per_week", "required_room_type"],
        "batches": ["batch_id", "department", "student_count", "subjects"]
    }
    
    sample_data = {
        "timeslots": [
            [1, "08:00", "09:00"], [2, "09:00", "10:00"], [3, "10:00", "11:00"],
            [4, "11:30", "12:30"], [5, "12:30", "13:30"], [6, "14:30", "15:30"],
            [7, "15:30", "16:30"], [8, "16:30", "17:30"]
        ],
        "rooms": [
            ["R101", 60, "classroom", "CSE"], ["R102", 60, "classroom", "ECE"],
            ["LAB01", 30, "computer_lab", "CSE"], ["LAB02", 30, "computer_lab", "ECE"]
        ],
        "faculties": [
            ["CSE01", "Dr. Alan Grant", "phd,computer_science", "Programming,Algorithms", 20, ""],
            ["CSE02", "Prof. Ellie Sattler", "mtech,computer_science", "Databases,Software", 18, ""],
            ["ECE01", "Dr. Ian Malcolm", "phd,electronics", "VLSI,Digital_Signal_Processing", 16, ""],
            ["PHY01", "Prof. Robert Muldoon", "msc,physics", "Physics,Quantum_Mechanics", 20, ""]
        ],
        "subjects": [
            ["CS101", "Intro to Programming", 3, True, 2, "computer_lab"],
            ["CS301", "Database Systems", 4, True, 2, "computer_lab"],
            ["EC101", "Basic Electronics", 4, True, 2, "computer_lab"],
            ["PH100", "Engineering Physics", 4, False, 0, "classroom"],
            ["MA100", "Engineering Maths-I", 5, False, 0, "classroom"]
        ],
        "batches": [
            ["CSE-SEM1-A", "CSE", 45, "CS101,PH100,MA100"],
            ["ECE-SEM1-A", "ECE", 40, "EC101,PH100,MA100"],
            ["CSE-SEM5-A", "CSE", 35, "CS301"]
        ]
    }
    
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for sheet_name, columns in template_sheets.items():
            if sheet_name in sample_data:
                df = pd.DataFrame(sample_data[sheet_name], columns=columns)
            else:
                df = pd.DataFrame(columns=columns)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"INFO: Created simplified template at '{filename}' with sample data")

def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", str(name)).replace(" ", "_")

def write_simple_output(output: Dict[str, Any], scheduler: SimplifiedScheduler, out_dir: str):
    if "error" in output and output.get("statistics", {}).get("scheduled_sessions", 0) == 0:
        print(f"ERROR: {output['error']}")
        return
    
    os.makedirs(out_dir, exist_ok=True)
    
    # Faculty timetables
    for faculty_name, grid in output.get("faculty_timetables", {}).items():
        filename = os.path.join(out_dir, f"faculty_{sanitize_filename(faculty_name)}.xlsx")
        pd.DataFrame(grid).to_excel(filename)
    
    # Room timetables by department
    rooms_by_dept = defaultdict(list)
    for room in scheduler.rooms:
        rooms_by_dept[room.department].append(room.room_id)
    
    for dept, room_ids in rooms_by_dept.items():
        filename = os.path.join(out_dir, f"rooms_{sanitize_filename(dept)}.xlsx")
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            for room_id in sorted(room_ids):
                if room_id in output.get("room_timetables", {}):
                    sheet_name = sanitize_filename(room_id)[:31]
                    pd.DataFrame(output["room_timetables"][room_id]).to_excel(writer, sheet_name=sheet_name)
    
    # Batch timetables
    filename = os.path.join(out_dir, "batch_timetables.xlsx")
    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        for batch_id, grid in output.get("batch_timetables", {}).items():
            sheet_name = sanitize_filename(batch_id)[:31]
            pd.DataFrame(grid).to_excel(writer, sheet_name=sheet_name)
    
    # Summary
    summary_file = os.path.join(out_dir, "scheduling_summary.xlsx")
    with pd.ExcelWriter(summary_file, engine="openpyxl") as writer:
        stats_df = pd.DataFrame([output.get("statistics", {})])
        stats_df.to_excel(writer, sheet_name="Statistics", index=False)
        
        solver_df = pd.DataFrame([output.get("solver_info", {})])
        solver_df.to_excel(writer, sheet_name="Solver_Info", index=False)
    
    print(f"INFO: Output files written to '{out_dir}' directory")
    stats = output.get("statistics", {})
    print(f"INFO: Scheduled {stats.get('scheduled_sessions', 0)}/{stats.get('total_sessions', 0)} sessions ({stats.get('success_rate', 0):.1f}% success)")

def main():
    INPUT_FILENAME = "simple_timetable_input.xlsx"
    OUTPUT_DIRECTORY = "simple_timetable_output"

    print("=" * 60)
    print("        SIMPLIFIED TIMETABLE GENERATOR")
    print("=" * 60)
    print("Features: Basic scheduling without credit validation")
    print("Focus: Conflict-free timetable generation")
    print("=" * 60)

    if not os.path.exists(INPUT_FILENAME):
        print(f"\nINFO: Input file '{INPUT_FILENAME}' not found.")
        create_simple_template(INPUT_FILENAME)
        print(f"\nNEXT STEPS:")
        print(f"1. Open '{INPUT_FILENAME}' in Excel")
        print(f"2. Update the data with your institution's information")
        print(f"3. Run this script again to generate the timetable")
        return

    print(f"INFO: Found input file: '{INPUT_FILENAME}'")
    
    try:
        # Parse input data
        print("INFO: Loading input data...")
        adapter = SimplifiedExcelAdapter(INPUT_FILENAME)
        parsed_data = adapter.parse()
        
        # Create scheduler
        print("INFO: Initializing scheduler...")
        scheduler = SimplifiedScheduler(**parsed_data)
        
        # Generate schedule
        print("INFO: Generating timetable...")
        result = scheduler.schedule(time_limit=60)
        
        # Display results
        print("\n" + "=" * 50)
        print("            RESULTS")
        print("=" * 50)
        
        if "error" in result:
            print(f"ERROR: {result['error']}")
        else:
            stats = result.get("statistics", {})
            print(f"Total Sessions: {stats.get('total_sessions', 0)}")
            print(f"Successfully Scheduled: {stats.get('scheduled_sessions', 0)}")
            print(f"Success Rate: {stats.get('success_rate', 0):.1f}%")
            
            if stats.get('unscheduled_sessions', 0) > 0:
                print(f"Unscheduled Sessions: {stats.get('unscheduled_sessions', 0)}")
                print("TIP: Try adding more faculty or rooms, or reduce course hours")
            
            print("=" * 50)
            
            # Write output files
            write_simple_output(result, scheduler, OUTPUT_DIRECTORY)
            
            if stats.get('success_rate', 0) > 80:
                print(f"\nSUCCESS: Timetable generated successfully!")
                print(f"Check '{OUTPUT_DIRECTORY}' for output files")
            else:
                print(f"\nPARTIAL SUCCESS: Some sessions couldn't be scheduled")
                print(f"Check constraints and try adjusting input data")

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        print(f"\nTroubleshooting:")
        print(f"1. Check Excel file format and column names")
        print(f"2. Ensure numeric fields contain valid numbers") 
        print(f"3. Verify subject codes match between subjects and batches sheets")

if __name__ == "__main__":
    main()
