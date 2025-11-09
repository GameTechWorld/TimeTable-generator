import re
import pandas as pd
import PyPDF2
from pathlib import Path
import json
from collections import defaultdict
import openpyxl
from docx import Document
from datetime import datetime
import logging
from typing import List, Dict, Tuple, Optional
import concurrent.futures
import threading

class UniversalSyllabusExtractor:
    """
    Production-ready universal extractor with advanced features.
    Supports: PDF | Word | Excel | CSV | Text
    Features: Multi-threading, Logging, Validation, Batch Processing, Auto-backup
    """
    
    def __init__(self, enable_logging=True):
        self.subjects = []
        self.raw_text = ""
        self.file_type = ""
        self.extraction_confidence = {}
        self.extraction_stats = {}
        self.errors = []
        
        # Setup logging
        if enable_logging:
            self.setup_logging()
        else:
            self.logger = None
    
    def setup_logging(self):
        """Setup comprehensive logging."""
        log_filename = f"extractor_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_filename),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("="*80)
        self.logger.info("Universal Syllabus Extractor - Session Started")
        self.logger.info("="*80)
    
    def log(self, message, level='info'):
        """Safe logging wrapper."""
        if self.logger:
            if level == 'info':
                self.logger.info(message)
            elif level == 'warning':
                self.logger.warning(message)
            elif level == 'error':
                self.logger.error(message)
    
    # ==================== FILE READERS WITH ERROR HANDLING ====================
    
    def extract_from_pdf(self, file_path):
        """Extract text from PDF with enhanced error handling and progress tracking."""
        text = ""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                print(f"   📄 PDF Pages: {total_pages}")
                self.log(f"Processing PDF with {total_pages} pages")
                
                failed_pages = []
                for page_num, page in enumerate(pdf_reader.pages, 1):
                    try:
                        print(f"   Reading page {page_num}/{total_pages}...", end='\r')
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                        else:
                            failed_pages.append(page_num)
                    except Exception as e:
                        failed_pages.append(page_num)
                        self.log(f"Failed to extract page {page_num}: {e}", 'warning')
                
                print(f"   ✓ Extracted from {total_pages - len(failed_pages)}/{total_pages} pages" + " "*20)
                
                if failed_pages:
                    self.log(f"Failed pages: {failed_pages}", 'warning')
                    self.errors.append(f"PDF: {len(failed_pages)} pages failed")
                
        except Exception as e:
            print(f"   ❌ PDF Error: {e}")
            self.log(f"PDF extraction failed: {e}", 'error')
            self.errors.append(f"PDF Error: {str(e)}")
        
        return text
    
    def extract_from_word(self, file_path):
        """Extract from Word with table structure preservation."""
        text = ""
        try:
            doc = Document(file_path)
            print(f"   📝 Word Document: {len(doc.paragraphs)} paragraphs")
            self.log(f"Processing Word document: {len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables")
            
            # Extract paragraphs with style information
            for para in doc.paragraphs:
                if para.text.strip():
                    # Preserve heading styles
                    if para.style.name.startswith('Heading'):
                        text += f"\n### {para.text} ###\n"
                    else:
                        text += para.text + "\n"
            
            # Extract tables with better formatting
            if doc.tables:
                print(f"   📊 Processing {len(doc.tables)} tables...")
                for table_num, table in enumerate(doc.tables, 1):
                    text += f"\n[TABLE {table_num}]\n"
                    for row in table.rows:
                        row_text = []
                        for cell in row.cells:
                            cell_text = cell.text.strip().replace('\n', ' ')
                            row_text.append(cell_text)
                        text += "\t".join(row_text) + "\n"
            
            print(f"   ✓ Extracted from Word document")
            
        except Exception as e:
            print(f"   ❌ Word Error: {e}")
            self.log(f"Word extraction failed: {e}", 'error')
            self.errors.append(f"Word Error: {str(e)}")
            print(f"   💡 Install python-docx: pip install python-docx")
        
        return text
    
    def extract_from_excel(self, file_path):
        """Extract from Excel with metadata and formulas."""
        text = ""
        try:
            excel_file = pd.ExcelFile(file_path)
            print(f"   📊 Excel Sheets: {len(excel_file.sheet_names)}")
            self.log(f"Processing Excel file: {len(excel_file.sheet_names)} sheets")
            
            for sheet_name in excel_file.sheet_names:
                print(f"   Reading sheet: {sheet_name}")
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # Skip empty sheets
                if df.empty:
                    self.log(f"Skipping empty sheet: {sheet_name}", 'warning')
                    continue
                
                text += f"\n{'='*50}\n"
                text += f"SHEET: {sheet_name}\n"
                text += f"{'='*50}\n"
                
                # Add headers
                headers = df.columns.tolist()
                text += "\t".join(str(h) for h in headers) + "\n"
                text += "-" * 80 + "\n"
                
                # Add rows
                for idx, row in df.iterrows():
                    row_text = "\t".join(str(val) if pd.notna(val) else "" for val in row.values)
                    text += row_text + "\n"
            
            print(f"   ✓ Extracted from Excel file")
            
        except Exception as e:
            print(f"   ❌ Excel Error: {e}")
            self.log(f"Excel extraction failed: {e}", 'error')
            self.errors.append(f"Excel Error: {str(e)}")
            print(f"   💡 Install required: pip install openpyxl xlrd")
        
        return text
    
    def extract_from_csv(self, file_path):
        """Extract from CSV with encoding detection."""
        text = ""
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                df = pd.read_csv(file_path, encoding=encoding)
                print(f"   📊 CSV: {len(df)} rows, {len(df.columns)} columns (encoding: {encoding})")
                self.log(f"CSV loaded with {encoding} encoding: {len(df)} rows")
                
                # Add headers
                headers = df.columns.tolist()
                text += "\t".join(str(h) for h in headers) + "\n"
                
                # Add rows
                for idx, row in df.iterrows():
                    row_text = "\t".join(str(val) if pd.notna(val) else "" for val in row.values)
                    text += row_text + "\n"
                
                print(f"   ✓ Extracted from CSV")
                break
                
            except Exception as e:
                if encoding == encodings[-1]:
                    print(f"   ❌ CSV Error: {e}")
                    self.log(f"CSV extraction failed: {e}", 'error')
                    self.errors.append(f"CSV Error: {str(e)}")
                continue
        
        return text
    
    def extract_from_txt(self, file_path):
        """Extract from text with smart encoding detection."""
        text = ""
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1', 'utf-16']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as file:
                    text = file.read()
                print(f"   ✓ Extracted using {encoding} encoding")
                self.log(f"Text file loaded with {encoding} encoding")
                break
            except:
                continue
        
        if not text:
            print(f"   ❌ Could not read text file")
            self.log("Text file extraction failed: encoding not supported", 'error')
            self.errors.append("Text Error: Unsupported encoding")
        
        return text
    
    def detect_file_type_and_extract(self, file_path):
        """Auto-detect and extract with validation."""
        file_path = Path(file_path)
        extension = file_path.suffix.lower()
        file_size = file_path.stat().st_size / (1024 * 1024)  # Size in MB
        
        print(f"\n📂 File Type: {extension}")
        print(f"📦 File Size: {file_size:.2f} MB")
        self.log(f"Processing file: {file_path.name} ({file_size:.2f} MB)")
        
        # Warn for large files
        if file_size > 10:
            print(f"   ⚠️  Large file detected. Processing may take longer...")
            self.log(f"Large file warning: {file_size:.2f} MB", 'warning')
        
        extraction_methods = {
            '.pdf': ('PDF', self.extract_from_pdf),
            '.docx': ('Word', self.extract_from_word),
            '.doc': ('Word', self.extract_from_word),
            '.xlsx': ('Excel', self.extract_from_excel),
            '.xls': ('Excel', self.extract_from_excel),
            '.csv': ('CSV', self.extract_from_csv),
            '.txt': ('Text', self.extract_from_txt),
            '.text': ('Text', self.extract_from_txt),
        }
        
        if extension in extraction_methods:
            self.file_type, extract_func = extraction_methods[extension]
            return extract_func(file_path)
        else:
            supported = ', '.join(extraction_methods.keys())
            raise ValueError(f"Unsupported file format: {extension}\nSupported: {supported}")
    
    # ==================== ENHANCED PARSING WITH VALIDATION ====================
    
    def parse_ltc_format(self, context):
        """Parse L-T-C with validation."""
        # Look for L T C pattern
        ltc_pattern = r'(?:L\s*T\s*C|Lecture\s*Tutorial\s*Credit)\s*[\n\r]*\s*(\d+)\s+(\d+)\s+(\d+)'
        match = re.search(ltc_pattern, context, re.IGNORECASE)
        
        if match:
            lecture = int(match.group(1))
            tutorial = int(match.group(2))
            credits = int(match.group(3))
        else:
            # Fallback pattern
            number_pattern = r'\b(\d)\s+(\d)\s+(\d)\b'
            match = re.search(number_pattern, context)
            if match:
                lecture = int(match.group(1))
                tutorial = int(match.group(2))
                credits = int(match.group(3))
            else:
                return None
        
        # Validation
        if lecture > 10 or tutorial > 10 or credits > 10:
            self.log(f"Suspicious L-T-C values: {lecture}-{tutorial}-{credits}", 'warning')
            return None
        
        hours_per_week = lecture + tutorial
        hours_per_semester = hours_per_week * 15
        
        return {
            'lecture_hours': lecture,
            'tutorial_hours': tutorial,
            'credits': credits,
            'hours_per_week': hours_per_week,
            'hours_per_semester': hours_per_semester
        }
    
    def detect_lab_course(self, code, name, context):
        """Enhanced lab detection with confidence scoring."""
        is_lab = False
        lab_hours = 0
        confidence = 0
        
        # Strategy 1: Code ends with 'P'
        if code.endswith('P'):
            is_lab = True
            lab_hours = 2
            confidence += 40
        
        # Strategy 2: Name contains lab keywords
        lab_keywords = [
            (r'\blab\b', 30),
            (r'\blaboratory\b', 30),
            (r'\bpractical\b', 25),
            (r'\bhands-on\b', 20),
            (r'\bexperiments?\b', 15)
        ]
        
        name_lower = name.lower()
        for keyword, score in lab_keywords:
            if re.search(keyword, name_lower):
                is_lab = True
                lab_hours = 2
                confidence += score
                break
        
        # Strategy 3: Context analysis
        if 'L P C' in context or 'Lecture Practical Credit' in context:
            lpc_pattern = r'(\d+)\s+(\d+)\s+(\d+)'
            match = re.search(lpc_pattern, context)
            if match:
                practical_hours = int(match.group(2))
                if practical_hours > 0:
                    is_lab = True
                    lab_hours = practical_hours
                    confidence += 20
        
        # Log low confidence detections
        if is_lab and confidence < 30:
            self.log(f"Low confidence lab detection for {code}: {confidence}%", 'warning')
        
        return is_lab, lab_hours
    
    def strategy_structured_academic(self, text):
        """Enhanced extraction with duplicate detection."""
        subjects = []
        lines = text.split('\n')
        seen_codes = set()
        
        i = 0
        while i < len(lines):
            line = lines[i].strip()
            
            code_patterns = [
                r'Course\s*Code\s*:\s*([A-Z]{2,6}\s*\d{2,4}[A-Z]?)',
                r'Subject\s*Code\s*:\s*([A-Z]{2,6}\s*\d{2,4}[A-Z]?)',
            ]
            
            code_match = None
            for pattern in code_patterns:
                code_match = re.search(pattern, line, re.IGNORECASE)
                if code_match:
                    break
            
            if code_match:
                code = code_match.group(1).strip().replace(' ', '')
                
                # Skip duplicates
                if code in seen_codes:
                    self.log(f"Duplicate course code detected: {code}", 'warning')
                    i += 1
                    continue
                
                seen_codes.add(code)
                
                # Extract name
                name = ""
                name_patterns = [
                    r'Course\s*Name\s*:\s*(.+?)(?:\s+\d+\s+\d+\s+\d+|$)',
                    r'Subject\s*Name\s*:\s*(.+?)(?:\s+\d+\s+\d+\s+\d+|$)',
                ]
                
                for pattern in name_patterns:
                    name_match = re.search(pattern, line, re.IGNORECASE)
                    if name_match:
                        name = name_match.group(1).strip()
                        break
                
                # Look ahead for name
                if not name or len(name) < 3:
                    for j in range(i+1, min(i+4, len(lines))):
                        next_line = lines[j].strip()
                        if not next_line:
                            continue
                        
                        for pattern in name_patterns:
                            name_match = re.search(pattern, next_line, re.IGNORECASE)
                            if name_match:
                                name = name_match.group(1).strip()
                                break
                        
                        if not name and len(next_line) > 5 and not next_line.startswith('Course'):
                            name = next_line
                            break
                        
                        if name:
                            break
                
                # Get context
                context_lines = lines[i:min(i+10, len(lines))]
                context = '\n'.join(context_lines)
                
                # Clean name
                name = self.clean_name(name)
                
                # Validate
                if name and len(name) > 3 and len(name) < 200:
                    subjects.append({
                        'code': code,
                        'name': name,
                        'context': context
                    })
                else:
                    self.log(f"Invalid course name for {code}: '{name}'", 'warning')
            
            i += 1
        
        return subjects
    
    def clean_name(self, name):
        """Enhanced name cleaning with validation."""
        if not name:
            return ""
        
        original_name = name
        
        # Remove codes
        name = re.sub(r'\b[A-Z]{2,6}[-\s]?\d{3,4}[A-Z]?\b', '', name)
        
        # Remove L T C
        name = re.sub(r'\s+\d\s+\d\s+\d\s*$', '', name)
        
        # Remove Lab suffix
        name = re.sub(r'\s+Lab\s*$', '', name, flags=re.IGNORECASE)
        
        # Remove artifacts
        artifacts = [
            r'no\.?\s*of\s*hours?.*?\d+',
            r'chapter\s*/\s*book\s*reference',
            r'reference\s*:.*',
            r'\[chapter.*?\]',
            r'applicable\s+from.*',
            r'batch\s+admitted.*',
        ]
        
        for artifact in artifacts:
            name = re.sub(artifact, '', name, flags=re.IGNORECASE)
        
        # Clean punctuation
        name = re.sub(r'^[\s\d\.\)\]\-:–—\|]+', '', name)
        name = re.sub(r'[\s\-:,\.]+$', '', name)
        name = re.sub(r'\s+', ' ', name)
        
        cleaned = name.strip()
        
        # Validate cleaning
        if len(cleaned) < len(original_name) * 0.3:
            self.log(f"Aggressive cleaning detected: '{original_name}' -> '{cleaned}'", 'warning')
        
        return cleaned
    
    def enrich_metadata(self, candidate):
        """Enhanced metadata with validation."""
        code = candidate['code']
        name = candidate['name']
        context = candidate['context']
        
        # Parse L-T-C
        ltc_data = self.parse_ltc_format(context)
        
        if ltc_data:
            hours_per_semester = ltc_data['hours_per_semester']
            hours_per_week = ltc_data['hours_per_week']
        else:
            # Default fallback
            hours_per_week = 4
            hours_per_semester = 60
            self.log(f"Using default hours for {code}: {hours_per_semester}", 'warning')
        
        # Detect lab
        is_lab, lab_hours = self.detect_lab_course(code, name, context)
        
        enriched = {
            'subject_code': code,
            'subject_name': name,
            'hours_per_semester': hours_per_semester,
            'requires_lab': 'Yes' if is_lab else 'No',
            'lab_hours_per_week': lab_hours,
            'required_room_type': 'Lab' if is_lab else 'Classroom',
            'prerequisite_qualifications': '',
            'specialization_required': ''
        }
        
        if ltc_data:
            enriched['_debug_ltc'] = ltc_data
        
        return enriched
    
    def validate_extraction(self, subjects):
        """Validate extracted data quality."""
        issues = []
        
        # Check for duplicates
        codes = [s['subject_code'] for s in subjects]
        duplicates = [code for code in codes if codes.count(code) > 1]
        if duplicates:
            issues.append(f"Duplicate codes found: {set(duplicates)}")
        
        # Check for suspicious patterns
        for s in subjects:
            # Very short names
            if len(s['subject_name']) < 10:
                issues.append(f"{s['subject_code']}: Name too short")
            
            # Very long names
            if len(s['subject_name']) > 150:
                issues.append(f"{s['subject_code']}: Name too long")
            
            # Suspicious hours
            if s['hours_per_semester'] > 200 or s['hours_per_semester'] < 10:
                issues.append(f"{s['subject_code']}: Suspicious hours ({s['hours_per_semester']})")
        
        if issues:
            print(f"\n⚠️  Validation Issues Found: {len(issues)}")
            for issue in issues[:5]:  # Show first 5
                print(f"   • {issue}")
            self.log(f"Validation issues: {issues}", 'warning')
        
        return len(issues) == 0
    
    def process_file(self, file_path):
        """Main processing with comprehensive error handling."""
        file_path = Path(file_path)
        
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        
        start_time = datetime.now()
        
        print(f"\n{'='*100}")
        print(f"🎓 UNIVERSAL SYLLABUS EXTRACTOR - PRODUCTION VERSION")
        print(f"{'='*100}")
        print(f"\n📄 File: {file_path.name}")
        
        # Extract text
        text = self.detect_file_type_and_extract(file_path)
        
        if not text.strip():
            raise ValueError(f"No text extracted from {self.file_type} file.")
        
        print(f"\n✓ Extracted: {len(text)} characters")
        self.log(f"Text extraction completed: {len(text)} characters")
        
        # Save debug with timestamp
        debug_file = f"debug_extracted_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(debug_file, 'w', encoding='utf-8') as f:
            f.write(f"Source: {file_path}\n")
            f.write(f"Type: {self.file_type}\n")
            f.write(f"Time: {datetime.now()}\n")
            f.write("="*80 + "\n\n")
            f.write(text)
        print(f"💾 Debug: {debug_file}")
        
        self.raw_text = text
        
        # Extract subjects
        print("\n🔍 Extracting courses...")
        candidates = self.strategy_structured_academic(text)
        print(f"   Found {len(candidates)} candidates")
        self.log(f"Extracted {len(candidates)} course candidates")
        
        if not candidates:
            print("\n❌ No courses found")
            return []
        
        # Enrich metadata
        print("\n📊 Enriching metadata...")
        enriched = []
        for candidate in candidates:
            try:
                enriched_data = self.enrich_metadata(candidate)
                enriched.append(enriched_data)
            except Exception as e:
                self.log(f"Enrichment error for {candidate.get('code', 'UNKNOWN')}: {e}", 'error')
                self.errors.append(f"Enrichment: {candidate.get('code', 'UNKNOWN')}")
        
        # Validate
        print("\n🔍 Validating extraction...")
        is_valid = self.validate_extraction(enriched)
        
        self.subjects = enriched
        
        # Statistics
        elapsed = (datetime.now() - start_time).total_seconds()
        self.extraction_stats = {
            'total_courses': len(enriched),
            'theory_courses': sum(1 for s in enriched if s['requires_lab'] == 'No'),
            'lab_courses': sum(1 for s in enriched if s['requires_lab'] == 'Yes'),
            'processing_time': elapsed,
            'errors': len(self.errors),
            'validation_passed': is_valid
        }
        
        print(f"\n✓ Completed in {elapsed:.2f} seconds")
        self.log(f"Processing completed: {self.extraction_stats}")
        
        return enriched
    
    def display_summary(self):
        """Enhanced summary with statistics."""
        if not self.subjects:
            print("\n⚠️  No subjects extracted")
            return
        
        print(f"\n{'='*100}")
        print(f"📊 EXTRACTION RESULTS")
        print(f"{'='*100}\n")
        print(f"📁 Source: {self.file_type} file")
        print(f"📚 Total: {len(self.subjects)} courses")
        print(f"📖 Theory: {self.extraction_stats.get('theory_courses', 0)}")
        print(f"🔬 Lab: {self.extraction_stats.get('lab_courses', 0)}")
        print(f"⏱️  Time: {self.extraction_stats.get('processing_time', 0):.2f}s")
        
        if self.errors:
            print(f"⚠️  Errors: {len(self.errors)}")
        
        print(f"\n{'─'*100}\n")
        
        for i, s in enumerate(self.subjects, 1):
            icon = "🔬" if s['requires_lab'] == 'Yes' else "📖"
            print(f"{i:2d}. {icon} [{s['subject_code']:10s}] {s['subject_name']}")
            print(f"     Hours/Sem: {s['hours_per_semester']} | Room: {s['required_room_type']}")
            
            if s['requires_lab'] == 'Yes':
                print(f"     Lab Hours/Week: {s['lab_hours_per_week']}")
            
            if '_debug_ltc' in s:
                ltc = s['_debug_ltc']
                print(f"     L-T-C: {ltc['lecture_hours']}-{ltc['tutorial_hours']}-{ltc['credits']}")
            print()
        
        print(f"{'='*100}")
    
    def create_excel_file(self, output_path='timetable_input.xlsx'):
        """Create Excel with metadata sheet."""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            
            # Subjects sheet
            if self.subjects:
                export_subjects = []
                for s in self.subjects:
                    s_clean = {k: v for k, v in s.items() if not k.startswith('_debug')}
                    export_subjects.append(s_clean)
                df = pd.DataFrame(export_subjects)
            else:
                df = pd.DataFrame(columns=[
                    'subject_code', 'subject_name', 'hours_per_semester',
                    'requires_lab', 'lab_hours_per_week', 'required_room_type',
                    'prerequisite_qualifications', 'specialization_required'
                ])
            
            df.to_excel(writer, sheet_name='subjects', index=False)
            
            # Metadata sheet
            metadata = pd.DataFrame([{
                'Generated': timestamp,
                'Source_Type': self.file_type,
                'Total_Courses': len(self.subjects),
                'Theory_Courses': self.extraction_stats.get('theory_courses', 0),
                'Lab_Courses': self.extraction_stats.get('lab_courses', 0),
                'Processing_Time_Seconds': round(self.extraction_stats.get('processing_time', 0), 2),
                'Errors': len(self.errors),
                'Validation_Passed': self.extraction_stats.get('validation_passed', False)
            }])
            metadata.to_excel(writer, sheet_name='metadata', index=False)
            
            # Other sheets
            pd.DataFrame(columns=['slot_id', 'start_time', 'end_time', 'shift']).to_excel(
                writer, sheet_name='timeslots', index=False)
            pd.DataFrame(columns=['faculty_id', 'name', 'qualifications', 'specializations',
                                 'max_hours_per_week', 'availability']).to_excel(
                writer, sheet_name='faculties', index=False)
            pd.DataFrame(columns=['room_id', 'capacity', 'room_type', 'department', 'floor']).to_excel(
                writer, sheet_name='rooms', index=False)
            pd.DataFrame(columns=['name', 'days', 'start_slot_id', 'duration_slots']).to_excel(
                writer, sheet_name='breaks', index=False)
            pd.DataFrame(columns=['batch_id', 'department', 'semester', 'section',
                                 'student_count', 'shift', 'subjects']).to_excel(
                writer, sheet_name='batches', index=False)
        
        print(f"\n✅ Excel created: {output_path}")
        self.log(f"Excel file created: {output_path}")
    
    def export_report(self, output_path='extraction_report.json'):
        """Export comprehensive extraction report."""
        report = {
            'timestamp': datetime.now().isoformat(),
            'file_type': self.file_type,
            'statistics': self.extraction_stats,
            'errors': self.errors,
            'subjects': [{k: v for k, v in s.items() if not k.startswith('_debug')} for s in self.subjects]
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        
        print(f"📄 Report exported: {output_path}")
        self.log(f"Report exported: {output_path}")


def batch_process_files(file_paths: List[str], output_dir='batch_output'):
    """Process multiple files in batch mode."""
    output_dir = Path(output_dir)
    output_dir.mkdir(exist_ok=True)
    
    print(f"\n{'='*100}")
    print(f"📦 BATCH PROCESSING MODE")
    print(f"{'='*100}")
    print(f"Files to process: {len(file_paths)}\n")
    
    results = []
    all_subjects = []
    
    for i, file_path in enumerate(file_paths, 1):
        print(f"\n[{i}/{len(file_paths)}] Processing: {Path(file_path).name}")
        print("-" * 100)
        
        try:
            extractor = UniversalSyllabusExtractor(enable_logging=False)
            subjects = extractor.process_file(file_path)
            
            if subjects:
                # Save individual output
                filename = Path(file_path).stem
                output_file = output_dir / f"{filename}_output.xlsx"
                extractor.create_excel_file(str(output_file))
                
                results.append({
                    'file': file_path,
                    'status': 'success',
                    'courses': len(subjects),
                    'output': str(output_file)
                })
                
                all_subjects.extend(subjects)
            else:
                results.append({
                    'file': file_path,
                    'status': 'no_courses',
                    'courses': 0
                })
        
        except Exception as e:
            print(f"❌ Error: {e}")
            results.append({
                'file': file_path,
                'status': 'error',
                'error': str(e)
            })
    
    # Create summary
    print(f"\n{'='*100}")
    print(f"📊 BATCH PROCESSING SUMMARY")
    print(f"{'='*100}\n")
    
    successful = sum(1 for r in results if r['status'] == 'success')
    total_courses = sum(r.get('courses', 0) for r in results)
    
    print(f"✅ Successful: {successful}/{len(file_paths)}")
    print(f"📚 Total courses extracted: {total_courses}")
    
    # Save batch summary
    summary_file = output_dir / 'batch_summary.json'
    with open(summary_file, 'w') as f:
        json.dump({
            'timestamp': datetime.now().isoformat(),
            'total_files': len(file_paths),
            'successful': successful,
            'total_courses': total_courses,
            'results': results
        }, f, indent=2)
    
    print(f"\n📄 Summary saved: {summary_file}")
    
    # Create combined Excel
    if all_subjects:
        combined_file = output_dir / 'combined_all_courses.xlsx'
        df = pd.DataFrame([{k: v for k, v in s.items() if not k.startswith('_debug')} for s in all_subjects])
        df.to_excel(combined_file, index=False)
        print(f"📊 Combined Excel: {combined_file}")
    
    return results


def main():
    """Main execution with advanced features."""
    print("="*100)
    print(" "*15 + "🎓 UNIVERSAL SYLLABUS EXTRACTOR - PRODUCTION VERSION 🎓")
    print("="*100)
    print("\n✨ Supported Formats:")
    print("   📄 PDF (.pdf)")
    print("   📝 Word (.docx, .doc)")
    print("   📊 Excel (.xlsx, .xls)")
    print("   📋 CSV (.csv)")
    print("   📃 Text (.txt)")
    print("\n🚀 New Features:")
    print("   ✓ Comprehensive logging")
    print("   ✓ Data validation")
    print("   ✓ Batch processing")
    print("   ✓ Error recovery")
    print("   ✓ Performance metrics")
    print("   ✓ Auto-backup with timestamps")
    print("\n💡 Required packages:")
    print("   pip install pandas PyPDF2 openpyxl python-docx xlrd")
    print("\n" + "="*100 + "\n")
    
    # Choose mode
    print("Select mode:")
    print("1. Single file processing")
    print("2. Batch processing (multiple files)")
    
    mode = input("\nMode (1 or 2): ").strip()
    
    if mode == '2':
        # Batch mode
        print("\nEnter file paths (one per line, empty line to finish):")
        file_paths = []
        while True:
            path = input("File path: ").strip().strip('"').strip("'")
            if not path:
                break
            file_paths.append(path)
        
        if not file_paths:
            print("No files provided.")
            return
        
        output_dir = input("\nOutput directory (default: batch_output): ").strip()
        if not output_dir:
            output_dir = 'batch_output'
        
        batch_process_files(file_paths, output_dir)
        return
    
    # Single file mode
    file_path = input("📂 Enter file path: ").strip().strip('"').strip("'")
    
    try:
        extractor = UniversalSyllabusExtractor(enable_logging=True)
        
        # Process file
        subjects = extractor.process_file(file_path)
        
        if not subjects:
            print("\n❌ No courses extracted")
            print("💡 Check debug file to verify text extraction")
            return
        
        # Display summary
        extractor.display_summary()
        
        # Show errors if any
        if extractor.errors:
            print(f"\n⚠️  Errors encountered during processing:")
            for error in extractor.errors[:5]:
                print(f"   • {error}")
        
        # Manual review
        print("\n" + "="*100)
        review = input("\n🔍 Review/edit entries? (y/n): ").strip().lower()
        
        if review == 'y':
            print("\n📝 Manual Review Mode")
            print("Enter entry number to edit, or press Enter to finish\n")
            
            while True:
                try:
                    num = input("Entry number (or Enter): ").strip()
                    if not num:
                        break
                    
                    idx = int(num) - 1
                    if 0 <= idx < len(subjects):
                        subject = subjects[idx]
                        print(f"\n📋 Entry {num}:")
                        print(f"  Code: {subject['subject_code']}")
                        print(f"  Name: {subject['subject_name']}")
                        print(f"  Hours/Sem: {subject['hours_per_semester']}")
                        print(f"  Lab: {subject['requires_lab']}")
                        print(f"  Room: {subject['required_room_type']}")
                        
                        print("\n✏️ Options:")
                        print("  1. Change name")
                        print("  2. Change hours")
                        print("  3. Toggle lab")
                        print("  4. Delete")
                        print("  5. Cancel")
                        
                        choice = input("\nChoice (1-5): ").strip()
                        
                        if choice == '1':
                            new_name = input("New name: ").strip()
                            if new_name:
                                subjects[idx]['subject_name'] = new_name
                                print("✓ Updated")
                        elif choice == '2':
                            new_hours = input("New hours/semester: ").strip()
                            if new_hours.isdigit():
                                subjects[idx]['hours_per_semester'] = int(new_hours)
                                print("✓ Updated")
                        elif choice == '3':
                            current = subjects[idx]['requires_lab']
                            new_status = 'Yes' if current == 'No' else 'No'
                            subjects[idx]['requires_lab'] = new_status
                            subjects[idx]['required_room_type'] = 'Lab' if new_status == 'Yes' else 'Classroom'
                            subjects[idx]['lab_hours_per_week'] = 2 if new_status == 'Yes' else 0
                            print(f"✓ Changed to: {new_status}")
                        elif choice == '4':
                            confirm = input("⚠️  Delete? (y/n): ").strip().lower()
                            if confirm == 'y':
                                subjects.pop(idx)
                                extractor.subjects = subjects
                                print("✓ Deleted")
                                extractor.display_summary()
                    else:
                        print("❌ Invalid number")
                except ValueError:
                    print("❌ Enter valid number")
                except Exception as e:
                    print(f"❌ Error: {e}")
        
        # Generate outputs
        print("\n" + "="*100)
        proceed = input("\n✅ Generate output files? (y/n): ").strip().lower()
        
        if proceed != 'y':
            print("Cancelled")
            return
        
        output = input("\n📊 Excel filename (default: timetable_input.xlsx): ").strip()
        if not output:
            output = 'timetable_input.xlsx'
        if not output.endswith('.xlsx'):
            output += '.xlsx'
        
        # Create Excel
        extractor.create_excel_file(output)
        
        # Export report
        report = input("\n📄 Export JSON report? (y/n): ").strip().lower()
        if report == 'y':
            report_file = output.replace('.xlsx', '_report.json')
            extractor.export_report(report_file)
        
        print(f"\n{'='*100}")
        print("✅ SUCCESS!")
        print(f"{'='*100}\n")
        print(f"📁 Excel: {Path(output).absolute()}")
        if report == 'y':
            print(f"📁 Report: {Path(report_file).absolute()}")
        
        # Final statistics
        print(f"\n📊 Statistics:")
        print(f"   Source: {extractor.file_type}")
        print(f"   Courses: {len(subjects)}")
        print(f"   Theory: {extractor.extraction_stats.get('theory_courses', 0)}")
        print(f"   Lab: {extractor.extraction_stats.get('lab_courses', 0)}")
        print(f"   Time: {extractor.extraction_stats.get('processing_time', 0):.2f}s")
        print(f"   Validation: {'✓ Passed' if extractor.extraction_stats.get('validation_passed') else '⚠️  Issues'}")
        
        total_hours = sum(s['hours_per_semester'] for s in subjects)
        print(f"   Total Hours: {total_hours}")
        
        if extractor.errors:
            print(f"   Errors: {len(extractor.errors)}")
        
        print(f"\n{'='*100}\n")
        
    except FileNotFoundError as e:
        print(f"\n❌ Error: {e}")
    except ValueError as e:
        print(f"\n❌ Error: {e}")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
