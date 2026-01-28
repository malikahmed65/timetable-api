from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import os
from typing import Dict, List, Tuple, Set
from collections import defaultdict
import json
from datetime import datetime

app = FastAPI(title="Timetable Generator")

# ==================== DATA STRUCTURES ====================
class TimeSlot:
    def __init__(self, day: str, start_hour: int, end_hour: int):
        self.day = day
        self.start_hour = start_hour
        self.end_hour = end_hour
    
    def __str__(self):
        return f"{self.start_hour}:00-{self.end_hour}:00"

# ==================== TIMETABLE LOGIC ====================
class TimetableGenerator:
    def __init__(self):
        self.days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
        self.time_slots = self._generate_time_slots()
        self.status_messages = []
        
        # Structure: bookings[Day][Time] = {Set of occupied Room IDs}
        self.room_bookings = defaultdict(lambda: defaultdict(set))
    
    def log_status(self, message: str):
        """Log status messages for debugging"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        msg = f"[{timestamp}] {message}"
        self.status_messages.append(msg)
        print(msg)
    
    def _generate_time_slots(self) -> Dict[str, List[Tuple[int, int]]]:
        """Generate available time slots for each day"""
        slots = {}
        for day in self.days:
            slots[day] = []
            # 8 AM to 12 PM
            for hour in range(8, 12):
                slots[day].append((hour, hour + 1))
            
            # After lunch break
            if day == 'Friday':
                # 2 PM to 4 PM on Friday (after 12-2 break)
                for hour in range(14, 16):
                    slots[day].append((hour, hour + 1))
            else:
                # 1 PM to 4 PM on other days (after 12-1 break)
                for hour in range(13, 16):
                    slots[day].append((hour, hour + 1))
        
        return slots
    
    def parse_excel(self, file_content: bytes) -> Dict:
        """Parse Excel file and validate data"""
        try:
            self.log_status("📂 Reading Excel file...")
            excel_file = pd.ExcelFile(io.BytesIO(file_content), engine='openpyxl')
            
            required_sheets = ['session', 'Teacher', 'Sections', 'rooms']
            missing_sheets = [s for s in required_sheets if s not in excel_file.sheet_names]
            
            if missing_sheets:
                raise ValueError(f"❌ Missing sheets: {', '.join(missing_sheets)}")
            
            session_df = pd.read_excel(excel_file, sheet_name='session')
            teacher_df = pd.read_excel(excel_file, sheet_name='Teacher')
            sections_df = pd.read_excel(excel_file, sheet_name='Sections')
            rooms_df = pd.read_excel(excel_file, sheet_name='rooms')
            
            return {
                'session': session_df,
                'teacher': teacher_df,
                'sections': sections_df,
                'rooms': rooms_df
            }
        
        except Exception as e:
            self.log_status(f"❌ Error parsing Excel: {str(e)}")
            raise
    
    def build_teacher_mapping(self, teacher_df: pd.DataFrame) -> Dict[str, List[Dict]]:
        """Build mapping of subjects to teachers"""
        mapping = defaultdict(list)
        for idx, row in teacher_df.iterrows():
            teacher_name = str(row.get('Nmae', row.get('Name', 'Unknown'))).strip()
            courses = str(row.get('courses', '')).strip()
            credit_hours = row.get('credit hours', 0)
            
            if not teacher_name or pd.isna(teacher_name):
                continue
            
            course_list = [c.strip() for c in courses.split(',') if c.strip()]
            for course in course_list:
                mapping[course].append({
                    'name': teacher_name,
                    'credit_hours': credit_hours
                })
        return mapping

    def check_room_availability(self, room_id: str, day: str, times: List[str]) -> bool:
        """Check if a room is free for ALL requested times"""
        for time_slot in times:
            if room_id in self.room_bookings[day][time_slot]:
                return False
        return True

    def find_and_book_room(self, day: str, times: List[str], is_lab: bool, all_rooms: List[Dict]) -> str:
        """Finds a single room available for consecutive slots"""
        for room in all_rooms:
            r_id = str(room.get('room id', 'Unknown'))
            r_type = str(room.get('type', '')).lower()
            
            # 1. Check Type Match
            type_match = False
            if is_lab:
                if 'lab' in r_type or 'lab' in r_id.lower():
                    type_match = True
            else:
                if ('ke' in r_type or 'ke' in r_id.lower() or 'class' in r_type) and 'lab' not in r_type:
                    type_match = True
            
            if not type_match:
                continue

            # 2. Check Availability for ALL consecutive slots
            if self.check_room_availability(r_id, day, times):
                # Book it for all slots
                for t in times:
                    self.room_bookings[day][t].add(r_id)
                return r_id
        
        return "TBA"

    def generate_timetables(self, sections_df: pd.DataFrame, teacher_mapping: Dict, rooms_df: pd.DataFrame) -> Dict[str, List[Dict]]:
        try:
            self.log_status("📅 Generating timetables with Logic (Labs=3hrs)...")
            timetables = {}
            all_rooms = rooms_df.to_dict('records')
            self.room_bookings = defaultdict(lambda: defaultdict(set))
            
            # Master list of slots
            all_possible_slots = []
            for day in self.days:
                for slot in self.time_slots[day]:
                    all_possible_slots.append({'day': day, 'time': slot})
            
            total_slots = len(all_possible_slots)
            current_slot_index = 0
            
            for idx, row in sections_df.iterrows():
                section = str(row.get('Section', f'Section_{idx}')).strip()
                raw_subjects = str(row.get('Subject', '')).strip()
                
                if not section or not raw_subjects or pd.isna(raw_subjects):
                    continue
                    
                subject_list = [s.strip() for s in raw_subjects.split(',') if s.strip()]

                for subject in subject_list:
                    teachers = teacher_mapping.get(subject, [])
                    teacher_name = teachers[0]['name'] if teachers else "TBA"
                    try:
                        credit_hours = int(teachers[0].get('credit_hours', 3)) if teachers else 3
                    except:
                        credit_hours = 3

                    # === NEW LOGIC: DETECT LAB ===
                    # Check if subject ends with "Lab" (case insensitive)
                    is_lab = subject.lower().strip().endswith('lab') or 'lab' in subject.lower()

                    if is_lab:
                        # === LOGIC FOR LABS (3 Consecutive Hours) ===
                        lab_duration = 3
                        scheduled = False
                        
                        # Search for a block of 3 slots
                        # We try to start from current_slot_index to keep distribution fair, 
                        # but we scan forward until we find a fit.
                        attempts = 0
                        start_search_index = current_slot_index
                        
                        while attempts < total_slots:
                            idx_check = (start_search_index + attempts) % total_slots
                            
                            # Check boundaries: Can we fit 3 slots without wrapping days?
                            if idx_check + lab_duration > len(all_possible_slots):
                                attempts += 1
                                continue
                                
                            slots_to_check = all_possible_slots[idx_check : idx_check + lab_duration]
                            
                            # Verify all slots are on the same day
                            first_day = slots_to_check[0]['day']
                            if not all(s['day'] == first_day for s in slots_to_check):
                                attempts += 1
                                continue
                            
                            # Verify slots are consecutive in time (e.g., 8, 9, 10)
                            # (Our list is ordered, so indices usually guarantee this, but good to be safe)
                            
                            # Format Times
                            time_strings = [f"{s['time'][0]}:00-{s['time'][1]}:00" for s in slots_to_check]
                            
                            # Find Room for this BLOCK
                            room_id = self.find_and_book_room(first_day, time_strings, True, all_rooms)
                            
                            if room_id != "TBA":
                                # Found a spot! Schedule it.
                                display_room = "[LAB (Decide: LabE/Lab2)]"
                                
                                for i, slot_obj in enumerate(slots_to_check):
                                    t_str = time_strings[i]
                                    if section not in timetables: timetables[section] = []
                                    timetables[section].append({
                                        'subject': subject,
                                        'teacher': teacher_name,
                                        'day': first_day,
                                        'time': t_str,
                                        'room': display_room
                                    })
                                
                                scheduled = True
                                # Advance the global counter roughly past this block
                                current_slot_index = (idx_check + lab_duration) % total_slots
                                break
                            
                            attempts += 1
                        
                        if not scheduled:
                            self.log_status(f"⚠️ Could not find 3 consecutive slots for Lab: {subject}")

                    else:
                        # === LOGIC FOR THEORY (Normal Credit Hours) ===
                        for i in range(credit_hours):
                            safe_index = current_slot_index % total_slots
                            selected_slot = all_possible_slots[safe_index]
                            
                            day = selected_slot['day']
                            time_slot_tuple = selected_slot['time']
                            formatted_time = f"{time_slot_tuple[0]}:00-{time_slot_tuple[1]}:00"
                            
                            # Find Room (Theory) - 1 hour at a time
                            room_id = self.find_and_book_room(day, [formatted_time], False, all_rooms)
                            display_room = f"[{room_id}]"
                            
                            current_slot_index += 1
                            
                            if section not in timetables:
                                timetables[section] = []
                            
                            timetables[section].append({
                                'subject': subject,
                                'teacher': teacher_name,
                                'day': day,
                                'time': formatted_time,
                                'room': display_room
                            })
            
            self.log_status(f"✅ Generated timetables for {len(timetables)} sections")
            return timetables
        
        except Exception as e:
            self.log_status(f"❌ Error generating timetables: {str(e)}")
            raise
    
    def generate_word_document(self, timetables: Dict[str, List[Dict]]) -> bytes:
        """Generate Word document with timetables"""
        try:
            self.log_status("📄 Generating Word document...")
            doc = Document()
            
            title = doc.add_heading('University Timetable', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph(f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
            
            for section, classes in sorted(timetables.items()):
                doc.add_heading(f'Section: {section}', level=1)
                table = doc.add_table(rows=1, cols=6)
                table.style = 'Table Grid'
                
                header_cells = table.rows[0].cells
                header_cells[0].text = 'Time'
                for i, day in enumerate(self.days, 1):
                    header_cells[i].text = day
                
                schedule = defaultdict(lambda: defaultdict(str))
                all_times = set()
                
                for cls in classes:
                    entry_text = f"{cls['subject']}\n({cls['teacher']})\n{cls['room']}"
                    existing = schedule[cls['day']][cls['time']]
                    if existing:
                        schedule[cls['day']][cls['time']] = existing + f"\n\n{entry_text}"
                    else:
                        schedule[cls['day']][cls['time']] = entry_text
                    all_times.add(cls['time'])
                
                all_times.add('12:00-13:00')
                all_times.add('13:00-14:00') 
                times_list = sorted(list(all_times), key=lambda x: int(x.split(':')[0]))
                
                for time_slot in times_list:
                    row_cells = table.add_row().cells
                    row_cells[0].text = time_slot
                    start_hour = int(time_slot.split(':')[0])
                    
                    for day_idx, day in enumerate(self.days, 1):
                        if day == 'Friday' and 12 <= start_hour < 14:
                             row_cells[day_idx].text = 'JUMMAH BREAK'
                        elif 12 <= start_hour < 13:
                             row_cells[day_idx].text = 'BREAK'
                        else:
                            row_cells[day_idx].text = schedule[day].get(time_slot, '')
                doc.add_paragraph()
            
            doc_bytes = io.BytesIO()
            doc.save(doc_bytes)
            doc_bytes.seek(0)
            return doc_bytes.getvalue()
        except Exception as e:
            self.log_status(f"❌ Error generating Word document: {str(e)}")
            raise

# ==================== API ENDPOINTS ====================
@app.get("/")
def root():
    return {"message": "Timetable Generator API is Running", "version": "1.4"}

@app.post("/generate-timetable")
async def generate_timetable(file: UploadFile = File(...)):
    """Generate timetable from Excel file"""
    try:
        generator = TimetableGenerator()
        generator.log_status("🚀 Starting timetable generation...")
        
        file_content = await file.read()
        generator.log_status(f"📧 Received file: {file.filename}")
        
        data = generator.parse_excel(file_content)
        teacher_mapping = generator.build_teacher_mapping(data['teacher'])
        timetables = generator.generate_timetables(data['sections'], teacher_mapping, data['rooms'])
        
        if not timetables:
            return {
                "status": "error",
                "message": "No timetables could be generated.",
                "messages": generator.status_messages
            }
        
        word_content = generator.generate_word_document(timetables)
        return {
            "status": "success",
            "message": "Timetable generated successfully",
            "messages": generator.status_messages,
            "sections_processed": len(timetables)
        }
    
    except Exception as e:
        return {
            "status": "error",
            "message": f"Server Error: {str(e)}",
            "messages": generator.status_messages if 'generator' in locals() else []
        }

@app.post("/download-timetable")
async def download_timetable(file: UploadFile = File(...)):
    try:
        generator = TimetableGenerator()
        file_content = await file.read()
        data = generator.parse_excel(file_content)
        teacher_mapping = generator.build_teacher_mapping(data['teacher'])
        timetables = generator.generate_timetables(data['sections'], teacher_mapping, data['rooms'])
        
        if not timetables:
            raise HTTPException(status_code=400, detail="No timetables generated")
        
        word_content = generator.generate_word_document(timetables)
        output = io.BytesIO(word_content)
        output.seek(0)
        filename = f"Timetable_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        return StreamingResponse(
            iter([output.getvalue()]),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    print(f"🚀 Starting server on 0.0.0.0:{port}")
    uvicorn.run(app, host="0.0.0.0", port=port)