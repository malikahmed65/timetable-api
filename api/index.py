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
        
        # NEW: Track room bookings to prevent conflicts
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
            # Use openpyxl engine explicitly for .xlsx
            excel_file = pd.ExcelFile(io.BytesIO(file_content), engine='openpyxl')
            
            # Check required sheets
            required_sheets = ['session', 'Teacher', 'Sections', 'rooms']
            missing_sheets = [s for s in required_sheets if s not in excel_file.sheet_names]
            
            if missing_sheets:
                raise ValueError(f"❌ Missing sheets: {', '.join(missing_sheets)}")
            
            self.log_status("✅ All required sheets found")
            
            # Read sheets
            session_df = pd.read_excel(excel_file, sheet_name='session')
            teacher_df = pd.read_excel(excel_file, sheet_name='Teacher')
            sections_df = pd.read_excel(excel_file, sheet_name='Sections')
            rooms_df = pd.read_excel(excel_file, sheet_name='rooms')
            
            self.log_status(f"✅ Loaded {len(session_df)} sessions")
            self.log_status(f"✅ Loaded {len(teacher_df)} teachers")
            self.log_status(f"✅ Loaded {len(sections_df)} sections")
            self.log_status(f"✅ Loaded {len(rooms_df)} rooms")
            
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
        try:
            self.log_status("🔗 Building teacher-subject mapping...")
            mapping = defaultdict(list)
            
            for idx, row in teacher_df.iterrows():
                # Handle 'Nmae' typo from your specific Excel file
                teacher_name = str(row.get('Nmae', row.get('Name', 'Unknown'))).strip()
                courses = str(row.get('courses', '')).strip()
                credit_hours = row.get('credit hours', 0)
                
                if not teacher_name or pd.isna(teacher_name):
                    continue
                
                # Split courses by comma if multiple
                course_list = [c.strip() for c in courses.split(',') if c.strip()]
                
                for course in course_list:
                    mapping[course].append({
                        'name': teacher_name,
                        'credit_hours': credit_hours
                    })
            
            self.log_status(f"✅ Mapped {len(mapping)} subjects to teachers")
            return mapping
        
        except Exception as e:
            self.log_status(f"❌ Error building teacher mapping: {str(e)}")
            raise

    def get_available_room(self, day: str, time_slot: str, is_lab: bool, all_rooms: List[Dict]) -> str:
        """Finds a free room of the correct type (KE/Theory vs Lab)"""
        
        for room in all_rooms:
            r_id = str(room.get('room id', 'Unknown'))
            r_type = str(room.get('type', '')).lower()
            
            # 1. Check if Room Type matches the Subject Type
            type_match = False
            if is_lab:
                # If subject is a Lab, we need a room with 'lab' in its type or ID
                if 'lab' in r_type or 'lab' in r_id.lower():
                    type_match = True
            else:
                # If subject is Theory, we need 'KE', 'Class', etc. (Not a lab)
                if ('ke' in r_type or 'ke' in r_id.lower() or 'class' in r_type) and 'lab' not in r_type:
                    type_match = True
            
            if not type_match:
                continue

            # 2. Check Availability (Is this room free at this time?)
            if r_id not in self.room_bookings[day][time_slot]:
                # Book the room
                self.room_bookings[day][time_slot].add(r_id)
                return r_id
        
        return "TBA" # No room found

    def generate_timetables(self, sections_df: pd.DataFrame, teacher_mapping: Dict, rooms_df: pd.DataFrame) -> Dict[str, List[Dict]]:
        """
        Generate timetables for each section with Room Allocation.
        """
        try:
            self.log_status("📅 Generating timetables...")
            timetables = {}
            
            # Convert rooms dataframe to list of dicts for faster looping
            all_rooms = rooms_df.to_dict('records')

            # Reset bookings
            self.room_bookings = defaultdict(lambda: defaultdict(set))
            
            # 1. CREATE MASTER LIST OF ALL AVAILABLE SLOTS
            all_possible_slots = []
            for day in self.days:
                for slot in self.time_slots[day]:
                    all_possible_slots.append({'day': day, 'time': slot})
            
            total_slots = len(all_possible_slots)
            if total_slots == 0:
                raise ValueError("No time slots defined in logic!")

            # Global counter to distribute classes evenly
            current_slot_index = 0
            
            for idx, row in sections_df.iterrows():
                section = str(row.get('Section', f'Section_{idx}')).strip()
                raw_subjects = str(row.get('Subject', '')).strip()
                
                if not section or not raw_subjects or pd.isna(raw_subjects):
                    continue
                    
                subject_list = [s.strip() for s in raw_subjects.split(',') if s.strip()]

                for subject in subject_list:
                    # Get teacher
                    teachers = teacher_mapping.get(subject, [])
                    teacher_name = "TBA"
                    credit_hours = 3
                    
                    if teachers:
                        teacher_name = teachers[0]['name']
                        try:
                            credit_hours = int(teachers[0].get('credit_hours', 3))
                        except:
                            credit_hours = 3
                    else:
                        self.log_status(f"⚠️  No teacher found for {subject}")

                    # Check if this subject is a LAB
                    is_lab = 'lab' in subject.lower()

                    # Schedule multiple lectures based on credit hours
                    for i in range(credit_hours):
                        
                        # Select Slot (Round Robin)
                        safe_index = current_slot_index % total_slots
                        selected_slot = all_possible_slots[safe_index]
                        
                        day = selected_slot['day']
                        time_slot_tuple = selected_slot['time']
                        formatted_time = f"{time_slot_tuple[0]}:00-{time_slot_tuple[1]}:00"
                        
                        # === FIND A ROOM ===
                        assigned_room_id = self.get_available_room(day, formatted_time, is_lab, all_rooms)
                        
                        # Custom Display Logic for Labs
                        display_room = f"[{assigned_room_id}]"
                        if is_lab:
                            # Even if we booked a specific lab (e.g. LabE) to reserve space,
                            # we display the generic instruction as requested.
                            display_room = "[LAB (Decide: LabE/Lab2)]"
                        
                        # Increment counter
                        current_slot_index += 1
                        
                        if section not in timetables:
                            timetables[section] = []
                        
                        timetables[section].append({
                            'subject': subject,
                            'teacher': teacher_name,
                            'day': day,
                            'time': formatted_time,
                            'room': display_room  # Save room info
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
            
            # Add title
            title = doc.add_heading('University Timetable', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph(f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
            doc.add_paragraph()
            
            # Add timetable for each section
            for section, classes in sorted(timetables.items()):
                doc.add_heading(f'Section: {section}', level=1)
                
                # Create table (6 columns: time + 5 days)
                table = doc.add_table(rows=1, cols=6)
                table.style = 'Table Grid'
                
                # Header row
                header_cells = table.rows[0].cells
                header_cells[0].text = 'Time'
                for i, day in enumerate(self.days, 1):
                    header_cells[i].text = day
                
                # Organize data
                schedule = defaultdict(lambda: defaultdict(str))
                all_times = set()
                
                for cls in classes:
                    # Content format: Subject \n (Teacher) \n [Room]
                    entry_text = f"{cls['subject']}\n({cls['teacher']})\n{cls['room']}"
                    
                    existing = schedule[cls['day']][cls['time']]
                    if existing:
                        schedule[cls['day']][cls['time']] = existing + f"\n\n{entry_text}"
                    else:
                        schedule[cls['day']][cls['time']] = entry_text
                        
                    all_times.add(cls['time'])
                
                # Add Standard Breaks
                all_times.add('12:00-13:00')
                all_times.add('13:00-14:00') 
                
                # Sort times numerically
                times_list = sorted(list(all_times), key=lambda x: int(x.split(':')[0]))
                
                for time_slot in times_list:
                    row_cells = table.add_row().cells
                    row_cells[0].text = time_slot
                    
                    start_hour = int(time_slot.split(':')[0])
                    
                    for day_idx, day in enumerate(self.days, 1):
                        # Friday Prayer Logic
                        if day == 'Friday' and 12 <= start_hour < 14:
                             row_cells[day_idx].text = 'JUMMAH BREAK'
                        # General Lunch Logic
                        elif 12 <= start_hour < 13:
                             row_cells[day_idx].text = 'BREAK'
                        else:
                            row_cells[day_idx].text = schedule[day].get(time_slot, '')
                
                doc.add_paragraph()
            
            # Convert to bytes
            doc_bytes = io.BytesIO()
            doc.save(doc_bytes)
            doc_bytes.seek(0)
            
            self.log_status("✅ Word document generated successfully")
            return doc_bytes.getvalue()
        
        except Exception as e:
            self.log_status(f"❌ Error generating Word document: {str(e)}")
            raise

# ==================== API ENDPOINTS ====================
@app.get("/")
def root():
    return {"message": "Timetable Generator API is Running", "version": "1.3"}

@app.post("/generate-timetable")
async def generate_timetable(file: UploadFile = File(...)):
    """Generate timetable from Excel file"""
    try:
        generator = TimetableGenerator()
        generator.log_status("🚀 Starting timetable generation...")
        
        # Read file content
        file_content = await file.read()
        generator.log_status(f"📧 Received file: {file.filename}")
        
        # Parse Excel
        data = generator.parse_excel(file_content)
        
        # Build teacher mapping
        teacher_mapping = generator.build_teacher_mapping(data['teacher'])
        
        # Generate timetables (PASSING ROOMS NOW)
        timetables = generator.generate_timetables(data['sections'], teacher_mapping, data['rooms'])
        
        if not timetables:
            return {
                "status": "error",
                "message": "No timetables could be generated. Check input file format.",
                "messages": generator.status_messages
            }
        
        # Generate Word document
        word_content = generator.generate_word_document(timetables)
        
        # Save to temporary file
        temp_path = f"timetable_output.docx"
        with open(temp_path, 'wb') as f:
            f.write(word_content)
        
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
    """Generate and download timetable directly"""
    try:
        generator = TimetableGenerator()
        
        file_content = await file.read()
        data = generator.parse_excel(file_content)
        teacher_mapping = generator.build_teacher_mapping(data['teacher'])
        
        # Generate timetables (PASSING ROOMS NOW)
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

# ==================== MAIN EXECUTION (RAILWAY COMPATIBLE) ====================
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    print(f"🚀 Starting server on 0.0.0.0:{port}")
    uvicorn.run(app, host="0.0.0.0", port=port)