import os
import random
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# 1. SETUP BASE DATA ARRAYS FOR FAKE PII
first_names = ["John", "Jane", "Michael", "Emily", "David", "Sarah", "James", "Jessica", "Robert", "Karen"]
last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Miller", "Davis", "Garcia", "Rodriguez", "Wilson"]

def generate_ssn():
    return f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(1000, 9999)}"

def generate_dob():
    start_date = datetime(1970, 1, 1)
    end_date = datetime(2005, 12, 31)
    days_between = (end_date - start_date).days
    random_date = start_date + timedelta(days=random.randrange(days_between))
    return random_date.strftime("%Y-%m-%d")

def generate_phone():
    return f"({random.randint(100, 999)}) {random.randint(100, 999)}-{random.randint(1000, 9999)}"

def generate_email(name):
    domains = ["example.com", "testmail.net", "mockdata.org", "demo.co"]
    return f"{name.lower().replace(' ', '.')}@{random.choice(domains)}"


# Helper function to draw running decorations on every page
def draw_page_decorations(c, file_id, page_num, current_date):
    c.setFont("Courier", 8)
    
    # Running Header (Top of Page)
    c.drawString(50, 765, f"APEX GLOBAL SOLUTIONS INC. | AUDIT DIVISION")
    c.drawRightString(562, 765, f"RUN DATE: {current_date}")
    c.drawString(50, 755, "_" * 85)
    
    # Re-draw Tabular Content Headers
    c.setFont("Courier", 9)
    c.drawString(50, 735, f"DATA DUMP REPORT - SOURCE ID: {file_id}")
    c.drawString(50, 720, "=" * 85)
    
    headers = f"{'FULL NAME'.ljust(22)}{'SSN'.ljust(13)}{'DOB'.ljust(12)}{'PHONE'.ljust(16)}{'EMAIL'}"
    c.drawString(50, 700, headers)
    c.drawString(50, 690, "-" * 85)
    
    # Running Footer (Bottom of Page)
    c.setFont("Courier", 8)
    c.drawString(50, 40, "CONFIDENTIAL - INTERNAL DATA ENGINEERING RECORD USE ONLY")
    c.drawRightString(562, 40, f"PAGE {page_num}")
    c.drawString(50, 48, "_" * 85)


# 2. PDF GENERATION LOGIC WITH MULTI-PAGE MECHANICS
def create_variable_length_pdf(filename, file_id, target_pages):
    c = canvas.Canvas(filename, pagesize=letter)
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    # Determine records needed based on roughly 40 records fitting per page layout
    # We add variations to row count so it looks natural and hits target pages cleanly
    records_per_page = 42
    total_records = target_pages * records_per_page - random.randint(5, 15)
    
    page_num = 1
    draw_page_decorations(c, file_id, page_num, current_date)
    
    # Content starts right below column headers
    y_position = 670
    
    for i in range(total_records):
        name = f"{random.choice(first_names)} {random.choice(last_names)}"
        ssn = generate_ssn()
        dob = generate_dob()
        phone = generate_phone()
        email = generate_email(name)
        
        # --- INTENTIONAL ERROR INJECTION FOR PIPELINE TESTING ---
        if random.random() < 0.15:
            corruption_pivot = random.choice(['ssn', 'email', 'phone'])
            if corruption_pivot == 'ssn':
                ssn = ssn.replace("-", "")
            elif corruption_pivot == 'email':
                email = email.replace("@", "_at_")
            elif corruption_pivot == 'phone':
                phone = phone[:5]
        
        row_string = f"{name.ljust(22)}{ssn.ljust(13)}{dob.ljust(12)}{phone.ljust(16)}{email}"
        
        c.setFont("Courier", 9)
        c.drawString(50, y_position, row_string)
        y_position -= 15  # Spacing between lines
        
        # MONITOR PAGE OVERFLOW
        # If the next line drops into the footer area (below 60 pixels)
        if y_position < 60:
            # Finalize current page layer
            c.showPage()
            
            # Increment page count and initialize fresh page structures
            page_num += 1
            draw_page_decorations(c, file_id, page_num, current_date)
            
            # Reset text cursor back to the top of the content area for the new page
            y_position = 670
            
    c.save()


# 3. CONTROLLER LOOP WITH INCOHERENT PAGE ASSIGNMENT
if __name__ == "__main__":
    output_directory = "PDF Extraction"
    os.makedirs(output_directory, exist_ok=True)
    
    # A list of non-sequential, variable target page lengths for the 10 files
    page_targets = [3, 5, 2, 8, 4, 11, 3, 6, 12, 5]
    
    print("Beginning Generation of Multi-Page Pseudo-Tabular PDF Datasets...")
    for i in range(1, 11):
        file_id = f"REL{str(i).zfill(6)}"
        pdf_path = os.path.join(output_directory, f"{file_id}.pdf")
        
        # Get target page length from our array allocation
        assigned_pages = page_targets[i - 1]
        
        create_variable_length_pdf(pdf_path, file_id, assigned_pages)
        print(f" -> Successfully Written: {pdf_path} ({assigned_pages} Pages Target)")
        
    print(f"\nAll operations complete. 10 files saved to './{output_directory}/'")