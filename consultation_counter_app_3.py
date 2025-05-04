import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# Define parent service groupings
grouped_services = {
    "General OPD": [
        "General OPD Consultation Adult",
        "General OPD Consultation Child",
        "Review Consultation - Review",
        "Internal OPD Consultation"
    ],
    "Dressing / Minor Procedures": [
        "Dressing And Major Suturing >= 12 Yrs",
        "(GENERAL) Change of Dressing Minor",
        "(GENERAL) Change of Dressing < 12 Yrs",
        "Dressing And Minor Suturing >= 12 Yrs",
        "Dressing And Minor Suturing <12 Yrs",
        "IUD >=12Yrs",
        "Dressing And Major Suturing <12 Yrs",
        "(GENERAL) Change of Dressing Major"
    ],
    "Scans / Diagnostics": [
        "PELVIC SCAN",
        "ABDOMINOPELVIC",
        "ElectroCardlography (ECG)",
        "OBSTETRIC SCAN",
        "BREAST SCAN"
    ],
    "Maternal Health (ANC/PNC)": [
        "Antenatal Consultation",
        "Postnatal Consultation",
        "INTERNAL ANC CONSULATION"
    ],
    "Eye Clinic": [
        "Eye Consultation Adult",
        "Eye Consultation Child",
        "Review Consultation - Eye",
        "Internal Eye Consultation"
    ],
    "Mental Health": [
        "MENTAL HEALTH CONSULTATION ADULT",
        "MENTAL HEALTH CONSULTATION CHILD",
        "MENTAL HEALTH REVIEW CONSULTATION",
        "INTERNAL MENTAL HEALTH UNIT CONSULTATION"
    ],
    "ENT": [
        "ENT CONSULTATION",
        "ENT REVIEW CONSULTATION",
        "ENT INTERNAL CONSULTATION",
        "ENT SPECIALIST CONSULTATION"
    ],
    "Dental": [
        "Dental Consultation Adult",
        "Dental Consultation Child",
        "Dental Review Consultation",
        "Internal Dental Consultation"
    ],
    "Dietetic": [
        "DIETETIC CONSULTATION ADULT",
        "DIETETIC CONSULTATION CHILD",
        "DIETETIC REVIEW CONSULTATION",
        "INTERNAL DIETETIC CONSULTATION",
        "DIETETIC APPOINTMENT"
    ]
}

def process_excel(file_path):
    try:
        df = pd.read_excel(file_path, skiprows=3)
        procedure_col = 'Type Of Procedure(s) Requested'
        df[procedure_col] = df[procedure_col].astype(str).str.strip().str.upper()
        
        # Split multiple services in a single cell
        df[procedure_col] = df[procedure_col].str.split(',')
        df = df.explode(procedure_col)

        # Strip whitespace again after exploding
        df[procedure_col] = df[procedure_col].str.strip()


        result = "===== Mapped / Recognized Procedures =====\n\n"
        all_grouped = set()

        for category, services in grouped_services.items():
            count = 0
            for service in services:
                count += df[procedure_col].eq(service.upper()).sum()
                all_grouped.add(service.upper())
            if count > 0:
                result += f"📌 {category}: {count}\n"

        # Find unmapped services
        unique_services = set(df[procedure_col].unique())
        unmapped_services = unique_services - all_grouped

        if unmapped_services:
            result += "\n⚠️ Unmapped / Unrecognized Procedures:\n"
            for svc in sorted(unmapped_services):
                count = df[procedure_col].eq(svc).sum()
                result += f"  • {svc.title()}: {count}\n"

        result += "\n==========================================\n"
        return result

    except Exception as e:
        return f"❌ Error: {e}"


# GUI App
def browse_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        result = process_excel(filepath)
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, result)

# Build GUI
root = tk.Tk()
root.title("BETHEL CONSULT")
root.geometry("600x500")

browse_btn = tk.Button(root, text="Import Excel File", command=browse_file, font=("Arial", 12), bg="#4CAF50", fg="white")
browse_btn.pack(pady=10)

text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=25, font=("Consolas", 15))
text_area.pack(padx=10, pady=10)

root.mainloop()
