import os
import comtypes.client

# -----------------------------
# PROJECT PATHS
# -----------------------------

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

input_folder = os.path.join(BASE_DIR, "input")

output_folder = os.path.join(BASE_DIR, "output")

# Create output folder if missing
os.makedirs(output_folder, exist_ok=True)

# -----------------------------
# START POWERPOINT
# -----------------------------

powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

powerpoint.Visible = 1

# -----------------------------
# LOOP THROUGH ALL PPT FILES
# -----------------------------

for file_name in os.listdir(input_folder):

    # Process only PPTX files
    if file_name.endswith(".pptx"):

        ppt_path = os.path.join(input_folder, file_name)

        pdf_name = file_name.replace(".pptx", ".pdf")

        pdf_path = os.path.join(output_folder, pdf_name)

        print(f"\nConverting: {file_name}")

        # Open presentation
        presentation = powerpoint.Presentations.Open(ppt_path)

        # Save as PDF
        presentation.SaveAs(pdf_path, 32)

        # Close presentation
        presentation.Close()

        print("PDF Saved:")
        print(pdf_path)

# -----------------------------
# CLOSE POWERPOINT
# -----------------------------

powerpoint.Quit()

print("\nAll PPT files converted successfully!")