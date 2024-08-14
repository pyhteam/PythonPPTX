import datetime
import os
import sys
from powerpoint_helper import PowerPointHelper


def export_pptx(show_pptx):
    try:
        helper = PowerPointHelper(show_pptx)
        print("Created PowerPoint presentation.")
        output = helper.create_presentation()
        print("Presentation created successfully.")
        with open(show_pptx["FilePath"], "wb") as f:
            f.write(output.getbuffer())
        print(f"Presentation saved to {show_pptx['FilePath']}")
        os.startfile(show_pptx["FilePath"])
        print("Presentation opened successfully.")
    except Exception as e:
        print(f"An error occurred: {e}", file=sys.stderr)
        with open("error_py.log", "a") as f:
            f.write(f"[ERROR] {datetime.datetime.now()}: An error occurred: {e}\n")
