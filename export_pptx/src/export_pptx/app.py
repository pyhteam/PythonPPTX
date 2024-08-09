"""
Library for creating PowerPoint presentations
"""


import datetime
import json
import os
import sys
from .power_point_helper import PowerPointHelper


def main():
    print("Creating PowerPoint presentation...")
    try:
        # Read input JSON data from stdin
        print("Reading input data...")
        input_data = sys.stdin.read()
        if not input_data:
            print("No input provided")
        else:
            # Proceed with the rest of your code
            pass
        print("Input data: ", input_data)
        print("Input data read successfully.")
        show_pptx = json.loads(input_data)
        print("Converted JSON data to Python object.")

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
        # log error to file for debugging
        with open("error_py.log", "a") as f:
            f.write(f"[ERROR] {datetime.datetime.now()}: An error occurred: {e}\n")
        # exit application with error code
    sys.exit(1)
