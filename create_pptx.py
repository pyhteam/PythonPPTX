import datetime
import os
import sys
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import textwrap
import io
from pptx.enum.shapes import MSO_SHAPE


class PowerPointHelper:
    def __init__(self, show_pptx):
        self.show_pptx = show_pptx
        self.presentation = Presentation()

    def create_presentation(self):
        if not self.show_pptx["Verses"]:
            return io.BytesIO()  # Return an empty BytesIO object

        for verse in self.show_pptx["Verses"]:
            self.create_slide(verse)

        output = io.BytesIO()
        self.presentation.save(output)
        output.seek(0)
        return output

    def create_slide(self, verse):
        slide_layout = self.presentation.slide_layouts[5]
        slide = self.presentation.slides.add_slide(slide_layout)
        self.set_background(slide)
        title = f"{self.show_pptx['BookName']} {self.show_pptx['ChapterNumber']}:{verse['label']}"

        if self.show_pptx["Config"]:
            font_family = self.show_pptx["Config"].get("FontFamily", "Arial")
            font_size = Pt(self.show_pptx["Config"].get("FontSize", 40))
            if self.show_pptx["Config"].get("Color"):
                font_color = RGBColor(
                    self.show_pptx["Config"]["Color"]["R"],
                    self.show_pptx["Config"]["Color"]["G"],
                    self.show_pptx["Config"]["Color"]["B"],
                )
            else:
                font_color = RGBColor(169, 194, 26)

            text_align = getattr(
                PP_ALIGN, self.show_pptx["Config"].get("TextAlign", "CENTER").upper()
            )
        else:
            font_family = "Arial"
            font_size = Pt(40)
            font_color = RGBColor(169, 194, 26)
            text_align = PP_ALIGN.CENTER

        content = f"{verse['label']}.  {verse['content']}"
        paragraphs = self.split_text(content, font_size)

        top_inch = 1  # Starting top position for the first paragraph
        for paragraph in paragraphs:
            textbox = slide.shapes.add_textbox(
                Inches(1), Inches(top_inch), Inches(8), Inches(1)
            )
            text_frame = textbox.text_frame
            text_frame.word_wrap = True  # Enable word wrap
            p = text_frame.add_paragraph()
            p.text = paragraph
            p.font.name = font_family
            p.font.size = font_size
            p.font.color.rgb = font_color
            p.alignment = text_align
            top_inch += 0.5  # Move down for the next paragraph

        textbox = slide.shapes.add_textbox(Inches(1), Inches(6.5), Inches(8), Inches(1))
        text_frame = textbox.text_frame
        text_frame.word_wrap = True  # Enable word wrap
        p = text_frame.add_paragraph()
        line_shape = slide.shapes.add_shape(
            MSO_SHAPE.LINE_CALLOUT_1_NO_BORDER, Inches(1), Inches(6.4), Inches(8), Pt(1)
        )
        line_shape.line.color.rgb = RGBColor(0, 0, 0)  # Set line color to black

        p = text_frame.add_paragraph()
        p.text = title
        p.font.name = "Arial"
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(0, 0, 255)
        p.alignment = PP_ALIGN.CENTER

    def set_background(self, slide):
        if self.show_pptx["Config"] and self.show_pptx["Config"].get("ImagePath"):
            slide.shapes.add_picture(
                self.show_pptx["Config"]["ImagePath"],
                0,
                0,
                self.presentation.slide_width,
                self.presentation.slide_height,
            )
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
            slide.background.fill.transparency = 0.5  # Set opacity to 0.5
        else:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
            slide.background.fill.transparency = 0.5  # Set opacity to 0.5

    def split_text(self, text, font_size):
        max_chars_per_line = 60
        wrapped_text = textwrap.wrap(text, width=max_chars_per_line)
        paragraphs = []
        paragraph = ""
        for line in wrapped_text:
            if len(paragraph) + len(line) > max_chars_per_line:
                paragraphs.append(paragraph)
                paragraph = line
            else:
                paragraph += " " + line if paragraph else line
        paragraphs.append(paragraph)
        return paragraphs


if __name__ == "__main__":
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
            # [ERROR] 2021-09-29 12:00:00: An error occurred: <error message
            f.write(f"[ERROR] {datetime.datetime.now()}: An error occurred: {e}\n")
