# Version 1.2. See bottom for version notes
import os
import comtypes.client
from fpdf import FPDF

# Define named sections and their corresponding slide ranges
sections = {
    "Introduction": "1-23",
    "WPD": "24-34",
    "AMD": "35-41",
    "Compliance": "42-44",
    "Marketing": "45-49",
    "Opportunities": "50-53",
    "Important Facts": "54-61",
    "Our Capital Partners": "62-65",
    "Vision 2030": "66-70",
    "ALL": "1-70"
    # Add more sections as needed
}

def replace_placeholder(slide, placeholder, replacement):
    shapes = slide.Shapes
    for shape in shapes:
        if shape.HasTextFrame:
            text_frame = shape.TextFrame
            if text_frame.HasText and placeholder in text_frame.TextRange.Text:
                text_frame.TextRange.Text = text_frame.TextRange.Text.replace(placeholder, replacement)

def convert_ppt_to_pdf(ppt_file, output_folder, selected_sections, user_name):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = True

    presentation = powerpoint.Presentations.Open(ppt_file)

    # Get the first slide
    first_slide = presentation.Slides(1)

    # Replace placeholder on the first slide with user input
    placeholder = "[INSERT ADVISOR NAME]"
    replace_placeholder(first_slide, placeholder, user_name)

    slide_count = presentation.Slides.Count

    pdf = FPDF(orientation='L')

    for section_name in selected_sections:
        if section_name not in sections:
            print(f"Invalid section: {section_name}")
            continue

        slide_range = sections[section_name]
        start_slide, end_slide = map(int, slide_range.split('-'))

        if start_slide < 1 or end_slide > slide_count or start_slide > end_slide:
            print(f"Invalid slide range for section {section_name}: {slide_range}")
            continue

        for slide_number in range(start_slide, end_slide + 1):
            slide = presentation.Slides(slide_number)
            image_path = os.path.join(output_folder, f"slide_{slide_number}.png")

            slide.Export(image_path, "PNG", 1024, 768)

            pdf.add_page(orientation='L')
            pdf.image(image_path, 10, 10, 277, 190)

            os.remove(image_path)

    pdf_output_path = os.path.join(output_folder, "output.pdf")
    pdf.output(pdf_output_path)

    presentation.Close()
    powerpoint.Quit()

    print(f"PDF presentation saved to: {pdf_output_path}")

if __name__ == "__main__":
    ppt_file = r"C:\Users\Douglas Haught\Desktop\Python\AutoPDFBuilder\reference\advisorOnboardingMaster.pptx"
    output_folder = r"C:\Users\Douglas Haught\Desktop\Python\AutoPDFBuilder\output"

    user_name = input("Enter your name: ")

    print("Available sections:")
    for section_name in sections:
        print(section_name)

    selected_sections = input("Enter section(s) separated by commas (e.g., Introduction,Research): ")
    selected_sections = [section.strip() for section in selected_sections.split(',')]

    convert_ppt_to_pdf(ppt_file, output_folder, selected_sections, user_name)


# Version 1.2 Release notes:
#only works locally. This will be used as the base code with plans to eventually build a GUI to accept user input to feed data to the above program. To be hosted on Raspberry Pi
#Orginal use case is for Advisor Onboarding Presentation built over 2022-2023
