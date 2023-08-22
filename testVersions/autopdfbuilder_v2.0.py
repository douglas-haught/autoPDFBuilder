import os
import comtypes.client
from fpdf import FPDF
import datetime

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
}

def replace_placeholder(slide, placeholder, replacement):
    shapes = slide.Shapes
    for shape in shapes:
        if shape.HasTextFrame:
            text_frame = shape.TextFrame
            if text_frame.HasText and placeholder in text_frame.TextRange.Text:
                text_frame.TextRange.Text = text_frame.TextRange.Text.replace(placeholder, replacement)

def get_filename(base_folder, user_name):
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    return os.path.join(base_folder, f"{user_name}_{timestamp}.pdf")

def convert_ppt_to_pdf(ppt_file, output_folder, selected_sections, user_name):
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = True

        presentation = powerpoint.Presentations.Open(ppt_file)

        first_slide = presentation.Slides(1)
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

        pdf_output_path = get_filename(output_folder, user_name)
        pdf.output(pdf_output_path)

        presentation.Close()
        powerpoint.Quit()

        print(f"PDF presentation saved to: {pdf_output_path}")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    ppt_file = input(r"C:\Users\Douglas Haught\Desktop\General\PPT-PDF\Reference\advisorOnboardingMaster.pptx").strip()
    output_folder = input(r"C:\Users\Douglas Haught\Desktop\General\PPT-PDF\Output").strip()
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    user_name = input("Enter your name: ")

    print("Available sections:")
    for index, section_name in enumerate(sections, 1):
        print(f"{index}. {section_name}")

    selected_sections = input("Enter section numbers separated by commas (e.g., 1,2): ")
    selected_sections = [list(sections.keys())[int(section.strip())-1] for section in selected_sections.split(',')]

    convert_ppt_to_pdf(ppt_file, output_folder, selected_sections, user_name)
