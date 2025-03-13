from pptx import Presentation

def create_pptx(input_file, output_file):
    # Create Presentation
    prs = Presentation()

    # Read input text
    with open(input_file, 'r', encoding='utf-8') as file:
        content = file.read().strip().split("\n\n")

    # Process each slide
    for block in content:
        lines = block.strip().split("\n")
        if lines:
            slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title + Content layout
            title, body = lines[0], "\n".join(lines[1:])
            slide.shapes.title.text = title
            slide.placeholders[1].text = body

    # Save PPTX
    prs.save(output_file)
    print(f"PPTX file created successfully: {output_file}")

# Example Usage
create_pptx("input.txt", "GeneratedPresentation.pptx")
