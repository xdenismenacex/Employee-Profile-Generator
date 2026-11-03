!pip install -q python-pptx

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

people = [
    {"id": "ENTER VALUE", "name": "ENTER VALUE", "title": "ENTER VALUE", "dept": "ENTER VALUE", "start": "ENTER VALUE"},
    {"id": "ENTER VALUE", "name": "ENTER VALUE", "title": "ENTER VALUE", "dept": "ENTER VALUE", "start": "ENTER VALUE"},
    {"id": "ENTER VALUE", "name": "ENTER VALUE", "title": "ENTER VALUE", "dept": "ENTER VALUE", "start": "ENTER VALUE"},
    {"id": "ENTER VALUE", "name": "ENTER VALUE", "title": "ENTER VALUE", "dept": "ENTER VALUE", "start": "ENTER VALUE"},
    {"id": "ENTER VALUE", "name": "ENTER VALUE", "title": "ENTER VALUE", "dept": "ENTER VALUE", "start": "ENTER VALUE"},
    {"id": "ENTER VALUE", "name": "ENTER VALUE", "title": "ENTER VALUE", "dept": "ENTER VALUE", "start": "ENTER VALUE"},
]

blue = RGBColor(0x1F, 0x4E, 0x9C)
light = RGBColor(0xE3, 0xE8, 0xF0)
white = RGBColor(0xFF, 0xFF, 0xFF)
grey = RGBColor(0x66, 0x66, 0x66)
black = RGBColor(0x22, 0x22, 0x22)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

for p in people:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # white card in the middle
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.2), Inches(0.6), Inches(4.9), Inches(6.3))
    card.fill.solid()
    card.fill.fore_color.rgb = white
    card.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)

    # blue bar on top
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.2), Inches(0.6), Inches(4.9), Inches(0.8))
    bar.fill.solid()
    bar.fill.fore_color.rgb = blue
    bar.line.fill.background()
    bar.text_frame.text = "EMPLOYEE PROFILE"
    bar.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    run = bar.text_frame.paragraphs[0].runs[0]
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = white

    # circle with initials
    initials = ""
    for word in p["name"].split():
        initials = initials + word[0]
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.0), Inches(1.7), Inches(1.3), Inches(1.3))
    circle.fill.solid()
    circle.fill.fore_color.rgb = light
    circle.line.fill.background()
    circle.text_frame.text = initials
    run = circle.text_frame.paragraphs[0].runs[0]
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = blue

    # the text lines: (text, top position, size, bold, color)
    lines = [
        (p["name"], 3.2, 26, True, black),
        (p["title"], 3.9, 18, False, grey),
        ("ID: " + p["id"], 4.9, 16, False, black),
        ("Department: " + p["dept"], 5.5, 16, False, black),
        ("Start date: " + p["start"], 6.1, 16, False, black),
    ]
    for text, top, size, bold, color in lines:
        box = slide.shapes.add_textbox(Inches(4.2), Inches(top), Inches(4.9), Inches(0.6))
        paragraph = box.text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.add_run()
        run.text = text
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color

prs.save("profiles.pptx")

from google.colab import files
files.download("profiles.pptx")
