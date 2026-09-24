import fitz

# Create a clean A4 sample template
doc = fitz.open()
page = doc.new_page(width=595.27, height=841.89)  # standard A4

# Draw a decorative border
page.draw_rect(fitz.Rect(30, 30, 565, 811), color=(0.2, 0.4, 0.6), width=2)

# Insert header placeholder
page.insert_textbox(
    fitz.Rect(50, 50, 545, 90),
    "{{MONTH}}",
    fontsize=16,
    align=fitz.TEXT_ALIGN_CENTER,
)

# Insert Table Slot 1
page.draw_rect(fitz.Rect(50, 110, 545, 420), color=(0.7, 0.7, 0.7), width=1)
page.insert_textbox(
    fitz.Rect(60, 120, 300, 150),
    "{{DATE_1}}",
    fontsize=13,
    align=fitz.TEXT_ALIGN_LEFT,
)
page.insert_text(
    fitz.Point(60, 180),
    "Memorization / Revision Notes:",
    fontsize=11,
    color=(0.4, 0.4, 0.4),
)

# Insert Table Slot 2
page.draw_rect(fitz.Rect(50, 440, 545, 750), color=(0.7, 0.7, 0.7), width=1)
page.insert_textbox(
    fitz.Rect(60, 450, 300, 480),
    "{{DATE_2}}",
    fontsize=13,
    align=fitz.TEXT_ALIGN_LEFT,
)
page.insert_text(
    fitz.Point(60, 510),
    "Memorization / Revision Notes:",
    fontsize=11,
    color=(0.4, 0.4, 0.4),
)

# Insert Footer placeholder
page.insert_textbox(
    fitz.Rect(50, 770, 545, 800),
    "{{PAGE_NUM}}",
    fontsize=10,
    align=fitz.TEXT_ALIGN_CENTER,
)

doc.save("sample_template.pdf")
doc.close()
print("Saved sample_template.pdf")