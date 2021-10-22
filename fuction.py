import random
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT,WD_LINE_SPACING
from docx.shared import Pt,RGBColor
from docx.dml.color import ColorFormat


def working(main,font,baris,alenia):
  nulis = main.add_paragraph()
  nulis.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
  if alenia  == 'tengah':
   nulis.alignment=WD_PARAGRAPH_ALIGNMENT.CENTER
  elif alenia == 'kiri':
   nulis.alignment=WD_PARAGRAPH_ALIGNMENT.LEFT
  elif alenia == 'kanan':
   nulis.alignment=WD_PARAGRAPH_ALIGNMENT.RIGHT
  elif alenia == 'rata':
   nulis.alignment=WD_PARAGRAPH_ALIGNMENT.JUSTIFY
  for huruf in baris[1].strip():
   style = nulis.add_run(huruf)
   style.font.size = Pt(22)
   style.font.name = random.choice(font)
   style.font.color.rgb = RGBColor(21,6,120)