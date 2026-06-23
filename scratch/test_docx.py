import docx
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

doc = Document()
style = doc.styles["Normal"]
style.font.name = "Arial"
style.font.size = Pt(10)

rPr = style.element.get_or_add_rPr()
rFonts = rPr.get_or_add_rFonts()
rFonts.set(qn('w:ascii'), 'Arial')
rFonts.set(qn('w:hAnsi'), 'Arial')

doc.add_paragraph("Teste de relatorio")
doc.save("test_output.docx")
print("Success")
