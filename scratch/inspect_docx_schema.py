import sys
sys.path.insert(0, ".")
import inspect
from docx.oxml.text import font

print(inspect.getsource(font.CT_RPr))
