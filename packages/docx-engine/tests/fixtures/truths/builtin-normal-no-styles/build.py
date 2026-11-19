import sys, os; sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from _docx import *
body = para('Forty characters of body text in the line.') + para('Second paragraph')
write(out_path(), body, None)
