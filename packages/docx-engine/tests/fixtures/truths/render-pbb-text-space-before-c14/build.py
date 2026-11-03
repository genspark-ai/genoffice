import os, runpy, sys
# the compat-15 sibling's document under compatibilityMode 14
sys.argv = [sys.argv[0], sys.argv[1] if len(sys.argv) > 1 else 'probe.docx', '', '14']
runpy.run_path(os.path.join(os.path.dirname(os.path.abspath(__file__)), '../render-pbb-text-space-before/build.py'))
