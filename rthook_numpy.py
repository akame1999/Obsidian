# rthook_numpy.py
# PyInstaller runtime hook to fix numpy 1.x/2.x import issues in frozen apps

import sys
import os

# If we're running frozen (inside PyInstaller bundle)
if getattr(sys, 'frozen', False):
    base = sys._MEIPASS
    
    # NumPy looks for these paths at import time
    numpy_paths = [
        os.path.join(base, 'numpy'),
        os.path.join(base, 'numpy', '.libs'),
        os.path.join(base, 'numpy.libs'),
    ]
    
    for p in numpy_paths:
        if os.path.isdir(p) and p not in sys.path:
            # Append (not insert) so stdlib modules are found first
            # This prevents numpy.random from shadowing stdlib random
            sys.path.append(p)
