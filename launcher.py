#!/usr/bin/env python3
"""MCO Maker - Application Launcher.
Bootstraps the application by delegating directly to the compiled native binary module (main.pyd)
or the Python source module (main.py).
"""

import sys
import types

# Lightweight cv2 stub to prevent ddddocr 1.6.1 from importing 118MB OpenCV
sys.modules.setdefault("cv2", types.ModuleType("cv2"))

import main

if __name__ == "__main__":
    main.start_app()
