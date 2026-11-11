#!/usr/bin/env python3
"""MCO Maker - Application Launcher.
Bootstraps the application by delegating directly to the compiled native binary module (main.pyd)
or the Python source module (main.py).
"""

import sys
import main

if __name__ == "__main__":
    main.start_app()
