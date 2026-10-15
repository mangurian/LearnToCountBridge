"""Double-click launcher for Bridge Tracker v1.7 (no console window).

Run this file, or bridge_tracker.py, whichever you prefer.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import bridge_tracker

if __name__ == "__main__":
    sys.exit(bridge_tracker.main())
