"""KAAVAL — Knowledge & Advanced Analysis for Vigilance & Asset Loss
Tamil Nadu Police EOW | Single-click entry point
"""
import sys
import os

# Ensure the kaaval package directory is on the path
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, os.path.join(base_dir, "kaaval"))
os.chdir(os.path.join(base_dir, "kaaval"))

from app import main

if __name__ == "__main__":
    main()
