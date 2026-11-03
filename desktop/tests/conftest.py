import os
import sys

# Ensure desktop/ is on sys.path for test discovery
desktop_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if desktop_dir not in sys.path:
    sys.path.insert(0, desktop_dir)
