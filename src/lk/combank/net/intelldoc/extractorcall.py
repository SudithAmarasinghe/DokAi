import sys
from pathlib import Path
from extractor import ExtractorGUI

wd = Path(__file__).parent.parent.resolve()
sys.path.append(str(wd))

ext = ExtractorGUI()
ext.extract()