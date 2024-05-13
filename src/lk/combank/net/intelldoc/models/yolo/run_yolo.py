import sys
from pathlib import Path

wd = Path(__file__).parent.parent.resolve()
sys.path.append(str(wd))

from yolo import Yolo
     

model_path = "models/yolo/best.pt"
source_dir = "models/yolo/images"

yolo = Yolo(model_path, source_dir)
yolo.run()

