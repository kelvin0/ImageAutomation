import os
from psdbase_utils import Photoshop

curdir = os.path.abspath(os.path.dirname(__file__))
background_path = os.path.join(curdir,"background.psd")
star_path = os.path.join(curdir,"star.jpg")

ps = Photoshop()

all_open_psd =\
	ps.compose(	background_path,
				star_path,
				"C base",
				curdir,
				"final.jpg")
				
for open_psd in all_open_psd:
	ps.close(open_psd)

ps.shutdown()