import os
import win32com.client

SILENT_CLOSE = 2

curdir = os.path.abspath(os.path.dirname(__file__))
background_path = os.path.join(curdir,"background.psd")
ball_base_path = os.path.join(curdir,"C base.psd")
star_path = os.path.join(curdir,"star.jpg")
gen_path = os.path.join(curdir,"final.jpg")

# This actually fires up Photoshop if not already running.
ps = win32com.client.Dispatch("Photoshop.Application")
ps.DisplayDialogs = 3            # psDisplayNoDialogs
ps.Preferences.RulerUnits = 1    # psPixels

"""1. Open the background image in Photoshop (mountains)."""
bg = ps.Open(background_path)
background = bg.Duplicate() # Work with a clone
bg.Close(SILENT_CLOSE)

"""2. Open the default product image in Photoshop (ball)."""
ball = ps.Open(ball_base_path)
ball_layer = ball.ArtLayers.Item(1)

"""3. Open the desired product image in Photoshop (star)."""
target = ps.Open(star_path)
star = target.Duplicate()
target.Close(SILENT_CLOSE)
							  
"""4. Copy the desired product image into the default product image. 
This also updates our background image."""
# Place copy of desired product image on clipboard
star_layer = star.ArtLayers.Item(1)
star_layer.Copy() 
star.Close(SILENT_CLOSE)

# Set as active image in Photoshop
ps.ActiveDocument = ball          

# Paste 'star' image from clipboard 
pasted = ball.Paste()

# We apply new image to smart object layer. 
ball.Save()                               

"""5. This is our final image we want to generate with mountains and the star."""
jpgSaveOptions = win32com.client.Dispatch( "Photoshop.JPEGSaveOptions" )
ps.ActiveDocument = background
background.SaveAs(gen_path, jpgSaveOptions, True, 2)

background.Close(SILENT_CLOSE)
ball.Close(SILENT_CLOSE)

ps.Quit() # Stops the Photoshop application