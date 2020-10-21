# Automating Photoshop
This is a brief blog post describing my experience with automating Photoshop using Python.
I am an experienced software developer, but had never really used Photoshop before. As you can tell from my wonderful programmer art in this post ;)

Here's the [Adobe documentation](https://www.adobe.com/content/dam/acom/en/devnet/photoshop/pdfs/photoshop-cc-vbs-ref.pdf) my work was based on.

Here's a link to another [Github Project](https://github.com/loonghao/photoshop-python-api) that could also be of interest. **I have not used their code,** but it looks very promising.

The sample code and repo source code have been **tested on Python 2.7**, but should work fine on Python 3.

## Generating images
Some time ago we needed a solution to be able to quickly generate some product images using Photoshop.

The graphic designer wanted to combine 2 images into a final product image to be used for display their products online. 
- An environment image (PSD file).
- An object image (jpg).
- Combine the above images into a product image (jpg). 

This was done manually in Photoshop and as expected was very time consuming and error prone.

The scenario described for generating an image might seem very simple, **why use Photoshop at all, right?** PIL, Skimage, OpenCV would work fine! Well in this case, there were some very fine transformations and image processing being done in Photoshop and the graphic designer needed all these features required in order to generate high-quality images using filter, shears and other exotic (for me at least) visual effects.

![Combination of 2 images](https://github.com/kelvin0/ImageAutomation/blob/gh-pages/Combine_2_images.png?raw=true)

## Smart objects?
Prior to my involvement, the graphic designer had been looking for a way to simplify and automate this image generation process.
An important feature that would be key to this work is the concept of [Smart Objects](https://helpx.adobe.com/ca/photoshop/using/create-smart-objects.html)

**Smart objects** in Photoshop allow you to 'link' 2 or more PSD files. Any changes made to the linked PSD are automatically made to any PSD linking to it!
Basically you create a PSD file, and have one of the layers be a Smart Object. Then you link that Smart Object layer to another PSD.
Afterwards when you open the background image and the product image in Photoshop, **any changes you make to the product image, also are made in the background image**.

Another cool thing about Smart Objects: **all the transformations within the Smart Object layer are preserved**, regardless of the changes you make to the source PSD.

![Smart Objects workflow](https://github.com/kelvin0/ImageAutomation/blob/gh-pages/smart_objects_update.png?raw=true)

**This requires:**
- Each background image (PSD) must contain a layer with a Smart Object.
- The Smart Object layer has  to be linked to a default image (PSD).
- Works best if both PSD files reside in same directory.


**The manual steps for generating a final product image becomes:**
1. Open the background image in Photoshop (mountains).
2. Open the default product image in Photoshop (ball).
3. Open the desired product image in Photoshop (star).
4. Copy the desired product image into the default product image. This updates the Smart Object.
5. Save the background image as JPEG. This is our final image we want to generate with mountains and the star.
6. Repeat this for every background/product combination image we want to generate.

## Python and COM
As mentionned at the beginning, we will be using the Photoshop COM programming interface.
The [Photoshop reference PDF](https://www.adobe.com/content/dam/acom/en/devnet/photoshop/pdfs/photoshop-cc-vbs-ref.pdf) will be our guide in writing our automation scripts. Of course we could be doing this directly in VB script, but it is much more fun (and productive!) to use Python.

Here's a basic sample that opens an image in Photoshop.

```python
import win32com.client

# This actually fires up Photoshop if not already running.
ps = win32com.client.Dispatch("Photoshop.Application")

# Open an image file (PSD in our case)
doc = ps.Open(r"X:\Path\To\My.psd")
# ... do something ...
doc.Close()

ps.Quit() # Stops the Photoshop application
```
This works on Windows, but some other scripting language might be more appropriate for Mac OS.
We will not be covering other platforms.

**There is no headless mode when running Python/COM automation scripts.**

Each script command actually translates to an action you see happen on the screen.
I will get into this and other annoyances later. 

## Basic Recipe

Here is some basic sample code that illustrates the automated steps to generate our final product image, which is a star on a background of mountains.

Notice also that we duplicate the PSD documents once we open them. We do this in order not to accidentally change and save the original PSD files.

**Important:** working with Photoshop's object containers is different than native Python lists and tuples. The indices are **1-based**, so the first element of container has index=1 (as opposed to index=0 as per usual).

### basic_recipe.py
```python
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
```

## A step further
In order to make this a little less painful to use, we created a psd_utils.py source file.
This file contains contains the **Photoshop class** to alleviate some of the boilerplate code.

### ps_sample.py
```python
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
```

## Watch out!
As mentionned earlier, even though there are quite a few advantages to automating with Photoshop, there are also quite a few points to consider.

Photoshop scripts require running an actual instance of Photoshop and it's **main window will be visible on the desktop**.

The **Photoshop window should not be minimized while running a script.** This might actually block Photoshop, and prevent your automated task from running properly.

If you make use of **Copy/Paste commands in your script, this will hijack your clipboard**, and prevent any other user/application from using it properly.

**Photoshop tends to hang/freeze/crash periodically.** The crashes are frequent on big batches of images and don't seem to be related to RAM/CPU usage. Just restart your script and it will eventually run to completion just fine. Regardless of crashes, **you can still make huge productivity gains from automating some tasks.**

For all these reasons, we **highly recommend that any automated tasks you create should run on a dedicated Windows PC.** You don't need a high end PC for most tasks and this will definitely make everyone more productive.

## Hope this was useful
Of course, most of the code and samples discussed here are related to the specific use case described.
Almost all Photoshop commands can be scripted this way. The sample code should help you get started, and more details can be found in Photoshop scripting reference.
**If you need any help with your project, we will gladly share our expertise if required!**


