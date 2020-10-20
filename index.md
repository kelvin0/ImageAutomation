# Automating Photoshop
This is a brief blog post describing my experience with automating Photoshop using Python.
I am an experienced software developer, but had never really used Photoshop before.

Here's the [Adobe documentation](https://www.adobe.com/content/dam/acom/en/devnet/photoshop/pdfs/photoshop-cc-vbs-ref.pdf) my work was based on.

Here's a link to another [Github Project](https://github.com/loonghao/photoshop-python-api) that could also be of interest. I have not used their code, but it looks very promising.

## Generating images
Some time ago we needed a solution to be able to quickly generate some product images using Photoshop.

The graphic designer wanted to combine 2 images into a final product image to be used for display their products online. 
- An environment image (PSD file).
- An object image (jpg).
- Combine the above images into a product image (jpg). 

This was done manually in Photoshop and as expected was very time consuming and error prone.

![Combination of 2 images](https://github.com/kelvin0/ImageAutomation/blob/gh-pages/Combine_2_images.png?raw=true)

## Smart objects?
Prior to my involvement, the graphic designer had been looking for a way to simplify and automate this image generation process.
An important feature that would be key to this work is the concept of [Smart Objects](https://helpx.adobe.com/ca/photoshop/using/create-smart-objects.html)

**Smart objects** in Photoshop allow you to 'link' 2 or more PSD files. Any changes made to the linked PSD are automatically made to any PSD linking to it!
Basically you create a PSD file, and have one the layers be a Smart Object. Then you link that Smart Object layer to another PSD.
In Photoshop when you open the background image and the product image, any changes you make to the product image instantly are also made to the background image.

Another cool thing about Smart Objects: all the transformations are in the Smart Object layer are preserved, regardless of the changes you make to the source PSD.

![Smart Objects workflow](https://github.com/kelvin0/ImageAutomation/blob/gh-pages/smart_objects_update.png?raw=true)

**This requires:**
- Each background image (PSD) contain a Smart Object.
- The Smart Object is linked to a default image (PSD).
- Both PSD need to reside in same directory.


**The manual steps for generating a final product image becomes:**
1. Open the background image in Photoshop (mountains).
2. Open the default product image in Photoshop (ball).
3. Open the desired product image in Photoshop (star).
4. Copy the desired product image into the default product image. This updates the Smart Object.
5. Save the background image. This is our final image we want to generate with mountains and the star.

## Python and COM

## Basic Recipe

## Running the script

## Watch out!

## Hope this was useful


