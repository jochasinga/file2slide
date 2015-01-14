file2slide
==========

file2slide is a simple script written in Python to help ease my wife with her arcane Powerpoint work. I hope it 
will be of some use for anyone doing trivial, laborious insertion of image and PDF files into Powerpoint who doesn't want to write his/her own script.

Dependencies
------------ 

+ Python >= 2.7
+ [python-pptx >= 0.5.6](https://python-pptx.readthedocs.org/en/latest/index.html)
+ [Imagemagick](http://www.imagemagick.org/) (For converting PDF to image)
+ [Wand](http://docs.wand-py.org/en/latest/wand/image.html) (Imagemagick Python binder)

Install
-------

Make sure you have **Python** installed (which you should already), but check anyway if the version isn't too outdated with `$ python --version` on your command line.

Install **python-pptx** and **Wand** with `pip`

    $ pip install python-pptx
    $ pip install Wand

or install the easy way using requirements.txt file inside the directory.

    $ pip install -r requirements.txt
    
To install **Imagemagick**, either install using a [binary for your OS](http://www.imagemagick.org/script/binary-releases.php) or use your package manager like [Homebrew](http://brew.sh/) or [Macport](https://www.macports.org/).

    $ brew install imagemagick
    $ sudo port install ImageMagick
    
How to use
----------

Set up your image directory so it looks like this:

    images/
    ├── image1.jpg
    ├── image2.gif
    ├── image3.png
    ├── image4.pdf
    ├── ...
    └── crop/
        ├── image1.png
        ├── image2.pdf
        └── ...

The directory and the image names can be anything you like. The only requirement is the subdirectory where you place images you want to crop must be named *crop*.
	
Just run the script with Python and the prompt will pretty much lead you through the options.

    $ python file2slide.py
    
**Here is what the script basically does**

+ Takes a directory where all your image and PDF files are stored.
+ Creates a presentation file (.pptx) based on it.
+ Looks for a subdirectory named *crop*
+ Reads the files and convert PDFs to images accordingly. 
+ Lets you customize the margins for every image relative to the slide.
+ Lets you customize the cropping for every image in *crop/*.
+ Creates each slide for every image and save the file as .pptx.
+ Takes a directory where you want to save the presentation, or just
  save to the image directory as a default if only a filename is given.






    
    
    
    

