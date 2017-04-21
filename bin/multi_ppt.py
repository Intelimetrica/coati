#!/usr/bin/env python
from __future__ import print_function
import os
import sys
from tempfile import mkstemp
from glob import glob
from coati.settings import load
from coati.powerpoint import SlideshowJoiner, SlideSourceOrdering
from coati.builder import SlideBuilder

def slide_paths():
    directory = os.path.abspath('sources')
    ordering = SlideSourceOrdering()
    return ordering(os.path.join(directory, slideshow)
                    for slideshow in os.listdir(directory))

def destination_path():
    if len(sys.argv) > 1:
        return os.path.join(os.getcwd(), sys.argv[1])
    else:
        # If no destination file is passed, use a tempfile
        fd, path = mkstemp()
        return path

def main():
    load()
    
    paths = slide_paths()
    builders = [SlideBuilder(path) for path in paths]

    # Join slides
    joiner = SlideshowJoiner(builders)
    joiner.build()
    joiner.save(destination_path())
    
    builders[0].excelapp.DisplayAlerts = False
    builders[0].excelapp.Quit()

    
if __name__ == '__main__':
    main()
