# -*- coding: utf-8 -*-
from __future__ import print_function
import os
import sys
from officereports.builder import SlideBuilder


def main():
    if len(sys.argv) < 2:
        print('You need to specify a slide directory')
        sys.exit(1)

    # Load resource paths
    src_path = os.path.abspath('sources')
    builder = SlideBuilder(os.path.join(src_path, sys.argv[1]))

    builder.build()


if __name__ == '__main__':
    main()
