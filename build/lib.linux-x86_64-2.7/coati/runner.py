#!C:\Python27\python.exe

from __future__ import print_function
import sys
import traceback
from coati import cli
from coati.errors import CoatiException


def main():
    try:
        parser = cli.createparser()
        args = parser.parse_args()
        args.func(args)
    except CoatiException as e:
        print(str(e), file=sys.stderr)
    except Exception as e:
        print("An unknown error occurred:", file=sys.stderr)
        print(str(e), file=sys.stderr)
        print(traceback.format_exc(), file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
