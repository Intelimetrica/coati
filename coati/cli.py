"""Contain argument parsing logic for
the application executable."""
import argparse
import os
from tempfile import mkstemp
from coati.powerpoint import SlideshowJoiner, SlideSourceOrdering
from coati.builder import SlideBuilder
from coati.errors import CLIException


def createparser():
    parser = argparse.ArgumentParser()

    # General options
    #parser.add_argument('-e', '--envfile')

    subparsers = parser.add_subparsers(help='Specify program action')

    # build command
    build_parser = subparsers.add_parser(
        'build', help='Create a presentation from project')
    build_parser.set_defaults(func=build)
    build_parser.add_argument('-d', '--dir', default=os.getcwd(),
                              help='The directory where the project lives')
    build_parser.add_argument('-t', '--temporal', action='store_true',
                              help='Output to a temporal file')
    build_parser.add_argument('-o', '--output', default=os.path.abspath('out.pptx'),
                              help='Name/path of the final built presentation')
    build_parser.add_argument('-c', '--close', action='store_true',
                              help='Close powerpoint after creation')

    # test command
    test_parser = subparsers.add_parser(
        'test', help='Test output of data functions without slides')
    test_parser.set_defaults(func=test)
    test_parser.add_argument('-d', '--dir', default=os.getcwd(),
                             help='The directory where the project lives')
    test_parser.add_argument('target',
                             help='The name of the slide to be tested')

    # generate command
    generate_parser = subparsers.add_parser(
        'scaffold', help='Scaffold a project')

    return parser


def slidepaths(path):
    directory = os.path.abspath(path)
    ordering = SlideSourceOrdering()
    return ordering(os.path.join(directory, slideshow)
                    for slideshow in os.listdir(directory))


def destinationpath(args):
    if args.temporal:
        fs, path = mkstemp()
        return path
    else:
        return os.path.abspath(args.output)


def singleslide(args):
    path = os.path.abspath(args.dir)
    fullpath = os.path.join(path, args.single)

    if not os.path.isdir(fullpath):
        raise CLIException("'%s' is not present in project" % args.single)

    builder = SlideBuilder(fullpath)
    builder.build()
    close_excel(builder)

    if args.close:
        builder.template.clean()
        builder.template.app.Quit()


def build(args):
    builder = SlideBuilder(args.template)
    builder.loadtemplate()
    builder.loadconfig('./config.py')
    builder.build()


def test(args):
    path = os.path.join(os.path.abspath(args.dir),
                        args.target)
    builder = SlideBuilder(path)
    builder.fillexcel()


def close_excel(builder):
    if hasattr(builder, 'excelapp'):
        builder.excelapp.DisplayAlerts = False
        builder.excelapp.Quit()
