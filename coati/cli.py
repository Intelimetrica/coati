"""Contain argument parsing logic for
the application executable."""
import argparse
import os
from tempfile import mkstemp
from coati.builder import SlideBuilder
from coati.errors import CLIException
from coati.generator import generate


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
    build_parser.add_argument('-i', '--input', default='template.pptx',
                              help='Input template to be processed')
    build_parser.add_argument('-f', '--config', default='config.py',
                              help='Config file to be used when building (slides must be defined)')

    # singleslide command
    singleslide_parser = subparsers.add_parser(
        'singleslide', help='Run coati on only one slide')
    singleslide_parser.set_defaults(func=singleslide)
    singleslide_parser.add_argument('-d', '--dir', default=os.getcwd(),
                                    help='The directory where the project lives')
    singleslide_parser.add_argument('-t', '--temporal', action='store_true',
                                    help='Output to a temporal file')
    singleslide_parser.add_argument('-o', '--output', default=os.path.abspath('out.pptx'),
                                    help='Name/path of the final built presentation')
    singleslide_parser.add_argument('-i', '--input', default='template.pptx',
                                    help='Input template to be processed')
    singleslide_parser.add_argument('-n', '--number', default='0',
                                    help='Number of slide to process')
    singleslide_parser.add_argument('-f', '--config', default='config.py',
                                    help='Config file to be used when building (slides must be defined)')

    # generate command
    generate_parser = subparsers.add_parser(
        'new', help='Scaffold a new project ')
    generate_parser.set_defaults(func=generate_boilerplate)
    generate_parser.add_argument('project_name', metavar='N', type=str,
                        help='A string indicating the new project name')
    generate_parser.add_argument('-p','--path', type=str, help='path of pptx file')

    return parser


def generate_boilerplate(args):
    generate(args.project_name, args.path)


def destinationpath(args):
    if args.temporal:
        fs, path = mkstemp()
        return path
    else:
        return os.path.abspath(args.output)


def singleslide(args):

    builder = SlideBuilder(args.dir, args.input, args.output)
    builder.loadtemplate()
    builder.loadconfig(args.config)
    resources = builder.loadresources(slide=int(args.number))
    #builder.build(resources)
    builder.buildSingleSlide(resources)
    builder.finish(close=False)

def test(args):
    pass

def build(args):
    builder = SlideBuilder(args.dir, args.input, args.output)
    builder.loadtemplate()
    builder.loadconfig(args.config)
    #resources = builder.loadresources()     
    #builder.build(resources)    
    builder.build()    
    builder.finish(close=args.close)


