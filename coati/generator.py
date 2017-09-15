from coati.powerpoint import open_pptx, runpowerpoint
import os
import sys
import logging
from shutil import copyfile
from colorlog import ColoredFormatter

LOG_LEVEL = logging.DEBUG
LOGFORMAT = "%(asctime)s - %(log_color)s%(message)s"
logging.root.setLevel(LOG_LEVEL)
formatter = ColoredFormatter(LOGFORMAT)
stream = logging.StreamHandler()
stream.setLevel(LOG_LEVEL)
stream.setFormatter(formatter)
log = logging.getLogger('pythonConfig')
log.setLevel(LOG_LEVEL)
log.addHandler(stream)

this_dir = os.path.dirname(__file__)
template_path = os.path.join(this_dir, 'templates/slide_template.txt')
config_template_path = os.path.join(this_dir, 'templates/config_template.txt')
init_template_path = os.path.join(this_dir, 'templates/init_template.txt')

def _get_slides_shapes(ppt_path):
    pptapp = runpowerpoint()
    pptFile = open_pptx(pptapp, ppt_path)
    log.debug('Open Template successfully...')

    all_slide_shapes = []
    for slide in pptFile.Slides:
        shapes_in_slide = _get_shapes_in_slide(slide)
        all_slide_shapes.append(shapes_in_slide)

    pptFile.close()
    pptapp.Quit()
    log.debug('Finish reading template...')

    return all_slide_shapes

def _get_shapes_in_slide(slide):
    shapes_in_slide = {each_shape.name: () for each_shape in slide.shapes}
    return shapes_in_slide

def _generate_path(p):
    if not os.path.exists(os.path.dirname(p)):
        try:
            os.makedirs(os.path.dirname(p))
        except OSError as exc:
            if exc.errno != errno.EEXIST:
                raise

def _cp(src, dst, fn):
    source = open(src, 'r')
    result = fn(source.read())
    destination = open(dst, 'w')
    destination.write(result)
    source.close
    destination.close

def _insert_code(complete_text, text_to_insert, text_to_replace):
    ans = complete_text.replace(text_to_replace, text_to_insert)
    return ans

def _file_exists(ppt_path):
    if not (ppt_path.endswith('.pptx') or ppt_path.endswith('.ppt')):
        sys.exit('The file provided is not a PPT file')
    elif not os.path.isfile(ppt_path):
        sys.exit('The PPT file provided doesnt exist or is damaged')
    pass

def generate(project_name, ppt_path):
    _file_exists(ppt_path)
    path = os.path.abspath(project_name)

    spaces = " " * 12
    slide_tuples = '['

    #Generate builders/ folder prior slides creation
    path_builders = os.path.join(path, 'builders/')
    _generate_path(path_builders)
    log.info('create folder %s', "./builders/")
    for i, slide in enumerate(_get_slides_shapes(ppt_path)):
        slide_name = 'slide' + str(i+1)
        filename = path_builders + slide_name + '.py';

        #Create slide#.py with the template info
        _cp(template_path, filename, lambda source: _insert_code(
            source,
            str(slide).replace(", ",",\n" + spaces),
            '"_-{}-_"'))
        log.info('create %s', filename)
        #This line is in the for loop cause is gathering info for the config.py
        slide_tuples += ('\n' + spaces if i != 0 else '') + '(' + str(i+1) + ', ' + slide_name + '.build()),'

    #Generate config.py with already gathered info in slide_tuples
    config_filename = path + '/config.py'
    _cp(config_template_path, config_filename, lambda source: _insert_code(
        source,
        (slide_tuples[:-1] + ']'),
        '"_-{}-_"'))
    log.info('create %s', config_filename)

    #Create __init__ in builders
    init_file = path + '/builders/__init__.py'
    copyfile(init_template_path, init_file)
    log.info('create %s', init_file)

    #Copy original template file
    copy_ppt = path + '/' + str(os.path.split(ppt_path)[-1])
    _cp(ppt_path, copy_ppt  , lambda source: source)
    log.info('copy %s', copy_ppt)

    #Add images folder
    _generate_path(os.path.join(path, 'images/'))
    log.info('create folder %s', "./images/")
