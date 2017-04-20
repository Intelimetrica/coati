from coati.powerpoint import open_pptx, runpowerpoint
import os
import logging
import sys

colors = {'pink': '\033[95m', 'blue': '\033[94m', 'green': '\033[92m', 'yellow': '\033[93m', 'red': '\033[91m'}
logging.basicConfig(stream=sys.stdout, level=logging.INFO, format='%(asctime)s - %(message)s', datefmt='%I:%M:%S')

path = 'builders/'
template_path = 'coati/generator/slide_template.py'
config_template_path = 'coati/generator/config_template.py'

def _get_slides_shapes(ppt_path):
    pptapp = runpowerpoint()
    pptFile = open_pptx(pptapp, ppt_path)
    logging.info(colors['blue'] + 'Template opened successfully')

    all_slide_shapes = []
    for slide in pptFile.Slides:
        shapes_in_slide = _get_shapes_in_slide(slide)
        all_slide_shapes.append(shapes_in_slide)

    pptFile.close()
    pptapp.Quit()
    logging.info(colors['blue'] + 'Finished reading template')

    return all_slide_shapes

def _get_shapes_in_slide(slide):
    shapes_in_slide = {}
    for each_shape in slide.shapes:
        shapes_in_slide.update({each_shape.name: ()})
    return shapes_in_slide

def generate_path(p):
    if not os.path.exists(os.path.dirname(p)):
        try:
            os.makedirs(os.path.dirname(p))
        except OSError as exc:
            if exc.errno != errno.EEXIST:
                raise

def cp(src, dst, fn):
    source = open(src, 'r')
    result = fn(source.read())
    destination = open(dst, 'w')
    destination.write(result)
    source.close
    destination.close

def insert_code(complete_text, text_to_insert, text_to_replace):
    ans = complete_text.replace(text_to_replace, text_to_insert)
    return ans

def generate(ppt_path):
    spaces = " " * 12
    slide_tuples = '['
    config_filename = path + 'config.py'
    for i, slide in enumerate(_get_slides_shapes(ppt_path)):
        slide_name = 'slide' + str(i+1)
        slide_tuples += ('\n' + spaces if i != 0 else '') + '(' + str(i) + ', ' + slide_name + '.build()),'
        filename = path + slide_name + '.py';
        generate_path(path)
        cp(template_path, filename, lambda source: insert_code(
            source,
            str(slide).replace(", ",",\n" + spaces),
            '"_-{}-_"'))
        if i == 0:
            logging.info(colors['pink'] + 'created folder %s', path)
        logging.info(colors['green'] + 'created %s', filename)

    cp(config_template_path, config_filename, lambda source: insert_code(
        source,
        (slide_tuples[:-1] + ']'),
        '"_-{}-_"'))
    logging.info(colors['green'] + 'created %s', config_filename)
