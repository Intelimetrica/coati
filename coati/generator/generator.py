import win32com.client as w32
import os

path = 'builders/'
src = 'generator/slide_template.py'

def _get_slides_shapes(ppt_path):
    pptApp = w32.Dispatch('PowerPoint.Application')
    pptFile = pptApp.Presentations.Open(ppt_path)

    all_slide_shapes = []
    for slide in pptFile.Slides:
        shapes_in_slide = _get_shapes_in_slide(slide)
        all_slide_shapes.append(shapes_in_slide)

    pptFile.close()
    pptApp.Quit()

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

for i, slide in enumerate(_get_slides_shapes):
    filename = path + 'slide' + str(i+1) + '.py';
    spaces = " " * 12
    generate_path(path)
    cp(src, filename, lambda source: insert_code(
        source,
        str(slide).replace(", ",",\n" + spaces),
        '"_-{}-_"'))

