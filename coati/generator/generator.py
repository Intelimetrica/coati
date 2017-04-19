import win32com.client as w32

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
