
def grab_shape(element, shape_name):
    return element.Shapes(shape_name)

def grab_styles(shape):
    return {'width':  shape.Width,
            'height': shape.Height,
            'top':    shape.Top,
            'left':   shape.Left}

def apply_styles(shape, styles):
    shape.Top    = styles['top']
    shape.Left   = styles['left']
    shape.Height = styles['height']
    shape.Width  = styles['width']

def replace_text(label_shape, content):
    label_shape.TextFrame.TextRange.Text = content
