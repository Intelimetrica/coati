
def grab_shape(slide, shape_name):
    return slide.Shapes(shape_name)

def grab_styles(shape):
    return {'width':  shape.Width,
            'height': shape.Height,
            'top':    shape.Top,
            'left':   shape.Left,
            'lock_aspect_ratio': shape.LockAspectRatio}

def apply_styles(shape, styles):
    shape.LockAspectRatio = styles['lock_aspect_ratio']
    shape.Top    = styles['top']
    shape.Left   = styles['left']
    shape.Height = styles['height']
    shape.Width  = styles['width']

def replace_text(label_shape, content):
    label_shape.TextFrame.TextRange.Text = content

def transfer_properties(from_shape, to_shape):
    apply_styles(to_shape, grab_styles(from_shape))
    name = from_shape.Name
    from_shape.Delete()
    to_shape.Name = name

def insert_image(slide, shape_name, filename, w = 4, h = 4):
    picture = slide.Shapes.AddPicture(filename, 1, 1, w, h)
    placeholder = grab_shape(slide, shape_name)
    transfer_properties(placeholder, picture)

def flatten(x):
    if type(x) is list:
        return [a for i in x for a in flatten(i)]
    else:
        return [x]
