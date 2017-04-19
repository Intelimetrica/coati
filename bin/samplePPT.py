import win32com.client, sys
Application = win32com.client.Dispatch("PowerPoint.Application")
Application.Visible = True
Presentation = Application.Presentations.Open(sys.argv[1])
for slide in Presentation.Slides:
    for shape in slide.Shapes:
        try:
            shape.TextFrame.TextRange.Font.Name = "Arial"
        except:
            print("this shape is not a text")

Presentation.Save()
Application.Quit()
