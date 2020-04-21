''' .==========Books Sizes==========.
    [ Novel    =   480     x 767.24 ]
    [ Standard =   498.89  x 744.56 ]
    [ Demy     =   521.57  x 816.37 ]
    [ US Royal =   574.48  x 865.51 ]
    [ A4       =   793.70  x 1122.51]
    [ A5       =   559.37  x 793.70 ]
    [ Square   =   793.70  x 793.70 ]
    .===============================.
                                    '''

''' .====Cut, Sizes====.
    [  5mm = 28.38 px  ]
    [  10mm= 56.76 px  ]
    .==================.
                        '''

from comtypes.client import GetActiveObject , CreateObject

# SIZES LIBRARY



cut    = 37.8
spines = 75.2
width  = 559.37
height = 793.70


psReplaceSelection = 1

fcoverPATH  = r"PATH FRONT.jpg"
bcoverPATH  = r"PATH BACK.jpg"
spcoverPATH = r"PATH SPINE.jpg"
try:
    app = GetActiveObject("Photoshop.Application")
    # RulerUnits=1 Pixel.
    app.Preferences.RulerUnits = 1

except Exception as err:
    print("You must run photoshopfirst.")

    
# Dimension en pixel [A5]

#Define the color- in this case, flat normal blue
green_color = CreateObject('Photoshop.SolidColor')
green_color.rgb.red = 88
green_color.rgb.green = 56
green_color.rgb.blue = 123

print("front_back color selected")

#Define the color red- in this case, flat normal blue
red_color = CreateObject('Photoshop.SolidColor')
red_color.rgb.red = 66
red_color.rgb.green = 34
red_color.rgb.blue = 106
print("spines color selected")

''' .===Function Details===.
    [     CREATE LAYER     ]
    [    FILL LAYER AREA   ]
    [   PASTE IMAGE LAYER  ]
    [    TRANSFORM IMAGE   ]
    .======================.
                            '''

def bookCoverFront():
    ''' SET THE FRONT COVER AREA OF BOOK ''' 

    # OPEN IMAGE AS LAYER AND SELECT SET SIZES.
    try:
        frontcover=app.Open(fcoverPATH )
    except Exception as err:
        print("No Front Cover Found.")
    frontcover.ResizeImage(width, height , 72)
    frontcover.Selection.SelectAll()
    frontcover.Selection.Copy()
    frontcover.Close(2)
    
    front_area = ((cut/2,cut/2), (cut/2, height+cut/2), (width+cut/2, height+cut/2), (width+cut/2, cut/2))
    book.Selection.Select(front_area)
    print("front area sized")
    
    print("fron cover layer Created")
    frontcover = book.ArtLayers.Add()
    frontcover.name = 'FrontBook'
    
    book.Selection.Select(front_area)
    app.activeDocument.selection.Fill(green_color)
    print("front-area colored.")

    book.Paste()

def bookCoverBack():

    try:
        back_cover=app.Open( bcoverPATH )
    except Exception as err:
        print("No back image Found.")

    back_cover.ResizeImage(width, height , 72)
    back_cover.Selection.SelectAll()
    back_cover.Selection.Copy()
    back_cover.Close(2)

    back_area = (((width*2)+spines+cut/2,cut/2), ((width*2)+spines+cut/2, height+cut/2), (width+spines+cut/2, height+cut/2), (width+spines+cut/2, cut/2))
    book.Selection.Select(back_area)
    print("back area sized")

    print("back cover layer Created")
    backcover = book.ArtLayers.Add()
    backcover.name = 'BackBook'

    book.Selection.Select(back_area)
    app.activeDocument.selection.Fill(green_color)
    print("back_area colored.")

    book.Paste()

def bookSpines():
    
    try:
        spinescover=app.Open( spcoverPATH )
    except Exception as err:
        print("No back image Found.")

    spinescover.ResizeImage(spines, height , 72)
    spinescover.Selection.SelectAll()
    spinescover.Selection.Copy()
    spinescover.Close(2)
    
    spines_area = ((width+cut/2,cut/2), (width+cut/2, height+cut/2), (width+spines+cut/2, height+cut/2), (width+spines+cut/2, cut/2))
    book.Selection.Select(spines_area)
    print("Spines area sized")

    print("spine layer Created")
    spinecover = book.ArtLayers.Add()
    spinecover.name = 'SpineBook'

    book.Selection.Select(spines_area)
    app.activeDocument.selection.Fill(red_color)
    print("spines_area colored")
    book.Paste()

# CALLFUNCTIONS.
'''
wid : width
hei : height
np  : number of page
fc  : front cover
bc  : back cover
spc : spine cover
'''

book=app.documents.add((width*2)+spines+cut, height+cut, 72, "BookDimension" ,2,1,1)
bookCoverFront()
bookCoverBack()
bookSpines()
