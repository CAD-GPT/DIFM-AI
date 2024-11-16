#Author-
#Description-
from .Modules import xlrd as xlr
import adsk.core, adsk.fusion, adsk.cam, traceback


#ws = xw.Book("chair_AI_01.xlsx").sheets['Sheet1']
#workbook = xlr.open_workbook(r"chair_AI_001.xls")
workbook = xlr.open_workbook('D:\\PhD Review 3 & 4\\PhD Review 3 & 4\\Fusion_API_Testing\\NewScript1\\ai_chair_new-generated-demo_XLS.xls')
sheet = workbook.sheet_by_name("Sheet1")
row_count = sheet.nrows
col_count = sheet.ncols
# use sheet.row_len() to get the effective column length when you set ragged_rows = True

#cell = sheet.cell(0, 3).value


# Selecting data from
# a single cell
# v2 = ws.range("F5").value


profile_list = [] #global variable for storing sketch profiles
extrudes_list = [] #global variable for storing extrudes
bodies_list = []

def drawCircleXY(rootComp,x,y,r):

    sketches = rootComp.sketches   
    sketch = sketches.add(rootComp.xYConstructionPlane)
    sketchCircles = sketch.sketchCurves.sketchCircles
    centerPoint = adsk.core.Point3D.create(x, y, 0)
    circle = sketchCircles.addByCenterRadius(centerPoint, r)
    return sketch.profiles.item(0)

def drawCircleYZ(rootComp,x,y,r):

    sketches = rootComp.sketches   
    sketch = sketches.add(rootComp.yZConstructionPlane)
    sketchCircles = sketch.sketchCurves.sketchCircles
    centerPoint = adsk.core.Point3D.create(x, y, 0)
    circle = sketchCircles.addByCenterRadius(centerPoint, r)
    return sketch.profiles.item(0)

def drawCircleXZ(rootComp,x,y,r):

    sketches = rootComp.sketches   
    sketch = sketches.add(rootComp.xZConstructionPlane)
    sketchCircles = sketch.sketchCurves.sketchCircles
    centerPoint = adsk.core.Point3D.create(x, y, 0)
    circle = sketchCircles.addByCenterRadius(centerPoint, r)
    return sketch.profiles.item(0)

def addCenterPointRectangleXY(rootComp,x,y,L,H):

    sketches = rootComp.sketches
    sketch = sketches.add(rootComp.xYConstructionPlane)
    # Define two points
    centerPoint = adsk.core.Point3D.create(x, y, 0)
    cornerPoint = adsk.core.Point3D.create(L/2, H/2, 0)

    # Create the rectangle using the points
    sketchRectangles = sketch.sketchCurves.sketchLines 
    rectangle = sketchRectangles.addCenterPointRectangle(centerPoint, cornerPoint)
    return sketch.profiles.item(0)

def addTwoPointRectangleXY(rootComp,x,y,L,H):

    sketches = rootComp.sketches
    sketch = sketches.add(rootComp.xYConstructionPlane)
    # Define two points
    pointOne = adsk.core.Point3D.create(x, y, 0)
    pointTwo = adsk.core.Point3D.create(x+L, y+H, 0)

    # Create the rectangle using the points
    sketchRectangles = sketch.sketchCurves.sketchLines 
    rectangle = sketchRectangles.addTwoPointRectangle(pointOne, pointTwo)
    return sketch.profiles.item(0)

def addTwoPointRectangleYZ(rootComp,x,y,L,H):

    sketches = rootComp.sketches
    sketch = sketches.add(rootComp.yZConstructionPlane)
    # Define two points
    pointOne = adsk.core.Point3D.create(x, y, 0)
    pointTwo = adsk.core.Point3D.create(x+L, y+H, 0)

    # Create the rectangle using the points
    sketchRectangles = sketch.sketchCurves.sketchLines 
    rectangle = sketchRectangles.addTwoPointRectangle(pointOne, pointTwo)
    return sketch.profiles.item(0)

def addTwoPointRectangleXZ(rootComp,x,y,L,H):

    sketches = rootComp.sketches
    sketch = sketches.add(rootComp.xZConstructionPlane)
    # Define two points
    pointOne = adsk.core.Point3D.create(x, y, 0)
    pointTwo = adsk.core.Point3D.create(x+L, y+H, 0)

    # Create the rectangle using the points
    sketchRectangles = sketch.sketchCurves.sketchLines 
    rectangle = sketchRectangles.addTwoPointRectangle(pointOne, pointTwo)
    return sketch.profiles.item(0)

def extrudeProfile(prof,rootComp,d):

    distance = adsk.core.ValueInput.createByReal(d)
    # Get extrude features
    extrudes = rootComp.features.extrudeFeatures
    extrude1 = extrudes.addSimple(prof, distance, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
     
    # Get the extrusion body
    return extrude1.bodies.item(0)

def run(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
            
        # Create a document.
        doc = app.documents.add(adsk.core.DocumentTypes.FusionDesignDocumentType)
    
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)

        # Get the root component of the active design
        rootComp = design.rootComponent
                    

        # Important to Note: Default units in Fusion 360 API is cm and it cannot be changed to mm

    
        """
        # Create sketch     
        #circle = drawCircle(rootComp,5,5,1)
            
        # Get the profile defined by the circle
        profile_list.append(drawCircleXY(rootComp,5,5,1))
            
        # Extrude Sample 1: A simple way of creating typical extrusions (extrusion that goes from the profile plane the specified distance).
        # Define a distance extent of 5 cm

        body1 = extrudeProfile(profile_list[0],rootComp,5)
        body1.name = "simple"


        profile_list.append(addTwoPointRectangleXY(rootComp,d1,e1,f1,g1))
        body2 = extrudeProfile(profile_list[1],rootComp,2)
        body2.name = "simple2"

        profile_list.append(addTwoPointRectangleXY(rootComp,d2,e2,f2,g2))
        body3 = extrudeProfile(profile_list[2],rootComp,1)
        body3.name = "simple3"
        
        """

        for num in range(row_count):
            if sheet.cell(num, 0).value=='S':
                if sheet.cell(num, 2).value=='R':
                    if sheet.cell(num, 1).value=='XY':
                        x1 = sheet.cell(num, 3).value
                        y1 = sheet.cell(num, 4).value
                        le1 = sheet.cell(num, 5).value
                        ht1 = sheet.cell(num, 6).value
                        profile_list.append(addTwoPointRectangleXY(rootComp,x1,y1,le1,ht1))
                    elif sheet.cell(num, 1).value=='YZ':
                        x1 = sheet.cell(num, 3).value
                        y1 = sheet.cell(num, 4).value
                        le1 = sheet.cell(num, 5).value
                        ht1 = sheet.cell(num, 6).value
                        profile_list.append(addTwoPointRectangleXY(rootComp,x1,y1,le1,ht1))
                    elif sheet.cell(num, 1).value=='ZX':
                        x1 = sheet.cell(num, 3).value
                        y1 = sheet.cell(num, 4).value
                        le1 = sheet.cell(num, 5).value #length of rectangle
                        ht1 = sheet.cell(num, 6).value #height of rectangle
                        profile_list.append(addTwoPointRectangleXY(rootComp,x1,y1,le1,ht1))
                elif sheet.cell(num, 2).value=='C':
                    if sheet.cell(num, 1).value=='XY':
                        x1 = sheet.cell(num, 3).value
                        y1 = sheet.cell(num, 4).value
                        dm1 = sheet.cell(num, 5).value #diameter
                        profile_list.append(drawCircleXY(rootComp,x1,y1,dm1))
                    elif sheet.cell(num, 1).value=='YZ':
                        x1 = sheet.cell(num, 3).value
                        y1 = sheet.cell(num, 4).value
                        dm1 = sheet.cell(num, 5).value #diameter
                        profile_list.append(drawCircleYZ(rootComp,x1,y1,dm1))
                    elif sheet.cell(num, 1).value=='ZX':
                        x1 = sheet.cell(num, 3).value
                        y1 = sheet.cell(num, 4).value
                        dm1 = sheet.cell(num, 5).value #diameter
                        profile_list.append(drawCircleXZ(rootComp,x1,y1,dm1))
            elif sheet.cell(num, 0).value=='E':
                if sheet.cell(num, 1).value=='E':
                    ed1 = sheet.cell(num, 2).value
                    profileindex = num/2
                    profileindex = int(profileindex)
                    bodies_list.append(extrudeProfile(profile_list[profileindex],rootComp,ed1))




        # Get the state of timeline object
        timeline = design.timeline
        timelineObj = timeline.item(timeline.count - 1);
        health = timelineObj.healthState
        message = timelineObj.errorOrWarningMessage
     
        ui.messageBox('Body Created')

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))