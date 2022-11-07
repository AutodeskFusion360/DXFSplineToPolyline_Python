#Author-Xiaodong Liang, Autodesk
#A companion command of [Save As DXF] of sketch. Can convert spline to polyline

import adsk.core, adsk.fusion, traceback, os, time #, re, math, string   

commandIdOnPanel = 'id_DXFSplineToPolyline'
workspaceToUse = 'FusionSolidEnvironment'
panelToUse = 'SketchPanel'

placeHolderString = '{E91751B6-09C7-4E27-9A44-D0A77EB9EBB3}\n'
lastUsedTolerance_cm = 0.1
sketchObj = None

# global set of event handlers to keep them referenced for the duration of the command
handlers = []

def showDbgMsg(text):
    app = adsk.core.Application.get()
    ui  = app.userInterface 
    ui.messageBox(text)

####DXF export begin
# string is a number
def isNumber(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
    return False

#replace old DXF. remove spline section. Add polyline    
def replaceDxf(spline_polyline_map, oldDxfContent):
    
    newDxfContent = oldDxfContent;
    
    mapIndex = 0
    while True:  
        try:
            matches = [i for i,x in enumerate(newDxfContent) if x == 'SPLINE\n']      
        except ValueError:
            break 
        
        #print(matches) 
        
        if len(matches) < 1:
            break
          
        #since we will change the length of the list, 
        #check the first one only every time
        indexOfSpline = matches[0]
        
        #replace with LWPOLYLINE\n
        newDxfContent[indexOfSpline] = 'LWPOLYLINE\n'            
        
        #search the AcDbSpline section right after SPLINE\n
        matches = [i for i,x in enumerate(newDxfContent) if x == 'AcDbSpline\n']            
        if len(matches) < 1:
            pass
        
        indexOfAcDbSpline = matches[0]
        
        #replace with AcDbPolyline
        newDxfContent[indexOfAcDbSpline] ='AcDbPolyline\n' 
        
        #mark the rows of spline values until the next entity
        for k  in range(indexOfAcDbSpline + 1, len(newDxfContent)-1):                                
            if isNumber(newDxfContent[k+1].rstrip()) == False:
                break
            if isNumber(newDxfContent[k].rstrip()) and \
                isNumber(newDxfContent[k+1].rstrip()) : 				    
                #set this row with the guid. which means this row will be deleted later 
                newDxfContent[k] = placeHolderString 
                
        #add the rows of the polyline values
        for eachVertexItem in spline_polyline_map[mapIndex]:                    
             newDxfContent.insert(k, eachVertexItem)
             k += 1 
             
        mapIndex += 1

    #remove the rows that are marked
    newDxfContent[:] = (value for value in newDxfContent \
                    if value != placeHolderString) 
                    
    return newDxfContent
 
#convert spline to polyline    
def convertBSplineToLines(sketch,
                   unitsMgr,
                   spline,
                   tolerance_cm, 
                   splineCount,                                         
                   spline_polyline_map
                   ): 
     
     #get curve evaluator
     curveEvaluator = spline.evaluator 
     
     #get range of param
     (returnValue, minP, maxP) = \
         curveEvaluator.getParameterExtents()     
         
     (returnValue, points) = curveEvaluator.getStrokes(minP, maxP, tolerance_cm) 
     pointCount = len(points)

     #showDbgMsg('tolerance = ' + str(tolerance_cm) + ', number of points = ' + str(pointCount))
     
     #prepare DXF header of a polyline
     polyLineDxfValues = [] 
     polyLineDxfValues.append('90\n')
     polyLineDxfValues.append(str(pointCount) +'\n')
     polyLineDxfValues.append('43\n')
     polyLineDxfValues.append('0.0\n') 
      
     currentLenUnit = unitsMgr.defaultLengthUnits
     
     #showDbgMsg('currentLenUnit = ' + currentLenUnit + "; distanceDisplayUnits = " + str(unitsMgr.distanceDisplayUnits) + "; defaultLengthUnits = " + unitsMgr.defaultLengthUnits)
     
     #add points value one by one
     for k in range(0, pointCount): 
         eachPt = points[k]
         x = unitsMgr.convert(eachPt.x, 'cm', currentLenUnit) 
         y = unitsMgr.convert(eachPt.y, 'cm', currentLenUnit) 
         polyLineDxfValues.append('10\n')
         polyLineDxfValues.append(str(x) + '\n')
         polyLineDxfValues.append('20\n')
         polyLineDxfValues.append(str(y) + '\n') 
         
     #the polyline for this spline    
     spline_polyline_map[splineCount] = polyLineDxfValues  

#DXF command	 
def exportDxf(tolerance_cm):
    ui = None
    try:   
        #get Fusion app
        app = adsk.core.Application.get()
        ui  = app.userInterface

        #pop out file dialog to save the DXF  
        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = "Export to DXF"
        fileDialog.filter = 'DXF files (*.dxf)'
        fileDialog.filterIndex = 0    
        dialogResult = fileDialog.showSave()
        if dialogResult == adsk.core.DialogResults.DialogOK:
            filename = fileDialog.filename
        else:
            return   

        #default dxf saving 
        dxfFileName = os.path.join(os.path.dirname(__file__), 'fusion_temp.dxf')         
        sketchObj.saveAsDXF(dxfFileName);        
        dxfInput = open(dxfFileName, 'r')
        # this returns a list with the lines of the DXF file
        dxfContent = dxfInput.readlines() 
            
        #get design 
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)      
                    
        #####check spline################
        unitsMgr = design.fusionUnitsManager  
        
        spline_polyline_map = {}  
        splineCount = 0
        
        #check each spline             
        for eachSpline in sketchObj.sketchCurves.sketchFittedSplines:
            #fitted spline
            convertBSplineToLines(sketchObj,
                           unitsMgr,
                           eachSpline.geometry,
                           tolerance_cm,
                           splineCount,
                           spline_polyline_map)
            splineCount += 1
       
        for eachSpline in sketchObj.sketchCurves.sketchFixedSplines:
            #fixed spline
            #print("fixed spline")  
            convertBSplineToLines(sketchObj,
                           unitsMgr,
                           eachSpline.geometry,
                           tolerance_cm,
                           splineCount,
                           spline_polyline_map)
            splineCount += 1

        for eachSpline in sketchObj.sketchCurves.sketchControlPointSplines:
            #control point spline
            convertBSplineToLines(sketchObj,
                           unitsMgr,
                           eachSpline.geometry,
                           tolerance_cm,
                           splineCount,
                           spline_polyline_map)
            splineCount += 1
        #####end of checking spline######
            
        #replace DXF 
        if splineCount < 1:    
            #no spline to convert
            newDxfContent = dxfContent
        else:
            newDxfContent = replaceDxf(spline_polyline_map, dxfContent)
            
        #write the content to the new DXF
        output = open(filename, 'w')
        output.writelines(newDxfContent)
        output.close()
        
        ui.messageBox('File exported as "' + filename + '"')
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
            
####DXF export end

def commandDefinitionById(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandDefinition id is not specified')
        return None
    commandDefinitions_ = ui.commandDefinitions
    commandDefinition_ = commandDefinitions_.itemById(id)
    return commandDefinition_

def commandControlByIdForPanel(id):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not id:
        ui.messageBox('commandControl id is not specified')
        return None
    workspaces_ = ui.workspaces
    modelingWorkspace_ = workspaces_.itemById(workspaceToUse)
    toolbarPanels_ = modelingWorkspace_.toolbarPanels
    toolbarPanel_ = toolbarPanels_.itemById(panelToUse)
    toolbarControls_ = toolbarPanel_.controls
    toolbarControl_ = toolbarControls_.itemById(id)
    return toolbarControl_

def destroyObject(uiObj, tobeDeleteObj):
    if uiObj and tobeDeleteObj:
        if tobeDeleteObj.isValid:
            tobeDeleteObj.deleteMe()
        else:
            uiObj.messageBox('tobeDeleteObj is not a valid object')

def run(context):
    ui = None
    try:
        commandName = 'Export to DXF (Splines as Polylines)'
        commandDescription = 'Exports selected sketch to DXF and converts the ' \
        'splines in the created file to polylines\n'
        commandResources = './resources/command'

        app = adsk.core.Application.get()
        ui = app.userInterface

        class CommandExecuteHandler(adsk.core.CommandEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                try:
                    command = args.firingEvent.sender                    

                    #showDbgMsg('command: ' + command.parentCommandDefinition.id + ' executed successfully')
                    
                    inputs = command.commandInputs
                    
                    input = inputs.itemById('tolerance') 
                    
                    #let's store the last used value so that the default
                    #will be what the user last specified
                    global lastUsedTolerance_cm
                    lastUsedTolerance_cm = input.value
                    
                    # value always seems to be in internal unit = cm                                                            
                    exportDxf(input.value)
                except:
                    if ui:
                        ui.messageBox('command executed failed:\n{}'.format(traceback.format_exc()))

        class CommandCreatedEventHandlerPanel(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__() 
            def notify(self, args):
                try:
                    #get Fusion app
                    app = adsk.core.Application.get()
                    ui  = app.userInterface
                    
                    #get design 
                    product = app.activeProduct
                    design = adsk.fusion.Design.cast(product)      
             
                    if design == None:
                        ui.messageBox('No active Fusion design', 'No Design')
                        return  
                        
                    #selected sketch
                    global sketchObj
                    sketchObj = None

                    sketchObj = adsk.fusion.Sketch.cast(product.activeEditObject)
            
                    if not sketchObj:
                        ui.messageBox('Please activate a sketch before running this command!', 'No Sketch Selected')
                        return   

                    # if ui.activeSelections.count == 1:
                    #     sketchObj = ui.activeSelections.item(0).entity
                    #     if sketchObj.objectType != "adsk::fusion::Sketch":
                    #         sketchObj = None
            
                    # if not sketchObj:
                    #     ui.messageBox('Please select a sketch before running this command!', 'No Sketch Selected')
                    #     return   
                        
                    cmd = args.command
                    #cmd.setDialogInitialSize(300, 150)
                    onExecute = CommandExecuteHandler()
                    cmd.execute.add(onExecute) 
                    # keep a reference to it 
                    handlers.append(onExecute)

                    inputs = cmd.commandInputs
                    lengthUnit = design.fusionUnitsManager.defaultLengthUnits 
                    inputs.addValueInput('tolerance', 'Conversion tolerance', lengthUnit, adsk.core.ValueInput.createByReal(lastUsedTolerance_cm))
                    
                    #showDbgMsg('Panel command created successfully')
                except:
                    if ui:
                        ui.messageBox('Panel command created failed:\n{}'.format(traceback.format_exc()))

        commandDefinitions_ = ui.commandDefinitions 
		
        # add a command on create panel in modelling workspace	
        workspaces_ = ui.workspaces
        modelingWorkspace_ = workspaces_.itemById(workspaceToUse)
        toolbarPanels_ = modelingWorkspace_.toolbarPanels
        toolbarPanel_ = toolbarPanels_.itemById(panelToUse) 
        toolbarControlsPanel_ = toolbarPanel_.controls
        toolbarControlPanel_ = toolbarControlsPanel_.itemById(commandIdOnPanel)
        if not toolbarControlPanel_:
            commandDefinitionPanel_ = commandDefinitions_.itemById(commandIdOnPanel)
            if not commandDefinitionPanel_:
                commandDefinitionPanel_ = commandDefinitions_.addButtonDefinition(commandIdOnPanel, commandName, commandDescription, commandResources)
            onCommandCreated = CommandCreatedEventHandlerPanel()
            commandDefinitionPanel_.commandCreated.add(onCommandCreated)
            # keep the handler referenced beyond this function
            handlers.append(onCommandCreated)
            toolbarControlPanel_ = toolbarControlsPanel_.addCommand(commandDefinitionPanel_, '')
            toolbarControlPanel_.isVisible = True

    except:
        if ui:
            ui.messageBox('AddIn Start Failed:\n{}'.format(traceback.format_exc()))

def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
      
        objArrayPanel = []         

        commandControlPanel_ = commandControlByIdForPanel(commandIdOnPanel)
        if commandControlPanel_:
            objArrayPanel.append(commandControlPanel_)

        commandDefinitionPanel_ = commandDefinitionById(commandIdOnPanel)
        if commandDefinitionPanel_:
            objArrayPanel.append(commandDefinitionPanel_) 

        for obj in objArrayPanel:
            destroyObject(ui, obj)

    except:
        if ui:
            ui.messageBox('AddIn Stop Failed:\n{}'.format(traceback.format_exc()))
