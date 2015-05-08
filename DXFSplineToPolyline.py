#Author-Xiaodong Liang, Autodesk
#A companion command of [Save As DXF] of sketch. Can convert spline to polyline

import adsk.core, adsk.fusion, traceback, os #, re, math, string   

commandIdOnPanel = 'id_DXFSplineToPolyline'
workspaceToUse = 'FusionSolidEnvironment'
panelToUse = 'SketchPanel'

# global set of event handlers to keep them referenced for the duration of the command
handlers = []

####DXF export begin
# string is a number
def is_number(s):
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
def ReplaceDXF(spline_polyline_map,olddxf):
    
    updateStr = olddxf;
    #search the first spline
    matches =  [i for i,x in enumerate(updateStr) if x == 'SPLINE\n']
    
    mapIndex = 0
    while len(matches) >0 :            
            #since we will change the length of the list, 
            #check the first one only every time
            indexOf_Spline = matches[0]
            #replace with LWPOLYLINE\n
            updateStr[indexOf_Spline] ='LWPOLYLINE\n'            
            
            #search the AcDbSpline section right after SPLINE\n
            matches = [i for i,x in enumerate(updateStr) if x == 'AcDbSpline\n']            
            if len(matches) ==0:
                pass
            indexOf_AcDbSpline = matches[0]
            #replace with AcDbPolyline
            updateStr[indexOf_AcDbSpline] ='AcDbPolyline\n' 
            
            #mark the rows of spline values until the next entity
            for k  in range(indexOf_AcDbSpline + 1,len(updateStr)-1):                                
                if is_number(updateStr[k+1].rstrip())==False:
                    break;
                if is_number(updateStr[k].rstrip()) and \
                   is_number(updateStr[k+1].rstrip()) :
				    #set this row with the guid. which means this row will be deleted later 
                    updateStr[k]  ='{E91751B6-09C7-4E27-9A44-D0A77EB9EBB3}\n' 
                    
            #add the rows of the polyline values
            for eachVertexItem in spline_polyline_map[mapIndex]:                    
                 updateStr.insert(k,eachVertexItem)
                 k+=1 
            
            try:
                #search the next spline
                matches =  [i for i,x in enumerate(updateStr) if x == 'SPLINE\n']      
            except ValueError:
                pass 
            
            print(matches) 

    #remove the rows that are marked
    updateStr[:] = (value for value in updateStr \
                    if value != '{E91751B6-09C7-4E27-9A44-D0A77EB9EBB3}\n') 
                    
    return updateStr
 
#convert spline to polyline    
def BSplineToLines(oSketch,
                   unitMgr,
                   oSpline,
                   precision, 
                   nbSpline,                                         
                   spline_polyline_map
                   ): 
     
     #get curve evaluator
     oCurveEvaluator = oSpline.evaluator 
     #get range of param
     (ReturnValue, minP, maxP) = \
         oCurveEvaluator.getParameterExtents()     
         
     #get step
     (ReturnValue,length)= \
         oCurveEvaluator.getLengthAtParameter(minP, maxP)
     nbPoints =precision 
     distBetweenPoints = length/nbPoints
     
     params = []    
     for i in range(0, nbPoints):         
         pos = minP + i * distBetweenPoints
         param=0
         (ReturnValue,param) = oCurveEvaluator.getParameterAtLength(minP, pos)
         params.append(param)         
     params.append(maxP)
     
     #get points at params
     (ReturnValue,points) = oCurveEvaluator.getPointsAtParameters(params)    
     
     #prepare DXF header of a polyline
     polyLineDXFValues = [] 
     polyLineDXFValues.append('90\n')
     polyLineDXFValues.append(str(nbPoints) +'\n')
     polyLineDXFValues.append('43\n')
     polyLineDXFValues.append('0.0\n') 
    
     d = {0:'mm',
          1:'cm',
          2:'m',
          3:'inch',
          4:'foot'}     
     currentLenUnit = unitMgr.formatUnits( d[unitMgr.distanceDisplayUnits] )
     
     #add points value one by one
     for k in range(0,len(points)): 
         eachPt = points[k]
         x = unitMgr.convert(eachPt.x, 'cm',currentLenUnit ) 
         y = unitMgr.convert(eachPt.y,'cm',currentLenUnit  ) 
         polyLineDXFValues.append('10\n')
         polyLineDXFValues.append(str(x)+'\n')
         polyLineDXFValues.append('20\n')
         polyLineDXFValues.append(str(y)+'\n') 
         
     #the polyline for this spline    
     spline_polyline_map[nbSpline]  = polyLineDXFValues  

#DXF command	 
def DXFExportMain(precision):
    ui = None
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
        #active sketch
        activeObj = design.activeEditObject  

        if activeObj.objectType != "adsk::fusion::Sketch":
            ui.messageBox('Active object is NOT a sketch!', 'Not a Sketch')
            return        
        
        #default dxf saving 
        fn = os.path.join(os.path.dirname(__file__), 'fusion_temp.dxf')         
        activeObj.saveAsDXF(fn);        
        dxfinput = open(fn, 'r')
        olddxfstr = dxfinput.readlines() 
     
        #ask user to input precision        
#        inputPre = '50'
#        (ReturnValue, objIsCancelled) = ui.inputBox(
#            'Number of sections to split splines into',                                
#            'Precision', 
#            inputPre)
            
        #Exit the program if the dialog was cancelled.
#        if objIsCancelled:
            #use the default value
            #precision = 50
#            return
#        else:
#            precision = int(ReturnValue) 
            
        
        #####check spline################
        unitsMgr = design.fusionUnitsManager  
        
        spline_polyline_map = {}  
        nbSpline = 0
        #check each spline             
        for eachSpline in activeObj.sketchCurves.sketchFittedSplines:
            #fitted spline
            BSplineToLines(activeObj,
                           unitsMgr,
                           eachSpline.geometry,
                           precision,
                           nbSpline,
                           spline_polyline_map)
            nbSpline += 1
       
        for eachSpline in activeObj.sketchCurves.sketchFixedSplines:
            #fixed spline
            print("fixed spline")            
        #####end of checking spline######
            
        #replace DXF 
        updateStr = ReplaceDXF(spline_polyline_map,olddxfstr)
    
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
            
        #write the content to the new DXF
        output = open(filename, 'w')
        output.writelines(updateStr)
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
        commandDescription = 'Exports the active sketch to DXF and converts the ' \
        'splines in the created file to polylines'
        commandResources = './resources/command'

        app = adsk.core.Application.get()
        ui = app.userInterface

        class CommandExecuteHandler(adsk.core.CommandEventHandler):
            def __init__(self):
                super().__init__()
            def notify(self, args):
                try:
                    #command = args.firingEvent.sender
                    #ui.messageBox('command: ' + command.parentCommandDefinition.id + ' executed successfully')
                    command = args.firingEvent.sender                    
                    inputs = command.commandInputs
                    
                    precision = None
                    for input in inputs:
                        if input.id == 'numOfSections':                    
                            precision = int(input.valueOne)
                            
                    #ui.messageBox('%.2f' % (precision))         
                    DXFExportMain(precision)
                except:
                    if ui:
                        ui.messageBox('command executed failed:\n{}'.format(traceback.format_exc()))

        class CommandCreatedEventHandlerPanel(adsk.core.CommandCreatedEventHandler):
            def __init__(self):
                super().__init__() 
            def notify(self, args):
                try:
                    cmd = args.command
                    onExecute = CommandExecuteHandler()
                    cmd.execute.add(onExecute) 
                    # keep a reference to it 
                    handlers.append(onExecute)

                    inputs = cmd.commandInputs
                    #initNumOfSections = adsk.core.ValueInput.createByReal(50)
                    #inputs.addValueInput('numOfSections', 'Number of sections to split splines into', '', initNumOfSections)
                    numOfSections = inputs.addRangeCommandFloatInput('numOfSections', 'Number of sections to split splines into', '', 10, 500, False) 
                    numOfSections.valueOne = 50
                    numOfSections.spinStep = 10
                    #ui.messageBox('Panel command created successfully')
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
                commandDefinitionPanel_ = commandDefinitions_.addButtonDefinition(commandIdOnPanel, commandName, commandName, commandResources)
                commandDefinitionPanel_.tooltipDescription = commandDescription
            onCommandCreated = CommandCreatedEventHandlerPanel()
            commandDefinitionPanel_.commandCreated.add(onCommandCreated)
            # keep the handler referenced beyond this function
            handlers.append(onCommandCreated)
            toolbarControlPanel_ = toolbarControlsPanel_.addCommand(commandDefinitionPanel_, commandIdOnPanel)
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
