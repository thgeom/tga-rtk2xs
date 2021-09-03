import win32com.client                                                  # For Application connection
import pythoncom
import comtypes
from pkg02.util import *

acad = win32com.client.Dispatch("AutoCAD.Application")                  # AutoCAD connection
#doc = acad.ActiveDocument
#print(dir(acad))

# Verify AutoCAD connection
def is_cadconnected():
    global doc, acprompt, ms
    conn_ok = False
    try:
        doc = acad.ActiveDocument
        #print(doc)
        print('File {} connected.'.format(doc.Name))
        #doc.Utility.Prompt("Execute from python\n")
        acprompt = doc.Utility.Prompt                                           # ACAD prompt
        ms = doc.ModelSpace
        conn_ok = True
        return doc
    except AttributeError:
        print('Connect to AutoCAD failed.!!!')
        print('Press Esc on AutoCAD window then try again.')
        return conn_ok


#doc.Utility.Prompt("Execute from python\n")
#acprompt = doc.Utility.Prompt                                           # ACAD prompt
#ms = doc.ModelSpace
"""
layer_code = 'XS_Code'
layer_name = 'XS_Name'
layer_point = 'XS_point'
"""
# Point foe win32com
def vtpt(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def vtobj(obj):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

def vtFloat(lis):
    """ list converted to floating points"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lis)

def vtint(val):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, val)

def vtvariant(var):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, var)

#Convert pt to vtpt
def pt_vtpt(pt):
    return vtpt(pt[0], pt[1], pt[2])


# Draw points in CAD
def pts2ac(pts, code_lay, name_lay, point_lay):
    #To create entities of RTK points
    p1 = [0.0, 0.0, 0.0]
    i = 0
    print('Processing.', end='')
    acprompt('Points processing.')
    for pt in pts.restab.values:
        p1[0] = pt[2]
        p1[1] = pt[3]
        p1[2] = pt[4]
        code = pt[0]
        name = pt[1]
        #print('Processing point : ' + name)
        if (i % 20) == 0:
            print('.', end='')                                          # Print dot for every 20 points
            acprompt('.')

        p_code = ms.AddText(code, pt_vtpt(p1), 1.5)
        p_code.Layer = code_lay                                       # Define Code layer
        p_code.Rotation = math.pi * 0.45
        p_name = ms.AddText(name, pt_vtpt(p1), 2)
        p_name.Layer = name_lay                                        # Define Name layer
        p_pt = ms.AddPoint(pt_vtpt(p1))
        p_pt.Layer = point_lay                                         # Define Point layer
        i = i + 1
    msg = '\nTotal RTK {0:d} points imported to DWG.'.format(i)
    acprompt(msg + '\n')  # Echo to ACAD with format
    #print('\nTotal RTK point = ' + str(i) + ', imported to DWG.')
    show_message(msg)                   # Print with format

# Checking layer exist or not
def layerexist(lay):
    layers = doc.Layers
    #print(layers[1].Name)

    """
    try:
        doc.Layers.Add(layer_Code)
    except:
        print('{} already exist.'.format(layer_Code))
    """
    layers_nums = layers.count
    layers_names = [layers.Item(i).Name for i in range(layers_nums)]    # List of ACAD layers
    if lay in layers_names:
        return True
    else:
        return False

#==========
def rtk2ac(rtk1, code_lay, name_lay, point_lay):
    print('Drawing Name is {}'.format(doc.Name))                        # Print ACAD Dwg. name
    print('Import Field data from RTK')
    acprompt('Hi, from Python : To manage RTK\n')
    acprompt('RTK file importing...\n')
    #rtk1 = getRTK()

    if not layerexist(code_lay):
        doc.Layers.Add(code_lay)                                      # Add layer if not exist
    #else:
    #print('{} already exist.'.format(layer_code))

    if not layerexist(name_lay):
        doc.Layers.Add(name_lay)
    if not layerexist(point_lay):
        doc.Layers.Add(point_lay)


    pts2ac(rtk1, code_lay, name_lay, point_lay)
    print('Import Field Data Points : Completed.')
    #show_message('Import Field Data Points : Completed.')
    doc.Regen(1)
    acprompt('Import & Draw Points Completed.\n')

#========
# To create X-section line on "XS_Line" layer
def cadInput(xspoint_layer):
    doc.Utility.Prompt('Select X-section line : ')
    xsObjSel = doc.Utility.GetEntity()                  #Get XS_Line entity by pick
    #return (<COMObject GetEntity>, (506465.30556296057, 1861201.4573297906, 0.0))
    xsObj = xsObjSel[0]
    #print(xsObj.EndPoint)
    txtpt = None
    while txtpt is None:
        doc.Utility.Prompt('Select Point code of center line : ')
        pcObjSel = doc.Utility.GetEntity()                  #Get P_Code entity by pick
        pcObj = pcObjSel[0]
        if pcObj.Layer == xspoint_layer:
            txtpt = pcObj.Layer
    #doc.Utility.Prompt('Chainage or X-section name : ')
    chn = doc.Utility.GetString(1, 'Chainage or X-section name : ')     #Get CHN string
    return [xsObj, pcObj, chn]

def cadProc( xsObj, pcObj, chn, cadapp, xslinelay, chnlay):
    p2 = xsObj.EndPoint
    chnObj = ms.AddText(chn, pt_vtpt(p2), 5)            #Create CHN @EndPoint
    chnObj.Layer = chnlay
    chnObj.Rotation = xsObj.Angle
    dataType = (1001, 1000, 1005, 1005)                 #Define Xdata
    data = (cadapp, chn, chnObj.Handle, pcObj.Handle)
    dataType = vtint(dataType)                          #Converse dataType format
    data = vtvariant(data)
    xsObj.SetXData(dataType, data)                      #Setting Xdata of X-section
    # Return the xdata for the line
    #xtypeOut, xdataOut = xsObj.GetXData("rtk_xs")
    #print(xtypeOut)
    #print(xdataOut)
    xsObj.Layer = xslinelay                             #Set layer of X-section to "XS_Line"

# Create X-Section Line
def createxline(cadapp, xsllay, chnlay, xsptlay):
    ## Start to create X-Section line
    print('>>> Select X-Section on CAD WINDOW')
    try:
        [xsObj, pcObj, chn] = cadInput(xsptlay)                    #Call cadInput to select X-section,
    except:
        return
    #point of center line and define X-section name
    cadProc(xsObj, pcObj, chn, cadapp, xsllay, chnlay)                         #Call cadProc to manipulate XS_Line
    doc.Application.Update()                                #Redraw CAD window
    doc.Utility.Prompt('>>>> X-section : {} has been created.\n'.format(chn))
    show_message('>>>> X-section : {} has been created.\n'.format(chn))

