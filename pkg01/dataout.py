#import pythoncom
import pandas as pd
from sympy import Line                   # For geometry calculation

#from pkg02.util import *
from pkg01.cad import *

# Class Xs4Xls to prepare X-section data for Excel file
class Xs4Xls:
    def __init__(self, fdir, fout):
        self.fdir = fdir
        self.fout = fout
        self.data = []

    def xsAdd(self, xsdt):
        for i in xsdt:
            self.data.append(i)

    # Writing data to Excel file using Pandas Data Frame
    def xs2xlsFile(self):
        df = pd.DataFrame(self.data, columns=['Chainage', 'Offset', 'Elevation', 'Code']) # DataFrame to Excel
        with pd.ExcelWriter(self.fdir + self.fout, mode='w') as writer:
            df.to_excel(writer, sheet_name='Offset_Elevation')

# X-section class for data manipulation
class XsInfo:
    CHN = []                                                # Initialize variables
    ptc = []
    p1, p2 = [], []
    ofs_ele = []
    enz = []
    num_xs = 0

    def __init__(self, ename, buffer):
        self.ename = ename
        #self.ptc = ptc
        self.buffer = buffer
        self._initsl()
        self.dt2xls = []
        XsInfo.num_xs += 1

    # Initialize selection set "SS2"
    def _initsl(self):
        # Add the name "SS2" selection set
        try:
            doc.SelectionSets.Item("SS2").Delete()
        except:
            print("Delete selection failed")

        self.slpts = doc.SelectionSets.Add("SS2")

    # Compute boundary of X-section by buffer distance
    def XsBounds(self):
        self.p1 = self.ename.StartPoint
        self.p2 = self.ename.EndPoint
        self.bounds = line_bounds(self.ename, self.buffer)
        self.pnts = bounds2list(self.bounds)                # compute pnts for SelectByPolygon

    # Select XS point by XS_Line with Xdata attached,
    def getXsPoints(self):
        self.XsBounds()                                     # Compute XS boundary
        self.slpts.Clear()
        p1 = self.p1
        p2 = self.p2
        """
        cmd = 'Zoom W ' + str(p1).replace(" ", "") + ' ' + str(p2).replace(" ", "") + ' '
        cmd = cmd.replace("(", "")
        cmd = cmd.replace(")", "")                          # Create command line
        #cmd = 'Zoom E '             #Example for Zoom Extend
        doc.SendCommand(cmd)                                # Send command to CAD for Zoom window of XS line
        """
        doc.Application.ZoomWindow(pt_vtpt(p1), pt_vtpt(p2))    # Zoom window of XS line in CAD
        ftyp = [0, 8]                                           # Set up filter condition
        ftdt = ["Text", xscode_layer]
        """
        #Example format of pnts for SelectionByPolygon
        #pnts = [0, 0, 0, 750, 750, 0, 550, 900, 0, -180, 120, 0, 0, 0, 0]
        #pnts = vtFloat(pnts)
        """
        pnts = vtFloat(self.pnts)                           # pnts from Xsbounds
        filterType = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, ftyp)
        filterData = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ftdt)
        self.slpts.SelectByPolygon(6, pnts, filterType, filterData)      # mode = 6, SelectByPolygon
        #doc.SendCommand('regen ')
        doc.Regen(1)                                        # Regenerate drawing
        doc.Utility.Prompt("{0:d} points selected\n".format(self.slpts.count)) # echo to CAD command prompt
        #doc.Utility.Prompt(str(self.slpts.count) + " points selected\n")

    # Calculate offset & elevation of XS points, using XS line and its Xdata
    def calOfsEle(self):
        l1 = Line(self.p1, self.p2)                         # Create Line object for perpendicular
        self.enz = []
        self.ofs_ele = []
        xs_ang = self.ename.Angle
        xdata = self.ename.GetXData(cadapp)                 # Get Xdata from XS_Line
        """
        #Show XS line Xdata content
        #print(xdata)
        #xdata in format : ((1001, 1000, 1005, 1005), ('rtk_xs', 'X81', '3C22', '37E3'))
        #'rtk_xs' is an AppName in ACAD extended data
        """
        self.CHN = xdata[1][1]
        cc = doc.HandleToObject(xdata[1][3])                # Center point entity of X-section
        ptc = cc.InsertionPoint
        self.ptc = ptc
        s1 = l1.perpendicular_segment(ptc)                  # Compute CL point on XS line
        ptc = (float(s1.p2.x), float(s1.p2.y))              # Define CL point format
        for i in self.slpts:
            pt = i.InsertionPoint                           # Get point from Text entity
            s1 = l1.perpendicular_segment(pt)               # Compute point on XS line
            ptx = (float(s1.p2.x), float(s1.p2.y))
            el = pt[2]
            p_code = i.TextString                           # Get Code from Text entity
            di = distance(ptc, ptx)
            ang = angle(ptc, ptx)
            dang = abs(xs_ang - ang)
            if (dang > math.pi * 0.5) and (dang < math.pi * 1.5):
                di = 0 - di
            self.ofs_ele.append([p_code, (di, el)])         # Create list of Offset & Elevation with Code

            ptx = (float(s1.p2.x), float(s1.p2.y), el)      # Set format of ENZ
            #print(p_code, ptx)
            self.enz.append([p_code, ptx])                  # Create list of ENZ with Code
        self.ofs_ele = sort_rtk_x(self.ofs_ele)             # Sorting by Offset distance
        self.enz = sort_rtk_x(self.enz)                     # Sorting by East
        if (xs_ang > math.pi * 0.5) and (xs_ang < math.pi * 1.5):
            self.enz.reverse()

    # Writing XS points to CSV file
    def xs2csvFile(self, fdir, fout):
        f1 = open(fdir + fout, "a")                         # Open file output for Offset & Elavation
        f2 = open(fdir + 'xyz_' + fout, "a")                # Open file output for ENZ
        f1.write(self.CHN)
        f2.write(self.CHN + '\n')
        i = 0
        self.dt2xls = []                                         # For Excel
        for pt in self.ofs_ele:
            if i == 0:
                f1.write('      {0:0.0f}     {1:0.3f}    {2}\n'.format(pt[1][0], pt[1][1], pt[0]))
                ofs_ele = ([self.CHN, round(pt[1][0]), pt[1][1], pt[0]])      #Data format for Excel
                self.dt2xls.append(ofs_ele)
            else:
                f1.write('           {0:0.0f}     {1:0.3f}    {2}\n'.format(pt[1][0], pt[1][1], pt[0]))
                ofs_ele = (['        ', round(pt[1][0]), pt[1][1], pt[0]])
                self.dt2xls.append(ofs_ele)
            i += 1
        for pt in self.enz:
            f2.write('{0:0.3f}   {1:0.3f}   {2:0.3f}   {3}\n'.format(pt[1][0], pt[1][1], pt[1][2], pt[0]))
            #Write ENZ to file
        f2.write('#\n')
        f1.close()
        f2.close()
        print('Total points = {:d} '.format(i))
        for i in self.slpts:                            # Change color of XS_Code to completed-color
            i.Color = completed_color

        #self.ename.Layer = xsline_completed_layer      # Set XS_Line layer = "XS_Line_Completed"



"""
#object.Select( Mode , Point1 , Point2 , FilterType , FilterData )
#- Object : Object SelectionSet at The Applies to the this Method,..
#- Mode : selection mode, AcSelect enum, the specific meaning of the table.
#- Point1 : 3-dimensional coordinates.
#- Point2 : 3-dimensional coordinates.
#- FilterType : Group A code Specifying the DXF The type of filter to use.
#- FilterData : at The filter value to ON.
#————————————————
#：https://blog.csdn.net/Hulunbuir/article/details/95446723

#   Mode	            enum	Description
#------------------------------------------
# acSelectionSetWindow	    0	Selects all objects completely inside a rectangular area whose corners are 
                                defined by Point1 and Point2.
# acSelectionSetCrossing	1	Selects objects within and crossing a rectangular area whose corners are 
                                defined by Point1 and Point2.
# acSelectionSetPrevious	3	Selects the most recent selection set. This mode is ignored if you switch 
                                between paper space and model space and attempt to use the selection set.
# acSelectionSetLast	    4	Selects the most recently created visible objects.
# acSelectionSetAll	        5	Selects all objects
#————————————————
#：https://blog.csdn.net/Hulunbuir/article/details/95446723
"""
"""
# Add the name "SS1" selection set
try:
    doc.SelectionSets.Item("SS1").Delete()
except:
    print("Delete selection failed")
slcn = doc.SelectionSets.Add("SS1")
"""

#==========
def createxsfile(doci, proj_params, sta_label):
    global xscode_layer, completed_color, xsline_completed_layer
    global doc, cadapp

    doc = doci
    workdir = proj_params['WorkDirectory']
    outfile = proj_params['OutputCsvFile']
    xlsfname = proj_params['OutputXlsFile']

    cadapp = proj_params['CadApp']
    xsline_layer = proj_params['XSLineLayer']
    xscode_layer = proj_params['XSCodeLayer']
    completed_color = proj_params['CompletedColor']
    xsline_completed_layer = proj_params['XSLineCompletedLayer']
    Buffer = proj_params['Buffer']

    # Add the name "SS1" selection set
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        print("Delete selection failed")
    slcn = doc.SelectionSets.Add("SS1")


    #p1list = []
    #xlsdata = Xs4Xls('d:/TGA_Lisp/', 'xs-03.xlsx')
    #xlsfname = input('Name of Excel file : ')
    #xlsfname = doc.Utility.GetInput()
    try:
        #print('Interact with AutoCAD Window >>>')
        doc.Utility.Prompt('To Extract X-section Data >>>\n')
        #print(dir(doc.Utility))
        #doc.Utility.InitializeUserInput(1)
        keypress = doc.Utility.GetKeyword('Press [Enter] to continue, [Esc] to cancel : ')
        #doc.Regen(1)
        #doc.Application.Update()                                #Redraw CAD window
    except:
        return

    #xlsfdir = 'd:/TGA_Lisp/'                        # Folder for Output Excel file
    #xlsfname = None
    #while not xlsfname:
    #xlsfname = doc.Utility.GetString(1, 'Output Excel file name : ')
    xlsdata = Xs4Xls(workdir, xlsfname)    # Excel file name from proj_params
    print('>>>X-section extraction start')
    doc.Regen(1)
    # Filtered of Line on layer XS_Line
    ftyp = [0, 8]
    ftdt = ["Line", xsline_layer]                                  # Filter with Line & XS_Line layer

    filterType = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, ftyp)
    filterData = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ftdt)

    #slcn.Select(5)
    slcn.Select(5, 0, 0, filterType, filterData)            # Select all with filtering
    print('{:d} X-sections are filtered'.format(slcn.count))

    j = 0
    XsInfo.num_xs = 0           # initialize number of X-section
    for i in range(slcn.count):
        #Ename = slcn[i]
        #print(Ename.GetXData(cadapp))

        xdata = slcn[i].GetXData(cadapp)
        #print(xdata)
        if xdata != (None, None):                       # Xdata 'rtk_xs' attached to X-section line
            CHN = xdata[1][1]
            #cc = doc.HandleToObject(xdata[1][3])
            #ptc = cc.InsertionPoint
            #print('Processing chainage : {}'.format(CHN))
            msg = 'Processing chainage : {}'.format(CHN)
            doc.Utility.Prompt(msg + '\n')          # Echo to CAD command prompt
            statusbox(sta_label, msg)

            #Testing OBJECT
            #XSObj1 = XSInfo(slcn[0], [507063.544, 1860344.719], 10)
            xsObj1 = XsInfo(slcn[i], Buffer)
            """"
            # Try to create Object at once
            if i == 0:
                xsObj1 = XsInfo(slcn[i], Buffer)
            else:
                xsObj1.ename = slcn[i]
            """

            xsObj1.getXsPoints()                                # Call getXsPoints Function

            xsObj1.calOfsEle()                                  # Call calOfsEle Function
            #print(xsObj1.ofs_ele)
            #print(xsObj1.enz)

            #xsObj1.xs2csvFile('d:/TGA_Lisp/', 'xsec-0.csv')     # Call Xs2File by giving Directory & FileName
            xsObj1.xs2csvFile(workdir, outfile)     # Call Xs2File by giving Directory & FileName
            xlsdata.xsAdd(xsObj1.dt2xls)
            #slcn[i].Layer = 'XS_Line_Completed'                 # change XS_Line to completed
        j += 1
    #for
    xlsdata.xs2xlsFile()                                        # Call xlsdata.xs2xlsFile -> Data to Excel
    msg = '>>>> Total {:d} X-sections extraction completed.'.format(XsInfo.num_xs)
    doc.Utility.Prompt(msg + '\n')  # Echo to CAD command prompt
    show_message(msg)
