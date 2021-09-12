from tkinter import filedialog as fd
from pkg01.datain import *
#from pkg01.cad import *
from pkg01.dataout import *
from pkg02.util import *

top = Tk()

# Set up parameters
# :project parameter shall be utilized for data manipulation
def setparams():
    global doc, cadapp
    global workdir, rtkfile, rtkcolumns, rtkencoding
    global xsline_layer, chn_layer
    global xscode_layer, xsname_layer, xspoint_layer
    global xsline_completed_layer
    global completed_color, Buffer

    workdir = proj_params['WorkDirectory']
    rtkfile = proj_params['RTKDatatFile']
    rtkcolumns = proj_params['RTKColumns']
    rtkencoding = proj_params['RTKEncoding']
    #outfile = proj_params['OutputCsvFile']
    #xlsfile = proj_params['OutputXlsFile']
    cadapp = proj_params['CadApp']
    xsline_layer = proj_params['XSLineLayer']
    chn_layer = proj_params['ChainageLayer']
    xscode_layer = proj_params['XSCodeLayer']
    xsname_layer = proj_params['XSNameLayer']
    xspoint_layer = proj_params['XSPointLayer']
    #xscomppoint_layer = proj_params['XSComputedPointLayer']
    #drawxscomppoint = proj_params['DrawXSComputedPoint']
    completed_color = proj_params['CompletedColor']
    xsline_completed_layer = proj_params['XSLineCompletedLayer']
    Buffer = proj_params['Buffer']

    doc = is_cadopen()                          # Checking AutoCAD is opened or not?
    if doc is None:
        return False
    return doc

# Select parameter file
def selectfile():
    global proj_params

    statusbox(sta_label, 'Open parameter file.')
    parfile = fd.askopenfilename(title='Select Parameter File')
    if parfile == '':
        return
    proj_params = getProjParams('', parfile)
    #print(proj_params)
    if proj_params == {}:
        msg = 'Incorrect Parameter File format!!!'
        for i in range(4):
            cad.entryconfig(i, state=DISABLED)
        warn_message(msg)
        return
    conn_ok = setparams()               # Check parameters & CAD connection
    if rtkfile != '' and workdir != '' and conn_ok:
        cad.entryconfig(0, state=NORMAL)
    if cadapp != '' and workdir != '' and conn_ok:
        cad.entryconfig(1, state=NORMAL)
        cad.entryconfig(2, state=NORMAL)
        statusbox(sta_label, 'AutoCAD connected.')

# To import points from CSV file
def importpoints():
    #show_message('>>>')
    #print('>>>')
    if not is_cadready():
        return False
    statusbox(sta_label, 'Importing points...')
    #statusbox2('Importing points...')
    rtkdata = getRTK(workdir, rtkfile, rtkcolumns, rtkencoding)     # Read data from CSV file
    rtk2ac(rtkdata, xscode_layer, xsname_layer, xspoint_layer)      # Create ACAD points

# Draw X-section line
def drawxline():
    doc = is_cadready()
    if not doc:
        return False
    statusbox(sta_label, 'Create Line of X-Section.')
    createxline(cadapp, xsline_layer, chn_layer, xscode_layer)

# To extract points from X-sections
def xs2dtabs():
    doc = is_cadready()
    if not doc:
        return False
    statusbox(sta_label, 'Extract X-Section...')
    statusbox(sta_label, 'Interact with AutoCAD Window >>>')
    #createxsfile(proj_params)
    xsdtab = create_xs_dtab(doc, proj_params, sta_label)
    cad.entryconfig(3, state=NORMAL)
    #xsdtab.xs_show_csvdata()
    #xsdtab.xs_show_xyzdata()

# To create files of X-sections
def xs2files():
    statusbox(sta_label, 'Create Files of X-section')
    create_xs_file(rtkencoding)

def main():
    global cad, sta_label

    menubar = Menu(top)
    file = Menu(menubar, tearoff=0)
    #file.add_command(label="New")
    file.add_command(label="Open", command=selectfile)
    #file.add_command(label="Save")
    #file.add_command(label="Save as...")
    #file.add_command(label="Close")

    file.add_separator()
    file.add_command(label="Exit", command=top.quit)

    menubar.add_cascade(label="File", menu=file)
    cad = Menu(menubar, tearoff=0)
    cad.add_command(label="Import Points", state=DISABLED, command=importpoints)
    cad.add_command(label="Draw X-Line", state=DISABLED, command=drawxline)
    cad.add_command(label="eXtract XS", state=DISABLED, command=xs2dtabs)
    cad.add_command(label="XS->Files", state=DISABLED, command=xs2files)

    cad.add_separator()

    #edit.add_command(label="Cut")
    #edit.add_command(label="Copy")
    #edit.add_command(label="Paste")
    #edit.add_command(label="Delete")
    cad.add_command(label="Select All", state=DISABLED)

    menubar.add_cascade(label="CAD", menu=cad)
    help = Menu(menubar, tearoff=0)
    help.add_command(label="About", state=DISABLED)
    menubar.add_cascade(label="Help", menu=help)

    top.config(menu=menubar)
    top.geometry('590x400')
    top.geometry('+150+100')                 # Position ('+Left+Top')
    top.title('THGeom Academy (RTK data to X-section: Offset & Elevation)')
    sta_label = Label(top, text=': ', width=40)
    #sta_label.place(x=-1.0, rely=1.0, anchor='sw')
    sta_label.pack()
    sta_label.place(relx=-0.1, rely=1.0, anchor=SW)
    #top.update()
    top.mainloop()




# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
