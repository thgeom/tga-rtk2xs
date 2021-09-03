import ast
import math
from tkinter import *
from tkinter import messagebox

# Get Project Parameters from .par file
def getProjParams(fdir, inpfile):
    fctr = open(fdir + inpfile, "r")
    try:
        contents = fctr.read()
        fctr.close()
    except:
        contents = {}
    try:
        proj_dict = ast.literal_eval(contents)
    except:
        proj_dict = {}
    return proj_dict

# Status message
def statusbox(label_id, msg, y=1.0):
    print(msg)
    label_id.configure(text=': '+msg, width=40, fg='#65B017')
    #label_id.pack()
    label_id.place(relx=-0.1, rely=y, anchor=SW)
    label_id.master.update()

# Status message2
def statusbox2(msg, y=0.5):
    win = Tk()
    win.geometry('300x150')
    win.geometry('+100+400')                 # Position ('+Left+Top')
    win.title('Info')

    label2 = Label(win, text=': ', width=40)
    label2.configure(text=': '+msg, width=40)
    label2.place(relx=0.5, rely=y, anchor=CENTER)
    win.mainloop()

# Function echo message
def show_message(msg, batch=False):
    print(msg)
    if not batch:
        messagebox.showinfo('Information', msg)

# Function warning message
def warn_message(msg, batch=False):
    if batch:
        print(msg)
        sys.exit(1)
    else:
        print(msg)
        messagebox.showwarning('Warning', msg)

# Polar function by giving point, angle, distance & Return 3D point(z=0)
def polar(p, a, d):
    x = p[0] + d * math.cos(a)
    y = p[1] + d * math.sin(a)
    return [x, y, 0.0]

# Distance function by giving 2D point1, point2
def distance(p, q):
    dx = p[0] - q[0]
    dy = p[1] - q[1]
    return math.sqrt(math.pow(dx, 2) + math.pow(dy, 2))

# Angle function by giving 2D point1, point2
def angle(p, q):
    dx = q[0] - p[0]
    dy = q[1] - p[1]
    return math.atan2(dy, dx)

# Compute boundary of giving Line entity & buffer
def line_bounds(e, b):
    al = e.Angle + math.pi * 0.5
    ar = e.Angle - math.pi * 0.5
    p1 = e.StartPoint
    p2 = e.EndPoint
    p11 = polar(p1, al, b)
    p12 = polar(p1, ar, b)
    p21 = polar(p2, al, b)
    p22 = polar(p2, ar, b)
    return [p11, p12, p22, p21, p11]

# Convert [[x1, y1, z1], [x2, y2, z2]] to [x1, y1, z1, x2, y2, z2] for selection boundary
def bounds2list(b):
    l = []
    for i in b:
        l += i
        """
        l.append(i[0])
        l.append(i[1])
        l.append(i[2])
        """
    return l

# To compare first element
def cmp(a, b):
    return lambda a, b: a[0] < b[0]

# Sorting list of XYZ by X element
def sort_x(lis):
    lis2 = sorted({tuple(x): x for x in lis}.values())
    return lis2

# Sorting list of points by specified element
def sort_rtk_x(lis):
    return sorted(lis, key=lambda e: e[1])
    #return lis2
