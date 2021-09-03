import pandas as pd

#top = Tk()

# Class of Field data from CSV file
class FieldDataCSV:
    def __init__(self,fname):
        #self.fdir = fdir
        #self.fname = fname
        #self.pathname = fdir + fname
        self.pathname = fname

    def getdata(self):
        self.restab = pd.read_csv(self.pathname, encoding='ANSI',
                                  usecols=['Code','Name','N','E','Z'])[['Code','Name','E','N','Z']]

    def show_data(self):
        print(self.restab.head(100))


# Class of Field data from Excel file
class FieldDataXLS:
    def __init__(self, fdir, fname, shname):
        self.fdir = fdir
        self.fname = fname
        self.pathname = fdir + fname
        self.shname = shname

    def getdata(self):
        self.restab = pd.read_excel(self.pathname, sheet_name=self.shname, index_col=None,
                                    dtype={'Name': str, 'East': float, 'North': float, 'Elev': float, 'Code': str})

    def show_data(self):
        for row in self.restab.values:
            print(row)

# Get data from Excel file
def getXLS():
    #print('Test FieldDataXLS')
    xls1 = FieldDataXLS('d:/TGA_Lisp/', 'pt_list-r4.xlsx', 'XYZ')
    xls1.getdata()
    xls1.show_data()
    return xls1

# Get data from CSV file
def getRTK(fdir, rtkfile):
    rtk = FieldDataCSV(fdir+rtkfile)
    rtk.getdata()                                                       # Get data from csv file
    rtk.show_data()
    return rtk
