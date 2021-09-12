# tga-rtk2xs
# This program shall be used for RTK data file in CSV format into AutoCAD drawing.
# The X-section line can be created with extended data in the AutoCAD
# From RTK points and X-section line, the program shall be extract X-section points to CSV file in X-section format
# To procees that RTK data, the parameter file is required for setting up the enveronment
# :pparams.par is an example file as shown:
{
"WorkDirectory" : "d:/TGA_TEST/RTK/",
"RTKDatatFile" : "RTK_X-sec.csv",
"RTKColumns" : ["Code", "Name", "N", "E", "Z"],
"RTKEncoding" : "TIS-620",
"DrawingFile" : "RTK_Point.dwg",
"OutputCsvFile" : "xsec-0.csv",
"OutputXlsFile" : "xsec-0.xlsx",
"CadApp" : "rtk_xs",
"XSLineLayer" : "XS_Line",
"ChainageLayer" : "CHN_Layer",
"XSCodeLayer" : "XS_Code",
"XSNameLayer" : "XS_Name",
"XSPointLayer" : "XS_point",
"XSComputedPointLayer" : "XS_Computed",
"DrawXSComputedPoint" : True,
"CompletedColor"  :  5,
"XSLineCompletedLayer" : "XS_Line_Completed",
"Buffer" : 10
}
# pandas, win32com are the library to connect with CSV, Excel and AutoCAD
# tkinter shall be used for GUI
