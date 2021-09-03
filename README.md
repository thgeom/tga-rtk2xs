# tga-rtk2xs
# This program shall be used for rtk data file in csv format into AutoCAD drawing.
# The X-section line can be created with extended data in the AutoCAD
# From rtk points and X-section line, the program shall be extract X-section points to CSV file in X-section format
# To procees that rtk data, the parameter file is required for ACAD drawing setting as shown:
{
"WorkDirectory" : "d:/TGA_TEST/RTK/",
"RTKDatatFile" : "RTK_X-sec.csv",
"DrawingFile" : "RTK_Point.dwg",
"OutputCsvFile" : "xsec-0.csv",
"OutputXlsFile" : "xsec-0.xlsx",
"CadApp" : "rtk_xs",
"XSLineLayer" : "XS_Line",
"ChainageLayer" : "CHN_Layer",
"XSCodeLayer" : "XS_Code",
"XSNameLayer" : 'XS_Name',
"XSPointLayer" : 'XS_point',
"CompletedColor"  :  5,
"XSLineCompletedLayer" : "XS_Line_Completed",
"Buffer" : 10
}
