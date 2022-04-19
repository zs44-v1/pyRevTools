# -*- coding: utf-8 -*-
__title__ = "Sheets from Excel"
__author__ = "Hanif Jeshani"
__doc__ = """ Create sheets in Revit from an Excel file. """

import clr

clr.AddReference('Microsoft.Office.Interop.Excel')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference("System.Windows.Forms")
clr.AddReference('System.Runtime.InteropServices')

from Autodesk.Revit.DB import *
from System import GC
from System.Runtime.InteropServices import Marshal
from Microsoft.Office.Interop import Excel
from pyrevit import forms, script
from Autodesk.Revit.UI import TaskDialog

form_title = __title__

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
selfdestruct_timer = 10

def main():
    TaskDialog.Show(form_title, "Function not implemented")


if __name__ == '__main__':
    main()
