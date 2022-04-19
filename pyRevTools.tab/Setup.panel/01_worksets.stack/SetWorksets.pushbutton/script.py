# -*- coding: utf-8 -*-
__title__ = "Worksets from Excel"
__author__ = "Hanif Jeshani"
__doc__ = """ Create worksets in Revit from an Excel file. """

import clr

clr.AddReference('Microsoft.Office.Interop.Excel')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
# clr.AddReference("System.Windows.Forms")
clr.AddReference('System.Runtime.InteropServices')

from Autodesk.Revit.DB import *
from System import GC
from System.Runtime.InteropServices import Marshal
from Microsoft.Office.Interop import Excel
from pyrevit import forms, script

# from Autodesk.Revit.UI import TaskDialog

form_title = __title__

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
selfdestruct_timer = 10


def main():
    wks_grid_levels = "Shared Levels and Grids"
    wks_default = "Workset1"
    if forms.check_modeldoc():
        # TaskDialog.Show(tskd_title, "Select the Excel file")
        open_file = forms.pick_excel_file(title="Select a file")
        if open_file:
            openfilename = open_file
            xl_app = Excel.ApplicationClass()
            xl_book = xl_app.Workbooks.Open(openfilename)
            xl_sheet = xl_book.Worksheets.Item('Worksets')
            # xlApp.Visible = True
            worksetrange = xl_sheet.Names.Item("worksets")
            wkslist = []
            if worksetrange.RefersToRange.Count > 1:
                for wks in worksetrange.RefersToRange.Cells.Value2:
                    wkslist.append(wks)
            else:
                wkslist.append(worksetrange.RefersToRange.Cells.Value2)
            xl_book.Close(False)
            xl_app.Quit()
            # Marshal.ReleaseComObject(xl_sheet)
            Marshal.ReleaseComObject(xl_book)
            Marshal.ReleaseComObject(xl_app)
            output_window = script.get_output()
            print("Creating worksets")
            if not doc.IsWorkshared and doc.CanEnableWorksharing:
                if len(wkslist) > 1:
                    wks_default = wkslist.pop(0)
                    wks_grid_levels = wkslist.pop(0)
                doc.EnableWorksharing(wks_grid_levels, wks_default)
            t = Transaction(doc, 'Create worksets')
            t.Start()
            for wks in wkslist:
                if WorksetTable.IsWorksetNameUnique(doc, wks):
                    Workset.Create(doc, wks)
                else:
                    output_window.log_warning("Duplicate worksets")
                    print(wks + " already exists")
                    # forms.alert(wks+" already exists", title= tskd_title)
                    # TaskDialog.Show(tskd_title, wks + " already exists")
            t.Commit()
            output_window.self_destruct(selfdestruct_timer)
        else:
            forms.alert("Could not open the file", title=form_title, exitscript=True)
    GC.Collect()
    GC.WaitForPendingFinalizers()


if __name__ == '__main__':
    main()
