# -*- coding: utf-8 -*-
__title__ = "Workset to Excel"
__author__ = "Hanif Jeshani"
__doc__ = """ List Revit worksets in an Excel file. """

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('Microsoft.Office.Interop.Excel')
clr.AddReference('System.Runtime.InteropServices')

# clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import *
from Microsoft.Office.Interop import Excel
from System import GC
from System.Runtime.InteropServices import Marshal
from Autodesk.Revit.UI import TaskDialog
from pyrevit import forms

# from Autodesk.Revit.UI import UIApplication

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

form_title = __title__


def main():
    if doc.IsWorkshared:
        worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)

        worksetlist = []
        for workset in worksets:
            worksetlist.append([workset.Name])

        flat_workset_list = []
        for elem in worksetlist:
            flat_workset_list.extend(elem)

        xl_app = Excel.ApplicationClass()
        xl_app.DisplayAlerts = False
        xl_book = xl_app.Workbooks.Add()
        xl_sheet = xl_book.ActiveSheet
        xl_sheet.Name = "Worksets"

        for xllpos, lstitem in enumerate(flat_workset_list):
            xl_sheet.Cells[xllpos + 1, 1].Value = lstitem
        worksetrange = xl_sheet.Range[xl_sheet.Cells[1, 1], xl_sheet.Cells[xllpos + 1, 1]]
        # worksetrange.RefersToR1C1 = '=OFFSET(Worksets!R1C1,0,0,COUNTA(Worksets!C1),1)'
        xl_sheet.Names.Add("worksets", worksetrange)

        excel_file = forms.save_excel_file("Save worksets")
        if excel_file:
            # xlApp.Visible = True
            xl_book.SaveAs(excel_file)
        xl_book.Close(False)
        xl_app.Quit()
        Marshal.ReleaseComObject(xl_sheet)
        Marshal.ReleaseComObject(xl_book)
        Marshal.ReleaseComObject(xl_app)
        GC.Collect()
        GC.WaitForPendingFinalizers()

    else:
        # TaskMsg = TaskDialog("Get worksets")
        # TaskMsg.MainInstruction = "This model is not workshared"
        # TaskMsg.MainContent = "Ending script"
        # TaskMsg.Show()
        forms.alert("This model is not workshared", title=form_title, exitscript=True)


if __name__ == '__main__':
    main()
