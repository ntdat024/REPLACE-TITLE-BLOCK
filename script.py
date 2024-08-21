#region library
import clr 
import os
import sys

clr.AddReference("System")
clr.AddReference("System.Data")
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference("System.Windows.Forms")

import math
import System
import RevitServices
import Autodesk
import Autodesk.Revit
import Autodesk.Revit.DB

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *

from System.Collections.Generic import *
from System.Windows import MessageBox, RoutedEventHandler
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Markup import XamlReader
from System.Windows.Controls import Button, ComboBox, TextBox

#endregion

#region revit infor
# Get the directory path of the script.py & the Window.xaml
dir_path = os.path.dirname(os.path.realpath(__file__))
xaml_file_path = os.path.join(dir_path, "Window.xaml")

#Get UIDocument, Document, UIApplication, Application
uidoc = __revit__.ActiveUIDocument
uiapp = UIApplication(uidoc.Document.Application)
app = uiapp.Application
doc = uidoc.Document
activeView = doc.ActiveView
#endregion

#region method
class Utils:
    def __init__(self):
        pass

    def get_all_title_blocks(self):
        collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
        names = []
        for ele in collector:
            name = ele.FamilyName + ": " + ele.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            names.append(name)
        names.sort()
        return names
    
    def get_all_view_sets(self):
        view_set = ["<Sheets in Model>"]
        collector = FilteredElementCollector(doc).OfClass(ViewSheetSet)
        for vs in collector:
            if vs.Views.Size > 0:
                view_set.append(vs.Name)
        return view_set
    
    def get_sheet_full_name(self, sheet):
        return sheet.SheetNumber + " - " + sheet.Name
        
        
    # end def
    def get_all_sheet_in_model (self):
        collector = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements()
        sheets = sorted(list(collector), key=lambda vs: vs.SheetNumber)
        sheet_names = []
        for vs in sheets:
            full_name = self.get_sheet_full_name(vs)
            sheet_names.append(full_name)
        return sheet_names
    
    
    def get_sheet_by_view_set (self, viewset_name):

        if str(viewset_name).__contains__("<"):
            return self.get_all_sheet_in_model()
        else:
            collector = FilteredElementCollector(doc).OfClass(ViewSheetSet)
            view_set = None
            for vs in collector:
                if vs.Name == viewset_name:
                    view_set = vs
                    break
            
            sheet_in_view_set = []
            if view_set is not None:
                for sheet in view_set.Views:
                    if str(sheet).__contains__("DB.ViewSheet"):
                        sheet_in_view_set.append(sheet)
                        
            
            sheet_names = []
            sheet_in_view_set = sorted(sheet_in_view_set, key=lambda vs: vs.SheetNumber)
            for sheet in sheet_in_view_set:
                full_name = self.get_sheet_full_name(sheet)
                sheet_names.append(full_name)

            return sheet_names
    
    def get_title_block_by_name (self, title_block_name):
        collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().ToElements()
        for ele in collector:
            name = ele.FamilyName + ": " + ele.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
            if title_block_name == name:
                return ele
        return None
    
    def get_sheet_element_by_name (self, list_sheet_names):
        collector = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements()
        sheet_elements = []
        for vs in collector:
            full_name = self.get_sheet_full_name(vs)
            if list(list_sheet_names).__contains__(full_name):
                sheet_elements.append(vs)
        return sheet_elements
        
    def replace_title_block (self, list_sheet_names, title_block_name):
        try:
            t = Transaction(doc, " ")
            t.Start()

            collector = list(FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements())
            title_block = self.get_title_block_by_name(title_block_name)
            for ele in collector:
                sheet = doc.GetElement(ele.OwnerViewId)
                full_name = self.get_sheet_full_name(sheet)
                if list(list_sheet_names).__contains__(full_name):
                    if ele.Symbol.Id != title_block.Id:
                        ele.Symbol = title_block
                        
            t.Commit()
            MessageBox.Show("Completed!","Message")
            
        except Exception as e:
            MessageBox.Show(str(e), "Message")
        
                


#region defind window
class WPFWindow:

    def load_window (self):
        
        #import window from .xaml file path
        file_stream = FileStream(xaml_file_path, FileMode.Open, FileAccess.Read)
        window = XamlReader.Load(file_stream)

        #controls
        self.cbb_SheetSet = window.FindName("cbb_SheetSet")
        self.cbb_TitleBlock = window.FindName("cbb_TitleBlock")
        self.tb_Filter = window.FindName("tb_Filter")
        self.lbx_Sheets = window.FindName("lbx_Sheets")
        self.tb_Rotate = window.FindName("tbRotate")
        self.bt_OK = window.FindName("bt_OK")
        self.bt_Cancel = window.FindName("bt_Cancel")

        #binding data
        self.binding_data()
        self.window = window

        return window


    def binding_data (self):

        self.cbb_TitleBlock.ItemsSource = Utils().get_all_title_blocks()
        self.cbb_SheetSet.ItemsSource = Utils().get_all_view_sets()
        self.cbb_TitleBlock.SelectedIndex = 0
        self.cbb_SheetSet.SelectedIndex = 0

        self.original_data = Utils().get_all_sheet_in_model()
        self.lbx_Sheets.ItemsSource = self.original_data

        #button
        self.bt_Cancel.Click += self.cancel_click
        self.bt_OK.Click += self.ok_click
        self.cbb_SheetSet.SelectionChanged += self.sheet_set_changed
        self.tb_Filter.TextChanged += self.tb_filter_Changed


    def ok_click(self, sender, e):
        
        list_sheet_names = self.lbx_Sheets.SelectedItems
        title_block_name =self.cbb_TitleBlock.SelectedItem
        if len(list_sheet_names) > 0:
            Utils().replace_title_block(list_sheet_names, title_block_name)
            self.window.Close()
            MessageBox.Show("Completed!", "Message")
        else: MessageBox.Show("Please choose sheets!", "Message")
        
        

    def cancel_click (self, sender, e):
        self.window.Close()
    
    def sheet_set_changed (self, sender, e):
        view_set_name = self.cbb_SheetSet.SelectedItem
        self.original_data = Utils().get_sheet_by_view_set(view_set_name)
        self.lbx_Sheets.ItemsSource = self.original_data
        self.tb_Filter.Text = ""
    
    def tb_filter_Changed (self, sender, e):
        filter_name = str(self.tb_Filter.Text)
        new_list = []
        if filter_name is not None or filter_name != "":
            name = filter_name.lower()
            for item_Name in self.original_data:
                if str(item_Name).lower().__contains__(name):
                    new_list.append(item_Name)
            self.lbx_Sheets.ItemsSource = new_list
        else: self.lbx_Sheets.ItemsSource = self.original_data

#endregion

def main_task():
    try:
        window = WPFWindow().load_window()
        window.ShowDialog()
    except Exception as e:
        MessageBox.Show(str(e), "Message")

if __name__ == "__main__":
    main_task()
        







