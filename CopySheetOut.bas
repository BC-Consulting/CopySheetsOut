Sub CopySheetOut()
' This macro is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This macro is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' Copy of the GNU General Public License is availabel at:
'           http://www.gnu.org/licenses/
'
' ©2019 - GeoProc.com
'----------------------------------------------------------------------

' Define variables
dim document   as object
dim oDoc       as object
dim dispatcher as object

dim args1(0) as new com.sun.star.beans.PropertyValue
dim args2(2) as new com.sun.star.beans.PropertyValue
dim args3(1) as new com.sun.star.beans.PropertyValue

'Output formats
dim types(4)
dim extns(4)
   types(0) = "MS Excel 97"                 ' MS Excel .xls
   types(1) = "Calc MS Excel 2007 XML"      ' MS Excel .xlsx
   types(2) = "calc8"                       ' LibreOffice/OO .ods
   types(3) = "Text - txt - csv (StarCalc)" '.csv file
   extns() = ".xls"
   extns() = ".xlsx"
   extns() = ".ods"
   extns() = ".csv"

' - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *
'Start customisation

'             Change variables below before runing the macro

   '----------------------------------------------------------------------
   'Full name of the folder to save the documents to.
   ' On Windows use / instead of \ to define the folder path.
   ' Terminate the path with /
   root = "C:/Your/Output/folder/here/"
   loca = "another/sub-folder/"
   ' We use a special naming convention: root/sheet_name/loca/sheet_name.extn
   '----------------------------------------------------------------------
   ' The output name will be the sheet name with some changes
   ' An optional suffix for the filename
   suff = ""
   '----------------------------------------------------------------------
   ' Output file format. See types() array for definition
   extn = 0

' End customisation
' - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - * - *

   '----------------------------------------------------------------------
   oDoc       = ThisComponent
   document   = ThisComponent.CurrentController.Frame
   dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

   ' Loop through all sheets and ...
   nSheets = oDoc.Sheets.Count
   for i = nSheets to 1 step -1
      '----------------------------------------------------------------------
      ' Generate output filename
      oSheet   = oDoc.Sheets(i-1)
      Sht_Name = oSheet.GetName
      If Sht_Name = "Sheet1" Then Sht_Name = "Sheet1New"
      ' Make some changes to the filename...
      Sht_Name = Replace(Sht_Name, " ", "")
      Sht_Name = Replace(Sht_Name, "-", "")
      Sht_Name = Replace(Sht_Name, "_", "")
      filename = Sht_Name & suff & extns(extn)
      ' Full path to the new file
      dire     = root & Sht_Name & "/" & loca
      '----------------------------------------------------------------------
      ' Create a Calc document
      oDocN    = StarDesktop.loadComponentFromURL( "private:factory/scalc", "_blank", 0, Array() )
      oDocNF   = oDocN.CurrentController.Frame
      ' Save new spreadheet
      If extn = 3 Then
        ' For CSV we need to provide more info...
         args2(0).Name  = "URL"
         args2(0).Value = "file:///" & dire & filename
         args2(1).Name  = "FilterName"
         args2(1).Value = types(extn)
         args2(2).Name = "FilterOptions"
         args2(2).Value = "44,34,76,1,,0,false,true,false,false,false"  'See readme.md for options
         dispatcher.executeDispatch(oDocNF, ".uno:SaveAs", "", 0, args2())
      Else
         args3(0).Name  = "URL"
         args3(0).Value = "file:///" & dire & filename
         args3(1).Name  = "FilterName"
         args3(1).Value = types(extn)
         dispatcher.executeDispatch(oDocNF, ".uno:SaveAs", "", 0, args3())
      End If

      '----------------------------------------------------------------------
      ' Activate sheet i
      args1(0).Name  = "Nr"
      args1(0).Value = i
      dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args1())
      '----------------------------------------------------------------------
      ' Copy sheet i to new spreadsheet file
      args2(0).Name  = "DocName"
      args2(0).Value = Sht_Name & suff
      args2(1).Name  = "Index"
      args2(1).Value = 1
      args2(2).Name  = "Copy"
      args2(2).Value = true
      dispatcher.executeDispatch(document, ".uno:Move", "", 0, args2())

      '----------------------------------------------------------------------
      ' Delete "Sheet1" from new document
      oShts = oDocN.Sheets
      oShts.removeByName("Sheet1")
      '----------------------------------------------------------------------
      ' Save & close new document
      dispatcher.executeDispatch(oDocNF, ".uno:Save", "", 0, Array())
      oDocN.close(true)
   next 'sheet
   print("done")

End Sub
