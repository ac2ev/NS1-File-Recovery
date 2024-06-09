Attribute VB_Name = "basExcelExport"
Option Explicit

'**************************************
' Name: Quickest way to export a listvie
'     w to Excel
' Description:This is a faster way to ta
'     ke a listview control and display its co
'     ntents in a new Excel workbook.
'
'A common mistake in using OLE to manipulate Excel is to send data values one cell at a time.
'However, if you are exporting listview, it is much faster to create a two-dimensional
'array of the data and then send the entire array to Excel all at once.
'This method can be applied to grids, recordsets, or any other table-like data.
'This code will also allow the user to select multiple,
'non-contiguous rows for export. Hidden columns are not exported, either.
'Also, if the ColumnHeader.Tag properties have been set to "string", "number", or
'"date", the Excel columns will be formatted as such.

' By: Brian Dunn
'
'
' Inputs:A reference to a ListView contr
'     ol.
'
' Returns:True if it worked, False if no
'     t
'
'Assumes:The listview allows multiple ro
'     w selection.
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.13733/lngWId.1/qx
'     /vb/scripts/ShowCode.htm
'for details.
'**************************************



Public Function ExportToExcel(fname As String, lvw As MSComctlLib.ListView, All As Boolean) As Boolean
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    Dim objWorksheet As Excel.Worksheet
    Dim objRange As Excel.Range
    
    Dim lngResults As Long
    Dim i As Long
    Dim intCounter As Long
    Dim intStartRow As Long
    Dim strArray() As String
    Dim intVisibleColumns() As Long
    Dim intColumns As Long
    Dim itm As ListItem
    'If there are no selected items in the l
    '     istview control



    If (lvw.SelectedItem Is Nothing And All = False) Or lvw.ListItems.count = 0 Then
        MsgBox "There aren't any items in the listview selected." _
        , vbOKOnly + vbInformation, "Export Failed"
        GoTo ExitFunction
    Else
    lngResults = vbNo
    
    End If
    'Ask the user if they want to export jus
    '     t the selected items
'    lngResults = MsgBox("Do you want to export ALL rows to Excel? " _
'    , vbYesNoCancel + vbQuestion, "Select Rows For Export")
'
'
'    If lngResults = vbCancel Then
'        GoTo ExitFunction
'    End If
'
    Screen.MousePointer = vbHourglass
    
    'Try to create an instance of Excel
    On Error Resume Next
    Set objExcel = New Excel.Application


    If Err.Number > 0 Then
        MsgBox "Microsoft Excel is not loaded on this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
        GoTo ExitFunction
    End If
    
    On Error GoTo HANDLE_ERROR
    ' Don't allow user to affect workbook
    objExcel.Interactive = False


    If objExcel.Visible = False Then
        objExcel.Visible = True
    End If
    
    objExcel.WindowState = xlMaximized
    
    Set objWorkbook = objExcel.Workbooks.Add
    
    Set objWorksheet = objWorkbook.Sheets(1)
   
    intCounter = 0
    Set objRange = objWorksheet.Rows(1)
    objRange.Font.Size = 10
    objRange.Font.Bold = True


    For i = 1 To lvw.ColumnHeaders.count


        If lvw.ColumnHeaders(i).Width <> 0 Then
            ' Create an array of visible column inde
            '     xes
            intColumns = intColumns + 1
            ReDim Preserve intVisibleColumns(1 To intColumns)
            intVisibleColumns(intColumns) = i
            objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text


            With objWorksheet.Columns(intColumns)


                Select Case LCase$(lvw.ColumnHeaders(i).Tag)
                    ' If tag is empty, format as text
                    Case "string", ""
                    .NumberFormat = "@"
                    Case "number"
                    .NumberFormat = "#,##0.00_);(#,##0.00)"
                    .HorizontalAlignment = xlRight
                    Case "date"
                    .NumberFormat = "mm/dd/yyyy"
                    .HorizontalAlignment = xlRight
                End Select

        End With
        
    End If
   
Next i
' Dimension array to number of listitems
'
ReDim strArray(1 To lvw.ListItems.count, 1 To intColumns)

intCounter = 0
intStartRow = 2


For Each itm In lvw.ListItems
    ' A response of vbNo meant to export all
    '     the items


    If lngResults = vbNo Or itm.Selected Then
        ' increment the number of selected rows
        intCounter = intCounter + 1


        For i = 1 To intColumns


            If intVisibleColumns(i) = 1 Then
                strArray(intCounter, 1) = itm.Text
            Else
                strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
            End If
        Next i
    End If
Next itm

' Send entire array to Excel range


With objWorksheet
    .Range(.Cells(2, 1), _
    .Cells(2 + intCounter - 1, intColumns)) = strArray
End With
    objExcel.Sheets(1).Name = "NS1 Export"
    With objExcel.Sheets(1).PageSetup
        .LeftHeader = ""
        .CenterHeader = "Kismet To Ns1 Conversion"
        .RightHeader = "&D&T" 'Date Time
        .LeftFooter = ""
        .CenterFooter = "NetStumbler File: " & fname
        .RightFooter = ""
        .Orientation = 2 'xlLandscape
        .PrintGridlines = True
    End With
objWorksheet.Columns.AutoFit
objExcel.Interactive = True
objWorkbook.SaveAs fname
'objExcel.ThisWorkbook.SaveAs fname
ExportToExcel = True
ExitFunction:
Screen.MousePointer = vbDefault
Exit Function
HANDLE_ERROR:
MsgBox "Export to Excel failed. Encountered thej following Error" & vbCrLf & vbCrLf & _
Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error Exporting To Excel"
Resume
Set objRange = Nothing
Set objWorksheet = Nothing
Set objWorkbook = Nothing
objExcel.Quit
GoTo ExitFunction
End Function

        




