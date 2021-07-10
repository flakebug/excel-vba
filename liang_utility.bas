Attribute VB_Name = "liang_utility"
Sub test()
    Dim n As String
    n = GetOpenFilename_Single("Excel (*.xlsx; *.xlsm),*.xlsx;*.xlsm", "Title")
    MsgBox GetExtensionWithoutFilenameFromPath(n)
End Sub

Private Sub FlatWorksheet(wks As Worksheet)
    'Author : Liang
    'Initial : 2021/7/10
    'Last update : 2021/7/10
    'Description : copy all sheet and paste in values
    wks.Activate
    wks.Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    wks.Cells(1, 1).Select
End Sub

Private Function GetExtensionWithoutFilenameFromPath(path As String) As String
    'Author : Liang
    'Initial : 2021/7/10
    'Last update : 2021/7/10
    GetExtensionWithoutFilenameFromPath = Right(path, Len(path) - InStrRev(path, "."))
End Function

Private Function GetFilenameWithoutExtensionFromPath(path As String) As String
    'Author : Liang
    'Initial : 2021/7/10
    'Last update : 2021/7/10
    Dim filename As String
    filename = GetFilenameFromPath(path)
    GetFilenameWithoutExtensionFromPath = Left(filename, InStr(filename, ".") - 1)
End Function

Private Function GetDirectoryFromPath(path As String) As String
    'https://stackoverflow.com/questions/418622/find-the-directory-part-minus-the-filename-of-a-full-path-in-access-97
   GetDirectoryFromPath = Left(path, InStrRev(path, Application.PathSeparator))
End Function

Private Function GetFilenameFromPath(ByVal strPath As String) As String
    'https://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path
    ' Returns the rightmost characters of a string upto but not including the rightmost '\'
    ' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function


Private Function Execution_Confirmation(procedure_name As String) As Boolean
    'Author : Liang
    'Initial : 2021/7/10
    'Last update : 2021/7/10
    'Description : show confirmation before execute subroutine (or function)
    Dim exec As Integer
    Dim result As Boolean
    exec = MsgBox(procedure_name & vbCrLf & "Proceed?", vbOKCancel + vbQuestion + vbDefaultButton2)
    If exec = vbCancel Then
        MsgBox "Operation cancelled", vbInformation
        result = False
    Else
        result = True
    End If
    Execution_Confirmation = result
End Function

Private Function GetOpenFilename_Single(FileFilter As String, Title As String) As String
    'Author : Liang
    'Initial : 2021/7/10
    'Last update : 2021/7/10
    'Description : show open file dialog and return string array, single select
    'Usage :
    'FileFilter example : "Excel (*.xlsx; *.xlsm),*.xlsx;*.xlsm"
    Dim result As Variant
    result = Application.GetOpenFilename(FileFilter:=FileFilter, Title:=Title, MultiSelect:=False)
    If result = False Then
        GetOpenFilename_Single = ""
    Else
        GetOpenFilename_Single = result
    End If
End Function


Private Function GetOpenFilename_Multiple(FileFilter As String, Title As String) As String()
    'Author : Liang
    'Initial : 2021/7/10
    'Last update : 2021/7/10
    'Description : show open file dialog and return string array, multiple select
    'Usage :
    'FileFilter example : "Excel (*.xlsx; *.xlsm),*.xlsx;*.xlsm"
    Dim files() As Variant
    Dim file As Variant
    Dim file_count As Integer
    Dim file_string_array() As String
    
    files = Application.GetOpenFilename(FileFilter:=FileFilter, Title:=Title, MultiSelect:=True)
    file_count = UBound(files)
    ReDim file_string_array(file_count)
    
    Dim indx As Integer
    For indx = 1 To file_count
        file_string_array(indx) = files(indx)
    Next
    Debug.Print UBound(files)
    GetOpenFilename_Multiple = file_string_array
End Function

Private Sub Set_ProgressBar(Total As Long, Current As Long, Optional Text As String)
    'Author : Liang
    'Initial : 2021/7/9
    'Last update : 2021/7/9
    'Description : progress bar
    frmProgress.Caption = Text
    Dim step As Single
    step = 200 / Total
    Dim barWidth As Single
    barWidth = Current * step
    frmProgress.bar.Width = barWidth
End Sub

Private Sub Init_ProgressBar()
    'Author : Liang
    'Initial : 2021/7/9
    'Last update : 2021/7/9
    'Description : progress bar
    frmProgress.bar.Width = 0
    frmProgress.Caption = ""
    frmProgress.Show
    frmProgress.Left = 100
    frmProgress.Top = 100
End Sub

Private Sub Close_ProgressBar()
    'Author : Liang
    'Initial : 2021/7/9
    'Last update : 2021/7/9
    'Description : progress bar
    frmProgress.bar.Width = 0
    frmProgress.Caption = ""
    frmProgress.Hide
End Sub

Private Function dirExists(s_directory As String) As Boolean
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    dirExists = oFSO.FolderExists(s_directory)
End Function



Private Sub Beta_FillEmptyCellsWithPreviousRowValue()
    'Author : Liang
    'Initial : 2021/7/7
    'Latest update : 2021/7/7
    'Description : if current cell is empty, then copy from previous cell
    'Usage : select a column of cells, the first row must contains valid value
    '           program will search the entire selection and fill with values
    Dim cl As Range
    Dim tval As String
    For Each cl In Selection
        If cl.Value = "" Then
            cl.Value = tval
        Else
            tval = cl.Value
        End If
    Next
End Sub






Sub S01_DirectoryAvailibilityCheck()
    'Author : Liang
    'Initial : 2021/7/3
    'Last update : 2021/7/5
    'Description : verify the path availability
    'Usage : select the cells that you want to perform path check, and execute program
    On Error Resume Next
    
    If Not Execution_Confirmation("S01_DirectoryAvailibilityCheck") Then
        Exit Sub
    End If
    
    Dim FilePath As String
    Dim strFolderName As String
    Dim strFolderExists As String
    Dim cell As Range
    Dim current_indx As Long
    Dim cell_count As Long
    cell_count = Selection.Cells.Count
    Init_ProgressBar
    For Each cell In Selection
        strFolderExists = Dir(cell.Value, vbDirectory)
        If strFolderExists = "" Then
              cell.Interior.ColorIndex = 3
        End If
        Set_ProgressBar cell_count, current_indx, "S01_DirectoryAvailibilityCheck"
        current_indx = current_indx + 1
        DoEvents
    Next
    Close_ProgressBar
    MsgBox "Done", vbInformation
End Sub


Public Sub S02_Flat_Excel_Workbooks()
    'Author : Liang
    'Initial : 2021/7/5
    'Last update : 2021/7/5
    'Description : transform excel formula cell to text cell, reduce the size and complexity
    'Usage : this subroutine only transform the cell for the filename consist with "_flatted"
    If Not Execution_Confirmation("S02_Flat_Excel_Workbooks") Then
        Exit Sub
    End If
    
    Dim xlsx_filename As String
    xlsx_filename = GetOpenFilename_Single("Excel (*.xlsx; *.xlsm),*.xlsx;*.xlsm", "Open Excel File")
    If xlsx_filename = "" Then
        Exit Sub
    End If
    
    Dim wkb As Workbook
    Set wkb = Workbooks.Open(xlsx_filename)
    Dim wks As Worksheet

    For Each wks In wkb.Worksheets
        FlatWorksheet wks
    Next
    
    Dim path_str As String
    Dim filename_str As String
    Dim extension_str As String
    path_str = GetDirectoryFromPath(xlsx_filename)
    filename_str = GetFilenameWithoutExtensionFromPath(xlsx_filename)
    extension_str = GetExtensionWithoutFilenameFromPath(xlsx_filename)
    
    wkb.SaveAs path_str & filename_str & "_flatted.xlsb", FileFormat:=xlExcel12
    wkb.Close
    MsgBox "Done", vbInformation
End Sub




Private Sub S03_EDC060_Link_Generator()
    'Author : Liang
    'Initial : 2021/7/6
    'Last update : 2021/7/9
    'Description : generate server link for EDC060.xlsm
    'Usage :
    
    Dim Execution_Confirmation As Integer
    Const subroutine_name As String = "S03_EDC060_Link_Generator"
    Execution_Confirmation = MsgBox(subroutine_name & vbCrLf & "Proceed?", vbOKCancel + vbQuestion + vbDefaultButton2)
    If Execution_Confirmation = vbCancel Then
        MsgBox "Operation cancelled", vbInformation
        Exit Sub
    End If
    
    Dim wkb As Workbook
    'Set wkb = Workbooks.Open("\\192.168.198.4\filesrv\B-Master Drawing\B-01-SPF Control Log\EDC060.xlsm")
    Set wkb = Workbooks.Open("d:\temp\EDC060.xlsm")
   
    Dim wks As Worksheet
    Set wks = wkb.Worksheets("Report")
    
    Dim lastrow As Long
    lastrow = wks.Range("A" & Rows.Count).End(xlUp).Row

    Dim report_date As Date
    report_date = wks.Range("U1")
    Dim report_date_str As String
    report_date_str = Format(report_date, "yyyymmdd")

    Dim row_indx As Long
    Dim clvalue As String
    Dim rank As String
    Init_ProgressBar
    For row_indx = 5 To lastrow
        rank = Cells(row_indx, 21).Value
    
        'transmit to client
        clvalue = Cells(row_indx, 13).Value
        If (clvalue <> "") Then
            wks.Hyperlinks.Add Anchor:=wks.Cells(row_indx, 13), Address:= _
                "\\192.168.198.4\filesrv\C-Correspondence\C-02-Transmittal\OUT\" & clvalue _
                , TextToDisplay:=clvalue
        End If
        
        'reply to client
        clvalue = Cells(row_indx, 15).Value
        If (clvalue <> "") Then
            wks.Hyperlinks.Add Anchor:=wks.Cells(row_indx, 15), Address:= _
                "\\192.168.198.4\filesrv\C-Correspondence\C-02-Transmittal\IN\" & clvalue _
                , TextToDisplay:=clvalue
        End If
        
       
        If rank > 1 Then
            wks.Cells(row_indx, 13).Font.Color = vbRed
            wks.Cells(row_indx, 15).Font.Color = vbRed
        End If
        
        If row_indx Mod 1000 = 0 Then
            Set_ProgressBar lastrow, row_indx, "S03_EDC060_Link_Generator"
            DoEvents
        End If
        If row_indx = lastrow Then
            Set_ProgressBar lastrow, lastrow, "S03_EDC060_Link_Generator"
            DoEvents
        End If
    Next
    Set_ProgressBar lastrow, row_indx
    wks.Range("E2").Value = "Liang, " & Date
    FlatWorksheet wks
    wks.Range("$A$4:$U$" & lastrow).AutoFilter Field:=21, Criteria1:="1"
    
    Application.DisplayAlerts = False
    wkb.Worksheets("Setting").Delete
    wkb.Worksheets("QueryParam").Delete
    wkb.Worksheets("Data1").Delete
    wkb.Worksheets("Data2").Delete
    'wkb.SaveAs "\\192.168.198.4\filesrv\X-File Exchange\Liang\EI Report\XR01_EDC060\EDC060_" & report_date_str & ".xlsb", FileFormat:=xlExcel12
    wkb.SaveAs "d:\temp\EDC060_" & report_date_str & ".xlsb", FileFormat:=xlExcel12
    Application.DisplayAlerts = True
    wkb.Close
    Close_ProgressBar
    MsgBox "Done, " & lastrow & " rows processed", vbInformation
    

End Sub



Private Sub S04_VDC050_Link_Generator()
    'Author : Liang
    'Initial : 2021/7/9
    'Last update : 2021/7/9
    'Description : generate server link for VDC050.xlsm
    'Usage :
    
    Dim Execution_Confirmation As Integer
    Const subroutine_name As String = "S04_VDC050_Link_Generator"
    Execution_Confirmation = MsgBox(subroutine_name & vbCrLf & "Proceed?", vbOKCancel + vbQuestion + vbDefaultButton2)
    If Execution_Confirmation = vbCancel Then
        MsgBox "Operation cancelled", vbInformation
        Exit Sub
    End If
    
    Dim wkb As Workbook
    'Set wkb = Workbooks.Open("\\192.168.198.4\filesrv\B-Master Drawing\B-01-SPF Control Log\EDC060.xlsm")
    Set wkb = Workbooks.Open("d:\temp\VDC050.xlsm")
   
    Dim wks As Worksheet
    Set wks = wkb.Worksheets("Report")
    
    Dim lastrow As Long
    lastrow = wks.Range("A" & Rows.Count).End(xlUp).Row

    Dim report_date As Date
    report_date = wks.Range("AG1")
    Dim report_date_str As String
    report_date_str = Format(report_date, "yyyymmdd")

    Dim row_indx As Long
    Dim clvalue As String
    Dim rank As String
    Init_ProgressBar
    For row_indx = 5 To lastrow
        rank = Cells(row_indx, 33).Value
        po_no = Cells(row_indx, 7).Value
    
        'receive from vendor
        clvalue = Cells(row_indx, 16).Value
        If (clvalue <> "") Then
            wks.Hyperlinks.Add Anchor:=wks.Cells(row_indx, 16), Address:= _
                "\\192.168.198.4\filesrv\B-Master Drawing\B-09-Vendor Document (By PO)\" & po_no & "\From Vendor\" & clvalue _
                , TextToDisplay:=clvalue
        End If
        
        'squad check transmittal
        'clvalue = Cells(row_indx, 19).Value
        'If (clvalue <> "") Then
        '    wks.Hyperlinks.Add Anchor:=wks.Cells(row_indx, 19), Address:= _
        '        "\\192.168.198.4\filesrv\C-Correspondence\C-02-Transmittal\OUT\" & clvalue _
        '        , TextToDisplay:=clvalue
        'End If
        
        'reply to vendor
        clvalue = Cells(row_indx, 22).Value
        If (clvalue <> "") Then
            wks.Hyperlinks.Add Anchor:=wks.Cells(row_indx, 22), Address:= _
                "\\192.168.198.4\filesrv\B-Master Drawing\B-09-Vendor Document (By PO)\" & po_no & "\To Vendor\" & clvalue _
                , TextToDisplay:=clvalue
        End If
                
        
        'transmit to client
        clvalue = Cells(row_indx, 25).Value
        If (clvalue <> "") Then
            wks.Hyperlinks.Add Anchor:=wks.Cells(row_indx, 25), Address:= _
                "\\192.168.198.4\filesrv\C-Correspondence\C-08-Vendor Transmittal\OUT\" & clvalue _
                , TextToDisplay:=clvalue
        End If
        
        'reply from client
        'clvalue = Cells(row_indx, 29).Value
        'If (clvalue <> "") Then
        '    wks.Hyperlinks.Add Anchor:=wks.Cells(row_indx, 29), Address:= _
        '        "\\192.168.198.4\filesrv\C-Correspondence\C-02-Transmittal\OUT\" & clvalue _
        '        , TextToDisplay:=clvalue
        'End If
        
        If rank > 1 Then
            wks.Cells(row_indx, 16).Font.Color = vbRed
            'wks.Cells(row_indx, 19).Font.Color = vbRed
            wks.Cells(row_indx, 22).Font.Color = vbRed
            wks.Cells(row_indx, 25).Font.Color = vbRed
            'wks.Cells(row_indx, 29).Font.Color = vbRed
        End If
        
        If row_indx Mod 1000 = 0 Then
            Set_ProgressBar lastrow, row_indx, "S04_VDC050_Link_Generator"
            DoEvents
        End If
        If row_indx = lastrow Then
            Set_ProgressBar lastrow, lastrow, "S04_VDC050_Link_Generator"
            DoEvents
        End If
    Next
    
    wks.Range("E2").Value = "Liang, " & Date
    FlatWorksheet wks
    
    wks.Range("$A$5:$AG$" & lastrow).AutoFilter Field:=33, Criteria1:="1"
    
    Application.DisplayAlerts = False
    wkb.Worksheets("Setting").Delete
    wkb.Worksheets("QueryParam").Delete
    wkb.Worksheets("Data1").Delete
    wkb.Worksheets("Data2").Delete
    wkb.Worksheets("¤u§@ªí1").Delete
    'wkb.SaveAs "\\192.168.198.4\filesrv\X-File Exchange\Liang\EI Report\XR01_EDC060\EDC060_" & report_date_str & ".xlsb", FileFormat:=xlExcel12
    wkb.SaveAs "d:\temp\VDC050_" & report_date_str & ".xlsb", FileFormat:=xlExcel12
    Application.DisplayAlerts = True
    wkb.Close
    Close_ProgressBar
    MsgBox "Done, " & lastrow & " rows processed", vbInformation
    

End Sub

