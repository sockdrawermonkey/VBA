Option Compare Database
Option Explicit


Public Function Maximum(ParamArray FieldArray() As Variant)
    ' Declare the two local variables.
    Dim I As Integer
    Dim currentVal As Variant

    ' Set the variable currentVal equal to the array of values.
    currentVal = FieldArray(0)

    ' Cycle through each value from the row to find the largest.

    For I = 0 To UBound(FieldArray)
        If FieldArray(I) > currentVal Then
            currentVal = FieldArray(I)
        End If
    Next I

    ' Return the maximum value found.
    Maximum = currentVal

End Function
Function GrabMFT()
    Dim file_path As String
    file_path = "J:\SessionData\2019\ChenMed_2018_2019\Technology\Automations\C.Grab_Inbound_MFT.bat"
    Call Shell(file_path, vbNormalFocus)
End Function
Function UpdateCovSum_NotSeen()
    Dim file_path As String
    file_path = "J:\SessionData\2019\ChenMed_2018_2019\Technology\Automations\U2X_Reports_Latest.bat"
    Call Shell(file_path, vbNormalFocus)
End Function
Function Rename_HRIS()
    Dim file_path As String
    file_path = "J:\SessionData\2019\ChenMed_2018_2019\Technology\2 PreEnroll\2 Client Final Data\HRIS\RenameToClientHRIS.bat"
    Call Shell(file_path, vbNormalFocus)
End Function
Function Proc_Run_Import()
    On Error GoTo Proc_Run_Import_Err
    Dim EEFilePath As String
    Dim DepFilePath As String


    EEFilePath = Dir("J:\SessionData\2019\ChenMed_2018_2019\Technology\2 PreEnroll\2 Client Final Data\HRIS\2691_HRIS_*.txt")
    'DepFilePath = Dir("C:\Users\A0711823\Desktop\2691_Dependent*.txt")

    DoCmd.OpenQuery "_001_Clear_EELayoutV8", acViewNormal, acEdit
    'DoCmd.OpenQuery "_002_Clear_DepLayoutV8", acViewNormal, acEdit

    DoCmd.TransferText acImportDelim, "_EELayout_Import", "_EELayoutV8", "J:\SessionData\2019\ChenMed_2018_2019\Technology\2 PreEnroll\2 Client Final Data\HRIS\" & EEFilePath, False, ""
    'DoCmd.TransferText acImportDelim, "_DepLayoutV8_Import", "_DepLayoutV8", "C:\Users\A0711823\Desktop\" & DepFilePath, False, ""

    DoCmd.OpenQuery "_005_Delete_Header_Footer_EELayout", acViewNormal, acEdit
    DoCmd.OpenQuery "_006_Delete_Header_Footer_DepLayoutV8", acViewNormal, acEdit

Proc_Run_Import_Exit:
    On Error GoTo 0
    Exit Function

Proc_Run_Import_Err:
    MsgBox Error$ & " - No file import"
    Resume Proc_Run_Import_Exit
End Function


Function Import_Census_Files()
    On Error GoTo ERR_RTN

    Dim strPath1 As String
    Dim strFile1 As String
    Dim strFile2 As String
    Dim strMsgBox As String

    Import_Census_Files = ""

    strPath1 = Replace(Replace(CurrentProject.Path, "\3 Working Database", "\2 Client Final Data\Census\"), "C:\", "J:\")
    strFile1 = GetLatestFile(strPath1, "2691_HRIS_", "xlsx")
    strFile2 = GetLatestFile(strPath1, "2691_Dependent_", "xlsx")
    strMsgBox = "Import these Census Files?  Please verify." & vbCrLf & vbCrLf & _
                "     " & strFile1 & vbCrLf & _
                "     " & strFile2 & vbCrLf

    If (MsgBox(strMsgBox, vbYesNo) <> vbYes) Then GoTo ERR_RTN

    DoCmd.OpenQuery "_001_Clear_EELayoutV8", acViewNormal, acEdit
    DoCmd.OpenQuery "_002_Clear_DEPLayoutV8", acViewNormal, acEdit

    DoCmd.TransferText acImportDelim, "_EELayout_Import", "_EELayoutV8", strPath1 & strFile1, False, ""
    DoCmd.TransferText acImportDelim, "_DepLayoutV8_Import", "_DepLayoutV8", strPath1 & strFile2, False, ""

    DoCmd.OpenQuery "_011_Delete_Header_Footer_EELayout", acViewNormal, acEdit
    DoCmd.OpenQuery "_012_Delete_Header_Footer_DEPLayoutV8", acViewNormal, acEdit

    Import_Census_Files = "Y"

EXIT_RTN:
    On Error GoTo 0
    Exit Function

ERR_RTN:
    MsgBox Err.Number & " - " & Err.Description & " - No file import"
    Resume EXIT_RTN
End Function

Public Function Import_U2X_Reports()
    On Error GoTo ERR_RTN

    Dim strPath1 As String
    Dim strPath2 As String
    Dim strFile1 As String
    Dim strFile2 As String
    Dim strMsgBox As String

    Import_U2X_Reports = ""

    strPath1 = "J:\Regional\U2X_Reports\PRD\2018\2691_ChenMed\BASE_Employee_Census\"
    strPath2 = "J:\Regional\U2X_Reports\PRD\2018\2691_ChenMed\BASE_Not_Seen_Report\"
    strFile1 = GetLatestFile(strPath1, "BASE_EmployeeCensus_", "xlsx")
    strFile2 = GetLatestFile(strPath2, "NotSeenReport_", "xlsx")
    strMsgBox = "Import these U2X Reports?  Please verify." & vbCrLf & vbCrLf & _
                "     " & strFile1 & vbCrLf & _
                "     " & strFile2 & vbCrLf

    If (MsgBox(strMsgBox, vbYesNo) <> vbYes) Then GoTo ERR_RTN

    DoCmd.OpenQuery "_021_Clear_U2X_Base_Census", acViewNormal, acEdit
    DoCmd.OpenQuery "_022_Clear_NotSeen_NH", acViewNormal, acEdit

    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "U2X_Base_Census", strPath1 & strFile1, True
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "NotSeen_NH", strPath2 & strFile2, True

    Import_U2X_Reports = "Y"

EXIT_RTN:
    On Error GoTo 0
    Exit Function

ERR_RTN:
    MsgBox Err.Number & " - " & Err.Description & " - No file import"
    Resume EXIT_RTN
End Function
Public Function Export_HRIS2()     'Export only if records exist
    Dim objExcel As Object
    Dim objBook  As Object
    Dim objSheet As Object
    Dim strFile, strSource  As String
    Dim rs As Recordset
    Dim objFSO As Object

    Set objFSO = CreateObject("Scripting.Filesystemobject")

    'strFile = Replace(Replace(CurrentProject.Path, "\3 Working Database", "\2 Client Final Data\HRIS\"), "C:\", "J:\")
    strFile = "J:\SessionData\2019\ChenMed_2018_2019\Technology\2 PreEnroll\2 Client Final Data\HRIS\"
    strFile = strFile & "2691_HRIS.xlsx"
    strSource = "_EELayoutV8_ToExport"

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & strSource)
    
    If rs.RecordCount > 0 Then
    
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strSource, strFile
    
        Set objExcel = CreateObject("Excel.Application")
        Set objBook = objExcel.Workbooks.Open(strFile)
        Set objSheet = objBook.Sheets(1)
    
        objSheet.name = "Import_File"
        objSheet.Cells.EntireColumn.AutoFit
    
        objBook.Save
        objBook.Saved = True
        objBook.Close
        Set objSheet = Nothing
        Set objBook = Nothing
        Set objExcel = Nothing
        
        objFSO.CopyFile strFile, Replace(strFile, ".xlsx", "_" & Format(Now(), "yyyymmdd") & ".xlsx")

        'MsgBox "Census exported.", vbOKOnly
    Else
        MsgBox "No records to export.", vbOKOnly
    End If
    
    rs.Close
    Set rs = Nothing
    Set objFSO = Nothing
    On Error GoTo 0
    
End Function
Function NewestNotSeenFile()

Dim FileName As String
Dim MostRecentFile As String
Dim MostRecentDate As Date
Dim FileSpec As String

'Specify the file type, if any
 FileSpec = "*.*"
'specify the directory
 Directory = "J:\Regional\U2X_Reports\PRD\2019\2691_ChenMed\BASE_Not_Seen_Report"
FileName = Dir(Directory & FileSpec)

If FileName <> "" Then
    MostRecentFile = FileName
    MostRecentDate = FileDateTime(Directory & FileName)
    Do While FileName <> ""
        If FileDateTime(Directory & FileName) > MostRecentDate Then
             MostRecentFile = FileName
             MostRecentDate = FileDateTime(Directory & FileName)
        End If
        FileName = Dir
    Loop
End If

NewestFile = MostRecentFile

End Function

Public Function Export_HRIS()     'Export only if records exist
    Dim objExcel As Object
    Dim objBook  As Object
    Dim objSheet As Object
    Dim strFile, strSource  As String
    Dim rs As Recordset
    Dim objFSO As Object

    Set objFSO = CreateObject("Scripting.Filesystemobject")

    'strFile = Replace(Replace(CurrentProject.Path, "\3 Working Database", "\2 Client Final Data\HRIS\"), "C:\", "J:\")
    strFile = "J:\SessionData\2018\ChenMed_2018_2019\Technology\2 PreEnroll\2 Client Final Data\HRIS\"
    strFile = strFile & "2691_HRIS.xlsx"
    strSource = "_EELayoutV8_ToExport"

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & strSource)
    
    If rs.RecordCount > 0 Then
    
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strSource, strFile
    
        Set objExcel = CreateObject("Excel.Application")
        Set objBook = objExcel.Workbooks.Open(strFile)
        Set objSheet = objBook.Sheets(1)
    
        objSheet.name = "Import_File"
        objSheet.Cells.EntireColumn.AutoFit
    
        objBook.Save
        objBook.Saved = True
        objBook.Close
        Set objSheet = Nothing
        Set objBook = Nothing
        Set objExcel = Nothing
        
       ' objFSO.CopyFile strFile, Replace(strFile, ".xlsx", "_" & Format(Now(), "yyyymmdd") & ".xlsx")

        MsgBox "Census exported.", vbOKOnly
    Else
        MsgBox "No records to export.", vbOKOnly
    End If
    
    rs.Close
    Set rs = Nothing
    Set objFSO = Nothing
    On Error GoTo 0
    
End Function

Function Import_Weekly_File()
   Dim strName As String
   Dim objFSO As New FileSystemObject
   Dim objFolder As String
    Dim objSubFolder As Object
   Dim objFile As String
   Dim vardate As String
    Dim PreFilePath As String
    PreFilePath = Application.CurrentProject.Path
    PreFilePath = Left(PreFilePath, InStrRev(PreFilePath, "\")) & "2 Client Final Data\"
    
    DoCmd.RunSQL ("DELETE [_EELayoutV8].* FROM _EELayoutV8")
   ' DoCmd.RunSQL ("DELETE [_pre_emp_normal].* FROM _pre_emp_normal;")
  '  DoCmd.RunSQL ("DELETE [_pre_emp_monthly].* FROM _pre_emp_monthly;")
  '  DoCmd.RunSQL ("DELETE [_pre_emp_temp].* FROM _pre_emp_temp;")
    
    
 '   Set objfso = CreateObject("Scripting.FileSystemObject")
   objFSO.GetFolder (PreFilePath)

   'For Each objFile In objFolder.Files
    '        If vardate = Empty Or vardate < objFile.DateLastModified Then
     '           vardate = objFile.DateLastModified
      '          strName = objFile.name
       '     End If
            

        


   ' Next objFile
    
    If InStr(strName, "2691_HRIS") > 0 Then
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "_EELayoutV8", PreFilePath & strName, False
    '     DoCmd.RunSQL ("DELETE [_pre_emp_temp].F1, [_pre_emp_temp].F2, [_pre_emp_temp].* FROM _pre_emp_temp WHERE ((([_pre_emp_temp].F1)='Code') AND (([_pre_emp_temp].F2)='SSN'));")
     '  DoCmd.OpenQuery ("_0_0_1_Clear_Blanks_From_Temp")
     '   DoCmd.OpenQuery ("_0_0_2_ClearDemoRefresh")
        '   DoCmd.OpenQuery ("0_0_2_MovetoZCensus1")
        
        
 '   Else
 '       DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel9, "_pre_emp_normal", PreFilePath & strName, True
    End If


 '   DoCmd.OpenQuery ("_0020_Append_pre_emp_monthly")
  '  DoCmd.OpenQuery ("_0021_Append_pre_emp_normal")

    
  '  DoCmd.RunSQL ("UPDATE _pre_emp SET [_pre_emp].F35 = '" & vardate & "';")





End Function

Function sort_move_files()


    Dim strName As String
    Set objFSO = Nothing
    Set objFolder = Nothing
    Dim objSubFolder As Object
    Set objFile = Nothing
    vardate = Empty
    
    
    Dim strFileName As String
    Dim strFileDate As String
    
    Dim PreFilePath As String
    PreFilePath = Application.CurrentProject.Path
    PreFilePath = Left(PreFilePath, InStrRev(PreFilePath, "\")) & "2 Client Final Data\"
    
    Dim teststring As String
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(PreFilePath)
    
    For Each objFile In objFolder.Files
           
                'varnotseenDate = objFile.DateLastModified
                'strNotseenName = objFile.name
            If Dir(PreFilePath & "Archive" & "\", vbDirectory) = "" Then
                MkDir (PreFilePath & "Archive")
            End If
                
                
                
            strFileName = objFile.name
            strFileDate = Format(objFile.DateLastModified, "yyyymmdd")
            If Dir(PreFilePath & "Archive\" & strFileDate & "\", vbDirectory) = "" Then
                MkDir (PreFilePath & "Archive\" & strFileDate)
            End If
            Call objFSO.CopyFile(PreFilePath & strFileName, PreFilePath & "Archive\" & strFileDate & "\")
            If Dir(PreFilePath & "Archive\" & strFileDate & "\" & strFileName) <> "" Then
                objFSO.DeleteFile (PreFilePath & strFileName)
            End If
       
    Next 'objFile


End Function
Public Function Export_BenLoad()     'Export only if records exist
    Dim objExcel As Object
    Dim objBook  As Object
    Dim objSheet As Object
    Dim strFile, strSource  As String
    Dim rs As Recordset
    Dim objFSO As Object

    Set objFSO = CreateObject("Scripting.Filesystemobject")

    'strFile = Replace(Replace(CurrentProject.Path, "\3 Working Database", "\2 Client Final Data\HRIS\"), "C:\", "J:\")
    strFile = "J:\SessionData\2018\ChenMed\Technology\2 PreEnroll\2 Client Final Data\HRIS\"
    strFile = strFile & "2691_Beneficiary.xlsx"
    strSource = "_BenefLayoutV8_ToExport"

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & strSource)
    
    If rs.RecordCount > 0 Then
    
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strSource, strFile
    
        Set objExcel = CreateObject("Excel.Application")
        Set objBook = objExcel.Workbooks.Open(strFile)
        Set objSheet = objBook.Sheets(1)
    
        objSheet.name = "Import_File"
        objSheet.Cells.EntireColumn.AutoFit
    
        objBook.Save
        objBook.Saved = True
        objBook.Close
        Set objSheet = Nothing
        Set objBook = Nothing
        Set objExcel = Nothing
        
        objFSO.CopyFile strFile, Replace(strFile, ".xlsx", "_" & Format(Now(), "yyyymmdd") & ".xlsx")

        'MsgBox "Census exported.", vbOKOnly
    Else
        MsgBox "No records to export.", vbOKOnly
    End If
    
    rs.Close
    Set rs = Nothing
    Set objFSO = Nothing
    On Error GoTo 0
    
End Function


Public Function Export_CovLoad()     'Export only if records exist
    Dim objExcel As Object
    Dim objBook  As Object
    Dim objSheet As Object
    Dim strFile, strSource  As String
    Dim rs As Recordset
    Dim objFSO As Object

    Set objFSO = CreateObject("Scripting.Filesystemobject")

    'strFile = Replace(Replace(CurrentProject.Path, "\3 Working Database", "\2 Client Final Data\HRIS\"), "C:\", "J:\")
    strFile = "J:\SessionData\2018\ChenMed\Technology\2 PreEnroll\2 Client Final Data\HRIS\"
    strFile = strFile & "2691_Coverage.xlsx"
    strSource = "_CoreCovLayoutV8_ToExport"

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & strSource)
    
    If rs.RecordCount > 0 Then
    
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strSource, strFile
    
        Set objExcel = CreateObject("Excel.Application")
        Set objBook = objExcel.Workbooks.Open(strFile)
        Set objSheet = objBook.Sheets(1)
    
        objSheet.name = "Import_File"
        objSheet.Cells.EntireColumn.AutoFit
    
        objBook.Save
        objBook.Saved = True
        objBook.Close
        Set objSheet = Nothing
        Set objBook = Nothing
        Set objExcel = Nothing
        
        objFSO.CopyFile strFile, Replace(strFile, ".xlsx", "_" & Format(Now(), "yyyymmdd") & ".xlsx")

        'MsgBox "Census exported.", vbOKOnly
    Else
        MsgBox "No records to export.", vbOKOnly
    End If
    
    rs.Close
    Set rs = Nothing
    Set objFSO = Nothing
    On Error GoTo 0
    
End Function


Public Function Export_DepLoad()     'Export only if records exist
    Dim objExcel As Object
    Dim objBook  As Object
    Dim objSheet As Object
    Dim strFile, strSource  As String
    Dim rs As Recordset
    Dim objFSO As Object

    Set objFSO = CreateObject("Scripting.Filesystemobject")

    'strFile = Replace(Replace(CurrentProject.Path, "\3 Working Database", "\2 Client Final Data\HRIS\"), "C:\", "J:\")
    strFile = "J:\SessionData\2018\ChenMed\Technology\2 PreEnroll\2 Client Final Data\HRIS\"
    strFile = strFile & "2691_Dependent.xlsx"
    strSource = "_DepLayoutV8_ToExport"

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & strSource)
    
    If rs.RecordCount > 0 Then
    
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strSource, strFile
    
        Set objExcel = CreateObject("Excel.Application")
        Set objBook = objExcel.Workbooks.Open(strFile)
        Set objSheet = objBook.Sheets(1)
    
        objSheet.name = "Import_File"
        objSheet.Cells.EntireColumn.AutoFit
    
        objBook.Save
        objBook.Saved = True
        objBook.Close
        Set objSheet = Nothing
        Set objBook = Nothing
        Set objExcel = Nothing
        
        objFSO.CopyFile strFile, Replace(strFile, ".xlsx", "_" & Format(Now(), "yyyymmdd") & ".xlsx")

        'MsgBox "Census exported.", vbOKOnly
    Else
        MsgBox "No records to export.", vbOKOnly
    End If
    
    rs.Close
    Set rs = Nothing
    Set objFSO = Nothing
    On Error GoTo 0
    
End Function

Public Function Export_NotSeen()     'Export only if records exist
    Dim objExcel As Object
    Dim objBook  As Object
    Dim objSheet As Object
    Dim strFile, strSource  As String
    Dim rs As Recordset

    strFile = "J:\SessionData\2017\Valvoline_2017_2018\Technology\3 PostEnroll\3 Working Database\1 Reports\NotSeenReport\PROD\ToClient\"
    strFile = strFile & "NotSeen_" & Format(Now(), "yyyymmdd") & ".xlsx"
    strSource = "qryNotSeen"
    
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM " & strSource)
    
    If rs.RecordCount > 0 Then
    
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, strSource, strFile
    
        Set objExcel = CreateObject("Excel.Application")
        Set objBook = objExcel.Workbooks.Open(strFile)
        Set objSheet = objBook.Sheets(1)
    
        objSheet.name = "NotSeen"
        objSheet.Cells.EntireColumn.AutoFit
    
        objBook.Save
        objBook.Saved = True
        objBook.Close
        Set objSheet = Nothing
        Set objBook = Nothing
        Set objExcel = Nothing
        
        Call SendEmail_NotSeen
    
    Else
        MsgBox "No records to export.", vbOKOnly
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function SendEmail_NotSeen()        'SendEmail parameters that are passed to SendEmail function
    Dim strEmailTo As String
    Dim strEmailCC As String
    Dim strSubject As String
    Dim strbody As String
    Dim strFile As String
    Dim blnOK As Boolean
    Dim strFilePath As String
   
    strFilePath = "J:\SessionData\2017\Valvoline_2017_2018\Technology\3 PostEnroll\3 Working Database\1 Reports\NotSeenReport\PROD\ToClient\NotSeen_" & Format(Date, "yyyymmdd") & ".xlsx"
    
    strEmailTo = "david.hazleton@aon.com; sean.marsh2@aon.com"
    strEmailCC = "alma.heard@aon.com"
    strSubject = "Valvoline Daily Not Seen Report"
    strbody = "Daily Not Seen is ready:<BR><BR>" & _
        "<a href=""" & strFilePath & """>" & strFilePath & "</a >"

    blnOK = SendEmail(strEmailTo, strEmailCC, strSubject, strbody, "", True, "N")

End Function

Public Function GetBoiler(ByVal sFile As String) As String     'Get the email signature of the user
    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close

End Function
