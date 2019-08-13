Attribute VB_Name = "Module1"
Public File, ownb As String, HostFolder As String
 
Sub exportPNG()
Dim FileSystem As Object

Set FileSystem = CreateObject("Scripting.FileSystemObject")

MsgBox "Select directory with .xlsx file with chart."
GetFolder


DoFolder FileSystem.GetFolder(HostFolder)

End Sub
Sub DoFolder(Folder)
    Dim SubFolder
    For Each SubFolder In Folder.SubFolders
        DoFolder SubFolder
    Next
    For Each File In Folder.Files
    Workbooks.Open File
    Call exportToPNG
    Next
End Sub
Sub exportToPNG()
    Dim objChrt As ChartObject
    Dim myChart As Chart
    ownb = File.Name
    Workbooks(ownb).Activate
    
    Sheets.Add After:=ActiveSheet
    Sheets(1).Select
    ActiveChart.ChartArea.Select
    ActiveChart.ChartArea.Copy
    Sheets("Sheet1").Select
    ActiveSheet.Paste
    
    
    Set objChrt = ActiveWorkbook.Sheets("Sheet1").ChartObjects(1)
    Set myChart = objChrt.Chart

    myFileName = "chart.png"

    On Error Resume Next
    Kill Workbooks(File).Path & "\" & myFileName
    On Error GoTo 0

    myChart.Export Filename:=Workbooks(ownb).Path & "\" & myFileName, Filtername:="PNG"
    
    Workbooks(ownb).Close SaveChanges = False
    
End Sub
Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        HostFolder = .SelectedItems(1)
    End With
NextCode:
    GetFolder = HostFolder
    Set fldr = Nothing
End Function
