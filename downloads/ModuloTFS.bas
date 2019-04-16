Attribute VB_Name = "ModuloTFS"
Function RefreshTeamQueryOnWorksheet(worksheetName As String) As String

    Dim activeSheet As Worksheet
    Dim teamQueryRange As Range
    Dim refreshControl As CommandBarControl

    Set refreshControl = FindTeamControl("IDC_REFRESH")

    If refreshControl Is Nothing Then
        RefreshTeamQueryOnWorksheet = "Could not find Team Foundation commands in Ribbon. Please make sure that the Team Foundation Excel plugin is installed."
        Exit Function
    End If
    
    On Error GoTo errorHandler

    Application.ScreenUpdating = False

    Set activeSheet = ActiveWorkbook.activeSheet
    
    If SheetNotExists(worksheetName) Then
        RefreshTeamQueryOnWorksheet = "Could not find the worksheet " & worksheetName
        Exit Function
    End If
    
    Set teamQueryRange = Worksheets(worksheetName).ListObjects(1).Range

    teamQueryRange.Worksheet.Select
    teamQueryRange.Select
    refreshControl.Execute

    activeSheet.Select

    Application.ScreenUpdating = True
    
    RefreshTeamQueryOnWorksheet = "Sucess"
    
    Exit Function
    
errorHandler:
    If Not activeSheet Is Nothing Then activeSheet.Select
    Application.ScreenUpdating = True
    
    RefreshTeamQueryOnWorksheet = "The following error occurred: " & Err.Number & " " & Err.Description
End Function

Private Function FindTeamControl(tagName As String) As CommandBarControl

    Dim commandBar As commandBar
    Dim teamCommandBar As commandBar
    Dim control As CommandBarControl

    For Each commandBar In Application.CommandBars
		If commandBar.Name = "Team" Or commandBar.Name = "Equipe" Then
            Set teamCommandBar = commandBar
            Exit For
        End If
    Next

    If Not teamCommandBar Is Nothing Then
        For Each control In teamCommandBar.Controls
            If InStr(1, control.Tag, tagName) Then
                Set FindTeamControl = control
                Exit Function
            End If
        Next
    End If
End Function

Private Function SheetNotExists(sheetToFind As String) As Boolean
    SheetNotExists = True
    For Each sheet In Worksheets
        If sheetToFind = sheet.Name Then
            SheetNotExists = False
            Exit Function
        End If
    Next sheet
End Function
