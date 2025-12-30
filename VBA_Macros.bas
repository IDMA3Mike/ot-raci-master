
Option Explicit

' =============================
' Global constants & helpers
' =============================
Private Const DEPT_SLICER_NAME As String = "slcDepartment"
Private Const DEPT_FIELD_CAPTION As String = "Department"

Public Sub btnFilterDepartment_Click()
    On Error GoTo ErrHandler
    Dim dept As String
    dept = Application.InputBox("Enter Department name to filter (or leave blank to reset)", "Filter by Department", Type:=2)
    If dept = "False" Then Exit Sub ' cancelled
    Call ApplyDepartmentFilter(dept)
    MsgBox IIf(Len(dept) > 0, "Filtered workbook to department: " & dept, "Reset all Department filters."), vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Filter failed: " & Err.Description, vbExclamation
End Sub

Public Sub ApplyDepartmentFilter(ByVal dept As String)
    Dim sc As SlicerCache
    Dim si As SlicerItem
    ' Try model slicer first
    For Each sc In ThisWorkbook.SlicerCaches
        If sc.Name = DEPT_SLICER_NAME Or sc.SourceName = DEPT_FIELD_CAPTION Then
            sc.ClearManualFilter
            If Len(dept) = 0 Then Exit For
            For Each si In sc.SlicerItems
                si.Selected = (LCase(si.Value) = LCase(dept))
            Next si
        End If
    Next sc

    ' Fallback: filter Tables on worksheet
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            On Error Resume Next
            If Len(dept) = 0 Then
                lo.AutoFilter.ShowAllData
            Else
                Dim colIdx As Long
                colIdx = GetColumnIndex(lo, DEPT_FIELD_CAPTION)
                If colIdx > 0 Then
                    lo.Range.AutoFilter Field:=colIdx, Criteria1:=dept
                End If
            End If
            On Error GoTo 0
        Next lo
    Next ws
End Sub

Private Function GetColumnIndex(lo As ListObject, headerName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If LCase(lo.ListColumns(i).Name) = LCase(headerName) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function

' =============================
' Export CSVs by Department
' =============================
Public Sub btnExportCSVByDepartment_Click()
    On Error GoTo ErrHandler
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Choose export folder"
        If .Show <> -1 Then Exit Sub
    End With
    Dim exportPath As String
    exportPath = fldr.SelectedItems(1)

    Dim departments As Collection
    Set departments = CollectDepartments()
    If departments.Count = 0 Then
        MsgBox "No departments found.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dept As Variant
    Dim manifest As Collection
    Set manifest = New Collection

    For Each dept In departments
        Call ApplyDepartmentFilter(CStr(dept))
        ' Export every master table ListObject
        For Each ws In ThisWorkbook.Worksheets
            For Each lo In ws.ListObjects
                If IsMasterTable(lo.Name) Then
                    Dim fileName As String
                    fileName = exportPath & Application.PathSeparator & CStr(dept) & "_" & CleanName(lo.Name) & ".csv"
                    Dim rowCount As Long
                    rowCount = ExportListObjectUTF8 lo, fileName
                    manifest.Add CStr(dept) & "," & lo.Name & "," & CStr(rowCount) & "," & Format(Now, "yyyy-mm-dd hh:nn:ss")
                End If
            Next lo
        Next ws
    Next dept

    ' Write manifest CSV
    Dim manifestFile As String
    manifestFile = exportPath & Application.PathSeparator & "Manifest.csv"
    WriteTextUTF8 manifestFile, "FileName,TableName,RowCount,GeneratedOn" & vbCrLf & JoinCollection(manifest, vbCrLf)

    Call ApplyDepartmentFilter("") ' reset
    MsgBox "Export complete. Files saved to: " & exportPath, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Export failed: " & Err.Description, vbExclamation
End Sub

Private Function IsMasterTable(name As String) As Boolean
    Dim masters As Variant
    masters = Array("Master_Activities","Master_RACI_Assignments","Master_Staffing","Staffing_Ratio_Models","Questionnaire_Responses","Dependencies_Register","Role_Map","OrgNodes","OrgEdges")
    Dim i As Long
    For i = LBound(masters) To UBound(masters)
        If LCase(name) = LCase(masters(i)) Then
            IsMasterTable = True
            Exit Function
        End If
    Next i
    IsMasterTable = False
End Function

Private Function CleanName(n As String) As String
    CleanName = Replace(Replace(n, " ", ""), "/", "-")
End Function

Private Function CollectDepartments() As Collection
    Dim coll As New Collection
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Gather from Role_Map and Master_Staffing if present
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase(lo.Name) = "role_map" Or LCase(lo.Name) = "master_staffing" Then
                Dim colIdx As Long: colIdx = GetColumnIndex(lo, DEPT_FIELD_CAPTION)
                If colIdx > 0 Then
                    Dim rng As Range: Set rng = lo.DataBodyRange.Columns(colIdx)
                    Dim cell As Range
                    For Each cell In rng.Cells
                        If Not dict.Exists(LCase(CStr(cell.Value))) And Len(Trim(CStr(cell.Value))) > 0 Then
                            dict.Add LCase(CStr(cell.Value)), True
                        End If
                    Next cell
                End If
            End If
        Next lo
    Next ws

    Dim k As Variant
    For Each k In dict.Keys
        coll.Add k
    Next k
    Set CollectDepartments = coll
End Function

' =============================
' UTF-8 CSV writer using ADODB.Stream
' =============================
Private Function ExportListObjectUTF8(lo As ListObject, filePath As String) As Long
    Dim tmp As String
    tmp = GetListObjectCSV(lo)
    WriteTextUTF8 filePath, tmp
    ExportListObjectUTF8 = lo.DataBodyRange.Rows.Count
End Function

Private Sub WriteTextUTF8(filePath As String, ByVal content As String)
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2 ' text
        .Charset = "utf-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2
        .Close
    End With
End Sub

Private Function GetListObjectCSV(lo As ListObject) As String
    Dim s As String
    Dim r As Range: Set r = lo.Range
    Dim i As Long, j As Long
    ' Header row
    For j = 1 To lo.ListColumns.Count
        s = s & EscapeCSV(lo.ListColumns(j).Name)
        If j < lo.ListColumns.Count Then s = s & ","
    Next j
    s = s & vbCrLf
    ' Data rows
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.DataBodyRange.Rows.Count
            For j = 1 To lo.ListColumns.Count
                s = s & EscapeCSV(lo.DataBodyRange.Cells(i, j).Value)
                If j < lo.ListColumns.Count Then s = s & ","
            Next j
            s = s & vbCrLf
        Next i
    End If
    GetListObjectCSV = s
End Function

Private Function EscapeCSV(v As Variant) As String
    Dim t As String
    t = CStr(v)
    If InStr(1, t, ",") > 0 Or InStr(1, t, "
") > 0 Or InStr(1, t, "") > 0 Or InStr(1, t, '"') > 0 Then
        t = '"' & Replace(t, '"', '""') & '"'
    End If
    EscapeCSV = t
End Function

' =============================
' Org chart & overlap diagram (basic)
' =============================
Public Sub BuildOrgChart()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("OrgChart")
    ws.Cells.Clear

    Dim loNodes As ListObject
    Set loNodes = FindListObject("OrgNodes")
    If loNodes Is Nothing Then
        MsgBox "OrgNodes table not found.", vbExclamation
        Exit Sub
    End If

    ' Create SmartArt Organization Chart
    Dim shp As Shape
    Set shp = ws.Shapes.AddSmartArt(Application.SmartArtLayouts(31), 50, 50, 800, 500) ' 31 is OrgChart in many builds

    ' Populate text (simple: Department nodes)
    Dim dictDept As Object: Set dictDept = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To loNodes.DataBodyRange.Rows.Count
        Dim dept As String
        dept = CStr(loNodes.DataBodyRange.Cells(i, loNodes.ListColumns("Department").Index).Value)
        If Len(dept) > 0 And Not dictDept.Exists(dept) Then dictDept.Add dept, 1
    Next i
    shp.TextFrame2.TextRange.Text = "OT Organization" & vbCrLf & Join(dictDept.Keys, vbCrLf)

    MsgBox "Org chart added (basic). For detailed role hierarchy, consider Visio export using OrgNodes/OrgEdges CSVs.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Org chart failed: " & Err.Description, vbExclamation
End Sub

Private Function FindListObject(name As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set FindListObject = ws.ListObjects(name)
        If Not FindListObject Is Nothing Then Exit Function
        On Error GoTo 0
    Next ws
End Function
