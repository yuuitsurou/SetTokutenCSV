Attribute VB_Name = "Main"
Option Explicit

Public Const G_COL_NEN = 2
Public Const G_COL_KUMI = 3
Public Const G_COL_BAN = 4
Public Const G_COL_SEI = 5
Public Const G_COL_MEI = 6

Public Const G_ROW_DAT_START = 18
Public Const G_ROW_DAT_END = 217
Public Const G_COL_DAT_END = 30

Public Const G_DAT_MAX = 200

Public Type Score
    Nen As String
    Kumi As String
    Ban As String
    Sei As String
    Mei As String
    Haiten As String
    Tokuten As String
    Kanten1 As String
    Kanten2 As String
End Type

Public Scs() As Score

'/////////////////////////////////////////////////////
'//
'// 得点データのクリア
'//
Public Sub DoClearData()
    
    On Error GoTo DoClearData_Error
    
    On Error Resume Next
    Dim result As Range
    Set result = Application.InputBox("得点をクリアする最初のセルをクリックしてください。", Type:=8)
    If Err.Number <> 0 Then
        Call MsgBox("キャンセルされました。")
        Exit Sub
    End If
    Err.Clear
    
    On Error GoTo DoClearData_Error
    Dim startColumn As Long: startColumn = result.Column
    Set result = Nothing
    
    Dim r As Range
    Set r = Range(Cells(G_ROW_DAT_START, startColumn), Cells(G_ROW_DAT_END, startColumn + 2))
    r.Clear
    Exit Sub
    
DoClearData_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "(" & Err.Number & ":" & Err.Description & ")")
    
End Sub

'/////////////////////////////////////////////////////
'//
'//生徒名リストのクリア
'//
Public Sub DoClearMeibo()

    On Error GoTo DoClearMeibo_Error
    
    Dim r As Range
    Set r = Range(Cells(G_ROW_DAT_START, G_COL_NEN), Cells(G_ROW_DAT_END, G_COL_MEI))
    r.Clear
    Exit Sub
    
DoClearMeibo_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "(" & Err.Number & ":" & Err.Description & ")")

End Sub

'/////////////////////////////////////////////////////
'//
'// SetTokutenCSV
'// リアテンダントからダウンロードしたCSVファイルの値を考査得点・クラス名票貼り付けシートに貼り付ける
'//
Public Sub SetTokutenCSV()

    On Error GoTo SetTokutenCSV_Error
    
    Dim fn As String
    'ファイルの文字コードを Shift_SJIS に変換したファイルを作成して、読み込み
    fn = SelectCSVFile()
    If fn = "" Then Exit Sub
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fs As TextStream
    Set fs = fso.OpenTextFile(fn, ForReading)
    
    Dim ii As Long: ii = 0
    Dim line As String
    Dim items() As String
    Dim seimei() As String
    '行をカンマで分解し、配列にセットする
    Do Until fs.AtEndOfStream
        line = fs.ReadLine
        If ii > 0 Then
            items = Split(line, ",")
            If items(0) <> "" Then
                ReDim Preserve Scs(ii - 1)
                Scs(ii - 1).Nen = items(0)
                Scs(ii - 1).Kumi = items(1)
                Scs(ii - 1).Ban = items(2)
                seimei = Split(items(3), " ")
                Scs(ii - 1).Sei = seimei(0)
                Scs(ii - 1).Mei = seimei(1)
                Scs(ii - 1).Haiten = items(5)
                Scs(ii - 1).Tokuten = items(6)
                Scs(ii - 1).Kanten1 = items(7)
                Scs(ii - 1).Kanten2 = items(8)
            Else
                ii = ii - 1
            End If
        End If
        ii = ii + 1
    Loop
    
    Sheets("考査得点・クラス名票貼り付け").Select
    Range("B18").Select
    
    On Error Resume Next
    Dim result As Range
    Set result = Application.InputBox("得点をセットする最初のセルをクリックしてください。", Type:=8)
    If Err.Number <> 0 Then
        Call MsgBox("キャンセルされました。")
        Exit Sub
    End If
    Err.Clear
    
    On Error GoTo SetTokutenCSV_Error
    Dim startColumn As Long: startColumn = result.Column
    Set result = Nothing
    
    Application.ScreenUpdating = False
    
    Dim c As Range
    ii = 0
    Dim ri As Long
    Dim setToCell As Boolean
    Dim newri As Long
    newri = -1
    Dim n() As Score
    For ii = 0 To UBound(Scs)
        ri = G_ROW_DAT_START - 1
        setToCell = False
        Do
            ri = ri + 1
            If Cells(ri, G_COL_NEN).Value = "" Then Exit Do
            If Cells(ri, G_COL_NEN).Value = Scs(ii).Nen _
                And Cells(ri, G_COL_KUMI).Value = Scs(ii).Kumi _
                And Cells(ri, G_COL_BAN).Value = Scs(ii).Ban _
                And Cells(ri, G_COL_SEI).Value = Scs(ii).Sei _
                And Cells(ri, G_COL_MEI).Value = Scs(ii).Mei Then
                Cells(ri, startColumn).Value = Scs(ii).Tokuten
                Cells(ri, startColumn + 1).Value = Scs(ii).Kanten1
                Cells(ri, startColumn + 2).Value = Scs(ii).Kanten2
                setToCell = True
                Exit Do
            End If
        Loop
        If Not setToCell Then
            newri = newri + 1
            ReDim Preserve n(newri)
            n(newri).Nen = Scs(ii).Nen
            n(newri).Kumi = Scs(ii).Kumi
            n(newri).Ban = Scs(ii).Ban
            n(newri).Sei = Scs(ii).Sei
            n(newri).Mei = Scs(ii).Mei
            n(newri).Haiten = Scs(ii).Haiten
            n(newri).Tokuten = Scs(ii).Tokuten
            n(newri).Kanten1 = Scs(ii).Kanten1
            n(newri).Kanten2 = Scs(ii).Kanten2
        End If
    Next
    If newri > -1 Then
        ri = G_ROW_DAT_START
        Do Until Cells(ri, 2).Value = ""
            ri = ri + 1
        Loop
        For ii = 0 To UBound(n)
            Cells(ri, G_COL_NEN).Value = n(ii).Nen
            Cells(ri, G_COL_KUMI).Value = n(ii).Kumi
            Cells(ri, G_COL_BAN).Value = n(ii).Ban
            Cells(ri, G_COL_SEI).Value = n(ii).Sei
            Cells(ri, G_COL_MEI).Value = n(ii).Mei
            Cells(ri, startColumn).Value = n(ii).Tokuten
            Cells(ri, startColumn + 1).Value = n(ii).Kanten1
            Cells(ri, startColumn + 2).Value = n(ii).Kanten2
            ri = ri + 1
        Next
    End If
    
    Application.ScreenUpdating = True
    
    Range(Cells(G_ROW_DAT_START - 1, 1), Cells(G_ROW_DAT_END, G_COL_DAT_END)).Select
    ActiveWorkbook.Worksheets("考査得点・クラス名票貼り付け").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("考査得点・クラス名票貼り付け").Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_NEN), Cells(G_ROW_DAT_END, 2)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("考査得点・クラス名票貼り付け").Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_KUMI), Cells(G_ROW_DAT_END, 3)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("考査得点・クラス名票貼り付け").Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_BAN), Cells(G_ROW_DAT_END, 4)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("考査得点・クラス名票貼り付け").Sort
        .SetRange Range(Cells(G_ROW_DAT_START - 1, 1), Cells(G_ROW_DAT_END, G_COL_DAT_END))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    For ii = 1 To G_DAT_MAX
        Cells(ii + 17, 1).Value = ii
    Next
    
    
    Exit Sub
    
SetTokutenCSV_Error:
    Application.ScreenUpdating = True
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "(" & Err.Number & ":" & Err.Description & ")")
    
End Sub

'/////////////////////////////////////////////////////
'// SelectCSVFile
'// リアテンダントからダウンロードしたファイルを読み込み
'// Shift_JIS に変換してファイルを作成する
'// 戻り値:
'// 作成した Shift_JIS ファイルのファイル名
'//
Private Function SelectCSVFile() As String

    On Error GoTo SelectCSVFile_Error
    
    Dim fn As String
    Dim dlg As FileDialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "リアテンダントからダウンロードしたCSVファイルを選択してください。"
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        .InitialFileName = Application.ActiveWorkbook.path
        .AllowMultiSelect = False
        If .Show = False Then
            Call MsgBox("キャンセルされました。")
            Exit Function
        Else
            fn = .SelectedItems(1)
        End If
    End With
    
    Dim rStream As New ADODB.Stream
    rStream.Type = adTypeText
    rStream.Charset = "UTF-8"
    rStream.Open
    Call rStream.LoadFromFile(fn)
    
    Dim t As Variant
    
    t = rStream.ReadText
    t = Replace(t, vbLf, vbCrLf)
    t = Replace(t, vbCr & vbCr, vbCr)
    
    Dim wStream As New ADODB.Stream
    wStream.Type = adTypeText
    wStream.Charset = "Shift-JIS"
    wStream.Open
    Call wStream.WriteText(t)
    fn = Left(fn, InStrRev(fn, ".") - 1) & "_SJIS.csv"
    Call wStream.SaveToFile(fn, adSaveCreateOverWrite)
    SelectCSVFile = fn
    
    Exit Function
    
SelectCSVFile_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "(" & Err.Number & ":" & Err.Description & ")")

End Function

