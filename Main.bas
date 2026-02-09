Attribute VB_Name = "Main"
'/////////////////////////////////////////////////////
'// Main.bas
'// リアテンダントの得点ファイルを取り込む
'//
'// 関数
'// DoClearData()
'// DoClearMeibo()
'// SetTokutenCSV()
'// CsvToScs()
'// IsSelectedFile(fn)
'//
'// 履歴
'// Ver.0.1                2026/02/06
Option Explicit

Private Const G_COL_NEN = 2
Private Const G_COL_KUMI = 3
Private Const G_COL_BAN = 4
Private Const G_COL_SEI = 5
Private Const G_COL_MEI = 6

Private Const G_ROW_DAT_START = 18
Private Const G_ROW_DAT_END = 217
Private Const G_COL_DAT_END = 30

Private Const G_DAT_MAX = 200

Private Type Score
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

Private Scs() As Score

Private Enum Sitms
   Nen = 0
   Kumi
   Ban
   Sei
   Mei
   Haiten
   Tokuten
   Kanten1
   Kanten2
End Enum

Private Const G_LIN_TITLE = 0
Private Const G_DATA_SET_SHEET = "考査得点・クラス名票貼り付け"
Private Const G_CONF_SHEET = "設定"
Private Const G_CELL_NEN = "A2"
Private Const G_ROW_FILE_START = 5
Private Const G_COL_FILE_START = 1

'/////////////////////////////////////////////////////
'// DoClearData
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
    Err.Clear
    
End Sub

'/////////////////////////////////////////////////////
'// DoClearMeibo
'// 生徒名リストのクリア
'//
Public Sub DoClearMeibo()

    On Error GoTo DoClearMeibo_Error
    
    Dim r As Range
    Set r = Range(Cells(G_ROW_DAT_START, G_COL_NEN), Cells(G_ROW_DAT_END, G_COL_MEI))
    r.Clear
    Exit Sub
    
DoClearMeibo_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "(" & Err.Number & ":" & Err.Description & ")")
    Err.Clear

End Sub

'/////////////////////////////////////////////////////
'//
'// SetTokutenCSV
'// リアテンダントからダウンロードしたCSVファイルの値を考査得点・クラス名票貼り付けシートに貼り付ける
'//
Public Sub SetTokutenCSV()

    On Error GoTo SetTokutenCSV_Error
    
    'ファイルの文字コードを Shift_SJIS に変換したファイルを作成して、読み込み
    '配列 Scs にセットする
    If Not CsvToScs() Then Exit Sub

    Sheets(G_DATA_SET_SHEET).Select
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
    Dim ii As Long: ii = 0
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
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_NEN), Cells(G_ROW_DAT_END, 2)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_KUMI), Cells(G_ROW_DAT_END, 3)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_BAN), Cells(G_ROW_DAT_END, 4)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort
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
'// CsvToScs
'// リアテンダントからダウンロードしたファイルを読み込み
'// Shift_JIS に変換してファイルを作成し、それを読み込んで
'// 配列 Scs にセットする
'// 戻り値:
'// 処理の成功か否か
'//
Private Function CsvToScs() As Boolean

    On Error GoTo CsvToScs_Error
    
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
	   CsvToScs = False
	   Exit Function
        Else
            fn = .SelectedItems(1)
        End If
    End With
    If IsSelectedFile(fn) Then
       CsvToScs = False
       Exit Function
    End If

    Dim fc As String
    fc = ReadFileToSJISText(fn)
    If Len(fc) <= 0 Then
       Call MsgBox("選択されたファイルに取り込めるデータがありません。" & vbCrLf & "(" & fn & ")")
       CsvToScs = False
    Else
       CsvToScs = True
    End If

    Dim lines() As String
    lines = Split(fc, vbCrLf)
    
    Dim ii As Long: ii = 0
    Dim scs_idx As Long: scs_idx = -1
    Dim items() As String
    Dim seimei() As String
    For ii = 0 To UBound(lines)
       If ii <> G_LIN_TITLE And lines(ii) = "" Then
	  items = Split(lines(ii), ",")
	  scs_idx = scs_idx + 1
	  If items(Sitms.Nen) <> "" Then
	     If items(Sitms.Nen) <> Sheets(G_CONF_SHEET).Range(G_CELL_NEN).Value Then
		Call MsgBox("選択されたファイルには学年が違うデータがあるようです。" & vbCrLf & _
			    CStr(ii + 1) & "行目 " & _
			    "想定されている学年: " & CStr(Sheets(G_CONF_SHEET).Range(G_CELL_NEN).Value) & " / " & "このファイルにあるデータ:" & items(Sitms.Nen))
		CsvToScs = False
		Exit Function
	     End If
	     ReDim Preserve Scs(scs_idx)
	     Scs(scs_idx).Nen = items(Sitms.Nen)
	     Scs(scs_idx).Kumi = items(Sitms.Kumi)
	     Scs(scs_idx).Ban = items(Sitms.Ban)
	     If items(Sitms.Mei) = "さん" Then
		seimei = Split(items(Sitms.Sei), " ")
		Scs(scs_idx).Sei = seimei(0)
		Scs(scs_idx).Mei = seimei(1)
	     Else
		Scs(scs_idx).Sei = items(Sitms.Sei)
		Scs(scs_idx).Mei = items(Sitms.Mei)
	     End If
	     Scs(scs_idx).Haiten = items(Sitms.Haiten)
	     Scs(scs_idx).Tokuten = items(Sitms.Tokuten)
	     Scs(scs_idx).Kanten1 = items(Sitms.Kanten1)
	     Scs(scs_idx).Kanten2 = items(Sitms.Kanten2)
	  Else
	     scs_idx = scs_idx - 1
	  End If
       End If
    Next
    
    Exit Function
    
CsvToScs_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "(" & Err.Number & ":" & Err.Description & ")")
    Err.Clear
    CsvToScs = False

End Function

'/////////////////////////////////////////////////////
'// IsSelectedFile
'// 選択されたファイルが以前に取り込まれたファイル名と同じかどうかをチェックする
'// 引数:
'// fn: 文字列 ファイル名
'// 戻り値:
'// 処理の成功か否か
'//
Private Function IsSelectedFile(ByVal fn As String) As Boolean

   On Error GoTo IsSelectedFile_Error

   IsSelectedFile = False
   If Len(fn) = 0 Then Exit Function
   Dim c As Long: c = G_COL_FILE_START
   Dim r As Long: r = G_ROW_FILE_START

   With Sheets(G_CONF_SHEET)
      Do
	 If .Cells(r, c).Value = "" Then
	    .Cells(r, c).Value = fn
	    IsSelectedFile = False
	    Exit Do
	 Elseif .Cells(r, c).Value = fn Then
	    Dim result As Long
	    result = MsgBox("このファイルは以前取り込んだことがあるファイルのようです。再度取り込みますか？" & vbCrLf & _
			    "ファイル名: " & fn , vbYesNo + vbQuestion + vbDefaultButton2, "確認")
	    If result = vbNo Then 
	       IsSelectedFile = True
	       Exit Do
	    Else
	       Do Until .Cells(r, c).Value = ""
		  r = r + 1
	       Loop
	       .Cells(r, c).Value = fn
	       IsSelectedFile = False
	       Exit Do
	    End If
	 End If
	 r = r + 1
      Loop
   End With
   
   Exit Function

IsSelectedFile_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "(" & Err.Number & ":" & Err.Description & ")")
    Err.Clear
    IsSelectedFile = False
   
End Function
