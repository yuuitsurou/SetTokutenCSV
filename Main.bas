Attribute VB_Name = "Main"
'/////////////////////////////////////////////////////
'// Main.bas
'// リアテンダントの得点ファイルを取り込む
'//
'// 関数
'// DoClearData()
'// DoClearMeibo()
'// DoClearKetuji()
'// SetTokutenCSV()
'// CsvToScs()
'// SetKetujiCSV()
'// CsvToKtjs()
'// SortDatas()
'// IsOperated(fn)
'// RecordFileName(fn)
'//
'// 履歴
'// 2026/02/06 Ver.0.1
'// 2026/02/09 Ver.1.0
'// 2026/02/11 Ver.1.1 欠時読み込みを追加
'// 2026/02/13 Ver.1.2 欠時データのクリア追加
Option Explicit

Private Const G_COL_NEN = 2
Private Const G_COL_KUMI = 3
Private Const G_COL_BAN = 4
Private Const G_COL_SEI = 5
Private Const G_COL_MEI = 6

Private Const G_ROW_DAT_START = 18
Private Const G_ROW_DAT_END = 217
Private Const G_COL_DAT_END = 34

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

Private Type Ketuji
    Nen As String
    Kumi As String
    Ban As String
    Sei As String
    Mei As String
    Nissu As String
End Type

Private Scs() As Score
Private Ktjs() As Ketuji

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

Private Enum Kitms
   Nen = 0
   Kumi
   Ban
   Sei
   Mei
   Nissu
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
        Call MsgBox("キャンセルされました。", title:="得点データのクリア")
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
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "DoClearData: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
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
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "DoClearMeibo: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
    Err.Clear

End Sub

'/////////////////////////////////////////////////////
'// DoClearKetuji
'// 欠時データのクリア
'//
Public Sub DoClearKetuji()
    
    On Error GoTo DoClearKetuji_Error
    
    On Error Resume Next
    Dim result As Range
    Set result = Application.InputBox("欠時をクリアする最初のセルをクリックしてください。", Type:=8)
    If Err.Number <> 0 Then
        Call MsgBox("キャンセルされました。", title:="欠時データのクリア")
        Exit Sub
    End If
    Err.Clear
    
    On Error GoTo DoClearKetuji_Error
    Dim startColumn As Long: startColumn = result.Column
    Set result = Nothing
    
    Dim r As Range
    Set r = Range(Cells(G_ROW_DAT_START, startColumn), Cells(G_ROW_DAT_END, startColumn))
    r.Clear
    Exit Sub
    
DoClearKetuji_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "DoClearKetuji: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
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
        Call MsgBox("キャンセルされました。", title:="得点データのセット")
        Exit Sub
    End If
    Err.Clear
    
    On Error GoTo SetTokutenCSV_Error
    Dim startColumn As Long: startColumn = result.Column
    Set result = Nothing
    
    Application.ScreenUpdating = False
    
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

    Call SortDatas()
    
    Exit Sub
    
SetTokutenCSV_Error:
    Application.ScreenUpdating = True
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "SetTokutenCSV: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
    Err.Clear
    
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
	   Call MsgBox("キャンセルされました。", title:="リアテンダントのCSVデータファイルの指定")
	   CsvToScs = False
	   Exit Function
        Else
            fn = .SelectedItems(1)
        End If
    End With
    If IsOperated(fn) Then
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
    Erase Scs
    For ii = 0 To UBound(lines)
       If ii <> G_LIN_TITLE And lines(ii) <> "" And Not IsEmpty(lines(ii)) Then
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
    CsvToScs = RecordFileName(fn)
    
    Exit Function
    
CsvToScs_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "CsvToScs: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
    Err.Clear
    CsvToScs = False

End Function

'/////////////////////////////////////////////////////
'// SetKetujiCSV
'// 欠時が記録されたCSVファイルの値を考査得点・クラス名票貼り付けシートに貼り付ける
'//
Public Sub SetKetujiCSV()
   
    On Error GoTo SetKetujiCSV_Error
    
    'ファイルの文字コードを Shift_SJIS に変換したファイルを作成して、読み込み
    '配列 Scs にセットする
    If Not CsvToKtjs() Then Exit Sub

    Sheets(G_DATA_SET_SHEET).Select
    Range("B18").Select
    
    On Error Resume Next
    Dim result As Range
    Set result = Application.InputBox("欠時をセットする最初のセルをクリックしてください。", Type:=8)
    If Err.Number <> 0 Then
        Call MsgBox("キャンセルされました。", title:="欠時データのセット")
        Exit Sub
    End If
    Err.Clear
    
    On Error GoTo SetKetujiCSV_Error
    Dim startColumn As Long: startColumn = result.Column
    Set result = Nothing
    
    Application.ScreenUpdating = False
    
    Dim ii As Long: ii = 0
    Dim ri As Long
    Dim setToCell As Boolean
    Dim newri As Long
    newri = -1
    Dim n() As Ketuji
    For ii = 0 To UBound(Ktjs)
        ri = G_ROW_DAT_START - 1
        setToCell = False
        Do
            ri = ri + 1
            If Cells(ri, G_COL_NEN).Value = "" Then Exit Do
            If Cells(ri, G_COL_NEN).Value = Ktjs(ii).Nen _
                And Cells(ri, G_COL_KUMI).Value = Ktjs(ii).Kumi _
                And Cells(ri, G_COL_BAN).Value = Ktjs(ii).Ban _
                And Cells(ri, G_COL_SEI).Value = Ktjs(ii).Sei _
                And Cells(ri, G_COL_MEI).Value = Ktjs(ii).Mei Then
                Cells(ri, startColumn).Value = Ktjs(ii).Nissu
                setToCell = True
                Exit Do
            End If
        Loop
        If Not setToCell Then
            newri = newri + 1
            ReDim Preserve n(newri)
            n(newri).Nen = Ktjs(ii).Nen
            n(newri).Kumi = Ktjs(ii).Kumi
            n(newri).Ban = Ktjs(ii).Ban
            n(newri).Sei = Ktjs(ii).Sei
            n(newri).Mei = Ktjs(ii).Mei
            n(newri).Nissu = Ktjs(ii).Nissu
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
            Cells(ri, startColumn).Value = n(ii).Nissu
            ri = ri + 1
        Next
    End If
    
    Application.ScreenUpdating = True

    Call SortDatas()
    
    Exit Sub
    
SetKetujiCSV_Error:
    Application.ScreenUpdating = True
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "SetKetujiCSV: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
    Err.Clear
End Sub

'/////////////////////////////////////////////////////
'// CsvToKtjs
'// 欠時を記録した CSV を読み込み
'// Shift_JIS に変換してファイルを作成し、それを読み込んで
'// 配列 Ktjs にセットする
'// 戻り値:
'// 処理の成功か否か
'//
Private Function CsvToKtjs() As Boolean

   On Error GoTo CsvToKtjs_Error
    Dim fn As String
    Dim dlg As FileDialog
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "欠時が記録されたCSVファイルを選択してください。"
        .Filters.Clear
        .Filters.Add "CSV", "*.csv"
        .InitialFileName = Application.ActiveWorkbook.path
        .AllowMultiSelect = False
        If .Show = False Then
	   Call MsgBox("キャンセルされました。", title:="欠時データCSVの選択")
	   CsvToKtjs = False
	   Exit Function
        Else
            fn = .SelectedItems(1)
        End If
    End With
    If IsOperated(fn) Then
       CsvToKtjs = False
       Exit Function
    End If

    Dim fc As String
    fc = ReadFileToSJISText(fn)
    If Len(fc) <= 0 Then
       Call MsgBox("選択されたファイルに取り込めるデータがありません。" & vbCrLf & "(" & fn & ")")
       CsvToKtjs = False
    Else
       CsvToKtjs = True
    End If

    Dim lines() As String
    lines = Split(fc, vbCrLf)
    
    Dim ii As Long: ii = 0
    Dim ktjs_idx As Long: ktjs_idx = -1
    Dim items() As String
    Dim seimei() As String
    Erase Ktjs
    For ii = 0 To UBound(lines)
       If ii <> G_LIN_TITLE And lines(ii) <> "" And Not IsEmpty(lines(ii)) Then
	  items = Split(lines(ii), ",")
	  ktjs_idx = ktjs_idx + 1
	  If items(Kitms.Nen) <> "" Then
	     If items(Kitms.Nen) <> Sheets(G_CONF_SHEET).Range(G_CELL_NEN).Value Then
		Call MsgBox("選択されたファイルには学年が違うデータがあるようです。" & vbCrLf & _
			    CStr(ii + 1) & "行目 " & _
			    "想定されている学年: " & CStr(Sheets(G_CONF_SHEET).Range(G_CELL_NEN).Value) & " / " & "このファイルにあるデータ:" & items(Kitms.Nen))
		CsvToKtjs = False
		Exit Function
	     End If
	     ReDim Preserve Ktjs(ktjs_idx)
	     Ktjs(ktjs_idx).Nen = items(Kitms.Nen)
	     Ktjs(ktjs_idx).Kumi = items(Kitms.Kumi)
	     Ktjs(ktjs_idx).Ban = items(Kitms.Ban)
	     If items(Kitms.Mei) = "さん" Then
		seimei = Split(items(Kitms.Sei), " ")
		Ktjs(ktjs_idx).Sei = seimei(0)
		Ktjs(ktjs_idx).Mei = seimei(1)
	     Else
		Ktjs(ktjs_idx).Sei = items(Kitms.Sei)
		Ktjs(ktjs_idx).Mei = items(Kitms.Mei)
	     End If
	     Ktjs(ktjs_idx).Nissu = items(Kitms.Nissu)
	  Else
	     ktjs_idx = ktjs_idx - 1
	  End If
       End If
    Next
    CsvToKtjs = RecordFileName(fn)
    
    Exit Function

CsvToKtjs_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "CsvToKtjs: " & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
    Err.Clear
    CsvToKtjs = False
   
End Function

'/////////////////////////////////////////////////////
'// IsOperated
'// 選択されたファイルが以前に取り込まれたファイル名と同じかどうかをチェックする
'// 引数:
'// fn: 文字列 ファイル名
'// 戻り値:
'// 処理の成功か否か
'//
Private Function IsOperated(ByVal fn As String) As Boolean

   On Error GoTo IsOperated_Error

   IsOperated = False
   If Len(fn) = 0 Then Exit Function
   Dim c As Long: c = G_COL_FILE_START
   Dim r As Long: r = G_ROW_FILE_START

   With Sheets(G_CONF_SHEET)
      Do
	 If .Cells(r, c).Value = "" Then
	    IsOperated = False
	    Exit Do
	 Elseif .Cells(r, c).Value = fn Then
	    Dim result As Long
	    result = MsgBox("このファイルは以前取り込んだことがあるファイルのようです。再度取り込みますか？" & vbCrLf & _
			    "ファイル名: " & fn , vbYesNo + vbQuestion + vbDefaultButton2, "確認")
	    IsOperated = (result = vbYes)
	    Exit Do
	 End If
	 r = r + 1
      Loop
   End With
   
   Exit Function

IsOperated_Error:
    Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "IsOperated:" & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
    Err.Clear
    IsOperated = False
   
End Function

'/////////////////////////////////////////////////////
'// RecordFileName
'// 処理されたファイルを記録する
'// 引数;
'// fn: 文字列 ファイル名
'//
Private Function RecordFileName(ByVal fn As String) As Boolean

   On Error GoTo RecordFileName_Error

   RecordFileName = False
   If Len(fn) = 0 Then Exit Function

   Dim c As Long: c = G_COL_FILE_START
   Dim r As Long: r = G_ROW_FILE_START

   With Sheets(G_CONF_SHEET)
      Do Until .Cells(r, c).Value = ""
	 r = r + 1
      Loop
      .Cells(r, c).Value = fn
      .Cells(r, c + 1).Value = FormatDateTime(Now(), vbGeneralDate)
      RecordFileName = True
   End With

   Exit Function

RecordFileName_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "RecordFileName:" & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
   RecordFileName = False
   
End Function

'/////////////////////////////////////////////////////
'// SortDatas
'// データを並び換える
'//
Private Sub SortDatas()

   On Error GoTo SortDatas_Error
   
   Range(Cells(G_ROW_DAT_START - 1, 1), Cells(G_ROW_DAT_END, G_COL_DAT_END)).Select
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_NEN), Cells(G_ROW_DAT_END, G_COL_NEN)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_KUMI), Cells(G_ROW_DAT_END, G_COL_KUMI)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort.SortFields.Add2 Key:=Range(Cells(G_ROW_DAT_START, G_COL_BAN), Cells(G_ROW_DAT_END, G_COL_BAN)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(G_DATA_SET_SHEET).Sort
        .SetRange Range(Cells(G_ROW_DAT_START - 1, 1), Cells(G_ROW_DAT_END, G_COL_DAT_END))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Dim ii As Long
    For ii = G_ROW_DAT_START To G_ROW_DAT_END
        Cells(ii, 1).Value = (ii - G_ROW_DAT_START + 1)
    Next

    Exit Sub
    
SortDatas_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf _
	       & "SortDatas:" & Err.Number & vbCrLf _
	       & "( " & Err.Description & " )")
   Err.Clear
    
End Sub

