Attribute VB_Name = "ReadFileToSJISTextModule"
'////////////////////////////////////////////////////////////
'// ReadFileToSJISText.bas
'// ファイルを読み込んで SJIS の文字列を返す標準モジュール
'//
'// 関数
'// ReadFileToSJISText(fn)
'// MojiCode(fn)
'//
'// 履歴
'// Ver.01.                         2026/02/09
Option Explicit

Private Const ENC_UNKNOWN = "Unicode"
Private Const ENC_UTF8 = "UTF-8"
Private Const ENC_UTF8N = "UTF-8"
Private Const ENC_UTF16LE = "UTF-16le"
Private Const ENC_UTF16BE = "UTF-16be"
Private Const ENC_SHIFT_JIS = "Shift_JIS"
'Private Const ENC_SHIFT_JIS = "CP932"

Private Const adTypeBinary = 1
Private Const adTypeText = 2

Private Const adSaveCreateNotExist = 1
Private Const adSaveCreateOverWrite = 2

Private Const adReadAll = -1
Private Const adReadLine = -2

'////////////////////////////////////////////////////////////
'// ReadFileToSJISText
'// ファイルを読み込んで SJIS の文字列を返す
'// 引数:
'// fn : ファイル名
'//      ex. C:\test\file.txt
'// 戻り値:
'// Shift_JIS の文字列(行区切りは CRLF)
'//
Public Function ReadFileToSJISText(ByVal fn As String) As String

   On Error GoTo ReadFileToSJISText_Error

   ReadFileToSJISText = ""
   Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
   If Not fso.FileExists(fn) Then
      Call MsgBox("ファイルが存在しません。" & vbCrLf & fn)
      Exit Function
   Else
      Set fso = Nothing
   End If
   Dim dest_fn As String: dest_fn = Left(fn, InstrRev(fn, ".") - 1) & "_SJIS" & "." & Right(fn, Len(fn) - InstrRev(fn, "."))
   
   Dim org_char As String: org_char = MojiCode(fn)
   Dim org As Object: Set org = CreateObject("ADODB.Stream")
   Dim dest As Object: Set dest = CreateObject("ADODB.Stream")
   With dest
      .Type = adTypeText
      .Charset = ENC_SHIFT_JIS
      .Open
   End With
   With org
      .Type = adTypeText
      .Charset = org_char
      .Open
      .LoadFromFile fn
      .CopyTo dest, -1
      .Close
   End With
   With dest
      .SaveToFile dest_fn, adSaveCreateOverWrite
      .Close
   End With
   With dest
      .Type = adTypeText
      .Charset = ENC_SHIFT_JIS
      .Open
      .LoadFromFile dest_fn
      ReadFileToSJISText = .ReadText(adReadAll)
      .Close
   End With
   
   Set org = Nothing
   Set dest = Nothing

   If Len(ReadFileToSJISText) > 0 Then
      ReadFileToSJISText = Replace(Replace(Replace(ReadFileToSJISText, vbCrLf, vbLf), vbCr, vbLf), vbLf, vbCrLf)
   End If
   
   Exit Function
   
ReadFileToSJISText_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "ReadFileToSJISText:(" & Err.Number & ":" & Err.Description & ")")
   Err.Clear
   ReadFileToSJISText = ""
   
End Function

'////////////////////////////////////////////////////////////
'// MojiCode
'// ファイルの文字コードを判定する
'// 引数:
'// fn : ファイル名
'//      ex. C:\text\file.txt
'// 戻り値:
'// 文字コード文字列(定数を参照)
Public Function MojiCode(ByVal fn As String) As String

   On Error GoTo MojiCode_Error

   MojiCode = ""
   Dim ds() As Byte
   Dim strm As Object
   Set strm = CreateObject("ADODB.Stream")
   With strm
      .Type = adTypeBinary
      .Open
      .LoadFromFile = fn
      MojiCode = ENC_UNKNOWN
      If .Size > 3 Then
         .Position = 0
         ds = .Read(1)
         Select Case ds(0)
            Case &HEF
               ' UTF-8(BOM)
               ds = .Read(2)
                If ds(0) = &HBB And ds(1) = &HBF Then MojiCode = ENC_UTF8
            Case &HFF
               'UTF-16LE
               ds = .Read(1)
               If ds(0) = &HFE Then MojiCode = ENC_UTF16LE
            Case &HFE
               ds = .Read(1)
               If ds(0) = &HFF Then MojiCode = ENC_UTF16BE
            Case Else
               MojiCode = ENC_UNKNOWN
         End Select
      End If
      If MojiCode = ENC_UNKNOWN Then
         'Shift_JIS の判定
         Dim isSJIS As Boolean: isSJIS = True
         .Position = 0
         Do While .Position < .Size
            ds = .Read(1)
            If ds(0) <= &H7F Or (ds(0) >= &HA1 And ds(0) <= &HDF) Then
               '1バイト文字
            ElseIf (ds(0) >= &H81 And ds(0) <= &H9F) Or (ds(0) >= &HE0 And ds(0) <= &HFC) Then
               If .Position < .Size Then
                  ds = .Read(1)
                  If ((ds(0) >= &H40 And ds(0) <= &H7E) Or (ds(0) >= &H80 And ds(0) <= &HFC)) Then
                     '2バイト文字の1バイト目
                  Else
                     isSJIS = False
                     Exit Do
                  End If
               Else
                  isSJIS = False
                  Exit Do
               End If
            Else
               isSJIS = False
               Exit Do
            End If
         Loop
         If isSJIS Then
            MojiCode = ENC_SHIFT_JIS
            .Close
            Exit Function
         End If
         'UTF-8(BOM無) の判定
         Dim isUTF8N As Boolean: isUTF8N = True
         .Position = 0
         Do While .Position < .Size
            ds = .Read(1)
            If ds(0) <= &H7F Then
               '1バイト文字
            ElseIf ds(0) >= &HC2 And ds(0) <= &HDF Then
               If .Position < .Size Then
                  ds = .Read(1)
                  If ds(0) >= &H80 And ds(0) <= &HBF Then
                     '2バイト文字
                  Else
                     isUTF8N = False
                     Exit Do
                  End If
               Else
                  isUTF8N = False
                  Exit Do
               End If
            ElseIf ds(0) >= &HE0 And ds(0) <= &HEF Then
               If .Position + 1 < .Size Then
                  ds = .Read(2)
                  If ds(0) >= &H80 And ds(0) <= &HBF And ds(1) >= &H80 And ds(1) <= &HBF Then
                     '3バイト文字
                  Else
                     isUTF8N = False
                     Exit Do
                  End If
               Else
                  isUTF8N = False
                  Exit Do
               End If
            ElseIf ds(0) >= &HF0 And ds(0) <= &HF4 Then
               If .Position + 2 < .Size Then
                  ds = .Read(3)
                  If ds(0) >= &H80 And ds(0) <= &HBF And _
                     ds(1) >= &H80 And ds(1) <= &HBF And _
                     ds(2) >= &H80 And ds(2) <= &HBF Then
                     '4バイト文字
                  Else
                     isUTF8N = False
                     Exit Do
                  End If
               Else
                  isUTF8N = False
                  Exit Do
               End If
            Else
               isUTF8N = False
               Exit Do
            End If
         Loop
         If isUTF8N Then
            MojiCode = ENC_UTF8N
            .Close
            Exit Function
        End If
      End If
      .Close
   End With

   Exit Function
   
MojiCode_Error:
   Call MsgBox("エラーが発生しました。システム管理者に連絡してください。" & vbCrLf & "MojiCode:(" & Err.Number & ":" & Err.Description & ")")
   Err.Clear
   MojiCode = ""
   
End Function
