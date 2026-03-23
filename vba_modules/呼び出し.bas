Attribute VB_Name = "呼び出し"
' 標準モジュール
' 全フォーム・モジュールから呼び出せる共通ユーティリティ関数を定義する。

Option Explicit

Sub StartSaitourokuProcess()
    Dim frmSelect As New 工事名称選択
    Dim frmSaitorokui As 再登録

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler_Saitouroku

    frmSelect.Show

    If Not frmSelect.Cancelled Then
        Set frmSaitorokui = New 再登録
        frmSaitorokui.SearchedKoujiName = frmSelect.selectedKoujiName
        frmSaitorokui.SelectedTantousha = frmSelect.SelectedTantousha
        frmSaitorokui.Show
    End If

    Unload frmSelect

Finalize_Saitouroku:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Set frmSelect = Nothing
    Set frmSaitorokui = Nothing
    Exit Sub

ErrorHandler_Saitouroku:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Finalize_Saitouroku
End Sub

Sub StartIraishoProcess()
    Dim frmSelect As New 工事名称選択
    Dim frmIraisho As 依頼書作成

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler_Iraisho

    frmSelect.Show

    If Not frmSelect.Cancelled Then
        Set frmIraisho = New 依頼書作成
        frmIraisho.SetupAndShow frmSelect.selectedKoujiName, frmSelect.SelectedTantousha
    End If

    Unload frmSelect

Finalize_Iraisho:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Set frmSelect = Nothing
    Set frmIraisho = Nothing
    Exit Sub

ErrorHandler_Iraisho:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Finalize_Iraisho
End Sub

'--------------------------------------------------------------------------------
' 共通ユーティリティ関数
'--------------------------------------------------------------------------------

' 外部マスターファイルのパスを返す。
' IS_TEST_MODE が True なら TEST_FILE_PATH、False なら「入力フォーム」シートから読み取る。
' 各フォーム・モジュールで同じ If IS_TEST_MODE ... パターンが重複していたため、ここに一元化。
Public Function GetTargetFilePath() As String
    If IS_TEST_MODE Then
        GetTargetFilePath = TEST_FILE_PATH
    Else
        GetTargetFilePath = Trim(CStr(ThisWorkbook.Sheets("入力フォーム").Range(PATH_CELL).Value))
    End If
End Function

' 指定したブック内に特定名称のシートが存在するかチェックする。
Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' 値が日付なら "yyyy/mm/dd" 形式の文字列を返す。日付でなければ空文字を返す。
Public Function FormatIfDate(ByVal Value As Variant) As String
    If IsDate(Value) Then
        FormatIfDate = Format(Value, "yyyy/mm/dd")
    Else
        FormatIfDate = ""
    End If
End Function

