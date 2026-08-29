Option Explicit

' =====================================================================
' UI Automation 要素記録・貼り付け専用の簡易版
' ---------------------------------------------------------------------
' 公開マクロは次の2つだけです。
'   1. UIA_初期化
'   2. UIA_要素を1件記録
'
' 記録キー:
'   対象へフォーカスして Ctrl + Alt + G
'   中止は Esc
'
' 前提:
'   Windows版Excel専用
'   VBE「ツール」→「参照設定」→「UIAutomationClient」
' =====================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
        ByVal vKey As Long) As Integer
#Else
    Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
    Private Declare Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" ( _
        ByVal vKey As Long) As Integer
#End If

Private Const OUTPUT_SHEET As String = "UIA_要素一覧"

' 記録キー Ctrl + Alt + G
Private Const VK_CONTROL As Long = 17
Private Const VK_MENU As Long = 18
Private Const VK_ESCAPE As Long = 27
Private Const VK_G As Long = 71

' UIA Pattern IDs
Private Const UIA_INVOKE_PATTERN_ID As Long = 10000
Private Const UIA_VALUE_PATTERN_ID As Long = 10002
Private Const UIA_EXPAND_COLLAPSE_PATTERN_ID As Long = 10005
Private Const UIA_SELECTION_ITEM_PATTERN_ID As Long = 10010
Private Const UIA_TEXT_PATTERN_ID As Long = 10014
Private Const UIA_TOGGLE_PATTERN_ID As Long = 10015

' ControlType IDs
Private Const UIA_CT_BUTTON As Long = 50000
Private Const UIA_CT_CALENDAR As Long = 50001
Private Const UIA_CT_CHECKBOX As Long = 50002
Private Const UIA_CT_COMBOBOX As Long = 50003
Private Const UIA_CT_EDIT As Long = 50004
Private Const UIA_CT_HYPERLINK As Long = 50005
Private Const UIA_CT_IMAGE As Long = 50006
Private Const UIA_CT_LISTITEM As Long = 50007
Private Const UIA_CT_LIST As Long = 50008
Private Const UIA_CT_MENUITEM As Long = 50011
Private Const UIA_CT_RADIOBUTTON As Long = 50013
Private Const UIA_CT_TABITEM As Long = 50019
Private Const UIA_CT_TEXT As Long = 50020
Private Const UIA_CT_TREEITEM As Long = 50024
Private Const UIA_CT_GROUP As Long = 50026
Private Const UIA_CT_DATAITEM As Long = 50029
Private Const UIA_CT_DOCUMENT As Long = 50030
Private Const UIA_CT_SPLITBUTTON As Long = 50031
Private Const UIA_CT_WINDOW As Long = 50032
Private Const UIA_CT_PANE As Long = 50033

Private mUIA As CUIAutomation


' =====================================================================
' 公開マクロ1：UIAと記録先シートを用意する
' 既存の記録行は削除しません。
' =====================================================================
Public Sub UIA_初期化()

    Dim outputBook As Workbook
    Dim outputSheet As Worksheet
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    Set outputBook = GetOutputBook()
    EnsureUIA

    Set outputSheet = EnsureOutputSheet(outputBook)
    SetupOutputSheet outputSheet
    ReturnToExcel outputBook, outputSheet, 1

    MsgBox _
        "初期化が完了しました。" & vbCrLf & _
        "記録先: " & OUTPUT_SHEET & vbCrLf & vbCrLf & _
        "記録するときは「UIA_要素を1件記録」を実行してください。", _
        vbInformation, "UIA初期化"
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    SetForegroundWindow Application.hwnd
    On Error GoTo 0

    MsgBox _
        "初期化できませんでした。" & vbCrLf & _
        errorNumber & ": " & errorDescription, _
        vbExclamation, "UIA初期化エラー"

End Sub


' =====================================================================
' 公開マクロ2：Ctrl + Alt + Gを押した時点の要素を1件だけ記録する
' =====================================================================
Public Sub UIA_要素を1件記録()

    Dim outputBook As Workbook
    Dim outputSheet As Worksheet
    Dim target As IUIAutomationElement
    Dim windowElement As IUIAutomationElement
    Dim nextRow As Long
    Dim windowTitle As String
    Dim automationId As String
    Dim elementName As String
    Dim controlTypeId As Long
    Dim controlType As String
    Dim className As String
    Dim frameworkId As String
    Dim patterns As String
    Dim canSetText As Boolean
    Dim previousCancelKey As XlEnableCancelKey
    Dim cancelKeyChanged As Boolean
    Dim captureRequested As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    Set outputBook = GetOutputBook()
    EnsureUIA

    Set outputSheet = EnsureOutputSheet(outputBook)
    SetupOutputSheet outputSheet

    MsgBox _
        "OKを押したあと、記録したい画面へ移動してください。" & _
        vbCrLf & vbCrLf & _
        "入力欄：クリックして文字カーソルを表示" & vbCrLf & _
        "ボタン等：Tabキーでフォーカス" & vbCrLf & vbCrLf & _
        "準備できたら Ctrl + Alt + G を押してください。" & vbCrLf & _
        "その時点の要素を1件だけ記録します。" & vbCrLf & _
        "中止はEscです。", _
        vbInformation, "UIA要素を1件記録"

    previousCancelKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlErrorHandler
    cancelKeyChanged = True
    Application.StatusBar = _
        "UIA記録待機中：Ctrl + Alt + G／Escで中止"

    captureRequested = WaitForRecordShortcut()

    If Not captureRequested Then
        Application.EnableCancelKey = previousCancelKey
        cancelKeyChanged = False
        Application.StatusBar = False
        ReturnToExcel outputBook, outputSheet, 1
        MsgBox "記録を中止しました。", vbInformation, "UIA要素記録"
        Exit Sub
    End If

    ' フォーカス中の要素を取得するのは、この1回だけ
    Set target = mUIA.GetFocusedElement

    If target Is Nothing Then
        Err.Raise vbObjectError + 4200, , _
                  "フォーカス中のUIA要素を取得できませんでした。"
    End If

    ' Excelへ戻す前に外部画面の情報を変数へ退避する
    Set windowElement = GetWindowElement(target)

    If Not windowElement Is Nothing Then
        windowTitle = SafeName(windowElement)
    End If

    automationId = SafeAutomationId(target)
    elementName = SafeName(target)
    controlTypeId = SafeControlType(target)
    controlType = ControlTypeName(controlTypeId)
    className = SafeClassName(target)
    frameworkId = SafeFrameworkId(target)
    patterns = SupportedPatterns(target)
    canSetText = CanSetTextValue(target)

    ' Ctrl、Alt、Gが離れてからExcelへ戻す
    WaitForRecordShortcutRelease

    Application.EnableCancelKey = previousCancelKey
    cancelKeyChanged = False
    Application.StatusBar = False

    nextRow = outputSheet.Cells(outputSheet.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    outputSheet.Cells(nextRow, 1).Value = Now
    WriteCellText outputSheet.Cells(nextRow, 2), windowTitle
    WriteCellText outputSheet.Cells(nextRow, 3), automationId
    WriteCellText outputSheet.Cells(nextRow, 4), elementName
    outputSheet.Cells(nextRow, 5).Value = controlTypeId
    outputSheet.Cells(nextRow, 6).Value = controlType
    WriteCellText outputSheet.Cells(nextRow, 7), className
    WriteCellText outputSheet.Cells(nextRow, 8), frameworkId
    outputSheet.Cells(nextRow, 9).Value = patterns
    outputSheet.Cells(nextRow, 10).Value = IIf(canSetText, "○", "×")

    ColorRecordRow outputSheet, nextRow, canSetText
    ReturnToExcel outputBook, outputSheet, nextRow
    Beep

    MsgBox _
        "1件記録しました。" & vbCrLf & vbCrLf & _
        "ControlType: " & controlType & vbCrLf & _
        "Name: " & elementName & vbCrLf & _
        "CanSetText: " & IIf(canSetText, "○", "×"), _
        vbInformation, "UIA記録結果"
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    WaitForKeyRelease VK_ESCAPE
    WaitForRecordShortcutRelease

    If cancelKeyChanged Then
        Application.EnableCancelKey = previousCancelKey
    End If

    Application.StatusBar = False

    If Not outputBook Is Nothing Then
        ReturnToExcel outputBook, outputSheet, 1
    Else
        SetForegroundWindow Application.hwnd
    End If
    On Error GoTo 0

    If errorNumber = 18 Then
        MsgBox "記録を中止しました。", vbInformation, "UIA要素記録"
        Exit Sub
    End If

    MsgBox _
        "要素を記録できませんでした。" & vbCrLf & _
        errorNumber & ": " & errorDescription, _
        vbExclamation, "UIA記録エラー"

End Sub


' =====================================================================
' Ctrl + Alt + Gを待つ。EscならFalseを返す。
' =====================================================================
Private Function WaitForRecordShortcut() As Boolean

    Dim shortcutWasDown As Boolean
    Dim shortcutIsDown As Boolean

    ' マクロ起動時に残っているキーが離れてから監視を開始する
    Do While AnyRecordKeyDown()
        If IsKeyDown(VK_ESCAPE) Then
            WaitForKeyRelease VK_ESCAPE
            WaitForRecordShortcut = False
            Exit Function
        End If

        DoEvents
        Sleep 20
    Loop

    Do
        ' 中止を先に判定する
        If IsKeyDown(VK_ESCAPE) Then
            WaitForKeyRelease VK_ESCAPE
            WaitForRecordShortcut = False
            Exit Function
        End If

        shortcutIsDown = RecordShortcutDown()

        If shortcutIsDown And Not shortcutWasDown Then
            WaitForRecordShortcut = True
            Exit Function
        End If

        shortcutWasDown = shortcutIsDown
        DoEvents
        Sleep 20
    Loop

End Function


Private Function RecordShortcutDown() As Boolean

    RecordShortcutDown = _
        IsKeyDown(VK_CONTROL) And _
        IsKeyDown(VK_MENU) And _
        IsKeyDown(VK_G)

End Function


Private Function AnyRecordKeyDown() As Boolean

    AnyRecordKeyDown = _
        IsKeyDown(VK_CONTROL) Or _
        IsKeyDown(VK_MENU) Or _
        IsKeyDown(VK_G) Or _
        IsKeyDown(VK_ESCAPE)

End Function


Private Sub WaitForRecordShortcutRelease()

    Do While _
        IsKeyDown(VK_CONTROL) Or _
        IsKeyDown(VK_MENU) Or _
        IsKeyDown(VK_G)

        DoEvents
        Sleep 20
    Loop

End Sub


Private Sub WaitForKeyRelease(ByVal virtualKey As Long)

    Do While IsKeyDown(virtualKey)
        DoEvents
        Sleep 20
    Loop

End Sub


Private Function IsKeyDown(ByVal virtualKey As Long) As Boolean

    IsKeyDown = (GetAsyncKeyState(virtualKey) < 0)

End Function


' =====================================================================
' 初期化と保存先
' =====================================================================
Private Sub EnsureUIA()

    If mUIA Is Nothing Then Set mUIA = New CUIAutomation

End Sub


Private Function GetOutputBook() As Workbook

    Dim activeBook As Workbook

    Set activeBook = Application.ActiveWorkbook

    If activeBook Is Nothing Then
        Err.Raise vbObjectError + 4210, , _
                  "記録先にするExcelブックが開かれていません。"
    End If

    If UCase$(activeBook.Name) = "PERSONAL.XLSB" Or activeBook.IsAddin Then
        Err.Raise vbObjectError + 4211, , _
                  "通常のExcelブックを前面にしてから実行してください。"
    End If

    Set GetOutputBook = activeBook

End Function


Private Function EnsureOutputSheet( _
    ByVal outputBook As Workbook) As Worksheet

    Dim existingSheet As Object

    On Error Resume Next
    Set EnsureOutputSheet = outputBook.Worksheets(OUTPUT_SHEET)
    On Error GoTo 0

    If Not EnsureOutputSheet Is Nothing Then Exit Function

    On Error Resume Next
    Set existingSheet = outputBook.Sheets(OUTPUT_SHEET)
    On Error GoTo 0

    If Not existingSheet Is Nothing Then
        Err.Raise vbObjectError + 4220, , _
                  "同名のシートがありますが、ワークシートではありません。"
    End If

    If outputBook.ProtectStructure Then
        Err.Raise vbObjectError + 4221, , _
                  "ブック構成が保護されているため、一覧シートを作成できません。"
    End If

    Set EnsureOutputSheet = outputBook.Worksheets.Add( _
        After:=outputBook.Sheets(outputBook.Sheets.Count))
    EnsureOutputSheet.Name = OUTPUT_SHEET

End Function


Private Sub SetupOutputSheet(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long
    Dim existingHeader As String

    headers = Array( _
        "CapturedAt", "WindowTitle", "AutomationId", "Name", _
        "ControlTypeId", "ControlType", "ClassName", "FrameworkId", _
        "Patterns", "CanSetText")

    For i = LBound(headers) To UBound(headers)
        existingHeader = CStr(ws.Cells(1, i + 1).Value)

        If Len(existingHeader) = 0 Then
            ws.Cells(1, i + 1).Value = headers(i)
        ElseIf StrComp(existingHeader, CStr(headers(i)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 4222, , _
                      "記録先シートの見出しが異なります: " & _
                      ws.Cells(1, i + 1).Address(False, False)
        End If
    Next i

    With ws.Range("A1:J1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .WrapText = True
    End With

    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 36
    ws.Columns("C:D").ColumnWidth = 28
    ws.Columns("E:F").ColumnWidth = 16
    ws.Columns("G:H").ColumnWidth = 20
    ws.Columns("I").ColumnWidth = 38
    ws.Columns("J").ColumnWidth = 14

    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:J1").AutoFilter
    On Error GoTo 0

End Sub


Private Sub WriteCellText( _
    ByVal targetCell As Range, _
    ByVal textValue As String)

    targetCell.NumberFormat = "@"
    targetCell.Value2 = textValue

End Sub


Private Sub ColorRecordRow( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal canSetText As Boolean)

    Dim rowRange As Range

    Set rowRange = ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, 10))

    If canSetText Then
        rowRange.Interior.Color = RGB(226, 239, 218)
    Else
        rowRange.Interior.Color = RGB(221, 235, 247)
    End If

End Sub


Private Sub ReturnToExcel( _
    ByVal outputBook As Workbook, _
    ByVal outputSheet As Worksheet, _
    ByVal rowNumber As Long)

    On Error Resume Next
    Application.Visible = True
    outputBook.Activate

    If Not outputSheet Is Nothing Then
        outputSheet.Activate
        Application.Goto outputSheet.Cells(rowNumber, 1), True
    End If

    SetForegroundWindow Application.hwnd
    On Error GoTo 0

End Sub


' =====================================================================
' 対象が属するウィンドウを探す
' =====================================================================
Private Function GetWindowElement( _
    ByVal target As IUIAutomationElement) As IUIAutomationElement

    Dim walker As IUIAutomationTreeWalker
    Dim current As IUIAutomationElement
    Dim parentElement As IUIAutomationElement
    Dim depth As Long

    On Error GoTo Failed

    Set walker = mUIA.ControlViewWalker
    Set current = target

    For depth = 0 To 40
        If SafeControlType(current) = UIA_CT_WINDOW Then
            Set GetWindowElement = current
            Exit Function
        End If

        Set parentElement = walker.GetParentElement(current)
        If parentElement Is Nothing Then Exit For

        Set current = parentElement
        Set parentElement = Nothing
    Next depth

    Exit Function

Failed:
    Err.Clear

End Function


' =====================================================================
' 対応Patternと入力可否
' =====================================================================
Private Function SupportedPatterns( _
    ByVal target As IUIAutomationElement) As String

    Dim result As String

    AddPattern result, target, UIA_VALUE_PATTERN_ID, "Value"
    AddPattern result, target, UIA_TOGGLE_PATTERN_ID, "Toggle"
    AddPattern result, target, UIA_INVOKE_PATTERN_ID, "Invoke"
    AddPattern result, target, UIA_SELECTION_ITEM_PATTERN_ID, "SelectionItem"
    AddPattern result, target, UIA_EXPAND_COLLAPSE_PATTERN_ID, "ExpandCollapse"
    AddPattern result, target, UIA_TEXT_PATTERN_ID, "Text"

    SupportedPatterns = result

End Function


Private Sub AddPattern( _
    ByRef result As String, _
    ByVal target As IUIAutomationElement, _
    ByVal patternId As Long, _
    ByVal patternName As String)

    If HasPattern(target, patternId) Then
        If Len(result) > 0 Then result = result & ","
        result = result & patternName
    End If

End Sub


Private Function HasPattern( _
    ByVal target As IUIAutomationElement, _
    ByVal patternId As Long) As Boolean

    Dim invokePattern As IUIAutomationInvokePattern
    Dim valuePattern As IUIAutomationValuePattern
    Dim expandPattern As IUIAutomationExpandCollapsePattern
    Dim selectionPattern As IUIAutomationSelectionItemPattern
    Dim textPattern As IUIAutomationTextPattern
    Dim togglePattern As IUIAutomationTogglePattern

    On Error GoTo NotAvailable

    Select Case patternId
        Case UIA_INVOKE_PATTERN_ID
            Set invokePattern = target.GetCurrentPattern(patternId)
            HasPattern = Not invokePattern Is Nothing

        Case UIA_VALUE_PATTERN_ID
            Set valuePattern = target.GetCurrentPattern(patternId)
            HasPattern = Not valuePattern Is Nothing

        Case UIA_EXPAND_COLLAPSE_PATTERN_ID
            Set expandPattern = target.GetCurrentPattern(patternId)
            HasPattern = Not expandPattern Is Nothing

        Case UIA_SELECTION_ITEM_PATTERN_ID
            Set selectionPattern = target.GetCurrentPattern(patternId)
            HasPattern = Not selectionPattern Is Nothing

        Case UIA_TEXT_PATTERN_ID
            Set textPattern = target.GetCurrentPattern(patternId)
            HasPattern = Not textPattern Is Nothing

        Case UIA_TOGGLE_PATTERN_ID
            Set togglePattern = target.GetCurrentPattern(patternId)
            HasPattern = Not togglePattern Is Nothing
    End Select

    Exit Function

NotAvailable:
    Err.Clear

End Function


Private Function CanSetTextValue( _
    ByVal target As IUIAutomationElement) As Boolean

    Dim valuePattern As IUIAutomationValuePattern

    On Error GoTo NotAvailable
    Set valuePattern = target.GetCurrentPattern(UIA_VALUE_PATTERN_ID)
    CanSetTextValue = Not valuePattern.CurrentIsReadOnly
    Exit Function

NotAvailable:
    Err.Clear

End Function


' =====================================================================
' UIAプロパティの安全な読取り
' =====================================================================
Private Function SafeName( _
    ByVal target As IUIAutomationElement) As String

    On Error GoTo Failed
    SafeName = target.CurrentName
    Exit Function

Failed:
    Err.Clear

End Function


Private Function SafeAutomationId( _
    ByVal target As IUIAutomationElement) As String

    On Error GoTo Failed
    SafeAutomationId = target.CurrentAutomationId
    Exit Function

Failed:
    Err.Clear

End Function


Private Function SafeClassName( _
    ByVal target As IUIAutomationElement) As String

    On Error GoTo Failed
    SafeClassName = target.CurrentClassName
    Exit Function

Failed:
    Err.Clear

End Function


Private Function SafeFrameworkId( _
    ByVal target As IUIAutomationElement) As String

    On Error GoTo Failed
    SafeFrameworkId = target.CurrentFrameworkId
    Exit Function

Failed:
    Err.Clear

End Function


Private Function SafeControlType( _
    ByVal target As IUIAutomationElement) As Long

    On Error GoTo Failed
    SafeControlType = target.CurrentControlType
    Exit Function

Failed:
    Err.Clear

End Function


Private Function ControlTypeName(ByVal controlTypeId As Long) As String

    Select Case controlTypeId
        Case UIA_CT_BUTTON: ControlTypeName = "Button"
        Case UIA_CT_CALENDAR: ControlTypeName = "Calendar"
        Case UIA_CT_CHECKBOX: ControlTypeName = "CheckBox"
        Case UIA_CT_COMBOBOX: ControlTypeName = "ComboBox"
        Case UIA_CT_EDIT: ControlTypeName = "Edit"
        Case UIA_CT_HYPERLINK: ControlTypeName = "Hyperlink"
        Case UIA_CT_IMAGE: ControlTypeName = "Image"
        Case UIA_CT_LISTITEM: ControlTypeName = "ListItem"
        Case UIA_CT_LIST: ControlTypeName = "List"
        Case UIA_CT_MENUITEM: ControlTypeName = "MenuItem"
        Case UIA_CT_RADIOBUTTON: ControlTypeName = "RadioButton"
        Case UIA_CT_TABITEM: ControlTypeName = "TabItem"
        Case UIA_CT_TEXT: ControlTypeName = "Text"
        Case UIA_CT_TREEITEM: ControlTypeName = "TreeItem"
        Case UIA_CT_GROUP: ControlTypeName = "Group"
        Case UIA_CT_DATAITEM: ControlTypeName = "DataItem"
        Case UIA_CT_DOCUMENT: ControlTypeName = "Document"
        Case UIA_CT_SPLITBUTTON: ControlTypeName = "SplitButton"
        Case UIA_CT_WINDOW: ControlTypeName = "Window"
        Case UIA_CT_PANE: ControlTypeName = "Pane"
        Case Else: ControlTypeName = "ControlType(" & controlTypeId & ")"
    End Select

End Function
