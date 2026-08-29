Attribute VB_Name = "UIA_OneFileRecorderRunner"
Option Explicit

' =====================================================================
' UI Automation レコーダー＋シナリオ実行（一体型・標準モジュール1個）
' ---------------------------------------------------------------------
' できること:
'   ・F8でマウス位置、F9でフォーカス中のUIA要素を直接記録
'   ・記録した行を選択して検索テスト
'   ・選択した1行だけ実行
'   ・有効な全行を上から順番に実行
'   ・SET_TEXT / CHECK_ON / CHECK_OFF / CLICK / SELECT
'     EXPAND / COLLAPSE / READ_VALUE / READ_TEXT / READ_NAME / WAIT
'
' 前提:
'   ・Windows版Excel専用です。
'   VBEの「ツール」→「参照設定」で
'   「UIAutomationClient」にチェックを入れてください。
'
' 最初に実行するマクロ:
'   UIA_メニュー
'
' 基本手順:
'   1. メニューの「1」で、F8（マウス）またはF9（直接）で記録
'   2. UIA_操作シートのAction・ValueOrSource・Destinationを確認
'   3. 行を選んで「2」で検索テスト、「3」でその行だけ実行
'   4. 問題がなければ「4」でEnabled=○の行をStep順に一括実行
'   緑行は直接取得か通常のマウス取得かつ一意、黄色行は要確認です。
'
' 重要:
'   ・本番前にテスト用画面で確認してください。
'   ・CLICKは既定で確認ありにしています。
'   ・SendKeysは使用しません。
'   ・画面遷移後はキャッシュを消して再取得します。
' =====================================================================

Private Type POINTAPI
    x As Long
    y As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
        ByRef lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
        ByVal vKey As Long) As Integer
    Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)

    #If Win64 Then
        Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" ( _
            ByVal packedPoint As LongPtr, _
            ByRef ppacc As Object, _
            ByRef pvarChild As Variant) As Long
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            ByRef Destination As Any, _
            ByRef Source As Any, _
            ByVal Length As LongPtr)
    #Else
        Private Declare PtrSafe Function AccessibleObjectFromPoint Lib "oleacc" ( _
            ByVal pointX As Long, _
            ByVal pointY As Long, _
            ByRef ppacc As Object, _
            ByRef pvarChild As Variant) As Long
    #End If
#Else
    Private Declare Function GetCursorPos Lib "user32" ( _
        ByRef lpPoint As POINTAPI) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" ( _
        ByVal vKey As Long) As Integer
    Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
    Private Declare Function AccessibleObjectFromPoint Lib "oleacc" ( _
        ByVal pointX As Long, _
        ByVal pointY As Long, _
        ByRef ppacc As Object, _
        ByRef pvarChild As Variant) As Long
#End If

' キー
Private Const VK_ESCAPE As Long = 27
Private Const VK_F8 As Long = 119
Private Const VK_F9 As Long = 120

' HRESULT
Private Const S_OK As Long = 0

' TreeScope
Private Const SCOPE_CHILDREN As Long = 2
Private Const SCOPE_SUBTREE As Long = 7

' UIA Property IDs
Private Const UIA_PROCESS_ID_PROPERTY_ID As Long = 30002
Private Const UIA_CONTROL_TYPE_PROPERTY_ID As Long = 30003
Private Const UIA_NAME_PROPERTY_ID As Long = 30005
Private Const UIA_AUTOMATION_ID_PROPERTY_ID As Long = 30011
Private Const UIA_CLASS_NAME_PROPERTY_ID As Long = 30012

' UIA Pattern IDs
Private Const UIA_INVOKE_PATTERN_ID As Long = 10000
Private Const UIA_VALUE_PATTERN_ID As Long = 10002
Private Const UIA_EXPAND_COLLAPSE_PATTERN_ID As Long = 10005
Private Const UIA_SELECTION_ITEM_PATTERN_ID As Long = 10010
Private Const UIA_TEXT_PATTERN_ID As Long = 10014
Private Const UIA_TOGGLE_PATTERN_ID As Long = 10015

' ControlType IDs
Private Const UIA_CT_ANY As Long = 0
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

' ToggleState / ExpandCollapseState
Private Const TOGGLE_OFF As Long = 0
Private Const TOGGLE_ON As Long = 1
Private Const EXPAND_COLLAPSED As Long = 0
Private Const EXPAND_EXPANDED As Long = 1

' シート名
Private Const SHEET_OPERATIONS As String = "UIA_操作"
Private Const SHEET_LOG As String = "UIA_ログ"

' UIA_操作シートの列
Private Const C_ENABLED As Long = 1
Private Const C_STEP As Long = 2
Private Const C_KEY As Long = 3
Private Const C_WINDOW As Long = 4
Private Const C_METHOD As Long = 5
Private Const C_CONTAINER_ID As Long = 6
Private Const C_AUTOMATION_ID As Long = 7
Private Const C_NAME As Long = 8
Private Const C_CONTROL_TYPE_ID As Long = 9
Private Const C_CONTROL_TYPE As Long = 10
Private Const C_CLASS_NAME As Long = 11
Private Const C_OCCURRENCE As Long = 12
Private Const C_MATCH_COUNT As Long = 13
Private Const C_PATTERNS As Long = 14
Private Const C_ACTION As Long = 15
Private Const C_VALUE_SOURCE As Long = 16
Private Const C_DESTINATION As Long = 17
Private Const C_TIMEOUT As Long = 18
Private Const C_CLEAR_CACHE As Long = 19
Private Const C_CONFIRM As Long = 20
Private Const C_STATUS As Long = 21
Private Const C_CAPTURED_AT As Long = 22
Private Const C_CAPTURE_SOURCE As Long = 23
Private Const C_PROCESS_ID As Long = 24

Private mUIA As CUIAutomation
Private mElementCache As Object
Private mWindowCache As Object
Private mLastError As String


' =====================================================================
' メインメニュー
' =====================================================================
Public Sub UIA_メニュー()

    Dim choice As Variant

    EnsureSheets

    choice = Application.InputBox( _
        Prompt:= _
            "実行する番号を入力してください。" & vbCrLf & vbCrLf & _
            "1  要素を連続記録（F8マウス／F9直接／Esc終了）" & vbCrLf & _
            "2  選択行の要素を検索テスト" & vbCrLf & _
            "3  選択行だけ実行" & vbCrLf & _
            "4  有効な全行を順番に実行" & vbCrLf & _
            "5  実行キャッシュを消去" & vbCrLf & _
            "6  操作・ログを初期化", _
        Title:="UI Automation レコーダー／実行", _
        Default:=1, _
        Type:=1)

    If VarType(choice) = vbBoolean Then Exit Sub

    Select Case CLng(choice)
        Case 1
            UIA_要素記録開始
        Case 2
            UIA_選択行を確認
        Case 3
            UIA_選択行を実行
        Case 4
            UIA_全行を実行
        Case 5
            UIA_実行キャッシュ消去
            MsgBox "実行キャッシュを消去しました。", vbInformation
        Case 6
            UIA_記録を初期化
        Case Else
            MsgBox "1～6を指定してください。", vbExclamation
    End Select

End Sub


' =====================================================================
' 参照設定と基本構文を確認するための簡易診断
' エラーが出なければ、必要シートとUIAオブジェクトを作成できます。
' =====================================================================
Public Sub UIA_自己診断()

    EnsureSheets
    EnsureRuntime
    ThisWorkbook.Worksheets(SHEET_OPERATIONS).Range("A1").Value = "Enabled"
    ThisWorkbook.Worksheets(SHEET_OPERATIONS).Range("U1").Value = "Status"

End Sub


' =====================================================================
' F8でマウス位置、F9でフォーカス中のUIA要素を直接登録
' =====================================================================
Public Sub UIA_要素記録開始()

    Dim wasF8Down As Boolean
    Dim wasF9Down As Boolean
    Dim f8Down As Boolean
    Dim f9Down As Boolean
    Dim captureRequested As Boolean
    Dim useFocusedElement As Boolean
    Dim captureCount As Long
    Dim ws As Worksheet
    Dim previousCancelKey As XlEnableCancelKey

    EnsureSheets
    EnsureRuntime
    Set ws = ThisWorkbook.Worksheets(SHEET_OPERATIONS)
    ws.Activate

    MsgBox _
        "記録モードを開始します。" & vbCrLf & vbCrLf & _
        "1. 対象画面へ移動します。" & vbCrLf & _
        "2. 次のどちらかで登録します。" & vbCrLf & _
        "   F8：対象へマウスを合わせて取得" & vbCrLf & _
        "   F9：フォーカス中のUIA要素を直接取得" & vbCrLf & _
        "3. 次の要素へ移動して同様に登録します。" & vbCrLf & _
        "4. 終了するときは Esc を押します。" & vbCrLf & vbCrLf & _
        "ボタンやチェックボックスをF9で取る場合は、" & vbCrLf & _
        "クリックせずTabキーでフォーカスを移すと安全です。", _
        vbInformation, "UIA要素記録"

    previousCancelKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled
    Application.StatusBar = _
        "UIA記録中：F8=マウス／F9=フォーカス要素を直接取得／Esc=終了"

    Do While KeyIsDown(VK_F8) Or KeyIsDown(VK_F9)
        DoEvents
        Sleep 30
    Loop

    On Error GoTo RecordError

    Do
        If KeyIsDown(VK_ESCAPE) Then Exit Do

        f8Down = KeyIsDown(VK_F8)
        f9Down = KeyIsDown(VK_F9)
        captureRequested = False
        useFocusedElement = False

        ' 同時押しの場合はUIA直接取得（F9）を優先する
        If f9Down And Not wasF9Down Then
            captureRequested = True
            useFocusedElement = True
        ElseIf f8Down And Not wasF8Down Then
            captureRequested = True
        End If

        If captureRequested Then
            If CaptureCurrentElement(ws, useFocusedElement) Then
                captureCount = captureCount + 1
                Beep
                Application.StatusBar = _
                    "UIA記録中：" & captureCount & _
                    "件登録済み／F8=マウス／F9=直接／Esc=終了"
            Else
                Beep
                Beep
            End If
        End If

        wasF8Down = f8Down
        wasF9Down = f9Down

        DoEvents
        Sleep 30
    Loop

RecordExit:
    Application.EnableCancelKey = previousCancelKey
    Application.StatusBar = False
    ws.Activate

    If captureCount > 0 Then
        MsgBox captureCount & "件の要素を登録しました。" & vbCrLf & _
               "UIA_操作シートでActionと値を確認してください。", _
               vbInformation
    End If
    Exit Sub

RecordError:
    Application.StatusBar = False
    MsgBox "記録中にエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbExclamation
    Resume RecordExit

End Sub


' =====================================================================
' 選択行が現在の画面で見つかるか確認（操作はしない）
' =====================================================================
Public Sub UIA_選択行を確認()

    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim elm As IUIAutomationElement

    If Not GetSelectedOperationRow(ws, rowNumber) Then Exit Sub

    EnsureRuntime
    UIA_実行キャッシュ消去

    On Error GoTo Failed
    Set elm = WaitRecordedElement(ws, rowNumber, GetTimeout(ws, rowNumber))

    If elm Is Nothing Then
        Err.Raise vbObjectError + 3000, , "指定要素が見つかりません。"
    End If

    ws.Cells(rowNumber, C_STATUS).Value = "確認OK"
    MsgBox _
        "要素を確認できました。" & vbCrLf & vbCrLf & _
        "Key: " & ws.Cells(rowNumber, C_KEY).Value & vbCrLf & _
        "Name: " & SafeName(elm) & vbCrLf & _
        "AutomationId: " & SafeAutomationId(elm) & vbCrLf & _
        "ControlType: " & ControlTypeName(SafeControlType(elm)) & vbCrLf & _
        "ClassName: " & SafeClassName(elm), _
        vbInformation
    Exit Sub

Failed:
    ws.Cells(rowNumber, C_STATUS).Value = "確認NG: " & Err.Description
    MsgBox "要素を確認できませんでした。" & vbCrLf & _
           Err.Description, vbExclamation

End Sub


' =====================================================================
' 選択行だけ実行
' =====================================================================
Public Sub UIA_選択行を実行()

    Dim ws As Worksheet
    Dim rowNumber As Long
    Dim runId As String

    If Not GetSelectedOperationRow(ws, rowNumber) Then Exit Sub

    EnsureRuntime
    UIA_実行キャッシュ消去
    runId = Format$(Now, "yyyymmdd_HHmmss")

    If RunOperationRow(ws, rowNumber, runId, True) Then
        MsgBox "選択行を実行しました。", vbInformation
    End If

End Sub


' =====================================================================
' 有効な行をStep順に実行。エラー時はそこで停止。
' =====================================================================
Public Sub UIA_全行を実行()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rows() As Long
    Dim steps() As Double
    Dim count As Long
    Dim i As Long, j As Long
    Dim tempRow As Long
    Dim tempStep As Double
    Dim runId As String

    On Error GoTo UnexpectedError

    EnsureSheets
    EnsureRuntime
    UIA_実行キャッシュ消去

    Set ws = ThisWorkbook.Worksheets(SHEET_OPERATIONS)
    lastRow = ws.Cells(ws.Rows.Count, C_KEY).End(xlUp).Row

    For i = 2 To lastRow
        If IsYes(ws.Cells(i, C_ENABLED).Value) And _
           Len(Trim$(CStr(ws.Cells(i, C_ACTION).Value))) > 0 Then
            count = count + 1
            ReDim Preserve rows(1 To count)
            ReDim Preserve steps(1 To count)
            rows(count) = i
            steps(count) = Val(ws.Cells(i, C_STEP).Value)
        End If
    Next i

    If count = 0 Then
        MsgBox "実行対象の行がありません。", vbExclamation
        Exit Sub
    End If

    ' 件数は通常少ないため、単純な並べ替えでStep順にする
    For i = 1 To count - 1
        For j = i + 1 To count
            If steps(j) < steps(i) Then
                tempStep = steps(i): steps(i) = steps(j): steps(j) = tempStep
                tempRow = rows(i): rows(i) = rows(j): rows(j) = tempRow
            End If
        Next j
    Next i

    If MsgBox(count & "件をStep順に実行します。" & vbCrLf & _
              "テスト環境であることを確認してください。", _
              vbOKCancel + vbExclamation, "UIA一括実行") <> vbOK Then
        Exit Sub
    End If

    runId = Format$(Now, "yyyymmdd_HHmmss")

    For i = 1 To count
        Application.StatusBar = _
            "UIA実行中 " & i & "/" & count & _
            "  Key=" & ws.Cells(rows(i), C_KEY).Value

        If Not RunOperationRow(ws, rows(i), runId, False) Then
            Application.StatusBar = False
            MsgBox _
                "処理を停止しました。" & vbCrLf & _
                "Step: " & ws.Cells(rows(i), C_STEP).Value & vbCrLf & _
                "Key: " & ws.Cells(rows(i), C_KEY).Value & vbCrLf & _
                "内容: " & mLastError, _
                vbExclamation
            Exit Sub
        End If
    Next i

    Application.StatusBar = False
    MsgBox count & "件の処理が完了しました。", vbInformation
    Exit Sub

UnexpectedError:
    Application.StatusBar = False
    MsgBox "一括実行中に予期しないエラーが発生しました。" & vbCrLf & _
           Err.Number & ": " & Err.Description, vbExclamation

End Sub


' =====================================================================
' 操作・ログシートを初期化
' =====================================================================
Public Sub UIA_記録を初期化()

    Dim ws As Worksheet

    If MsgBox("記録した操作とログをすべて削除します。", _
              vbYesNo + vbExclamation, "初期化確認") <> vbYes Then Exit Sub

    Set ws = EnsureWorksheet(SHEET_OPERATIONS)
    ws.Cells.Clear
    Set ws = EnsureWorksheet(SHEET_LOG)
    ws.Cells.Clear

    SetupOperationsSheet ThisWorkbook.Worksheets(SHEET_OPERATIONS)
    SetupLogSheet ThisWorkbook.Worksheets(SHEET_LOG)
    UIA_実行キャッシュ消去

    ThisWorkbook.Worksheets(SHEET_OPERATIONS).Activate
    MsgBox "初期化しました。", vbInformation

End Sub


' =====================================================================
' マウス位置またはフォーカス中のUIA要素を1件記録
' =====================================================================
Private Function CaptureCurrentElement( _
    ByVal ws As Worksheet, _
    ByVal useFocusedElement As Boolean) As Boolean

    Dim rawElement As IUIAutomationElement
    Dim target As IUIAutomationElement
    Dim windowElement As IUIAutomationElement
    Dim containerElement As IUIAutomationElement
    Dim captureSource As String
    Dim method As String
    Dim matchCount As Long
    Dim occurrence As Long
    Dim nextRow As Long
    Dim nextStep As Long
    Dim elementKey As String
    Dim automationId As String
    Dim elementName As String
    Dim className As String
    Dim containerId As String
    Dim controlTypeId As Long
    Dim patterns As String
    Dim defaultAction As String
    Dim windowTitle As String

    On Error GoTo Failed

    If useFocusedElement Then
        Set rawElement = FocusedElementDirect(captureSource)
    Else
        Set rawElement = ElementAtCursor(captureSource)
    End If

    If rawElement Is Nothing Then
        If useFocusedElement Then
            Err.Raise vbObjectError + 3010, , _
                      "フォーカス中のUIA要素を取得できませんでした。"
        Else
            Err.Raise vbObjectError + 3010, , _
                      "マウス位置の要素を取得できませんでした。"
        End If
    End If

    Set target = NormalizeActionElement(rawElement)
    GetElementContext target, windowElement, containerElement

    If windowElement Is Nothing Then
        Err.Raise vbObjectError + 3011, , _
                  "対象要素が属するウィンドウを特定できませんでした。"
    End If

    windowTitle = SafeName(windowElement)
    If Len(windowTitle) = 0 Then
        Err.Raise vbObjectError + 3012, , _
                  "ウィンドウ名が空のため、安全に再特定できません。"
    End If

    automationId = SafeAutomationId(target)
    elementName = SafeName(target)
    className = SafeClassName(target)
    controlTypeId = SafeControlType(target)
    patterns = SupportedPatterns(target)

    If Not containerElement Is Nothing Then
        containerId = SafeAutomationId(containerElement)
    End If

    ChooseSelector _
        target, windowElement, containerElement, _
        automationId, elementName, controlTypeId, className, _
        method, matchCount, occurrence

    nextRow = ws.Cells(ws.Rows.Count, C_KEY).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    nextStep = NextStepNumber(ws)
    elementKey = "E" & Format$(nextStep, "000")
    defaultAction = DefaultActionFromPatterns(patterns, controlTypeId)

    ws.Cells(nextRow, C_ENABLED).Value = "○"
    ws.Cells(nextRow, C_STEP).Value = nextStep
    ws.Cells(nextRow, C_KEY).Value = elementKey
    WriteCellText ws.Cells(nextRow, C_WINDOW), windowTitle
    ws.Cells(nextRow, C_METHOD).Value = method
    WriteCellText ws.Cells(nextRow, C_CONTAINER_ID), containerId
    WriteCellText ws.Cells(nextRow, C_AUTOMATION_ID), automationId
    WriteCellText ws.Cells(nextRow, C_NAME), elementName
    ws.Cells(nextRow, C_CONTROL_TYPE_ID).Value = controlTypeId
    ws.Cells(nextRow, C_CONTROL_TYPE).Value = ControlTypeName(controlTypeId)
    WriteCellText ws.Cells(nextRow, C_CLASS_NAME), className
    ws.Cells(nextRow, C_OCCURRENCE).Value = occurrence
    ws.Cells(nextRow, C_MATCH_COUNT).Value = matchCount
    ws.Cells(nextRow, C_PATTERNS).Value = patterns
    ws.Cells(nextRow, C_ACTION).Value = defaultAction
    ws.Cells(nextRow, C_TIMEOUT).Value = 10
    ws.Cells(nextRow, C_CLEAR_CACHE).Value = _
        IIf(defaultAction = "CLICK", "○", "")
    ws.Cells(nextRow, C_CONFIRM).Value = _
        IIf(defaultAction = "CLICK", "○", "")
    ws.Cells(nextRow, C_STATUS).Value = "記録済み"
    ws.Cells(nextRow, C_CAPTURED_AT).Value = Now
    ws.Cells(nextRow, C_CAPTURE_SOURCE).Value = captureSource
    ws.Cells(nextRow, C_PROCESS_ID).Value = SafeProcessId(windowElement)

    If matchCount = 1 And _
       (captureSource = "MOUSE" Or _
        captureSource = "UIA_FOCUSED_DIRECT") Then
        ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, C_PROCESS_ID)).Interior.Color = _
            RGB(226, 239, 218)
    Else
        ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, C_PROCESS_ID)).Interior.Color = _
            RGB(255, 242, 204)
    End If

    CaptureCurrentElement = True
    Exit Function

Failed:
    Application.StatusBar = "UIA記録失敗: " & Err.Description
    CaptureCurrentElement = False

End Function


' =====================================================================
' 現在キーボードフォーカスを持つUIA要素を直接取得
' =====================================================================
Private Function FocusedElementDirect( _
    ByRef captureSource As String) As IUIAutomationElement

    Dim elm As IUIAutomationElement

    EnsureRuntime

    On Error Resume Next
    Set elm = mUIA.GetFocusedElement
    On Error GoTo 0

    If Not elm Is Nothing Then
        captureSource = "UIA_FOCUSED_DIRECT"
        Set FocusedElementDirect = elm
    End If

End Function


' =====================================================================
' マウス座標からMSAA要素を取得し、UIA要素へ変換
' ElementFromPointのVBA構造体値渡し問題を避ける方式
' =====================================================================
Private Function ElementAtCursor( _
    ByRef captureSource As String) As IUIAutomationElement

    Dim pt As POINTAPI
    Dim accessible As Object
    Dim childVariant As Variant
    Dim childId As Long
    Dim result As Long
    Dim elm As IUIAutomationElement
    Dim usedParentFallback As Boolean

    EnsureRuntime

    If GetCursorPos(pt) = 0 Then Exit Function

    #If VBA7 And Win64 Then
        Dim packedPoint As LongPtr
        CopyMemory packedPoint, pt, LenB(pt)
        result = AccessibleObjectFromPoint( _
            packedPoint, accessible, childVariant)
    #Else
        result = AccessibleObjectFromPoint( _
            pt.x, pt.y, accessible, childVariant)
    #End If

    If result = S_OK And Not accessible Is Nothing Then
        If IsNumeric(childVariant) Then childId = CLng(childVariant)

        On Error Resume Next
        Set elm = mUIA.ElementFromIAccessible(accessible, childId)
        If elm Is Nothing And childId <> 0 Then
            Err.Clear
            usedParentFallback = True
            Set elm = mUIA.ElementFromIAccessible(accessible, 0)
        End If
        On Error GoTo 0

        If Not elm Is Nothing Then
            If usedParentFallback Then
                captureSource = "MOUSE_PARENT_FALLBACK"
            Else
                captureSource = "MOUSE"
            End If
            Set ElementAtCursor = elm
            Exit Function
        End If
    End If

    ' 一部の画面でMSAAからUIAへ変換できない場合の補助
    On Error Resume Next
    Set elm = mUIA.GetFocusedElement
    On Error GoTo 0

    If Not elm Is Nothing Then
        captureSource = "FOCUS_FALLBACK"
        Set ElementAtCursor = elm
    End If

End Function


' =====================================================================
' Textの内側を指した場合、操作可能な近い親要素へ補正
' =====================================================================
Private Function NormalizeActionElement( _
    ByVal rawElement As IUIAutomationElement) As IUIAutomationElement

    Dim walker As IUIAutomationTreeWalker
    Dim current As IUIAutomationElement
    Dim parent As IUIAutomationElement
    Dim depth As Long

    If HasActionPattern(rawElement) Then
        Set NormalizeActionElement = rawElement
        Exit Function
    End If

    If SafeControlType(rawElement) <> UIA_CT_TEXT Then
        Set NormalizeActionElement = rawElement
        Exit Function
    End If

    Set walker = mUIA.ControlViewWalker
    Set current = rawElement

    For depth = 1 To 5
        On Error Resume Next
        Set parent = walker.GetParentElement(current)
        On Error GoTo 0

        If parent Is Nothing Then Exit For
        If SafeControlType(parent) = UIA_CT_WINDOW Then Exit For

        If HasActionPattern(parent) Then
            Set NormalizeActionElement = parent
            Exit Function
        End If

        Set current = parent
        Set parent = Nothing
    Next depth

    Set NormalizeActionElement = rawElement

End Function


' =====================================================================
' 対象ウィンドウと、最も近いAutomationId付き親コンテナを取得
' =====================================================================
Private Sub GetElementContext( _
    ByVal target As IUIAutomationElement, _
    ByRef windowElement As IUIAutomationElement, _
    ByRef containerElement As IUIAutomationElement)

    Dim walker As IUIAutomationTreeWalker
    Dim current As IUIAutomationElement
    Dim parent As IUIAutomationElement
    Dim depth As Long

    Set walker = mUIA.ControlViewWalker
    Set current = target

    For depth = 1 To 40
        On Error Resume Next
        Set parent = walker.GetParentElement(current)
        On Error GoTo 0

        If parent Is Nothing Then Exit For

        If SafeControlType(parent) = UIA_CT_WINDOW Then
            Set windowElement = parent
            Exit For
        End If

        If containerElement Is Nothing Then
            If Len(SafeAutomationId(parent)) > 0 Then
                Set containerElement = parent
            End If
        End If

        Set current = parent
        Set parent = Nothing
    Next depth

    If windowElement Is Nothing Then
        Set windowElement = FindWindowByProcessId(SafeProcessId(target))
    End If

End Sub


' =====================================================================
' 記録時に最も簡潔な検索条件を選択
' =====================================================================
Private Sub ChooseSelector( _
    ByVal target As IUIAutomationElement, _
    ByVal windowElement As IUIAutomationElement, _
    ByVal containerElement As IUIAutomationElement, _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    ByVal className As String, _
    ByRef method As String, _
    ByRef matchCount As Long, _
    ByRef occurrence As Long)

    Dim matches As IUIAutomationElementArray
    Dim containerMatches As IUIAutomationElementArray
    Dim containerId As String
    Dim containerIsUnique As Boolean

    method = ""
    matchCount = 0
    occurrence = 0

    If Not containerElement Is Nothing Then
        containerId = SafeAutomationId(containerElement)

        If Len(containerId) > 0 Then
            Set containerMatches = FindMatches( _
                windowElement, containerId, "", UIA_CT_ANY, "")

            If ArrayLength(containerMatches) = 1 Then
                On Error Resume Next
                containerIsUnique = mUIA.CompareElements( _
                    containerMatches.GetElement(0), containerElement)
                On Error GoTo 0
            End If
        End If
    End If

    If Len(automationId) > 0 Then
        Set matches = FindMatches( _
            windowElement, automationId, "", controlTypeId, "")

        If ArrayLength(matches) = 1 Then
            occurrence = OccurrenceOf(matches, target)

            If occurrence = 1 Then
                method = "AUTOID_TYPE"
                matchCount = 1
                Exit Sub
            End If
        End If

        If containerIsUnique Then
            Set matches = FindMatches( _
                containerElement, automationId, "", controlTypeId, "")

            If ArrayLength(matches) > 0 Then
                occurrence = OccurrenceOf(matches, target)
                If occurrence > 0 Then
                    method = "CONTAINER_AUTOID_TYPE"
                    matchCount = ArrayLength(matches)
                    Exit Sub
                End If
            End If
        End If

        Set matches = FindMatches( _
            windowElement, automationId, "", controlTypeId, "")
        matchCount = ArrayLength(matches)
        occurrence = OccurrenceOf(matches, target)

        If matchCount > 0 And occurrence > 0 Then
            method = "AUTOID_TYPE"
            Exit Sub
        End If
    End If

    If Len(elementName) > 0 Then
        Set matches = FindMatches( _
            windowElement, "", elementName, controlTypeId, "")

        If ArrayLength(matches) = 1 Then
            occurrence = OccurrenceOf(matches, target)

            If occurrence = 1 Then
                method = "NAME_TYPE"
                matchCount = 1
                Exit Sub
            End If
        End If

        If containerIsUnique Then
            Set matches = FindMatches( _
                containerElement, "", elementName, controlTypeId, "")

            If ArrayLength(matches) > 0 Then
                occurrence = OccurrenceOf(matches, target)
                If occurrence > 0 Then
                    method = "CONTAINER_NAME_TYPE"
                    matchCount = ArrayLength(matches)
                    Exit Sub
                End If
            End If
        End If

        If Len(className) > 0 Then
            Set matches = FindMatches( _
                windowElement, "", elementName, controlTypeId, className)

            If ArrayLength(matches) > 0 Then
                occurrence = OccurrenceOf(matches, target)
                If occurrence > 0 Then
                    method = "NAME_TYPE_CLASS"
                    matchCount = ArrayLength(matches)
                    Exit Sub
                End If
            End If
        End If

        Set matches = FindMatches( _
            windowElement, "", elementName, controlTypeId, "")
        matchCount = ArrayLength(matches)
        occurrence = OccurrenceOf(matches, target)

        If matchCount > 0 And occurrence > 0 Then
            method = "NAME_TYPE"
            Exit Sub
        End If
    End If

    If controlTypeId <= 0 Or Len(className) = 0 Then
        Err.Raise vbObjectError + 3020, , _
                  "安全に再特定できるAutomationId・Name・ClassNameがありません。"
    End If

    Set matches = FindMatches( _
        windowElement, "", "", controlTypeId, className)
    matchCount = ArrayLength(matches)
    occurrence = OccurrenceOf(matches, target)

    If matchCount = 0 Or occurrence = 0 Then
        Err.Raise vbObjectError + 3021, , _
                  "記録した要素を安全な検索条件へ変換できませんでした。"
    End If

    method = "TYPE_CLASS"

End Sub


' =====================================================================
' 操作を1行実行
' =====================================================================
Private Function RunOperationRow( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal runId As String, _
    ByVal interactive As Boolean) As Boolean

    Dim elm As IUIAutomationElement
    Dim actionName As String
    Dim valueSpec As String
    Dim actualValue As String
    Dim destination As String
    Dim elementKey As String
    Dim detail As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    actionName = UCase$(Trim$(CStr(ws.Cells(rowNumber, C_ACTION).Value)))
    elementKey = CStr(ws.Cells(rowNumber, C_KEY).Value)
    valueSpec = CStr(ws.Cells(rowNumber, C_VALUE_SOURCE).Value)
    destination = CStr(ws.Cells(rowNumber, C_DESTINATION).Value)

    If Len(actionName) = 0 Then
        Err.Raise vbObjectError + 3100, , "Actionが空です。"
    End If

    If IsYes(ws.Cells(rowNumber, C_CONFIRM).Value) Then
        If MsgBox( _
            "次の操作を実行します。" & vbCrLf & vbCrLf & _
            "Step: " & ws.Cells(rowNumber, C_STEP).Value & vbCrLf & _
            "Key: " & elementKey & vbCrLf & _
            "Name: " & ws.Cells(rowNumber, C_NAME).Value & vbCrLf & _
            "Action: " & actionName, _
            vbYesNo + vbExclamation, "操作確認") <> vbYes Then

            ws.Cells(rowNumber, C_STATUS).Value = "スキップ"
            WriteLog runId, ws, rowNumber, "SKIP", "利用者が中止"
            RunOperationRow = True
            Exit Function
        End If
    End If

    ws.Cells(rowNumber, C_STATUS).Value = "実行中"
    Set elm = WaitRecordedElement(ws, rowNumber, GetTimeout(ws, rowNumber))

    If elm Is Nothing Then
        Err.Raise vbObjectError + 3101, , "対象要素が見つかりません。"
    End If

    Select Case actionName
        Case "SET_TEXT"
            actualValue = ResolveInputValue(valueSpec)
            SetElementValue elm, actualValue
            detail = "入力完了（文字数=" & Len(actualValue) & "）"

        Case "CHECK_ON"
            SetCheckState elm, True
            detail = "ON"

        Case "CHECK_OFF"
            SetCheckState elm, False
            detail = "OFF"

        Case "CLICK"
            InvokeElement elm
            detail = "Invoke"

        Case "SELECT"
            SelectElement elm
            detail = "Select"

        Case "EXPAND"
            SetExpandState elm, True
            detail = "Expand"

        Case "COLLAPSE"
            SetExpandState elm, False
            detail = "Collapse"

        Case "READ_VALUE"
            actualValue = ReadElementValue(elm)
            WriteDestination destination, actualValue
            detail = "取得完了（文字数=" & Len(actualValue) & "）"

        Case "READ_TEXT"
            actualValue = ReadElementText(elm)
            WriteDestination destination, actualValue
            detail = "文字数=" & Len(actualValue)

        Case "READ_NAME"
            actualValue = SafeName(elm)
            WriteDestination destination, actualValue
            detail = "取得完了（文字数=" & Len(actualValue) & "）"

        Case "WAIT"
            detail = "要素出現を確認"

        Case Else
            Err.Raise vbObjectError + 3102, , _
                      "未対応のActionです: " & actionName
    End Select

    If IsYes(ws.Cells(rowNumber, C_CLEAR_CACHE).Value) Then
        UIA_実行キャッシュ消去
    End If

    ws.Cells(rowNumber, C_STATUS).Value = "成功"
    WriteLog runId, ws, rowNumber, "SUCCESS", detail
    RunOperationRow = True
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    mLastError = errorNumber & ": " & errorDescription

    ' エラー表示やログ書込み自体が失敗しても、元のエラーを失わない
    On Error Resume Next
    ws.Cells(rowNumber, C_STATUS).Value = "エラー: " & errorDescription
    WriteLog runId, ws, rowNumber, "ERROR", mLastError

    If interactive Then
        MsgBox "実行できませんでした。" & vbCrLf & _
               mLastError, vbExclamation
    End If

    On Error GoTo 0

    RunOperationRow = False

End Function


' =====================================================================
' 記録された条件から要素を待って取得
' =====================================================================
Private Function WaitRecordedElement( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal timeoutSeconds As Double) As IUIAutomationElement

    Dim startedAt As Double
    Dim elm As IUIAutomationElement
    Dim cacheKey As String

    EnsureRuntime
    cacheKey = RecordedElementCacheKey(ws, rowNumber)

    Set elm = GetAliveCachedElement(cacheKey)
    If Not elm Is Nothing Then
        Set WaitRecordedElement = elm
        Exit Function
    End If

    startedAt = Timer

    Do
        Set elm = FindRecordedElementOnce(ws, rowNumber)

        If Not elm Is Nothing Then
            If mElementCache.Exists(cacheKey) Then mElementCache.Remove cacheKey
            mElementCache.Add cacheKey, elm
            Set WaitRecordedElement = elm
            Exit Function
        End If

        DoEvents
        Sleep 50
    Loop While ElapsedSeconds(startedAt) < timeoutSeconds

End Function


Private Function RecordedElementCacheKey( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long) As String

    RecordedElementCacheKey = _
        CStr(rowNumber) & "|" & _
        CStr(ws.Cells(rowNumber, C_WINDOW).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_METHOD).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_CONTAINER_ID).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_AUTOMATION_ID).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_NAME).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_CONTROL_TYPE_ID).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_CLASS_NAME).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_OCCURRENCE).Value) & "|" & _
        CStr(ws.Cells(rowNumber, C_PROCESS_ID).Value)

End Function


' =====================================================================
' 記録されたセレクターで1回だけ直接検索
' =====================================================================
Private Function FindRecordedElementOnce( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long) As IUIAutomationElement

    Dim windowElement As IUIAutomationElement
    Dim searchRoot As IUIAutomationElement
    Dim container As IUIAutomationElement
    Dim cond As IUIAutomationCondition
    Dim matches As IUIAutomationElementArray
    Dim method As String
    Dim windowTitle As String
    Dim occurrence As Long
    Dim occurrenceValue As Variant

    windowTitle = CStr(ws.Cells(rowNumber, C_WINDOW).Value)
    method = UCase$(Trim$(CStr(ws.Cells(rowNumber, C_METHOD).Value)))
    occurrenceValue = ws.Cells(rowNumber, C_OCCURRENCE).Value

    If Not IsNumeric(occurrenceValue) Then
        Err.Raise vbObjectError + 3109, , _
                  "Occurrenceは1以上を指定してください。"
    End If

    If CDbl(occurrenceValue) < 1 Or _
       CDbl(occurrenceValue) <> Fix(CDbl(occurrenceValue)) Then
        Err.Raise vbObjectError + 3109, , _
                  "Occurrenceは1以上の整数を指定してください。"
    End If

    occurrence = CLng(occurrenceValue)

    Set windowElement = FindTopWindow( _
        windowTitle, Val(ws.Cells(rowNumber, C_PROCESS_ID).Value))
    If windowElement Is Nothing Then Exit Function
    Set searchRoot = windowElement

    If Left$(method, 10) = "CONTAINER_" Then
        Set container = FindByAutomationId( _
            windowElement, CStr(ws.Cells(rowNumber, C_CONTAINER_ID).Value))
        If container Is Nothing Then Exit Function
        Set searchRoot = container
    End If

    Set cond = ConditionFromRecordedRow(ws, rowNumber)

    If occurrence = 1 Then
        Set FindRecordedElementOnce = _
            searchRoot.FindFirst(SCOPE_SUBTREE, cond)
        Exit Function
    End If

    Set matches = searchRoot.FindAll(SCOPE_SUBTREE, cond)
    If ArrayLength(matches) >= occurrence Then
        Set FindRecordedElementOnce = matches.GetElement(occurrence - 1)
    End If

End Function


' =====================================================================
' 記録行から検索Conditionを作成
' =====================================================================
Private Function ConditionFromRecordedRow( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long) As IUIAutomationCondition

    Dim method As String
    Dim cond As IUIAutomationCondition
    Dim part As IUIAutomationCondition
    Dim controlTypeId As Long
    Dim automationId As String
    Dim elementName As String
    Dim className As String

    method = UCase$(Trim$(CStr(ws.Cells(rowNumber, C_METHOD).Value)))
    controlTypeId = Val(ws.Cells(rowNumber, C_CONTROL_TYPE_ID).Value)
    automationId = CStr(ws.Cells(rowNumber, C_AUTOMATION_ID).Value)
    elementName = CStr(ws.Cells(rowNumber, C_NAME).Value)
    className = CStr(ws.Cells(rowNumber, C_CLASS_NAME).Value)

    Select Case method
        Case "AUTOID_TYPE", "CONTAINER_AUTOID_TYPE"
            If Len(automationId) = 0 Then
                Err.Raise vbObjectError + 3110, , _
                          "AutomationIdが空です。"
            End If

        Case "NAME_TYPE", "CONTAINER_NAME_TYPE"
            If Len(elementName) = 0 Then
                Err.Raise vbObjectError + 3111, , _
                          "Nameが空です。"
            End If

        Case "NAME_TYPE_CLASS"
            If Len(elementName) = 0 Or Len(className) = 0 Then
                Err.Raise vbObjectError + 3112, , _
                          "NameまたはClassNameが空です。"
            End If

        Case "TYPE_CLASS"
            If Len(className) = 0 Then
                Err.Raise vbObjectError + 3113, , _
                          "ClassNameが空です。"
            End If

        Case Else
            Err.Raise vbObjectError + 3114, , _
                      "未対応のSelectorMethodです: " & method
    End Select

    If controlTypeId <= 0 Then
        Err.Raise vbObjectError + 3115, , _
                  "ControlTypeIdが未設定です。"
    End If

    Set cond = mUIA.CreateTrueCondition

    Select Case method
        Case "AUTOID_TYPE", "CONTAINER_AUTOID_TYPE"
            Set part = mUIA.CreatePropertyCondition( _
                UIA_AUTOMATION_ID_PROPERTY_ID, automationId)

        Case "NAME_TYPE", "CONTAINER_NAME_TYPE", "NAME_TYPE_CLASS"
            Set part = mUIA.CreatePropertyCondition( _
                UIA_NAME_PROPERTY_ID, elementName)

        Case "TYPE_CLASS"
            ' ControlTypeとClassNameだけで検索する
    End Select

    If Not part Is Nothing Then
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If controlTypeId > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_CONTROL_TYPE_PROPERTY_ID, controlTypeId)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If method = "NAME_TYPE_CLASS" Or method = "TYPE_CLASS" Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_CLASS_NAME_PROPERTY_ID, className)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    Set ConditionFromRecordedRow = cond

End Function


' =====================================================================
' 条件一致要素を取得（記録・一意性判定のときだけ使用）
' =====================================================================
Private Function FindMatches( _
    ByVal searchRoot As IUIAutomationElement, _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    ByVal className As String) As IUIAutomationElementArray

    Dim cond As IUIAutomationCondition
    Dim part As IUIAutomationCondition

    Set cond = mUIA.CreateTrueCondition

    If Len(automationId) > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_AUTOMATION_ID_PROPERTY_ID, automationId)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If Len(elementName) > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_NAME_PROPERTY_ID, elementName)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If controlTypeId > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_CONTROL_TYPE_PROPERTY_ID, controlTypeId)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If Len(className) > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_CLASS_NAME_PROPERTY_ID, className)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    Set FindMatches = searchRoot.FindAll(SCOPE_SUBTREE, cond)

End Function


Private Function OccurrenceOf( _
    ByVal matches As IUIAutomationElementArray, _
    ByVal target As IUIAutomationElement) As Long

    Dim i As Long

    If matches Is Nothing Then Exit Function

    For i = 0 To matches.Length - 1
        On Error Resume Next
        If mUIA.CompareElements(matches.GetElement(i), target) Then
            OccurrenceOf = i + 1
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next i

End Function


Private Function ArrayLength( _
    ByVal items As IUIAutomationElementArray) As Long

    If items Is Nothing Then Exit Function
    ArrayLength = items.Length

End Function


' =====================================================================
' ウィンドウ・コンテナ検索
' =====================================================================
Private Function FindTopWindow( _
    ByVal windowTitle As String, _
    Optional ByVal expectedProcessId As Long = 0) As IUIAutomationElement

    Dim cached As IUIAutomationElement
    Dim root As IUIAutomationElement
    Dim condType As IUIAutomationCondition
    Dim condName As IUIAutomationCondition
    Dim condAnd As IUIAutomationCondition
    Dim exactMatches As IUIAutomationElementArray
    Dim windows As IUIAutomationElementArray
    Dim elm As IUIAutomationElement
    Dim candidate As IUIAutomationElement
    Dim processCandidate As IUIAutomationElement
    Dim i As Long
    Dim currentName As String
    Dim matchCount As Long
    Dim processMatchCount As Long
    Dim cacheKey As String

    If Len(windowTitle) = 0 Then Exit Function
    cacheKey = windowTitle & "|" & CStr(expectedProcessId)

    Set cached = GetAliveCachedWindow(cacheKey)
    If Not cached Is Nothing Then
        Set FindTopWindow = cached
        Exit Function
    End If

    Set root = mUIA.GetRootElement
    Set condType = mUIA.CreatePropertyCondition( _
        UIA_CONTROL_TYPE_PROPERTY_ID, UIA_CT_WINDOW)
    Set condName = mUIA.CreatePropertyCondition( _
        UIA_NAME_PROPERTY_ID, windowTitle)
    Set condAnd = mUIA.CreateAndCondition(condType, condName)

    ' 完全一致でも同名ウィンドウがあり得るため、一意性を確認する
    Set exactMatches = root.FindAll(SCOPE_CHILDREN, condAnd)
    Set elm = PickUniqueWindow(exactMatches, expectedProcessId)

    If elm Is Nothing Then
        ' タイトルが少し変わった場合だけ部分一致。複数候補なら実行しない。
        Set windows = root.FindAll(SCOPE_CHILDREN, condType)

        For i = 0 To windows.Length - 1
            currentName = SafeName(windows.GetElement(i))

            If Len(currentName) > 0 Then
                If InStr(1, currentName, windowTitle, vbTextCompare) > 0 Or _
                   InStr(1, windowTitle, currentName, vbTextCompare) > 0 Then
                    matchCount = matchCount + 1
                    Set candidate = windows.GetElement(i)

                    If expectedProcessId > 0 And _
                       SafeProcessId(candidate) = expectedProcessId Then
                        processMatchCount = processMatchCount + 1
                        Set processCandidate = candidate
                    End If
                End If
            End If
        Next i

        If matchCount = 1 Then
            Set elm = candidate
        ElseIf processMatchCount = 1 Then
            Set elm = processCandidate
        End If
    End If

    If Not elm Is Nothing Then
        If mWindowCache.Exists(cacheKey) Then mWindowCache.Remove cacheKey
        mWindowCache.Add cacheKey, elm
        Set FindTopWindow = elm
    End If

End Function


Private Function PickUniqueWindow( _
    ByVal windows As IUIAutomationElementArray, _
    ByVal expectedProcessId As Long) As IUIAutomationElement

    Dim i As Long
    Dim processMatchCount As Long
    Dim processCandidate As IUIAutomationElement
    Dim candidate As IUIAutomationElement

    If windows Is Nothing Then Exit Function

    If windows.Length = 1 Then
        Set PickUniqueWindow = windows.GetElement(0)
        Exit Function
    End If

    If expectedProcessId <= 0 Then Exit Function

    For i = 0 To windows.Length - 1
        Set candidate = windows.GetElement(i)

        If SafeProcessId(candidate) = expectedProcessId Then
            processMatchCount = processMatchCount + 1
            Set processCandidate = candidate
        End If
    Next i

    If processMatchCount = 1 Then
        Set PickUniqueWindow = processCandidate
    End If

End Function


Private Function FindWindowByProcessId( _
    ByVal processId As Long) As IUIAutomationElement

    Dim root As IUIAutomationElement
    Dim condType As IUIAutomationCondition
    Dim condProcess As IUIAutomationCondition
    Dim condAnd As IUIAutomationCondition
    Dim matches As IUIAutomationElementArray

    If processId <= 0 Then Exit Function

    Set root = mUIA.GetRootElement
    Set condType = mUIA.CreatePropertyCondition( _
        UIA_CONTROL_TYPE_PROPERTY_ID, UIA_CT_WINDOW)
    Set condProcess = mUIA.CreatePropertyCondition( _
        UIA_PROCESS_ID_PROPERTY_ID, processId)
    Set condAnd = mUIA.CreateAndCondition(condType, condProcess)

    Set matches = root.FindAll(SCOPE_CHILDREN, condAnd)

    If ArrayLength(matches) = 1 Then
        Set FindWindowByProcessId = matches.GetElement(0)
    End If

End Function


Private Function FindByAutomationId( _
    ByVal searchRoot As IUIAutomationElement, _
    ByVal automationId As String) As IUIAutomationElement

    Dim cond As IUIAutomationCondition
    Dim matches As IUIAutomationElementArray

    If Len(automationId) = 0 Then Exit Function

    Set cond = mUIA.CreatePropertyCondition( _
        UIA_AUTOMATION_ID_PROPERTY_ID, automationId)
    Set matches = searchRoot.FindAll(SCOPE_SUBTREE, cond)

    If ArrayLength(matches) = 1 Then
        Set FindByAutomationId = matches.GetElement(0)
    End If

End Function


' =====================================================================
' UIAアクション
' =====================================================================
Private Sub SetElementValue( _
    ByVal elm As IUIAutomationElement, _
    ByVal newValue As String)

    Dim valuePattern As IUIAutomationValuePattern

    Set valuePattern = elm.GetCurrentPattern(UIA_VALUE_PATTERN_ID)
    If valuePattern.CurrentIsReadOnly Then
        Err.Raise vbObjectError + 3200, , "対象は読み取り専用です。"
    End If
    valuePattern.SetValue newValue

End Sub


Private Sub SetCheckState( _
    ByVal elm As IUIAutomationElement, _
    ByVal desiredOn As Boolean)

    Dim togglePattern As IUIAutomationTogglePattern
    Dim desiredState As Long
    Dim attempt As Long

    Set togglePattern = elm.GetCurrentPattern(UIA_TOGGLE_PATTERN_ID)
    desiredState = IIf(desiredOn, TOGGLE_ON, TOGGLE_OFF)

    ' 3状態チェックボックス（OFF→ON→中間など）にも対応する
    For attempt = 1 To 3
        If togglePattern.CurrentToggleState = desiredState Then Exit For
        togglePattern.Toggle
    Next attempt

    If togglePattern.CurrentToggleState <> desiredState Then
        Err.Raise vbObjectError + 3201, , _
                  "チェック状態を変更できませんでした。"
    End If

End Sub


Private Sub InvokeElement(ByVal elm As IUIAutomationElement)

    Dim invokePattern As IUIAutomationInvokePattern
    Set invokePattern = elm.GetCurrentPattern(UIA_INVOKE_PATTERN_ID)
    invokePattern.Invoke

End Sub


Private Sub SelectElement(ByVal elm As IUIAutomationElement)

    Dim selectionPattern As IUIAutomationSelectionItemPattern
    Set selectionPattern = elm.GetCurrentPattern( _
        UIA_SELECTION_ITEM_PATTERN_ID)
    selectionPattern.Select

End Sub


Private Sub SetExpandState( _
    ByVal elm As IUIAutomationElement, _
    ByVal desiredExpanded As Boolean)

    Dim expandPattern As IUIAutomationExpandCollapsePattern

    Set expandPattern = elm.GetCurrentPattern( _
        UIA_EXPAND_COLLAPSE_PATTERN_ID)

    If desiredExpanded Then
        If expandPattern.CurrentExpandCollapseState <> EXPAND_EXPANDED Then
            expandPattern.Expand
        End If
    Else
        If expandPattern.CurrentExpandCollapseState <> EXPAND_COLLAPSED Then
            expandPattern.Collapse
        End If
    End If

    If desiredExpanded Then
        If expandPattern.CurrentExpandCollapseState <> EXPAND_EXPANDED Then
            Err.Raise vbObjectError + 3202, , _
                      "要素を展開状態にできませんでした。"
        End If
    Else
        If expandPattern.CurrentExpandCollapseState <> EXPAND_COLLAPSED Then
            Err.Raise vbObjectError + 3203, , _
                      "要素を折りたたみ状態にできませんでした。"
        End If
    End If

End Sub


Private Function ReadElementValue( _
    ByVal elm As IUIAutomationElement) As String

    Dim valuePattern As IUIAutomationValuePattern
    Set valuePattern = elm.GetCurrentPattern(UIA_VALUE_PATTERN_ID)
    ReadElementValue = valuePattern.CurrentValue

End Function


Private Function ReadElementText( _
    ByVal elm As IUIAutomationElement) As String

    Dim textPattern As IUIAutomationTextPattern
    Dim textRange As IUIAutomationTextRange

    Set textPattern = elm.GetCurrentPattern(UIA_TEXT_PATTERN_ID)
    Set textRange = textPattern.DocumentRange
    ReadElementText = textRange.GetText(-1)

End Function


' =====================================================================
' 入出力セル指定
' ValueOrSource:
'   文字列           → そのまま入力
'   @Sheet1!A2       → セルの値を入力
'   @@abc            → @abcという文字列を入力
'
' Destination:
'   Sheet1!B2 または @Sheet1!B2
'   取得結果は単一セルへ文字列として書き込みます。
' =====================================================================
Private Function ResolveInputValue(ByVal valueSpec As String) As String

    If Left$(valueSpec, 2) = "@@" Then
        ResolveInputValue = Mid$(valueSpec, 2)
    ElseIf Left$(valueSpec, 1) = "@" Then
        ResolveInputValue = CStr(ReadCellReference(Mid$(valueSpec, 2)))
    Else
        ResolveInputValue = valueSpec
    End If

End Function


Private Sub WriteDestination( _
    ByVal destination As String, _
    ByVal valueToWrite As String)

    Dim target As Range

    If Left$(destination, 1) = "@" Then destination = Mid$(destination, 2)
    If Len(destination) = 0 Then
        Err.Raise vbObjectError + 3210, , "Destinationが空です。"
    End If

    Set target = GetCellReference(destination)

    If target.Cells.CountLarge <> 1 Then
        Err.Raise vbObjectError + 3212, , _
                  "Destinationは単一セルを指定してください。"
    End If

    ' Web等から取得した「=」始まりの文字列を数式として実行させない
    WriteCellText target, valueToWrite

End Sub


Private Sub WriteCellText( _
    ByVal target As Range, _
    ByVal textValue As String)

    target.NumberFormat = "@"
    target.Value2 = textValue

End Sub


Private Function ReadCellReference(ByVal referenceText As String) As Variant

    ReadCellReference = GetCellReference(referenceText).Value

End Function


Private Function GetCellReference( _
    ByVal referenceText As String) As Range

    Dim bangPosition As Long
    Dim sheetName As String
    Dim addressText As String

    bangPosition = InStrRev(referenceText, "!")
    If bangPosition <= 1 Then
        Err.Raise vbObjectError + 3211, , _
                  "セル指定は Sheet1!A1 の形式にしてください: " & referenceText
    End If

    sheetName = Left$(referenceText, bangPosition - 1)
    addressText = Mid$(referenceText, bangPosition + 1)
    sheetName = Replace(sheetName, "'", "")

    Set GetCellReference = _
        ThisWorkbook.Worksheets(sheetName).Range(addressText)

End Function


' =====================================================================
' パターン・既定Action
' =====================================================================
Private Function SupportedPatterns( _
    ByVal elm As IUIAutomationElement) As String

    Dim result As String

    AddPatternName result, elm, UIA_VALUE_PATTERN_ID, "Value"
    AddPatternName result, elm, UIA_TOGGLE_PATTERN_ID, "Toggle"
    AddPatternName result, elm, UIA_INVOKE_PATTERN_ID, "Invoke"
    AddPatternName result, elm, UIA_SELECTION_ITEM_PATTERN_ID, "SelectionItem"
    AddPatternName result, elm, UIA_EXPAND_COLLAPSE_PATTERN_ID, "ExpandCollapse"
    AddPatternName result, elm, UIA_TEXT_PATTERN_ID, "Text"

    SupportedPatterns = result

End Function


Private Sub AddPatternName( _
    ByRef result As String, _
    ByVal elm As IUIAutomationElement, _
    ByVal patternId As Long, _
    ByVal patternName As String)

    If HasPattern(elm, patternId) Then
        If Len(result) > 0 Then result = result & ","
        result = result & patternName
    End If

End Sub


Private Function HasPattern( _
    ByVal elm As IUIAutomationElement, _
    ByVal patternId As Long) As Boolean

    Dim patternObject As Object

    On Error GoTo NotAvailable
    Set patternObject = elm.GetCurrentPattern(patternId)
    HasPattern = Not patternObject Is Nothing
    Exit Function

NotAvailable:
    Err.Clear

End Function


Private Function HasActionPattern( _
    ByVal elm As IUIAutomationElement) As Boolean

    HasActionPattern = _
        HasPattern(elm, UIA_VALUE_PATTERN_ID) Or _
        HasPattern(elm, UIA_TOGGLE_PATTERN_ID) Or _
        HasPattern(elm, UIA_INVOKE_PATTERN_ID) Or _
        HasPattern(elm, UIA_SELECTION_ITEM_PATTERN_ID) Or _
        HasPattern(elm, UIA_EXPAND_COLLAPSE_PATTERN_ID)

End Function


Private Function DefaultActionFromPatterns( _
    ByVal patterns As String, _
    ByVal controlTypeId As Long) As String

    If InStr(patterns, "Value") > 0 Then
        DefaultActionFromPatterns = "SET_TEXT"
    ElseIf InStr(patterns, "Toggle") > 0 Then
        DefaultActionFromPatterns = "CHECK_ON"
    ElseIf InStr(patterns, "Invoke") > 0 Then
        DefaultActionFromPatterns = "CLICK"
    ElseIf InStr(patterns, "SelectionItem") > 0 Then
        DefaultActionFromPatterns = "SELECT"
    ElseIf InStr(patterns, "ExpandCollapse") > 0 Then
        DefaultActionFromPatterns = "EXPAND"
    ElseIf InStr(patterns, "Text") > 0 Or _
           controlTypeId = UIA_CT_TEXT Then
        DefaultActionFromPatterns = "READ_NAME"
    Else
        DefaultActionFromPatterns = "READ_NAME"
    End If

End Function


' =====================================================================
' シート作成・書式
' =====================================================================
Private Sub EnsureSheets()

    Dim operationSheet As Worksheet
    Dim logSheet As Worksheet

    Set operationSheet = EnsureWorksheet(SHEET_OPERATIONS)
    Set logSheet = EnsureWorksheet(SHEET_LOG)

    SetupOperationsSheet operationSheet
    SetupLogSheet logSheet

End Sub


Private Function EnsureWorksheet( _
    ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set EnsureWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If

End Function


Private Sub SetupOperationsSheet(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Enabled", "Step", "Key", "WindowTitle", "SelectorMethod", _
        "ContainerAutomationId", "AutomationId", "Name", _
        "ControlTypeId", "ControlType", "ClassName", "Occurrence", _
        "MatchCount", "Patterns", "Action", "ValueOrSource", _
        "Destination", "TimeoutSec", "ClearCache", "Confirm", _
        "Status", "CapturedAt", "CaptureSource", "ProcessId")

    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i

        With ws.Range(ws.Cells(1, 1), ws.Cells(1, C_PROCESS_ID))
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
        End With

        ws.Range(ws.Cells(1, 1), ws.Cells(1, C_PROCESS_ID)).AutoFilter
        ws.Columns(C_ENABLED).ColumnWidth = 9
        ws.Columns(C_STEP).ColumnWidth = 7
        ws.Columns(C_KEY).ColumnWidth = 10
        ws.Columns(C_WINDOW).ColumnWidth = 32
        ws.Columns(C_METHOD).ColumnWidth = 24
        ws.Columns(C_CONTAINER_ID).ColumnWidth = 22
        ws.Columns(C_AUTOMATION_ID).ColumnWidth = 22
        ws.Columns(C_NAME).ColumnWidth = 24
        ws.Columns(C_CONTROL_TYPE).ColumnWidth = 16
        ws.Columns(C_PATTERNS).ColumnWidth = 30
        ws.Columns(C_ACTION).ColumnWidth = 18
        ws.Columns(C_VALUE_SOURCE).ColumnWidth = 24
        ws.Columns(C_DESTINATION).ColumnWidth = 20
        ws.Columns(C_STATUS).ColumnWidth = 28
    End If

    ApplyValidations ws

End Sub


Private Sub ApplyValidations(ByVal ws As Worksheet)

    On Error Resume Next

    With ws.Range(ws.Cells(2, C_ACTION), ws.Cells(5000, C_ACTION)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="SET_TEXT,CHECK_ON,CHECK_OFF,CLICK,SELECT,EXPAND,COLLAPSE,READ_VALUE,READ_TEXT,READ_NAME,WAIT"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With ws.Range(ws.Cells(2, C_ENABLED), ws.Cells(5000, C_ENABLED)).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="○,×"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With ws.Range(ws.Cells(2, C_CLEAR_CACHE), ws.Cells(5000, C_CLEAR_CACHE)).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="○,×"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    With ws.Range(ws.Cells(2, C_CONFIRM), ws.Cells(5000, C_CONFIRM)).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="○,×"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    On Error GoTo 0

End Sub


Private Sub SetupLogSheet(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "RunId", "Time", "Step", "Key", "Action", "Result", "Detail")

    If Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i

        With ws.Range("A1:G1")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(68, 68, 68)
        End With

        ws.Columns("A:G").ColumnWidth = 20
        ws.Columns("G").ColumnWidth = 50
    End If

End Sub


Private Sub WriteLog( _
    ByVal runId As String, _
    ByVal operationSheet As Worksheet, _
    ByVal operationRow As Long, _
    ByVal result As String, _
    ByVal detail As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_LOG)
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(nextRow, 1).Value = runId
    ws.Cells(nextRow, 2).Value = Now
    ws.Cells(nextRow, 3).Value = operationSheet.Cells(operationRow, C_STEP).Value
    ws.Cells(nextRow, 4).Value = operationSheet.Cells(operationRow, C_KEY).Value
    ws.Cells(nextRow, 5).Value = operationSheet.Cells(operationRow, C_ACTION).Value
    ws.Cells(nextRow, 6).Value = result
    ws.Cells(nextRow, 7).Value = detail

End Sub


' =====================================================================
' 実行環境・キャッシュ
' =====================================================================
Private Sub EnsureRuntime()

    If mUIA Is Nothing Then Set mUIA = New CUIAutomation
    If mElementCache Is Nothing Then
        Set mElementCache = CreateObject("Scripting.Dictionary")
    End If
    If mWindowCache Is Nothing Then
        Set mWindowCache = CreateObject("Scripting.Dictionary")
    End If

End Sub


Public Sub UIA_実行キャッシュ消去()

    If Not mElementCache Is Nothing Then mElementCache.RemoveAll
    If Not mWindowCache Is Nothing Then mWindowCache.RemoveAll

End Sub


Private Function GetAliveCachedElement( _
    ByVal cacheKey As String) As IUIAutomationElement

    Dim elm As IUIAutomationElement
    Dim dummy As Long

    If mElementCache Is Nothing Then Exit Function
    If Not mElementCache.Exists(cacheKey) Then Exit Function

    Set elm = mElementCache.Item(cacheKey)

    On Error GoTo Stale
    dummy = elm.CurrentControlType
    Set GetAliveCachedElement = elm
    Exit Function

Stale:
    Err.Clear
    mElementCache.Remove cacheKey

End Function


Private Function GetAliveCachedWindow( _
    ByVal cacheKey As String) As IUIAutomationElement

    Dim elm As IUIAutomationElement
    Dim dummy As Long

    If mWindowCache Is Nothing Then Exit Function
    If Not mWindowCache.Exists(cacheKey) Then Exit Function

    Set elm = mWindowCache.Item(cacheKey)

    On Error GoTo Stale
    dummy = elm.CurrentProcessId
    Set GetAliveCachedWindow = elm
    Exit Function

Stale:
    Err.Clear
    mWindowCache.Remove cacheKey

End Function


' =====================================================================
' 共通補助
' =====================================================================
Private Function GetSelectedOperationRow( _
    ByRef ws As Worksheet, _
    ByRef rowNumber As Long) As Boolean

    EnsureSheets
    Set ws = ThisWorkbook.Worksheets(SHEET_OPERATIONS)

    If Not ActiveSheet Is ws Then
        MsgBox "UIA_操作シートの実行したい行を選択してください。", vbExclamation
        ws.Activate
        Exit Function
    End If

    rowNumber = ActiveCell.Row
    If rowNumber < 2 Or _
       Len(CStr(ws.Cells(rowNumber, C_KEY).Value)) = 0 Then
        MsgBox "実行したいデータ行を選択してください。", vbExclamation
        Exit Function
    End If

    GetSelectedOperationRow = True

End Function


Private Function NextStepNumber(ByVal ws As Worksheet) As Long

    Dim lastRow As Long
    Dim i As Long
    Dim maximumValue As Long

    lastRow = ws.Cells(ws.Rows.Count, C_STEP).End(xlUp).Row

    For i = 2 To lastRow
        If Val(ws.Cells(i, C_STEP).Value) > maximumValue Then
            maximumValue = Val(ws.Cells(i, C_STEP).Value)
        End If
    Next i

    NextStepNumber = maximumValue + 1

End Function


Private Function GetTimeout( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long) As Double

    GetTimeout = Val(ws.Cells(rowNumber, C_TIMEOUT).Value)
    If GetTimeout <= 0 Then GetTimeout = 10

End Function


Private Function IsYes(ByVal value As Variant) As Boolean

    Dim textValue As String
    textValue = UCase$(Trim$(CStr(value)))

    IsYes = (textValue = "○" Or textValue = "〇" Or _
             textValue = "1" Or textValue = "TRUE" Or _
             textValue = "YES" Or textValue = "ON")

End Function


Private Function KeyIsDown(ByVal virtualKey As Long) As Boolean

    KeyIsDown = (GetAsyncKeyState(virtualKey) < 0)

End Function


Private Function ElapsedSeconds(ByVal startedAt As Double) As Double

    Dim currentValue As Double
    currentValue = Timer

    If currentValue >= startedAt Then
        ElapsedSeconds = currentValue - startedAt
    Else
        ElapsedSeconds = (86400# - startedAt) + currentValue
    End If

End Function


' =====================================================================
' 安全なUIAプロパティ取得
' =====================================================================
Private Function SafeName(ByVal elm As IUIAutomationElement) As String
    On Error GoTo Failed
    SafeName = elm.CurrentName
    Exit Function
Failed:
    Err.Clear
End Function


Private Function SafeAutomationId(ByVal elm As IUIAutomationElement) As String
    On Error GoTo Failed
    SafeAutomationId = elm.CurrentAutomationId
    Exit Function
Failed:
    Err.Clear
End Function


Private Function SafeClassName(ByVal elm As IUIAutomationElement) As String
    On Error GoTo Failed
    SafeClassName = elm.CurrentClassName
    Exit Function
Failed:
    Err.Clear
End Function


Private Function SafeControlType(ByVal elm As IUIAutomationElement) As Long
    On Error GoTo Failed
    SafeControlType = elm.CurrentControlType
    Exit Function
Failed:
    Err.Clear
End Function


Private Function SafeProcessId(ByVal elm As IUIAutomationElement) As Long
    On Error GoTo Failed
    SafeProcessId = elm.CurrentProcessId
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
