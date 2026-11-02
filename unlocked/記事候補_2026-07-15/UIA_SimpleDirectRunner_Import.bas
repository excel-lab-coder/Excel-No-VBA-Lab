Attribute VB_Name = "UIA_SimpleDirectRunner"
Option Explicit

' =====================================================================
' UI Automation 直接実行・簡易版
' ---------------------------------------------------------------------
' UIA_SimpleFocusedInspectorで1件ずつ確認した要素を、
' BuildScenarioの設定部へAddScenarioStepとして追加し、上から実行します。
'
' 前提:
'   Windows版Excel専用
'   VBE「ツール」→「参照設定」→「UIAutomationClient」
'
' 実行するマクロ:
'   UIA_登録要素の一致数を確認   ← 最初に実行
'   UIA_登録要素を順番に実行
'
' 対応Action:
'   SET_TEXT / CLICK / CHECK_ON / CHECK_OFF / SELECT
'   EXPAND / COLLAPSE / WAIT（要素出現待ち）
'
' ValueOrSource:
'   文字列       → そのまま入力
'   @入力!A2     → 実行開始時のActiveWorkbookから取得
'   @@abc        → @abcという文字列を入力
'
' 重要:
'   ・マウス座標、SendKeys、全要素の手動列挙は使いません。
'   ・Occurrence=1はFindFirstで直接取得します。
'   ・AutomationIdが空ならName＋ControlTypeで直接検索します。
'   ・ClassNameが動的に見える場合は空欄にしてください。
' =====================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" ( _
        ByVal dwMilliseconds As Long)
#End If

' TreeScope
Private Const SCOPE_CHILDREN As Long = 2
Private Const SCOPE_SUBTREE As Long = 7

' UIA Property IDs
Private Const UIA_CONTROL_TYPE_PROPERTY_ID As Long = 30003
Private Const UIA_NAME_PROPERTY_ID As Long = 30005
Private Const UIA_AUTOMATION_ID_PROPERTY_ID As Long = 30011
Private Const UIA_CLASS_NAME_PROPERTY_ID As Long = 30012

' UIA Pattern IDs
Private Const UIA_INVOKE_PATTERN_ID As Long = 10000
Private Const UIA_VALUE_PATTERN_ID As Long = 10002
Private Const UIA_EXPAND_COLLAPSE_PATTERN_ID As Long = 10005
Private Const UIA_SELECTION_ITEM_PATTERN_ID As Long = 10010
Private Const UIA_TOGGLE_PATTERN_ID As Long = 10015

' ControlType / ToggleState
Private Const UIA_CT_WINDOW As Long = 50032
Private Const TOGGLE_OFF As Long = 0
Private Const TOGGLE_ON As Long = 1
Private Const EXPAND_COLLAPSED As Long = 0
Private Const EXPAND_EXPANDED As Long = 1

' 1ステップを表すArrayの添字
Private Const S_STEP_NAME As Long = 0
Private Const S_ACTION As Long = 1
Private Const S_WINDOW_KEYWORD As Long = 2
Private Const S_AUTOMATION_ID As Long = 3
Private Const S_ELEMENT_NAME As Long = 4
Private Const S_CONTROL_TYPE_ID As Long = 5
Private Const S_CLASS_NAME As Long = 6
Private Const S_OCCURRENCE As Long = 7
Private Const S_VALUE_SOURCE As Long = 8
Private Const S_TIMEOUT As Long = 9

Private mUIA As CUIAutomation
Private mDataBook As Workbook
Private mLastError As String


' =====================================================================
' 実行する要素を1件ずつ追加する設定部
' =====================================================================
Private Function BuildScenario() As Collection

    Dim steps As New Collection

    ' ---------------------------------------------------------------
    ' 現在取得できたYahoo! JAPAN検索欄の例
    ' 「ここに入力する文字」を実際の文字、または @入力!A2 に変更します。
    ' ---------------------------------------------------------------
    AddScenarioStep steps, _
        "Yahoo検索欄へ入力", _
        "SET_TEXT", _
        "Yahoo! JAPAN", _
        "", _
        "検索したいキーワードを入力してください", _
        50004, _
        "", _
        1, _
        "ここに入力する文字", _
        10

    ' 次の要素は、このAddScenarioStep一式を複製して下へ追加します。
    ' ClassNameがランダム文字列に見える場合は "" のままにします。

    Set BuildScenario = steps

End Function


Private Sub AddScenarioStep( _
    ByVal steps As Collection, _
    ByVal stepName As String, _
    ByVal actionName As String, _
    ByVal windowKeyword As String, _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    ByVal className As String, _
    ByVal occurrence As Long, _
    ByVal valueOrSource As String, _
    ByVal timeoutSeconds As Double)

    steps.Add Array( _
        stepName, UCase$(Trim$(actionName)), windowKeyword, _
        automationId, elementName, controlTypeId, className, _
        occurrence, valueOrSource, timeoutSeconds)

End Sub


' =====================================================================
' 公開マクロ：設定した全ステップを上から実行
' =====================================================================
Public Sub UIA_登録要素を順番に実行()

    Dim steps As Collection
    Dim stepData As Variant
    Dim i As Long
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo UnexpectedError

    Set mDataBook = Application.ActiveWorkbook

    If mDataBook Is Nothing Then
        Err.Raise vbObjectError + 4200, , _
                  "Excelブックを1つ開いてから実行してください。"
    End If

    If UCase$(mDataBook.Name) = "PERSONAL.XLSB" Or mDataBook.IsAddin Then
        Err.Raise vbObjectError + 4201, , _
                  "通常のExcelブックを前面にしてから実行してください。"
    End If

    EnsureUIA
    Set steps = BuildScenario

    If steps.Count = 0 Then
        MsgBox "BuildScenarioに実行ステップがありません。", vbExclamation
        Exit Sub
    End If

    If MsgBox( _
        steps.Count & "件のUIA操作を上から順に実行します。" & vbCrLf & _
        "テスト画面であることを確認してください。", _
        vbOKCancel + vbExclamation, "UIA直接実行") <> vbOK Then
        Exit Sub
    End If

    For i = 1 To steps.Count
        stepData = steps(i)
        Application.StatusBar = _
            "UIA実行中 " & i & "/" & steps.Count & _
            "：" & CStr(stepData(S_STEP_NAME))

        If Not ExecuteStep(stepData) Then
            Application.StatusBar = False
            MsgBox _
                "処理を停止しました。" & vbCrLf & vbCrLf & _
                "Step: " & CStr(stepData(S_STEP_NAME)) & vbCrLf & _
                "内容: " & mLastError, _
                vbExclamation, "UIA実行エラー"
            Exit Sub
        End If
    Next i

    Application.StatusBar = False
    MsgBox steps.Count & "件の処理が完了しました。", vbInformation
    Exit Sub

UnexpectedError:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Application.StatusBar = False
    MsgBox _
        "実行準備中にエラーが発生しました。" & vbCrLf & _
        errorNumber & ": " & errorDescription, _
        vbExclamation, "UIA実行エラー"

End Sub


' =====================================================================
' 設定確認：操作せず、各ステップの一致件数だけ確認
' =====================================================================
Public Sub UIA_登録要素の一致数を確認()

    Dim steps As Collection
    Dim stepData As Variant
    Dim i As Long
    Dim matchCount As Long
    Dim occurrence As Long
    Dim resultText As String
    Dim judgment As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    Set mDataBook = Application.ActiveWorkbook

    If mDataBook Is Nothing Then
        Err.Raise vbObjectError + 4202, , _
                  "Excelブックを1つ開いてから実行してください。"
    End If

    If UCase$(mDataBook.Name) = "PERSONAL.XLSB" Or mDataBook.IsAddin Then
        Err.Raise vbObjectError + 4203, , _
                  "通常のExcelブックを前面にしてから実行してください。"
    End If

    EnsureUIA
    Set steps = BuildScenario

    For i = 1 To steps.Count
        stepData = steps(i)
        occurrence = ValidatedOccurrence(stepData(S_OCCURRENCE))
        matchCount = CountMatchingElements(stepData)

        If matchCount = 1 And occurrence = 1 Then
            judgment = "OK"
        ElseIf matchCount >= occurrence And matchCount > 1 Then
            judgment = "要確認"
        Else
            judgment = "NG"
        End If

        resultText = resultText & _
            i & ". " & CStr(stepData(S_STEP_NAME)) & _
            "：" & matchCount & "件（" & judgment & "）" & vbCrLf
    Next i

    MsgBox _
        "要素の一致件数です。操作は実行していません。" & _
        vbCrLf & vbCrLf & resultText, _
        vbInformation, "UIA設定確認"
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MsgBox _
        "設定を確認できませんでした。" & vbCrLf & _
        errorNumber & ": " & errorDescription, _
        vbExclamation, "UIA設定確認"

End Sub


' =====================================================================
' 1ステップ実行
' =====================================================================
Private Function ExecuteStep(ByVal stepData As Variant) As Boolean

    Dim target As IUIAutomationElement
    Dim actionName As String
    Dim inputValue As String
    Dim occurrence As Long

    On Error GoTo Failed
    mLastError = ""

    actionName = UCase$(Trim$(CStr(stepData(S_ACTION))))

    occurrence = ValidatedOccurrence(stepData(S_OCCURRENCE))

    Set target = WaitForElement(stepData)

    If target Is Nothing Then
        Err.Raise vbObjectError + 4211, , _
                  "指定時間内に対象要素が見つかりませんでした。"
    End If

    Select Case actionName
        Case "SET_TEXT"
            inputValue = ResolveInputValue( _
                CStr(stepData(S_VALUE_SOURCE)))
            SetTextValue target, inputValue

        Case "CLICK"
            InvokeElement target

        Case "CHECK_ON"
            SetCheckState target, True

        Case "CHECK_OFF"
            SetCheckState target, False

        Case "SELECT"
            SelectElement target

        Case "EXPAND"
            SetExpandState target, True

        Case "COLLAPSE"
            SetExpandState target, False

        Case "WAIT"
            ' 要素が見つかった時点で成功

        Case Else
            Err.Raise vbObjectError + 4212, , _
                      "未対応のActionです: " & actionName
    End Select

    DoEvents
    Sleep 100
    ExecuteStep = True
    Exit Function

Failed:
    mLastError = Err.Number & ": " & Err.Description
    ExecuteStep = False

End Function


Private Function CountMatchingElements( _
    ByVal stepData As Variant) As Long

    Dim windowElement As IUIAutomationElement
    Dim condition As IUIAutomationCondition
    Dim matches As IUIAutomationElementArray

    Set windowElement = FindTopWindow( _
        CStr(stepData(S_WINDOW_KEYWORD)))

    If windowElement Is Nothing Then Exit Function

    Set condition = BuildElementCondition( _
        CStr(stepData(S_AUTOMATION_ID)), _
        CStr(stepData(S_ELEMENT_NAME)), _
        CLng(stepData(S_CONTROL_TYPE_ID)), _
        CStr(stepData(S_CLASS_NAME)))

    Set matches = windowElement.FindAll(SCOPE_SUBTREE, condition)

    If Not matches Is Nothing Then
        CountMatchingElements = matches.Length
    End If

End Function


Private Function ValidatedOccurrence( _
    ByVal occurrenceValue As Variant) As Long

    Dim numericValue As Double

    If Not IsNumeric(occurrenceValue) Then
        Err.Raise vbObjectError + 4210, , _
                  "Occurrenceは1以上の整数を指定してください。"
    End If

    numericValue = CDbl(occurrenceValue)

    If numericValue < 1 Or numericValue <> Fix(numericValue) Then
        Err.Raise vbObjectError + 4210, , _
                  "Occurrenceは1以上の整数を指定してください。"
    End If

    ValidatedOccurrence = CLng(numericValue)

End Function


' =====================================================================
' 指定要素が現れるまで待機
' =====================================================================
Private Function WaitForElement( _
    ByVal stepData As Variant) As IUIAutomationElement

    Dim startedAt As Double
    Dim timeoutSeconds As Double
    Dim target As IUIAutomationElement

    timeoutSeconds = CDbl(stepData(S_TIMEOUT))
    If timeoutSeconds <= 0 Then timeoutSeconds = 10

    startedAt = Timer

    Do
        Set target = FindElementOnce(stepData)

        If Not target Is Nothing Then
            Set WaitForElement = target
            Exit Function
        End If

        DoEvents
        Sleep 50
    Loop While ElapsedSeconds(startedAt) < timeoutSeconds

End Function


' =====================================================================
' Window＋AutomationId/Name＋ControlTypeで直接検索
' =====================================================================
Private Function FindElementOnce( _
    ByVal stepData As Variant) As IUIAutomationElement

    Dim windowElement As IUIAutomationElement
    Dim condition As IUIAutomationCondition
    Dim matches As IUIAutomationElementArray
    Dim occurrence As Long

    Set windowElement = FindTopWindow( _
        CStr(stepData(S_WINDOW_KEYWORD)))

    If windowElement Is Nothing Then Exit Function

    Set condition = BuildElementCondition( _
        CStr(stepData(S_AUTOMATION_ID)), _
        CStr(stepData(S_ELEMENT_NAME)), _
        CLng(stepData(S_CONTROL_TYPE_ID)), _
        CStr(stepData(S_CLASS_NAME)))

    occurrence = CLng(stepData(S_OCCURRENCE))

    If occurrence = 1 Then
        Set FindElementOnce = _
            windowElement.FindFirst(SCOPE_SUBTREE, condition)
        Exit Function
    End If

    Set matches = windowElement.FindAll(SCOPE_SUBTREE, condition)

    If Not matches Is Nothing Then
        If matches.Length >= occurrence Then
            Set FindElementOnce = matches.GetElement(occurrence - 1)
        End If
    End If

End Function


Private Function BuildElementCondition( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    ByVal className As String) As IUIAutomationCondition

    Dim condition As IUIAutomationCondition
    Dim part As IUIAutomationCondition

    If Len(automationId) = 0 And Len(elementName) = 0 Then
        Err.Raise vbObjectError + 4220, , _
                  "AutomationIdとNameの両方が空です。"
    End If

    If controlTypeId <= 0 Then
        Err.Raise vbObjectError + 4221, , _
                  "ControlTypeIdが未設定です。"
    End If

    Set condition = mUIA.CreateTrueCondition

    If Len(automationId) > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_AUTOMATION_ID_PROPERTY_ID, automationId)
    Else
        Set part = mUIA.CreatePropertyCondition( _
            UIA_NAME_PROPERTY_ID, elementName)
    End If
    Set condition = mUIA.CreateAndCondition(condition, part)

    Set part = mUIA.CreatePropertyCondition( _
        UIA_CONTROL_TYPE_PROPERTY_ID, controlTypeId)
    Set condition = mUIA.CreateAndCondition(condition, part)

    If Len(className) > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_CLASS_NAME_PROPERTY_ID, className)
        Set condition = mUIA.CreateAndCondition(condition, part)
    End If

    Set BuildElementCondition = condition

End Function


' =====================================================================
' トップレベルWindowをタイトルの完全一致または部分一致で一意取得
' =====================================================================
Private Function FindTopWindow( _
    ByVal windowKeyword As String) As IUIAutomationElement

    Dim root As IUIAutomationElement
    Dim typeCondition As IUIAutomationCondition
    Dim windows As IUIAutomationElementArray
    Dim candidate As IUIAutomationElement
    Dim exactCandidate As IUIAutomationElement
    Dim partialCandidate As IUIAutomationElement
    Dim currentTitle As String
    Dim exactCount As Long
    Dim partialCount As Long
    Dim i As Long

    If Len(Trim$(windowKeyword)) = 0 Then Exit Function

    Set root = mUIA.GetRootElement
    Set typeCondition = mUIA.CreatePropertyCondition( _
        UIA_CONTROL_TYPE_PROPERTY_ID, UIA_CT_WINDOW)
    Set windows = root.FindAll(SCOPE_CHILDREN, typeCondition)

    If windows Is Nothing Then Exit Function

    For i = 0 To windows.Length - 1
        Set candidate = windows.GetElement(i)
        currentTitle = SafeName(candidate)

        If StrComp(currentTitle, windowKeyword, vbTextCompare) = 0 Then
            exactCount = exactCount + 1
            Set exactCandidate = candidate
        ElseIf InStr(1, currentTitle, windowKeyword, vbTextCompare) > 0 Then
            partialCount = partialCount + 1
            Set partialCandidate = candidate
        End If
    Next i

    If exactCount = 1 Then
        Set FindTopWindow = exactCandidate
    ElseIf exactCount = 0 And partialCount = 1 Then
        Set FindTopWindow = partialCandidate
    ElseIf exactCount > 1 Or partialCount > 1 Then
        Err.Raise vbObjectError + 4222, , _
                  "WindowKeywordに一致するウィンドウが複数あります: " & _
                  windowKeyword
    End If

End Function


' =====================================================================
' UIAアクション
' =====================================================================
Private Sub SetTextValue( _
    ByVal target As IUIAutomationElement, _
    ByVal inputValue As String)

    Dim valuePattern As IUIAutomationValuePattern

    Set valuePattern = target.GetCurrentPattern(UIA_VALUE_PATTERN_ID)

    If valuePattern.CurrentIsReadOnly Then
        Err.Raise vbObjectError + 4230, , _
                  "対象は読み取り専用です。"
    End If

    valuePattern.SetValue inputValue

End Sub


Private Sub InvokeElement(ByVal target As IUIAutomationElement)

    Dim invokePattern As IUIAutomationInvokePattern
    Set invokePattern = target.GetCurrentPattern(UIA_INVOKE_PATTERN_ID)
    invokePattern.Invoke

End Sub


Private Sub SetCheckState( _
    ByVal target As IUIAutomationElement, _
    ByVal desiredOn As Boolean)

    Dim togglePattern As IUIAutomationTogglePattern
    Dim desiredState As Long
    Dim attempt As Long

    Set togglePattern = target.GetCurrentPattern(UIA_TOGGLE_PATTERN_ID)
    desiredState = IIf(desiredOn, TOGGLE_ON, TOGGLE_OFF)

    For attempt = 1 To 3
        If togglePattern.CurrentToggleState = desiredState Then Exit For
        togglePattern.Toggle
    Next attempt

    If togglePattern.CurrentToggleState <> desiredState Then
        Err.Raise vbObjectError + 4231, , _
                  "チェック状態を変更できませんでした。"
    End If

End Sub


Private Sub SelectElement(ByVal target As IUIAutomationElement)

    Dim selectionPattern As IUIAutomationSelectionItemPattern
    Set selectionPattern = target.GetCurrentPattern( _
        UIA_SELECTION_ITEM_PATTERN_ID)
    selectionPattern.Select

End Sub


Private Sub SetExpandState( _
    ByVal target As IUIAutomationElement, _
    ByVal desiredExpanded As Boolean)

    Dim expandPattern As IUIAutomationExpandCollapsePattern

    Set expandPattern = target.GetCurrentPattern( _
        UIA_EXPAND_COLLAPSE_PATTERN_ID)

    If desiredExpanded Then
        If expandPattern.CurrentExpandCollapseState <> EXPAND_EXPANDED Then
            expandPattern.Expand
        End If

        If expandPattern.CurrentExpandCollapseState <> EXPAND_EXPANDED Then
            Err.Raise vbObjectError + 4232, , _
                      "要素を展開状態にできませんでした。"
        End If
    Else
        If expandPattern.CurrentExpandCollapseState <> EXPAND_COLLAPSED Then
            expandPattern.Collapse
        End If

        If expandPattern.CurrentExpandCollapseState <> EXPAND_COLLAPSED Then
            Err.Raise vbObjectError + 4233, , _
                      "要素を折りたたみ状態にできませんでした。"
        End If
    End If

End Sub


' =====================================================================
' 入力値：文字列またはActiveWorkbookのセル
' =====================================================================
Private Function ResolveInputValue( _
    ByVal valueOrSource As String) As String

    If Left$(valueOrSource, 2) = "@@" Then
        ResolveInputValue = Mid$(valueOrSource, 2)
    ElseIf Left$(valueOrSource, 1) = "@" Then
        ResolveInputValue = CStr( _
            GetCellReference(Mid$(valueOrSource, 2)).Value2)
    Else
        ResolveInputValue = valueOrSource
    End If

End Function


Private Function GetCellReference( _
    ByVal referenceText As String) As Range

    Dim bangPosition As Long
    Dim sheetName As String
    Dim addressText As String
    Dim target As Range

    bangPosition = InStrRev(referenceText, "!")

    If bangPosition <= 1 Then
        Err.Raise vbObjectError + 4240, , _
                  "セル指定は 入力!A2 の形式にしてください。"
    End If

    sheetName = Replace(Left$(referenceText, bangPosition - 1), "'", "")
    addressText = Mid$(referenceText, bangPosition + 1)
    Set target = mDataBook.Worksheets(sheetName).Range(addressText)

    If target.Cells.CountLarge <> 1 Then
        Err.Raise vbObjectError + 4241, , _
                  "入力元は単一セルを指定してください。"
    End If

    Set GetCellReference = target

End Function


' =====================================================================
' 共通補助
' =====================================================================
Private Sub EnsureUIA()

    If mUIA Is Nothing Then Set mUIA = New CUIAutomation

End Sub


Private Function ElapsedSeconds(ByVal startedAt As Double) As Double

    Dim currentValue As Double
    currentValue = Timer

    If currentValue >= startedAt Then
        ElapsedSeconds = currentValue - startedAt
    Else
        ElapsedSeconds = (86400# - startedAt) + currentValue
    End If

End Function


Private Function SafeName(ByVal target As IUIAutomationElement) As String

    On Error GoTo Failed
    SafeName = target.CurrentName
    Exit Function

Failed:
    Err.Clear

End Function
