Attribute VB_Name = "modUIADirect"
Option Explicit

' ============================================================
' UI Automation 高速・直接指定版
' ------------------------------------------------------------
' 前提:
'   VBEの「ツール」→「参照設定」で
'   「UIAutomationClient」にチェックを入れてください。
'
' 方針:
'   1. 本番操作では全要素を列挙しない
'   2. AutomationId / Name / ControlType / ClassName の条件を作る
'   3. FindFirstで対象を直接取得する
'   4. 同じ画面の要素はキャッシュして再利用する
'   5. 画面遷移後は UIA_ClearCache を呼び、古い要素を捨てる
'
' 注意:
'   ・調査時の通し番号は画面更新で変わる可能性があります。
'   ・本番ではAutomationIdを優先し、Nameは補助にします。
'   ・SendKeysは使いません。対応するUIAパターンがなければ停止します。
' ============================================================

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' TreeScope
Private Const SCOPE_CHILDREN As Long = 2
Private Const SCOPE_SUBTREE As Long = 7

' Property IDs
Private Const UIA_PROCESS_ID_PROPERTY_ID As Long = 30002
Private Const UIA_CONTROL_TYPE_PROPERTY_ID As Long = 30003
Private Const UIA_NAME_PROPERTY_ID As Long = 30005
Private Const UIA_IS_ENABLED_PROPERTY_ID As Long = 30010
Private Const UIA_AUTOMATION_ID_PROPERTY_ID As Long = 30011
Private Const UIA_CLASS_NAME_PROPERTY_ID As Long = 30012

' Pattern IDs
Private Const UIA_INVOKE_PATTERN_ID As Long = 10000
Private Const UIA_VALUE_PATTERN_ID As Long = 10002
Private Const UIA_EXPAND_COLLAPSE_PATTERN_ID As Long = 10005
Private Const UIA_SELECTION_ITEM_PATTERN_ID As Long = 10010
Private Const UIA_TEXT_PATTERN_ID As Long = 10014
Private Const UIA_TOGGLE_PATTERN_ID As Long = 10015

' よく使うControlType IDs
Public Const UIA_CT_ANY As Long = 0
Public Const UIA_CT_BUTTON As Long = 50000
Public Const UIA_CT_CHECKBOX As Long = 50002
Public Const UIA_CT_COMBOBOX As Long = 50003
Public Const UIA_CT_EDIT As Long = 50004
Public Const UIA_CT_HYPERLINK As Long = 50005
Public Const UIA_CT_LISTITEM As Long = 50007
Public Const UIA_CT_RADIOBUTTON As Long = 50013
Public Const UIA_CT_TEXT As Long = 50020
Public Const UIA_CT_WINDOW As Long = 50032

' ToggleState
Private Const TOGGLE_OFF As Long = 0
Private Const TOGGLE_ON As Long = 1

Private mUIA As CUIAutomation
Private mWindow As IUIAutomationElement
Private mElementCache As Object  ' Scripting.Dictionary（遅延バインディング）


' ============================================================
' 最初に1回だけ実行：対象ウィンドウを取得
' exactTitle=Trueなら、タイトル完全一致をFindFirstで直接検索します。
' exactTitle=Falseなら、デスクトップ直下のWindowだけを調べて部分一致します。
' ============================================================
Public Function UIA_Start( _
    ByVal windowTitle As String, _
    Optional ByVal exactTitle As Boolean = False, _
    Optional ByVal timeoutSeconds As Double = 10, _
    Optional ByVal processId As Long = 0) As Boolean

    Dim startedAt As Double

    Set mUIA = New CUIAutomation
    Set mElementCache = CreateObject("Scripting.Dictionary")
    Set mWindow = Nothing

    startedAt = Timer

    Do
        Set mWindow = FindTopWindow(windowTitle, exactTitle, processId)

        If Not mWindow Is Nothing Then
            UIA_Start = True
            Exit Function
        End If

        DoEvents
        Sleep 50
    Loop While ElapsedSeconds(startedAt) < timeoutSeconds

    UIA_Start = False

End Function


' ============================================================
' AutomationId / Name / ControlType / ClassName から直接取得
' ------------------------------------------------------------
' occurrence=1  : FindFirst（最速）
' occurrence>=2 : 条件に一致した要素だけFindAllし、指定順位を取得
'
' containerAutomationIdを指定すると、そのコンテナ内だけを探します。
' 大きなWeb画面では検索範囲を絞るほど効果があります。
' ============================================================
Public Function UIA_GetElement( _
    Optional ByVal automationId As String = "", _
    Optional ByVal elementName As String = "", _
    Optional ByVal controlTypeId As Long = UIA_CT_ANY, _
    Optional ByVal className As String = "", _
    Optional ByVal occurrence As Long = 1, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal useCache As Boolean = True) As IUIAutomationElement

    Dim searchRoot As IUIAutomationElement
    Dim container As IUIAutomationElement
    Dim found As IUIAutomationElement
    Dim cacheKey As String

    EnsureStarted

    If Len(automationId) = 0 And _
       Len(elementName) = 0 And _
       controlTypeId = UIA_CT_ANY And _
       Len(className) = 0 Then
        Err.Raise vbObjectError + 2100, "UIA_GetElement", _
                  "検索条件が空です。AutomationId、Name、ControlType、ClassNameのいずれかを指定してください。"
    End If

    If occurrence < 1 Then occurrence = 1

    cacheKey = MakeCacheKey( _
        automationId, elementName, controlTypeId, className, _
        occurrence, containerAutomationId)

    If useCache Then
        Set found = GetAliveCachedElement(cacheKey)
        If Not found Is Nothing Then
            Set UIA_GetElement = found
            Exit Function
        End If
    End If

    Set searchRoot = mWindow

    If Len(containerAutomationId) > 0 Then
        Set container = FindDirect( _
            mWindow, containerAutomationId, "", UIA_CT_ANY, "", 1)

        If container Is Nothing Then
            Err.Raise vbObjectError + 2101, "UIA_GetElement", _
                      "検索コンテナが見つかりません。AutomationId=" & containerAutomationId
        End If

        Set searchRoot = container
    End If

    Set found = FindDirect( _
        searchRoot, automationId, elementName, controlTypeId, className, occurrence)

    If found Is Nothing Then
        Err.Raise vbObjectError + 2102, "UIA_GetElement", _
                  "対象要素が見つかりません。" & SelectorText( _
                      automationId, elementName, controlTypeId, className, occurrence)
    End If

    If useCache Then
        If mElementCache.Exists(cacheKey) Then mElementCache.Remove cacheKey
        mElementCache.Add cacheKey, found
    End If

    Set UIA_GetElement = found

End Function


' ============================================================
' テキストボックスへ直接入力（ValuePattern）
' ============================================================
Public Sub UIA_SetText( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal newValue As String, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1)

    Dim elm As IUIAutomationElement
    Dim valuePattern As IUIAutomationValuePattern

    Set elm = UIA_GetElement( _
        automationId, elementName, UIA_CT_EDIT, "", _
        occurrence, containerAutomationId, True)

    Set valuePattern = elm.GetCurrentPattern(UIA_VALUE_PATTERN_ID)

    If valuePattern.CurrentIsReadOnly Then
        Err.Raise vbObjectError + 2110, "UIA_SetText", _
                  "対象は読み取り専用です: " & elm.CurrentName
    End If

    valuePattern.SetValue newValue

End Sub


' ============================================================
' チェックボックスを指定状態へ合わせる
' Toggleするだけではなく、現在状態を見て必要な場合だけ変更します。
' ============================================================
Public Sub UIA_SetCheck( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal desiredOn As Boolean, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1)

    Dim elm As IUIAutomationElement
    Dim togglePattern As IUIAutomationTogglePattern
    Dim desiredState As Long

    Set elm = UIA_GetElement( _
        automationId, elementName, UIA_CT_CHECKBOX, "", _
        occurrence, containerAutomationId, True)

    Set togglePattern = elm.GetCurrentPattern(UIA_TOGGLE_PATTERN_ID)

    If desiredOn Then
        desiredState = TOGGLE_ON
    Else
        desiredState = TOGGLE_OFF
    End If

    If togglePattern.CurrentToggleState <> desiredState Then
        togglePattern.Toggle
    End If

    If togglePattern.CurrentToggleState <> desiredState Then
        Err.Raise vbObjectError + 2120, "UIA_SetCheck", _
                  "チェック状態を変更できませんでした: " & elm.CurrentName
    End If

End Sub


' ============================================================
' チェックボックスの現在状態を取得
' ============================================================
Public Function UIA_ReadCheck( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1) As Boolean

    Dim elm As IUIAutomationElement
    Dim togglePattern As IUIAutomationTogglePattern

    Set elm = UIA_GetElement( _
        automationId, elementName, UIA_CT_CHECKBOX, "", _
        occurrence, containerAutomationId, True)

    Set togglePattern = elm.GetCurrentPattern(UIA_TOGGLE_PATTERN_ID)
    UIA_ReadCheck = (togglePattern.CurrentToggleState = TOGGLE_ON)

End Function


' ============================================================
' ボタン・リンク等を直接実行（InvokePattern）
' ============================================================
Public Sub UIA_Invoke( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    Optional ByVal controlTypeId As Long = UIA_CT_BUTTON, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1, _
    Optional ByVal clearCacheAfter As Boolean = False)

    Dim elm As IUIAutomationElement
    Dim invokePattern As IUIAutomationInvokePattern

    Set elm = UIA_GetElement( _
        automationId, elementName, controlTypeId, "", _
        occurrence, containerAutomationId, True)

    Set invokePattern = elm.GetCurrentPattern(UIA_INVOKE_PATTERN_ID)
    invokePattern.Invoke

    If clearCacheAfter Then UIA_ClearCache

End Sub


' ============================================================
' ラジオボタンやリスト項目を選択（SelectionItemPattern）
' ============================================================
Public Sub UIA_SelectItem( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1)

    Dim elm As IUIAutomationElement
    Dim selectionPattern As IUIAutomationSelectionItemPattern

    Set elm = UIA_GetElement( _
        automationId, elementName, controlTypeId, "", _
        occurrence, containerAutomationId, True)

    Set selectionPattern = elm.GetCurrentPattern(UIA_SELECTION_ITEM_PATTERN_ID)
    selectionPattern.Select

End Sub


' ============================================================
' 入力欄などの値を取得（ValuePattern）
' ============================================================
Public Function UIA_ReadValue( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    Optional ByVal controlTypeId As Long = UIA_CT_EDIT, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1) As String

    Dim elm As IUIAutomationElement
    Dim valuePattern As IUIAutomationValuePattern

    Set elm = UIA_GetElement( _
        automationId, elementName, controlTypeId, "", _
        occurrence, containerAutomationId, True)

    Set valuePattern = elm.GetCurrentPattern(UIA_VALUE_PATTERN_ID)
    UIA_ReadValue = valuePattern.CurrentValue

End Function


' ============================================================
' 文書・長文領域を取得（TextPattern）
' ============================================================
Public Function UIA_ReadText( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    Optional ByVal controlTypeId As Long = UIA_CT_ANY, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1) As String

    Dim elm As IUIAutomationElement
    Dim textPattern As IUIAutomationTextPattern
    Dim documentRange As IUIAutomationTextRange

    Set elm = UIA_GetElement( _
        automationId, elementName, controlTypeId, "", _
        occurrence, containerAutomationId, True)

    Set textPattern = elm.GetCurrentPattern(UIA_TEXT_PATTERN_ID)
    Set documentRange = textPattern.DocumentRange
    UIA_ReadText = documentRange.GetText(-1)

End Function


' ============================================================
' Nameプロパティ（画面上の表示名）を取得
' ============================================================
Public Function UIA_ReadName( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    Optional ByVal controlTypeId As Long = UIA_CT_ANY, _
    Optional ByVal containerAutomationId As String = "", _
    Optional ByVal occurrence As Long = 1) As String

    Dim elm As IUIAutomationElement

    Set elm = UIA_GetElement( _
        automationId, elementName, controlTypeId, "", _
        occurrence, containerAutomationId, True)

    UIA_ReadName = elm.CurrentName

End Function


' ============================================================
' 特定要素が現れるまで直接検索して待つ
' 固定時間待機ではなく、要素の存在を完了条件にします。
' ============================================================
Public Function UIA_WaitElement( _
    Optional ByVal automationId As String = "", _
    Optional ByVal elementName As String = "", _
    Optional ByVal controlTypeId As Long = UIA_CT_ANY, _
    Optional ByVal timeoutSeconds As Double = 10, _
    Optional ByVal containerAutomationId As String = "") As IUIAutomationElement

    Dim startedAt As Double
    Dim elm As IUIAutomationElement
    Dim searchRoot As IUIAutomationElement

    EnsureStarted
    startedAt = Timer

    Do
        Set searchRoot = mWindow

        If Len(containerAutomationId) > 0 Then
            Set searchRoot = FindDirect( _
                mWindow, containerAutomationId, "", UIA_CT_ANY, "", 1)
        End If

        If Not searchRoot Is Nothing Then
            Set elm = FindDirect( _
                searchRoot, automationId, elementName, controlTypeId, "", 1)
        Else
            Set elm = Nothing
        End If

        If Not elm Is Nothing Then
            Set UIA_WaitElement = elm
            Exit Function
        End If

        DoEvents
        Sleep 50
    Loop While ElapsedSeconds(startedAt) < timeoutSeconds

End Function


' ============================================================
' 画面遷移・再読込・一覧更新後に呼ぶ
' ============================================================
Public Sub UIA_ClearCache()

    If mElementCache Is Nothing Then
        Set mElementCache = CreateObject("Scripting.Dictionary")
    Else
        mElementCache.RemoveAll
    End If

End Sub


' ============================================================
' セッション終了
' ============================================================
Public Sub UIA_Close()

    If Not mElementCache Is Nothing Then mElementCache.RemoveAll
    Set mElementCache = Nothing
    Set mWindow = Nothing
    Set mUIA = Nothing

End Sub


' ============================================================
' 使用例
' ------------------------------------------------------------
' 既存記事のローカルHTMLデモを開いた状態で実行します。
' EdgeでHTMLのidがAutomationIdとして公開されない環境では、
' 第1引数を空文字にしてNameで直接検索してください。
' ============================================================
Public Sub 使用例_名前で直接操作()

    If Not UIA_Start("UIA入力練習フォーム", False, 10) Then
        MsgBox "対象ウィンドウが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' AutomationIdが分からない場合：Name + ControlTypeでFindFirst
    UIA_SetText "", "社員番号", "A001"
    UIA_SetText "", "氏名", "山田太郎"
    UIA_SetText "", "金額", "12500"
    UIA_SetText "", "備考", "全要素列挙なしで直接入力"

    ' 登録後に画面構造が変わる場合はTrueを指定してキャッシュを消す
    UIA_Invoke "", "登録", UIA_CT_BUTTON, "", 1, True

    UIA_Close

End Sub


Public Sub 使用例_AutomationIdで直接操作()

    If Not UIA_Start("対象システムの画面タイトル", False, 10) Then
        MsgBox "対象ウィンドウが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' AutomationIdが安定している場合はこちらを優先
    UIA_SetText "empCode", "", _
        ThisWorkbook.Worksheets("操作").Range("A2").Value
    UIA_SetText "employeeName", "", _
        ThisWorkbook.Worksheets("操作").Range("B2").Value
    UIA_SetCheck "agree", "", True

    ' 画面に表示された値をExcelへ戻す
    ThisWorkbook.Worksheets("操作").Range("D2").Value = _
        UIA_ReadValue("resultValue", "", UIA_CT_EDIT)

    ' 確定ボタンは検証後に有効化する
    ' UIA_Invoke "registerButton", "", UIA_CT_BUTTON, "", 1, True

    UIA_Close

End Sub


' ============================================================
' 以下は内部処理
' ============================================================
Private Function FindTopWindow( _
    ByVal windowTitle As String, _
    ByVal exactTitle As Boolean, _
    ByVal processId As Long) As IUIAutomationElement

    Dim root As IUIAutomationElement
    Dim cond As IUIAutomationCondition
    Dim part As IUIAutomationCondition
    Dim windows As IUIAutomationElementArray
    Dim elm As IUIAutomationElement
    Dim i As Long
    Dim currentName As String

    Set root = mUIA.GetRootElement
    Set cond = mUIA.CreatePropertyCondition( _
        UIA_CONTROL_TYPE_PROPERTY_ID, UIA_CT_WINDOW)

    If processId > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_PROCESS_ID_PROPERTY_ID, processId)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If exactTitle Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_NAME_PROPERTY_ID, windowTitle)
        Set cond = mUIA.CreateAndCondition(cond, part)
        Set FindTopWindow = root.FindFirst(SCOPE_CHILDREN, cond)
        Exit Function
    End If

    ' 部分一致はUIAのPropertyConditionだけでは指定できないため、
    ' デスクトップ直下のWindowだけに限定して確認します。
    Set windows = root.FindAll(SCOPE_CHILDREN, cond)

    For i = 0 To windows.Length - 1
        Set elm = windows.GetElement(i)
        currentName = SafeCurrentName(elm)

        If InStr(1, currentName, windowTitle, vbTextCompare) > 0 Then
            Set FindTopWindow = elm
            Exit Function
        End If
    Next i

End Function


Private Function FindDirect( _
    ByVal searchRoot As IUIAutomationElement, _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    ByVal className As String, _
    ByVal occurrence As Long) As IUIAutomationElement

    Dim cond As IUIAutomationCondition
    Dim part As IUIAutomationCondition
    Dim matches As IUIAutomationElementArray

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

    If controlTypeId <> UIA_CT_ANY Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_CONTROL_TYPE_PROPERTY_ID, controlTypeId)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If Len(className) > 0 Then
        Set part = mUIA.CreatePropertyCondition( _
            UIA_CLASS_NAME_PROPERTY_ID, className)
        Set cond = mUIA.CreateAndCondition(cond, part)
    End If

    If occurrence <= 1 Then
        Set FindDirect = searchRoot.FindFirst(SCOPE_SUBTREE, cond)
        Exit Function
    End If

    ' 全要素ではなく「指定条件に一致した要素だけ」を取得します。
    Set matches = searchRoot.FindAll(SCOPE_SUBTREE, cond)

    If matches.Length >= occurrence Then
        Set FindDirect = matches.GetElement(occurrence - 1)
    End If

End Function


Private Function GetAliveCachedElement( _
    ByVal cacheKey As String) As IUIAutomationElement

    Dim elm As IUIAutomationElement
    Dim dummy As Long

    If mElementCache Is Nothing Then Exit Function
    If Not mElementCache.Exists(cacheKey) Then Exit Function

    Set elm = mElementCache.Item(cacheKey)

    On Error GoTo StaleElement
    dummy = elm.CurrentControlType
    Set GetAliveCachedElement = elm
    Exit Function

StaleElement:
    Err.Clear
    mElementCache.Remove cacheKey

End Function


Private Function MakeCacheKey( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    ByVal className As String, _
    ByVal occurrence As Long, _
    ByVal containerAutomationId As String) As String

    MakeCacheKey = _
        automationId & ChrW$(30) & _
        elementName & ChrW$(30) & _
        CStr(controlTypeId) & ChrW$(30) & _
        className & ChrW$(30) & _
        CStr(occurrence) & ChrW$(30) & _
        containerAutomationId

End Function


Private Function SelectorText( _
    ByVal automationId As String, _
    ByVal elementName As String, _
    ByVal controlTypeId As Long, _
    ByVal className As String, _
    ByVal occurrence As Long) As String

    SelectorText = _
        " AutomationId=""" & automationId & """" & _
        ", Name=""" & elementName & """" & _
        ", ControlType=" & controlTypeId & _
        ", ClassName=""" & className & """" & _
        ", occurrence=" & occurrence

End Function


Private Function SafeCurrentName( _
    ByVal elm As IUIAutomationElement) As String

    On Error GoTo Failed
    SafeCurrentName = elm.CurrentName
    Exit Function

Failed:
    Err.Clear

End Function


Private Function ElapsedSeconds(ByVal startedAt As Double) As Double

    Dim currentValue As Double
    currentValue = Timer

    If currentValue >= startedAt Then
        ElapsedSeconds = currentValue - startedAt
    Else
        ' 午前0時をまたいだ場合
        ElapsedSeconds = (86400# - startedAt) + currentValue
    End If

End Function


Private Sub EnsureStarted()

    If mUIA Is Nothing Then
        Err.Raise vbObjectError + 2199, "modUIADirect", _
                  "先にUIA_Startで対象ウィンドウを取得してください。"
    End If

    If mWindow Is Nothing Then
        Err.Raise vbObjectError + 2199, "modUIADirect", _
                  "先にUIA_Startで対象ウィンドウを取得してください。"
    End If

End Sub
