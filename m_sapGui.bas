Attribute VB_Name = "m_sapGui"
Option Explicit


Sub MainPage_OpenTransaction(Session As SAPFEWSELib.GuiSession, TransactionCode As String)

    Dim NavigationPanel As String: NavigationPanel = "wnd[0]/tbar[0]/okcd"
    
    Session.FindById(NavigationPanel).Text = TransactionCode
    Call SendEnter(Session)

End Sub
Sub SaveTransaction(Session As SAPFEWSELib.GuiSession, Optional ExitWindowFlag As Boolean = False)
    
    Dim ElementPath As String: ElementPath = "wnd[0]/tbar[1]/btn[23]"
    
    Session.FindById(ElementPath).Press
    
    ' Save transaction result message.
    Ws_Requests.Range("O" & ArrCounter + 2).Value = GetStatusBarProperty(Session, "Text")
    
    If ExitWindowFlag = True Then
        Call ExitWindow(Session)
    End If
    
End Sub

Function CreateSession() As SAPFEWSELib.GuiSession

    Dim SAP         As Variant
    Dim Connection  As SAPFEWSELib.GuiConnection
    Dim Appl        As SAPFEWSELib.GuiApplication
    
    On Error Resume Next
    Set SAP = GetObject("sapgui")
    Set Appl = SAP.GetScriptingEngine()
    Set Connection = Appl.Children(0)
    Set CreateSession = Connection.Children(0)

End Function

Function IsElementExists(Session As SAPFEWSELib.GuiSession, ElementPath As String) As Boolean
    
    On Error Resume Next
        Dim TempObject As Object
        Set TempObject = Session.FindById(ElementPath)
    On Error GoTo 0
    
    If Not (TempObject Is Nothing) Then
        IsElementExists = True
    Else
        IsElementExists = False
    End If
    
End Function

Sub ClosePopUpChangeDefaultValue(Session As SAPFEWSELib.GuiSession, Optional DoNotShowAgain As Boolean = False)

    Dim WindowPopUpChangeDefaultValue As Object
    
    If IsElementExists(Session, "wnd[1]") Then
        
        If DoNotShowAgain = True Then
            Call SetCheckbox(Session, "wnd[1]/usr/chkG_TIP_DONT_SHOW_AGAIN", True)
        End If
    
        ' Press ok button and close window.
        Session.FindById("wnd[1]/tbar[0]/btn[0]").Press
        
    Else
        GoTo ErrorHandler
    End If
    
    
ErrorHandler:
    ' Here will be error handling function.
    ' MsgBox ("Can't find wnd[1] PopUp!")
    
End Sub

Sub SetCheckbox(Session As SAPFEWSELib.GuiSession, CheckboxPath As String, Checked As Boolean)
    
    Dim CheckBox As Object

    If IsElementExists(Session, CheckboxPath) Then
    
        Set CheckBox = Session.FindById(CheckboxPath)
        
        If Checked = True Then
            
            If CheckBox.Selected <> True Then
                CheckBox.Selected = True
            Else
                ' Already checked.
            End If
    
        ElseIf Checked = False Then
            
            If CheckBox.Selected = True Then
                CheckBox.Selected = False
            Else
                ' Already unchecked.
            End If
        
        End If
        
    Else
    
        GoTo ErrorHandler
        
    End If
    
ErrorHandler:
    ' Here will be error handling function.
    ' MsgBox ("Can't find " & CheckboxPath & " checkbox!")

End Sub

Sub SetDropDownMenu(Session As SAPFEWSELib.GuiSession, ElementPath As String, Value As String)

    Session.FindById(ElementPath).Key = Value
    
End Sub

Sub SetTextboxValue(Session As SAPFEWSELib.GuiSession, TextboxPath As String, Value As Variant)

    Dim errMessage As String
    errMessage = Session.FindById("wnd[0]/sbar").Text
    
    If IsElementExists(Session, TextboxPath) Then

        'On Error Resume Next
        If Trim(Value) = "" Then
            Session.FindById(TextboxPath).Text = ""
        ElseIf IsNumeric(Value) Then
            Session.FindById(TextboxPath).Text = Replace(Value, ".", ",")
        Else
            Session.FindById(TextboxPath).Text = CStr(Trim(Value))
        End If
    
    Else
    
        GoTo ErrorHandler
    
    End If
    
ErrorHandler:

    ' Here will be error handler function.
    ' Msgbox("No textbox for value: " & value & "!")
    
End Sub

Function GetStatusBarProperty(Session As SAPFEWSELib.GuiSession, PropertyName As String) As String
    ' Returns text or message type from window status bar.
    ' There are next returning message type:
    ' S - Success, W - warning, E - Error, A - Abort, I - Information

    Dim StatusBarPath As String
    StatusBarPath = "wnd[0]/sbar"

    If IsElementExists(Session, StatusBarPath) = True Then
    
        Dim StatusBar As Object
        Set StatusBar = Session.FindById(StatusBarPath)
        
        If PropertyName = "Text" Then
        
            GetStatusBarProperty = StatusBar.Text
            
        ElseIf PropertyName = "Message Type" Then
        
            GetStatusBarProperty = StatusBar.MessageType
            
        End If
    
    Else
    
        Dim ErrorFlag As Boolean
        ErrorFlag = True
        GoTo ErrorHandler
    
    End If
    
ErrorHandler:

If ErrorFlag = True Then
    MsgBox ("Error with status bar handling!")
End If

End Function

Public Sub SendEnter(Session As SAPFEWSELib.GuiSession)

    Session.FindById("wnd[0]").SendVKey 0

End Sub
