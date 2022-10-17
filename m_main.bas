Attribute VB_Name = "m_main"
Option Explicit

Const rsrvNumColumn     As String = "A"
Const rsrvPosColumn     As String = "B"
Const qtyColumn         As String = "C"
Const docDateColumn     As String = "D"
Const postDateColumn    As String = "E"
Const docMatColumn      As String = "F"
Const txtDocHeadColumn  As String = "G"
Const doverFIOColumn    As String = "H"
Const movTypeColumn     As String = "I"
Const messageColumn     As String = "J"

Const migoCode          As String = "MIGO"
Const operationCodeCode As String = "A07"
Const refDocCode        As String = "R09"

Const firstRow          As Long = 2

Const wsMainName        As String = "main"
Const wsSettingsName    As String = "settings"

Public wb               As Workbook
Public wsMain           As Worksheet
Public wsSettings       As Worksheet



Sub init()

    Set wb = ThisWorkbook
    Set wsMain = wb.Sheets(wsMainName)
    Set wsSettings = wb.Sheets(wsSettingsName)

End Sub


Sub main()

    Call ImprovePerformance(True)
    Call init
    
    Dim Session                 As SAPFEWSELib.GuiSession
    Dim rowCounter              As Long
    
    Dim rsrvNum     As String
    Dim rsrvPos     As String
    Dim qty         As String
    Dim docDate     As String
    Dim postDate   As String
    Dim docMat      As String
    Dim txtDocHead  As String
    Dim movType     As String
    Dim message     As String
    Dim doverFIO    As String
    
    Dim counter     As Integer
    Dim errorsTexts As String
    Dim errorText   As String
    
    Dim lr As Long
    lr = GetBorders("LR", wsMainName)
    
    Set Session = CreateSession

    
    With wsMain
        For rowCounter = firstRow To lr
        
            errorsTexts = ""
            message = ""
            
            Call MainPage_OpenTransaction(Session, migoCode)
    
            Dim MIGO_GeneralLayoutPath  As String
            MIGO_GeneralLayoutPath = MIGO_InitGeneralLayoutPath(Session)

            ' Get sheet data.
            rsrvNum = .Range(rsrvNumColumn & rowCounter).Value
            rsrvPos = .Range(rsrvPosColumn & rowCounter).Value
            qty = .Range(qtyColumn & rowCounter).Value
            docDate = .Range(docDateColumn & rowCounter).Value
            postDate = .Range(postDateColumn & rowCounter).Value
            docMat = .Range(docMatColumn & rowCounter).Value
            txtDocHead = .Range(txtDocHeadColumn & rowCounter).Value
            movType = .Range(movTypeColumn & rowCounter).Value
            message = .Range(messageColumn & rowCounter).Value
            doverFIO = .Range(doverFIOColumn & rowCounter).Value
            
            ' Set Operation
            Call SetDropDownMenu(Session, MIGO_GeneralLayoutPath & OPERATION_PATH, operationCodeCode)
            Call SetDropDownMenu(Session, MIGO_GeneralLayoutPath & REF_DOC_PATH, refDocCode)
            
            ' Set move type.
            Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & MOVE_TYPE_PATH, movType)
            
            ' Enter res data and Enter.
            Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & RSRV_NUM_PATH, rsrvNum)
            Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & RSRV_POS_PATH, rsrvPos)
            Call SendEnter(Session)
        
            ' Fill data.
            If docDate <> "" Then
                Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & DOC_DATE_PATH, docDate)
            End If
            
            If postDate <> "" Then
                Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & POST_DATE_PATH, postDate)
            End If
            
            Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & DOC_MAT_PATH, docMat)
            Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & TEXT_HEADER_DOC_PATH, txtDocHead)
            Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & QTY_PATH, qty)
            Call SetTextboxValue(Session, MIGO_GeneralLayoutPath & DOVER_FIO_PATH, doverFIO)
             
            Call SetCheckbox(Session, MIGO_GeneralLayoutPath & POS_OK_PATH, True)
            
            ' Save transaction.
            Call Session.FindById(EXECUTE_PATH).Press
            
            ' Check for error window.
            Dim errorWindowsExists As Boolean
            errorWindowsExists = IsElementExists(Session, ERROR_WINDOW_ERROR_BTN_PATH)
            If errorWindowsExists = True Then
                
                Dim countOfErrors   As Integer
                Dim countOfWrns     As Integer
                Dim countOfStops    As Integer
                Dim countOfInfs     As Integer
                Dim countOfMessages As Integer
                
                countOfErrors = Session.FindById(ERROR_WINDOW_ERROR_BTN_PATH).Text
                countOfWrns = Session.FindById(ERROR_WINDOW_WRN_BTN_PATH).Text
                countOfStops = Session.FindById(ERROR_WINDOW_STOP_BTN_PATH).Text
                countOfInfs = Session.FindById(ERROR_WINDOW_INFS_BTN_PATH).Text
                countOfMessages = countOfInfs + countOfStops + countOfWrns + countOfErrors
                
                If countOfErrors > 0 Then
                
                    For counter = 1 To countOfMessages
                        Dim textPosNumber As Integer
                        textPosNumber = counter + 2
                        If Session.FindById(Replace(ERROR_WINDOW_ERROR_ICON_PATH, "@", textPosNumber)).ColorIndex = 4 Then
                            errorText = Session.FindById(Replace(ERROR_WINDOW_ERROR_TEXT_PATH, "@", textPosNumber)).Text
                            errorsTexts = errorsTexts & counter & ". " & errorText & Chr(10)
                        End If
                    Next counter
                
                End If
                
                message = Left(errorsTexts, Len(errorsTexts) - 1)
                .Range(messageColumn & rowCounter).Value = message
                Call Session.FindById(ERROR_WINDOW_BTN_OK_PATH).Press
                Call ExitWindow(Session)
                
            Else
                
                message = GetStatusBarProperty(Session, "Text")
                .Range(messageColumn & rowCounter).Value = message
                Call ExitWindow(Session)
                
            End If

        
        Next rowCounter
    End With
    
    Call ImprovePerformance(False)

End Sub

Function MIGO_InitGeneralLayoutPath(Session As SAPFEWSELib.GuiSession) As String

    Dim GeneralLayout_Path As String
    GeneralLayout_Path = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007"
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0007"
    End If
    
        If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0006"
    End If
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0009"
    End If
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003"
    End If
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002"
    End If
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0001"
    End If
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0004"
    End If
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0005"
    End If
    
    If IsElementExists(Session, "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0010") = True Then
        MIGO_InitGeneralLayoutPath = "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0010"
    End If
    
    
'    If MIGO_InitGeneralLayoutPath = "" Then
'        MIGO_InitGeneralLayoutPath = "ssubSUB_MAIN_CARRIER not found!"
'        ErrorFlag = True
'    End If

End Function

Sub MIGO_SetOpertaionType(Session As SAPFEWSELib.GuiSession, ParentGuiElement As String, OperationTypeCode As String)

    Dim OPERATION_TYPE_COMBOBOX_PATH As String
    OPERATION_TYPE_COMBOBOX_PATH = ParentGuiElement & "/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-ACTION"
    
    If Session.FindById(OPERATION_TYPE_COMBOBOX_PATH).Changeable = True Then
        Session.FindById(OPERATION_TYPE_COMBOBOX_PATH).Key = OperationTypeCode
    End If

End Sub

Sub MIGO_SetLinkedDocument(Session As SAPFEWSELib.GuiSession, ParentGuiElement As String, OperationTypeCode As String)

    Dim LINKED_DOCUMENT_COMBOBOX_PATH As String
    LINKED_DOCUMENT_COMBOBOX_PATH = ParentGuiElement & "/subSUB_FIRSTLINE:SAPLMIGO:0011/cmbGODYNPRO-REFDOC"
    
    If Session.FindById(LINKED_DOCUMENT_COMBOBOX_PATH).Changeable = True Then
        Session.FindById(LINKED_DOCUMENT_COMBOBOX_PATH).Key = OperationTypeCode
    End If
    
End Sub

Sub ExitWindow(Session As SAPFEWSELib.GuiSession, Optional QtyOfRepeats As Integer = 1)
    
    Dim ExitTransactionButtonPath   As String
    Dim CloseAreYouSurePopUp        As String
    
    ExitTransactionButtonPath = "wnd[0]/tbar[0]/btn[15]"
    CloseAreYouSurePopUp = "wnd[1]/usr/btnSPOP-OPTION2"
    
    Dim counter As Integer
    For counter = 1 To QtyOfRepeats
        
        Session.FindById(ExitTransactionButtonPath).Press
        
        Dim ErrorMessage As String
        On Error Resume Next
        ErrorMessage = Session.FindById(CloseAreYouSurePopUp).ID
        
        If ErrorMessage <> "" Then
            Session.FindById(CloseAreYouSurePopUp).Press
        End If
        
    Next counter
    
End Sub

