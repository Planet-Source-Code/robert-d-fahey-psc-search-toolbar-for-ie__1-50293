VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Planet Source Code"
'*********************************************************************************************
'
' Explorer Toolbar
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 03/12/2000
'
'*********************************************************************************************
Option Explicit

Implements olelib.IOleWindow
Implements olelib.IDeskBand
Implements olelib.IObjectWithSite
Implements olelib.IInputObject
Implements olelib.IPersist
Implements olelib2.IPersistStream
Implements IInputObjectCallback

Private m_oSite As olelib.IUnknown           ' Band site object
Private m_oIOS As olelib.IInputObjectSite
Private WithEvents m_oIE As InternetExplorer ' IE Object
Attribute m_oIE.VB_VarHelpID = -1

Private m_lMsgWnd As Long     ' Toolbar parent window
Private m_lToolWnd As Long    ' Toolbar window
Private m_lEdit As Long       ' Textbox window
Private m_oFont As IFont      ' Textbox font
Private m_bHasFocus As Boolean

Const ID_SEARCH = 1

Private Sub AddToolbarButtons()
    
    Dim tTB(1 To 10) As TBBUTTON
      
    tTB(1).fsStyle = BTNS_SEP
    tTB(1).iBitmap = 102
    
    tTB(2).idCommand = 1
    tTB(2).iString = StrPtr("Search")
    tTB(2).fsStyle = BTNS_BUTTON Or BTNS_AUTOSIZE
    tTB(2).fsState = TBSTATE_ENABLED
    tTB(2).iBitmap = 0

    SendMessage m_lToolWnd, TB_ADDBUTTONSW, 2, tTB(1)
   
End Sub

Public Sub CreateToolbar()
    
    Dim tTB As TBBUTTON
    Dim lImgList As Long
    Dim x As Integer
    
    ' Create the toolbar parent. This window
    ' will receive the toolbar messages
    m_lMsgWnd = CreateWindowEx(0, "#32770", vbNullString, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, App.hInstance, ByVal 0&)
    
    ' Create the toolbar window
    m_lToolWnd = CreateToolbarEx(m_lMsgWnd, CCS_NODIVIDER Or CCS_NORESIZE Or _
                            TBSTYLE_LIST Or TBSTYLE_TOOLTIPS Or TBSTYLE_FLAT Or TBSTYLE_TRANSPARENT, _
                            0, 0, 0, 0, tTB, 0, 16, 16, 16, 16, LenB(tTB))
    
    AddToolbarIcons
    AddToolbarButtons
    CreateTextBox
    
    ' Subclass the toolbar parent
    SubClass m_lMsgWnd, Me
   
End Sub


Private Sub AddToolbarIcons(Optional ID As Integer)
    Dim lImgList As Long
    Dim lIcon As Long
    
    ' Create a ImageList
    lImgList = ImageList_Create(16, 16, ILC_COLOR8 Or ILC_MASK, 1, 1)
    
    ' Load the icon
     Select Case ID
         Case 1
             lIcon = LoadIcon(App.hInstance, 1&)
         Case 2
             lIcon = LoadIcon(App.hInstance, 2&)
         Case 3
             lIcon = LoadIcon(App.hInstance, 101&)
         Case Else
             lIcon = LoadIcon(App.hInstance, 5&)
     End Select
    
    ' Add the icon to the image list
    ImageList_ReplaceIcon lImgList, -1, lIcon
    
    ' Set the toolbar imagelist
    SendMessage m_lToolWnd, TB_SETIMAGELIST, ByVal 0, ByVal lImgList
   
End Sub

Private Sub CreateTextBox()
   
   ' Create the text box window
   m_lEdit = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "", ES_AUTOHSCROLL Or WS_VISIBLE Or WS_CHILD, 0, 0, 100, 22, m_lToolWnd, 1, App.hInstance, ByVal 0&)
   
   ' Set the textbox font
   Set m_oFont = New StdFont
   With m_oFont
      .Size = 8
      .Name = "Trebuchet MS"
   End With
   SendMessage m_lEdit, WM_SETFONT, m_oFont.hFont, ByVal 0&

End Sub



Public Sub DestroyToolbar()

   UnsubClass m_lMsgWnd
   
   DestroyWindow m_lEdit
   DestroyWindow m_lToolWnd
   DestroyWindow m_lMsgWnd
   
End Sub


Private Sub IDeskBand_CloseDW(ByVal dwReserved As Long)
   
   ' Destroy the windows
   DestroyToolbar

End Sub

Private Sub IDeskBand_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)
   Err.Raise E_NOTIMPL
End Sub

Private Sub IDeskBand_GetBandInfo(ByVal dwBandID As Long, ByVal dwViewMode As olelib.GetBandInfo_ViewModes, pdbi As olelib.DESKBANDINFO)
Dim sTitle As String

   On Error Resume Next
   
   With pdbi
      
      If (.dwMask And DBIM_MINSIZE) = DBIM_MINSIZE Then
         .ptMinSize.x = 100 '320
         .ptMinSize.y = 22
      End If
      
      If (.dwMask And DBIM_MAXSIZE) = DBIM_MAXSIZE Then
         .ptMaxSize.y = -1
         .ptMaxSize.x = -1
      End If
      
      If (.dwMask And DBIM_ACTUAL) = DBIM_ACTUAL Then
         .ptActual.x = 100 '320
         .ptActual.y = 22
      End If
      
      If (.dwMask And DBIM_INTEGRAL) = DBIM_INTEGRAL Then
         .ptIntegral.x = 1
         .ptIntegral.y = 1
      End If
      
      If (.dwMask And DBIM_TITLE) = DBIM_TITLE Then
         sTitle = "PSC Search:"
         MoveMemory .wszTitle(0), ByVal StrPtr(sTitle), LenB(sTitle) + 2
      End If
      
      If (.dwMask And DBIM_MODEFLAGS) = DBIM_MODEFLAGS Then
         .dwModeFlags = DBIMF_NORMAL
      End If
   
   End With

End Sub
Private Function IDeskBand_GetWindow() As Long

   IDeskBand_GetWindow = m_lToolWnd
   
End Function


Private Sub IDeskBand_ResizeBorderDW(prcBorder As olelib.RECT, ByVal punkToolbarSite As Long, ByVal fReserved As olelib.BOOL)
   Err.Raise E_NOTIMPL
End Sub


Private Sub IDeskBand_ShowDW(ByVal fShow As olelib.BOOL)
   
   If fShow Then
      ShowWindow m_lToolWnd, SW_SHOWNOACTIVATE
   Else
      ShowWindow m_lToolWnd, SW_HIDE
   End If
   
End Sub

Private Sub IInputObject_HasFocusIO()
   ' This method is replaced by mdlBandProc.IInputObject_HasFocusIO
End Sub

Private Sub IInputObject_TranslateAcceleratorIO(lpmsg As olelib.MSG)
   ' This method is replaced by mdlBandProc.IInputObject_TranslateAcceleratorIO
End Sub

Private Sub IInputObject_UIActivateIO(ByVal fActivate As olelib.BOOL, lpmsg As olelib.MSG)

   If fActivate Then SetFocus m_lEdit
   
End Sub

Private Sub IInputObjectCallback_ButtonClicked(ByVal ID As Long)
    Dim sSearchStr As String
    Dim lIdx As Long
    Dim strTarget As String
    
    ' Get the text
    sSearchStr = GetText(m_lEdit)
    
    If sSearchStr = "" And ID <> 4 Then Exit Sub
    
    For lIdx = 1 To Len(sSearchStr)
       If Mid$(sSearchStr, lIdx, 1) = " " Then
          Mid$(sSearchStr, lIdx, 1) = "+"
       End If
    Next
    
    'If intNewWindow = 1 Then
        strTarget = "_Blank"
    'Else
    '    strTarget = ""
    'End If
    
    Select Case ID
        Case 1
            m_oIE.Navigate "http://planetsourcecode.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&txtCriteria=" & sSearchStr & "&lngWId=1&B1=Quick+Search", , strTarget
    End Select
   
End Sub

Private Function IInputObjectCallback_HasFocus() As Boolean

   If m_bHasFocus Then
      IInputObjectCallback_HasFocus = S_OK
   Else
      IInputObjectCallback_HasFocus = S_FALSE
   End If
   
End Function

Private Sub IInputObjectCallback_onFocus(ByVal Focused As Boolean)
      
   On Error Resume Next
   
   m_bHasFocus = Focused
   m_oIOS.OnFocusChangeIS Me, Focused
   
End Sub

Private Function IInputObjectCallback_TranslateAccelerator(MSG As olelib.MSG) As Long
Dim lProc As Long

   On Error Resume Next
   
   IInputObjectCallback_TranslateAccelerator = S_FALSE
   
   lProc = GetWindowLong(MSG.hwnd, GWL_WNDPROC)
   
   Select Case MSG.message
      
      Case WM_KEYUP
      
         Select Case MSG.wParam
         
            Case vbKeyLeft, vbKeyRight, vbKeyEnd, vbKeyHome, _
                 vbKeyPageUp, vbKeyPageDown
               
               CallWindowProc lProc, MSG.hwnd, MSG.message, MSG.wParam, MSG.lParam
               
               IInputObjectCallback_TranslateAccelerator = S_OK
         
         End Select
         
      Case WM_KEYDOWN
         
         Select Case MSG.wParam
         
            Case vbKeyLeft, vbKeyRight, vbKeyEnd, vbKeyHome, _
                 vbKeyPageUp, vbKeyPageDown
               
               CallWindowProc lProc, MSG.hwnd, MSG.message, MSG.wParam, MSG.lParam
               
               IInputObjectCallback_TranslateAccelerator = S_OK
               
            Case vbKeyBack
            
               ' Maybe this is not the best way
               ' to process the backspace key
               ' but I haven't found another
               
               Dim lStart As Long, lEnd As Long
               Dim sText As String
               
               ' Get the selection
               lEnd = SendMessage(MSG.hwnd, EM_GETSEL, 0, ByVal 0&)
               lStart = lEnd And &HFFFF&
               lEnd = lEnd \ &H10000
               sText = GetText(MSG.hwnd)
               
               If lStart = lEnd Then
                  If lStart > 0 Then
                     m_oIE.StatusText = lEnd
                     sText = Left$(sText, lStart - 1) & Mid$(sText, lEnd + 1)
                     lStart = lStart - 1
                  End If
               Else
                  sText = Left$(sText, lStart) & Mid$(sText, lEnd + 1)
               End If
            
               SendMessage MSG.hwnd, WM_SETTEXT, 0, ByVal sText
               SendMessage MSG.hwnd, EM_SETSEL, lStart, ByVal lStart
               
               IInputObjectCallback_TranslateAccelerator = S_OK
            
         End Select
         
      Case WM_CHAR
      
         If MSG.wParam = vbKeyReturn Then IInputObjectCallback_ButtonClicked 1
         
   End Select
   
   
End Function

Private Sub IObjectWithSite_GetSite(riid As olelib.UUID, ppvSite As stdole.IUnknown)
Dim lErr As Long

   ' Get the requested interface
   lErr = m_oSite.QueryInterface(riid, ppvSite)
   
   If lErr Then Err.Raise lErr
   
End Sub

Private Sub IObjectWithSite_SetSite(ByVal pUnkSite As stdole.IUnknown)
Dim oSiteOW As IOleWindow
   
   On Error Resume Next

   ' Store the new site object
   Set m_oSite = pUnkSite
   Set m_oIOS = pUnkSite
         
   If Not m_oSite Is Nothing Then
   
      ' Get the IE instance
      Set m_oIE = FindIESite(m_oSite)
      
      ' Create the toolbar
      CreateToolbar
            
   Else
            
      Set m_oIE = Nothing
      
   End If

End Sub
Private Sub IOleWindow_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)
   
   Err.Raise E_NOTIMPL

End Sub

Private Function IOleWindow_GetWindow() As Long
   
   IOleWindow_GetWindow = m_lToolWnd
   
End Function

Private Sub IPersist_GetClassID(pClassID As olelib.UUID)
   
   ' Return the CLSID of this class
   CLSIDFromProgID "PSCToolbar.Toolbar", pClassID
 
End Sub

Private Sub IPersistStream_GetClassID(pClassID As olelib.UUID)

   IPersist_GetClassID pClassID
   
End Sub


Private Function IPersistStream_GetSizeMax() As Currency
   
   Err.Raise E_NOTIMPL

End Function

Private Sub IPersistStream_IsDirty()

   Err.Raise E_NOTIMPL
   
End Sub

Private Sub Class_Initialize()
    Dim oIO As olelib.IInputObject

     'Load Toolbar Option Settings from vctb.sys
     'LoadSettings
     
    ' Replace the IInputObject:HasFocusIO method
    Set oIO = Me
    ReplaceVTableEntry ObjPtr(oIO), 5, AddressOf mdlBandProc.IInputObject_HasFocusIO
    ReplaceVTableEntry ObjPtr(oIO), 6, AddressOf mdlBandProc.IInputObject_TranslateAcceleratorIO
      
End Sub

Private Sub IPersistStream_Load(ByVal pStm As olelib2.IStream)

End Sub

Private Sub IPersistStream_Save(ByVal pStm As olelib2.IStream, ByVal fClearDirty As olelib.BOOL)

End Sub

