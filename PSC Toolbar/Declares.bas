Attribute VB_Name = "mdlDeclares"
'*********************************************************************************************
'
' Shell Bands
'
' Declarations module
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 03/12/2000
' Updates: 03/21/2000:
'                      * FindIESite now uses the
'                        IServiceProvider interface
'                        of the band site to get
'                        the IE window.
'
'*********************************************************************************************
Option Explicit

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_HIDE = 0
Public Const SW_SHOW = 1
Public Const SW_SHOWNOACTIVATE = 1

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = (-4)

Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const WS_EX_CLIENTEDGE = &H200&

Public Const CCS_NORESIZE = &H4&
Public Const CCS_NODIVIDER = &H40&

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_ALLCHILDREN = &H80

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EN_SETFOCUS = &H100
Public Const EN_KILLFOCUS = &H200
Public Const ES_AUTOHSCROLL = &H80&

Public Const HWND_MESSAGE = -3

Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETTEXT = &HC
Public Const WM_SETFONT = &H30
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const PAGE_EXECUTE_READWRITE = &H40

Public Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Public Const PROP_PREVPROC = "WinProc"
Public Const PROP_OBJECT = "Object"

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Any) As Long

Public Declare Sub InitCommonControls Lib "comctl32" ()

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Type TBBUTTON
   iBitmap As Long
   idCommand As Long
   fsState As Byte
   fsStyle As Byte
   bReserved(0 To 1) As Byte
   dwData As Long
   iString As Long
End Type

Public intDefault As Integer
Public strDefaultValue As String
Public intDB_Search As Integer
Public intKioskID_Search As Integer
Public intStreet_Search As Integer
Public intNewWindow As Integer
Public intLateList As Integer
Public intRepairList As Integer
Public intHDManagementList As Integer
Public intHDRetrievalList As Integer
Public intKioskID_ID As Integer
Public intDB_ID As Integer
Public intStreet_ID As Integer
Public intLateList_ID As Integer
Public intRepairList_ID As Integer
Public intHDManagement_ID As Integer
Public intHDRetrieval_ID As Integer
Public intLBound As Integer
Public intUBound As Integer
Public intButtonsToShow As Integer

Public Const TBSTYLE_TOOLTIPS = &H100&
Public Const TBSTYLE_FLAT = &H800&
Public Const TBSTYLE_LIST = &H1000&
Public Const TBSTYLE_TRANSPARENT = &H8000&

Public Const BTNS_BUTTON = &H0
Public Const BTNS_SEP = &H1
Public Const BTNS_AUTOSIZE = &H10

Public Const TBSTATE_ENABLED = &H4

Public Const TB_SETIMAGELIST = &H400 + 48
Public Const TB_ADDBUTTONSW = &H400 + 68

Public Const ILC_MASK = &H1&
Public Const ILC_COLOR8 = &H8&

Private Const strINI_PATH As String = "C:\Windows\vctb.sys"

Public Declare Function ImageList_Create Lib "comctl32" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal himl As Long, ByVal I As Long, ByVal hicon As Long) As Long

Public Declare Function CreateToolbarEx Lib "comctl32" ( _
   ByVal hwnd As Long, _
   ByVal ws As Long, _
   ByVal wID As Long, _
   ByVal nBitmaps As Long, _
   ByVal hBMInst As Long, _
   ByVal wBMID As Long, _
   lpButtons As Any, _
   ByVal iNumButtons As Long, _
   ByVal dxButton As Long, _
   ByVal dyButton As Long, _
   ByVal dxBitmap As Long, _
   ByVal dyBitmap As Long, _
   ByVal uStructSize As Long) As Long


'
' FindIESite
'
' Returns the explorer window that contains
' the band site
'
' Parameters:
'
' BandSite    IOleWindow interface of the band site
'
Public Function FindIESite(ByVal BandSite As olelib.IServiceProvider) As IWebBrowserApp
Dim IID_IWebBrowserApp As olelib.UUID
Dim SID_SInternetExplorer As olelib.UUID
  
   ' Convert IID and SID
   ' from strings to UUID UDTs
   CLSIDFromString IIDSTR_IWebBrowserApp, IID_IWebBrowserApp
   CLSIDFromString SIDSTR_SInternetExplorer, SID_SInternetExplorer
   
   ' Get the InternetExplorer
   ' object through IServiceProvider
   BandSite.QueryService SID_SInternetExplorer, IID_IWebBrowserApp, FindIESite
         
End Function
Sub Main()

   ' Initialize common controls
'   InitCommonControls
   
End Sub

Public Function LoadSettings()
    
    intButtonsToShow = 1
    intDefault = ReadINI("Text", "DefaultValue")
    strDefaultValue = ReadINI("Text", "Value")
    intDB_Search = ReadINI("SearchButtons", "DB")
    intKioskID_Search = ReadINI("SearchButtons", "KioskID")
    intStreet_Search = ReadINI("SearchButtons", "Street")
    intNewWindow = ReadINI("SearchButtons", "NewWindow")
    intLateList = ReadINI("TickerOptions", "LateList")
    intRepairList = ReadINI("TickerOptions", "RepairList")
    intHDManagementList = ReadINI("TickerOptions", "HDManagement")
    intHDRetrievalList = ReadINI("TickerOptions", "HDRetrieval")
    intKioskID_ID = ReadINI("ButtonID", "KioskID_ID")
    intDB_ID = ReadINI("ButtonID", "DB_ID")
    intStreet_ID = ReadINI("ButtonID", "Street_ID")
    intLateList_ID = ReadINI("ButtonID", "LateList_ID")
    intRepairList_ID = ReadINI("ButtonID", "RepairList_ID")
    intHDManagement_ID = ReadINI("ButtonID", "HDManagement_ID")
    intHDRetrieval_ID = ReadINI("ButtonID", "HDRetrieval_ID")
    intLBound = ReadINI("ButtonArray", "Min")
    intUBound = ReadINI("ButtonArray", "Max")
    
'    If intDB_Search = 1 Then
'        intButtonsToShow = intButtonsToShow + 1
'    Else
'        intButtonsToShow = intButtonsToShow - 1
'    End If
'
'    If intKioskID_Search = 1 Then
'        intButtonsToShow = intButtonsToShow + 1
'    Else
'        intButtonsToShow = intButtonsToShow - 1
'    End If
'
'    If intStreet_Search = 1 Then
'        intButtonsToShow = intButtonsToShow + 1
'    Else
'        intButtonsToShow = intButtonsToShow - 1
'    End If
'
'    If intButtonsToShow < 0 Then
'        intButtonsToShow = 0
'    End If
'
'    If intButtonsToShow = 4 Then
'        intButtonsToShow = intButtonsToShow - 1
'    End If
    
End Function

Public Function SaveSettings()
    
    intDefault = WriteINI("Text", "DefaultValue", CStr(intDefault))
    strDefaultValue = WriteINI("Text", "Value", strDefaultValue)
    'intDB_Search = WriteINI("SearchButtons", "DB", CStr(intDB_Search))
    'intKioskID_Search = WriteINI("SearchButtons", "KioskID", CStr(intKioskID_Search))
    'intStreet_Search = WriteINI("SearchButtons", "Street", CStr(intStreet_Search))
    intNewWindow = WriteINI("SearchButtons", "NewWindow", CStr(intNewWindow))
    'intLateList = WriteINI("TickerOptions", "LateList", CStr(intLateList))
    'intRepairList = WriteINI("TickerOptions", "RepairList", CStr(intRepairList))
    'intHDManagementList = WriteINI("TickerOptions", "HDManagement", CStr(intHDManagementList))
    'intHDRetrievalList = WriteINI("TickerOptions", "HDRetrieval", CStr(intHDRetrievalList))
'    intKioskID_ID = WriteINI("ButtonID", "KioskID_ID", CStr(intKioskID_ID))
'    intDB_ID = WriteINI("ButtonID", "DB_ID", CStr(intDB_ID))
'    intStreet_ID = WriteINI("ButtonID", "Street_ID", CStr(intStreet_ID))
'    intLateList_ID = WriteINI("ButtonID", "LateList_ID", CStr(intLateList_ID))
'    intRepairList_ID = WriteINI("ButtonID", "RepairList_ID", CStr(intRepairList_ID))
'    intHDManagement_ID = WriteINI("ButtonID", "HDManagement_ID", CStr(intHDManagement_ID))
'    intHDRetrieval_ID = WriteINI("ButtonID", "HDRetrieval_ID", CStr(intHDRetrieval_ID))
    
    'Load the new data back into the module variables.
    LoadSettings
    
End Function

Public Function ReadINI(ByVal strSection As String, ByVal strKeyName As String)
    
    Dim strRet As String
    
    If Not FileExists(strINI_PATH) Then
        Exit Function
    End If
    
    strRet = String(255, Chr(0))
    ReadINI = Left(strRet, GetPrivateProfileString(strSection, strKeyName, "", strRet, Len(strRet), strINI_PATH))

End Function

Public Function WriteINI(ByVal strSection As String, ByVal strKeyName As String, ByVal strNewString As String) As Integer
    
    Dim intRetVal
    
    If Not FileExists(strINI_PATH) Then
        Exit Function
    End If
    
    intRetVal = WritePrivateProfileString(strSection, strKeyName, strNewString, strINI_PATH)
    
End Function

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
Public Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    'Attempt to open the file, the return value of this function is False
    'if an error occurs on open, True otherwise
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err = 0, True, False)

    Close intFileNum

    Err = 0
End Function
