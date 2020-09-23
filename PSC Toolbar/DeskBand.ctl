VERSION 5.00
Begin VB.UserControl Deskband 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "DeskBand.ctx":0000
   ScaleHeight     =   65
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   263
End
Attribute VB_Name = "Deskband"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "&Recent VB Projects - Desktop Band Sample"
'*********************************************************************************************
'
' Shell Desktop Band
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
Implements olelib.IDockingWindow
Implements olelib.IDeskBand
Implements olelib.IObjectWithSite
Implements olelib.IPersist
Implements olelib2.IPersistStream
Implements olelib.IObjectSafety

' Band site object
Dim m_Site As olelib.IUnknown

' Band window
Dim m_Band As frmDeskBand
Private Sub IDeskBand_CloseDW(ByVal dwReserved As Long)

   ' Call IDockingWindow implementation
   IDockingWindow_CloseDW dwReserved
   
End Sub

Private Sub IDeskBand_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)
   
   ' Not implemented
   Err.Raise E_NOTIMPL

End Sub

Private Sub IDeskBand_GetBandInfo(ByVal dwBandID As Long, ByVal dwViewMode As olelib.GetBandInfo_ViewModes, pdbi As olelib.DESKBANDINFO)
Dim sTitle As String
     
   With pdbi
      
      If (.dwMask And DBIM_MINSIZE) = DBIM_MINSIZE Then
         ' Set minimum size
         .ptMinSize.x = 0
         .ptMinSize.y = 22
      End If
      
      If (.dwMask And DBIM_MAXSIZE) = DBIM_MAXSIZE Then
         ' Set maximum size
         .ptMaxSize.y = -1
         .ptMaxSize.x = -1
      End If
      
      If (.dwMask And DBIM_ACTUAL) = DBIM_ACTUAL Then
         ' Set ideal size
         .ptActual.x = 100
         .ptActual.y = 22
      End If
      
      If (.dwMask And DBIM_INTEGRAL) = DBIM_INTEGRAL Then
         .ptIntegral.x = 1
         .ptIntegral.y = 1
      End If
      
      If (.dwMask And DBIM_TITLE) = DBIM_TITLE Then
         sTitle = m_Band.Caption & vbNullChar
         MoveMemory .wszTitle(0), ByVal StrPtr(sTitle), LenB(sTitle)
      End If
      
      If (.dwMask And DBIM_BKCOLOR) = DBIM_BKCOLOR Then
         .crBkgnd = &HD5AC90
      End If
      
      If (.dwMask And DBIM_MODEFLAGS) = DBIM_MODEFLAGS Then
         ' Set flags
         .dwModeFlags = DBIMF_BKCOLOR Or DBIMF_VARIABLEHEIGHT
      End If
   
   End With

End Sub

Private Function IDeskBand_GetWindow() As Long

   ' Call IDockingWindow implementation
   
   IDeskBand_GetWindow = IDockingWindow_GetWindow
   
End Function


Private Sub IDeskBand_ResizeBorderDW(prcBorder As olelib.RECT, ByVal punkToolbarSite As Long, ByVal fReserved As olelib.BOOL)
   
   ' Not implemented
   Err.Raise E_NOTIMPL

End Sub

Private Sub IDeskBand_ShowDW(ByVal fShow As olelib.BOOL)

   ' Call IDockingWindow implementation
   IDockingWindow_ShowDW fShow
   
End Sub


Private Sub IDockingWindow_CloseDW(ByVal dwReserved As Long)

   ' Destroy the window
   Set m_Band = Nothing
   
End Sub

Private Sub IDockingWindow_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)

   ' Not implemented
   Err.Raise E_NOTIMPL

End Sub


Private Function IDockingWindow_GetWindow() As Long

   ' Call IOleWindow implementation
   
   IDockingWindow_GetWindow = IOleWindow_GetWindow
   
End Function


Private Sub IDockingWindow_ResizeBorderDW(prcBorder As olelib.RECT, ByVal punkToolbarSite As Long, ByVal fReserved As olelib.BOOL)
   
   ' Not implemented
   
   Err.Raise E_NOTIMPL

End Sub

Private Sub IDockingWindow_ShowDW(ByVal fShow As olelib.BOOL)

   ' Show/Hide the window
   If fShow Then
      ShowWindow m_Band.hwnd, SW_SHOW
   Else
      ShowWindow m_Band.hwnd, SW_HIDE
   End If
   
End Sub


Private Sub IObjectSafety_GetInterfaceSafetyOptions(riid As olelib.UUID, pdwSupportedOptions As olelib.OBJSAFE_Flags, pdwEnabledOptions As olelib.OBJSAFE_Flags)
   pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or INTERFACESAFE_FOR_UNTRUSTED_DATA
   pdwEnabledOptions = pdwSupportedOptions
End Sub



Private Sub IObjectSafety_SetInterfaceSafetyOptions(riid As olelib.UUID, ByVal dwOptionSetMask As olelib.OBJSAFE_Flags, ByVal dwEnabledOptions As olelib.OBJSAFE_Flags)
'
End Sub


Private Sub IObjectWithSite_GetSite(riid As olelib.UUID, ppvSite As stdole.IUnknown)
Dim lErr As Long

   ' Get the requested interface
   lErr = m_Site.QueryInterface(riid, ppvSite)
   
   If lErr Then Err.Raise lErr

End Sub

Private Sub IObjectWithSite_SetSite(ByVal pUnkSite As stdole.IUnknown)
Dim oSiteOW As IOleWindow
   
   On Error Resume Next

   ' Store the new site object
   Set m_Site = pUnkSite
   
   If Not m_Site Is Nothing Then
            
      ' Create the band window
      Set m_Band = New frmDeskBand
      
      Set m_Band.IEWindow = FindIESite(m_Site)
      
      ' Get the IOleWindow
      ' interface of the band site
      Set oSiteOW = m_Site
      
      ' Move the window
      ' to the band site
      SetParent m_Band.hwnd, oSiteOW.GetWindow()
   
   Else
      
      ' Destroy the window
      Set m_Band = Nothing
      
   End If
      

End Sub

Private Sub IOleWindow_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)

   Err.Raise E_NOTIMPL

End Sub

Private Function IOleWindow_GetWindow() As Long
   
   IOleWindow_GetWindow = m_Band.hwnd
   
End Function

Private Sub IPersist_GetClassID(pClassID As olelib.UUID)
   
   ' Return the CLSID of this class
   CLSIDFromProgID "PSCToolbar.DeskBand", pClassID
 
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






Private Sub IPersistStream_Load(ByVal pStm As olelib.IStream)

End Sub


Private Sub IPersistStream_Save(ByVal pStm As olelib.IStream, ByVal fClearDirty As olelib.BOOL)

End Sub


Private Sub UserControl_Show()
Dim oIE As InternetExplorer
Dim tCLSID As UUID
Dim sCLSID As String

   On Error Resume Next
   
   ' Get the IE instance
   Set oIE = FindIESite(Parent)

   ' Get the explorer band CLSID
   CLSIDFromProgID "PSCToolbar.ExplBand", tCLSID
   sCLSID = Space$(38)
   StringFromGUID2 tCLSID, sCLSID, 39
   
   ' Show the explorer band
   oIE.ShowBrowserBar sCLSID, True
   
   ' Set the control size
   UserControl.Width = 3945
   UserControl.Height = 975
   
End Sub


