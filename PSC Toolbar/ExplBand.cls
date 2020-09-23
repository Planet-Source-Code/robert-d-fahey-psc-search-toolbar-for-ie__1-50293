VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExplBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "&Page Links"
'*********************************************************************************************
'
' Explorer Band
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 03/12/2000
' Updates: 03/21/2000:
'                     * The band window was moved to the
'                       frmExplorerBand form. This fixes
'                       2 bugs: the Resize event is not
'                       executed and a GPF occurs when
'                       opening IE when the is shown.
'
'*********************************************************************************************
Option Explicit

Implements olelib.IOleWindow
Implements olelib.IDeskBand
Implements olelib.IObjectWithSite
Implements olelib.IPersist
Implements olelib2.IPersistStream
Implements olelib.IContextMenu
Implements IContextMenuCallback

' Band site object
Dim m_Site As olelib.IUnknown

' Band window
Dim m_Band As frmExplorerBand

Private Sub Class_Initialize()

   ReplaceContextMenu Me
   
End Sub

Private Sub IContextMenu_GetCommandString(ByVal idCmd As Long, ByVal uType As olelib.GetCommandStringFlags, pwReserved As Long, ByVal pszName As Long, ByVal cchMax As Long)
   Err.Raise E_FAIL
End Sub

Private Sub IContextMenu_InvokeCommand(lpici As olelib.CMINVOKECOMMANDINFO)

   Select Case lpici.lpVerb
   
      Case 0
         MsgBox "This band was created with Visual Basic 6", vbInformation
   
   End Select
   
End Sub

Private Sub IContextMenu_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As olelib.QueryContextMenuFlags)

End Sub

Private Function IContextMenuCallback_QueryContextMenu(ByVal hMenu As Long, ByVal indexMenu As Long, ByVal idCmdFirst As Long, ByVal idCmdLast As Long, ByVal uFlags As olelib.QueryContextMenuFlags) As Long

   InsertMenu hMenu, 0, MF_BYPOSITION, idCmdFirst, "&About this band"
   InsertMenu hMenu, 1, MF_BYPOSITION Or MF_SEPARATOR, 0, 0&
   
   IContextMenuCallback_QueryContextMenu = 1
   
End Function


Private Sub IDeskBand_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)

   ' Not implemented
   Err.Raise E_NOTIMPL
   
End Sub

Private Sub IDeskBand_GetBandInfo(ByVal dwBandID As Long, ByVal dwViewMode As olelib.GetBandInfo_ViewModes, pdbi As olelib.DESKBANDINFO)
Dim sTitle As String
   
   On Error Resume Next
     
   With pdbi
      
      If (.dwMask And DBIM_MINSIZE) = DBIM_MINSIZE Then
         ' Set minimum size
         .ptMinSize.x = 100
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
         .ptActual.y = 100
      End If
      
      If (.dwMask And DBIM_INTEGRAL) = DBIM_INTEGRAL Then
         .ptIntegral.x = 1
         .ptIntegral.y = 1
      End If
      
      If (.dwMask And DBIM_TITLE) = DBIM_TITLE Then
         
         ' Set band title
         sTitle = m_Band.Caption & vbNullChar
         MoveMemory .wszTitle(0), ByVal StrPtr(sTitle), LenB(sTitle)
         
      End If
      
      If (.dwMask And DBIM_MODEFLAGS) = DBIM_MODEFLAGS Then
         ' Set flags
         .dwModeFlags = DBIMF_VARIABLEHEIGHT
      End If

   End With
   
End Sub

Private Function IDeskBand_GetWindow() As Long

   ' Call IOleWindow implementation
   IDeskBand_GetWindow = IOleWindow_GetWindow
   
End Function

Private Sub IDeskBand_ResizeBorderDW(prcBorder As olelib.RECT, ByVal punkToolbarSite As Long, ByVal fReserved As olelib.BOOL)

   ' Not implemented
   Err.Raise E_NOTIMPL
   
End Sub

Private Sub IDeskBand_CloseDW(ByVal dwReserved As Long)
   
   On Error Resume Next

   ' Hide the UserControl
   Set m_Band = Nothing
   
End Sub

Private Sub IDeskBand_ShowDW(ByVal fShow As olelib.BOOL)
   
   On Error Resume Next

   ' Show/Hide the window
   If fShow Then
      ShowWindow m_Band.hwnd, SW_SHOW
   Else
      ShowWindow m_Band.hwnd, SW_HIDE
   End If
   
End Sub

Private Sub IOleWindow_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)
   Err.Raise E_NOTIMPL
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
      Set m_Band = New frmExplorerBand
      
      ' Get the IE instance
      Set m_Band.IEWindow = FindIESite(m_Site)
      
      ' Get the IOleWindow interface of the band site
      Set oSiteOW = m_Site
      
   Else
      
      ' Destroy the window
      Set m_Band = Nothing
      
   End If
      

End Sub

Private Function IOleWindow_GetWindow() As Long
   
   On Error Resume Next
   
   IOleWindow_GetWindow = m_Band.hwnd
   
End Function

Private Sub IPersist_GetClassID(pClassID As olelib.UUID)
   
   On Error Resume Next
   
   ' Return the CLSID of this class
   CLSIDFromProgID "PSCToolbar.ExplBand", pClassID
 
End Sub

Private Sub IPersistStream_GetClassID(pClassID As olelib.UUID)

   On Error Resume Next
   
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
