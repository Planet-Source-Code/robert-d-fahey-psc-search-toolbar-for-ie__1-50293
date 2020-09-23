VERSION 5.00
Begin VB.Form frmExplorerBand 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Page Links"
   ClientHeight    =   3825
   ClientLeft      =   3735
   ClientTop       =   3510
   ClientWidth     =   1770
   LinkTopic       =   "Form1"
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   118
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLinks 
      Height          =   1740
      IntegralHeight  =   0   'False
      Left            =   15
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1230
      Width           =   1620
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   345
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Navigates to the selected link"
      Top             =   2970
      Width           =   1635
   End
   Begin VB.CommandButton cmdHideMe 
      Caption         =   "Hide Me"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   930
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Hides this band"
      Top             =   3540
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Location:"
      Height          =   210
      Left            =   15
      TabIndex        =   7
      Top             =   30
      Width           =   1245
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   15
      TabIndex        =   6
      Top             =   270
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Document Title:"
      Height          =   210
      Left            =   15
      TabIndex        =   5
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   15
      TabIndex        =   4
      Top             =   750
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Page Links:"
      Height          =   210
      Left            =   15
      TabIndex        =   3
      Top             =   990
      Width           =   825
   End
End
Attribute VB_Name = "frmExplorerBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' Vertical Explorer Band Window
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Created: 03/21/2000
'
'*********************************************************************************************
Option Explicit

' Explorer window reference
Public WithEvents IEWindow As InternetExplorer
Attribute IEWindow.VB_VarHelpID = -1

Public Sub UpdatePageInfo()
Dim lIdx As Long

   On Error Resume Next
   
   With IEWindow
   
      ' Update URL
      lblURL.Caption = .LocationURL
      lblURL.ToolTipText = .Document.Title
      
      ' Update Name
      lblName.Caption = .LocationName
      
      ' Update Links
      
      ' Clear link list
      lstLinks.Clear
      
      For lIdx = 0 To .Document.links.length
         
         With .Document.links(lIdx)
         
            If Trim$(.innerText) = "" Then
               lstLinks.AddItem .href
            Else
               lstLinks.AddItem Trim$(.innerText)
            End If
            
            lstLinks.ItemData(lstLinks.NewIndex) = lIdx
            
         End With
         
         DoEvents
         
      Next
   
   End With
   
End Sub



Private Sub Form_Load()

   ' Add child style to the window
   SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_CHILD

End Sub

Private Sub IEWindow_DocumentComplete(ByVal pDisp As Object, URL As Variant)
   
   On Error Resume Next
   
   UpdatePageInfo
   
End Sub


Private Sub Form_Resize()
   
   On Error Resume Next

   With cmdHideMe
      .Move ScaleWidth - .Width - 2, ScaleHeight - .Height - 2
   End With
   
   With cmdGo
      .Width = ScaleWidth - 4
      .Top = cmdHideMe.Top - .Height - 2
   End With
   
   With lstLinks
      .Width = ScaleWidth - 4
      .Height = cmdGo.Top - .Top
   End With
   
   RedrawWindow Me.hwnd, ByVal 0&, 0, RDW_ERASE Or RDW_ERASENOW Or RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN
   
End Sub



Private Sub cmdGo_Click()
Dim lIdx As Long

   On Error Resume Next
   
   If lstLinks.ListIndex > -1 Then
   
      ' Get the link index
      lIdx = lstLinks.ItemData(lstLinks.ListIndex)
         
      IEWindow.Navigate IEWindow.Document.links(lIdx).href, , "_top"
      
   End If
   
End Sub


Private Sub cmdHideMe_Click()
Dim tCLSID As UUID
Dim sCLSID As String

   CLSIDFromProgID "PSCToolbar.ExplBand", tCLSID
   sCLSID = Space$(38)
   StringFromGUID2 tCLSID, sCLSID, 39
   
   IEWindow.ShowBrowserBar sCLSID, False
   
End Sub


Private Sub lstLinks_DblClick()

   cmdGo_Click
   
End Sub


