VERSION 5.00
Begin VB.Form frmInfoBand 
   BorderStyle     =   0  'None
   Caption         =   "HTML Source"
   ClientHeight    =   1200
   ClientLeft      =   2655
   ClientTop       =   4995
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   387
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdateHTML 
      Caption         =   "&Update HTML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4905
      TabIndex        =   2
      Top             =   60
      Width           =   885
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4890
      TabIndex        =   1
      Top             =   585
      Width           =   885
   End
   Begin VB.TextBox txtHTML 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   4845
   End
   Begin VB.Menu mnuSubMenu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuStation 
         Caption         =   "Station Search"
      End
      Begin VB.Menu mnuDB 
         Caption         =   "DB Search"
      End
   End
End
Attribute VB_Name = "frmInfoBand"
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

Dim m_bModified As Boolean
Private Sub cmdUpdate_Click()
   
   IEWindow.Document.body.innerHTML = txtHTML.Text
   
   cmdUpdateHTML_Click
   
End Sub

Private Sub cmdUpdateHTML_Click()

   txtHTML.Text = IEWindow.Document.body.innerHTML
   
End Sub


Private Sub Form_Load()
   
   ' Add child style to the window
   SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_CHILD

End Sub

Private Sub Form_Resize()
   
   On Error Resume Next

   cmdUpdate.Move ScaleWidth - cmdUpdate.Width - 2, 2, cmdUpdate.Width, (ScaleHeight - 4) / 2
   cmdUpdateHTML.Move cmdUpdate.Left, cmdUpdate.Top + cmdUpdate.Height, cmdUpdate.Width, cmdUpdate.Height
   txtHTML.Move 2, 2, cmdUpdate.Left - 4, ScaleHeight - 4
   
End Sub


Private Sub IEWindow_DocumentComplete(ByVal pDisp As Object, URL As Variant)
   
   On Error Resume Next
   
   cmdUpdateHTML_Click

End Sub


Private Sub txtHTML_Change()

   If m_bModified = False Then
      m_bModified = True
      cmdUpdate.Enabled = True
   End If

End Sub


