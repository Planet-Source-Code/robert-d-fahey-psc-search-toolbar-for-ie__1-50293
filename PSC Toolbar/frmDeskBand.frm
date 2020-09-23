VERSION 5.00
Begin VB.Form frmDeskBand 
   BorderStyle     =   0  'None
   Caption         =   "Recent VB Projects"
   ClientHeight    =   330
   ClientLeft      =   3405
   ClientTop       =   5940
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   349
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.ComboBox cmbProjects 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4320
   End
End
Attribute VB_Name = "frmDeskBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
'
' ToolBand Window
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

Public IEWindow As InternetExplorer
Attribute IEWindow.VB_VarHelpID = -1

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long

Const HKEY_CURRENT_USER = &H80000001
Const KEY_ALL_ACCESS = &HF003F
Const REG_SZ = 1
Private Sub LoadVBProjects(ByVal VBVer As String)
Dim hKey As Long, lIdx As Long
Dim sProject As String, lProject As Long

   On Error Resume Next
      
   ' Open the RecentFiles key
   If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\" & VBVer & "\RecentFiles", 0, KEY_ALL_ACCESS, hKey) = 0 Then
   
      lIdx = 1
      
      sProject = String$(260, 0)
      lProject = 260
      
      ' Enumerate values until an error is found
      Do While RegQueryValueExStr(hKey, CStr(lIdx), 0, 0, sProject, lProject) = 0
     
         ' Clear the Err object
         Err.Clear
         
         ' Try to open the file
         Open sProject For Input As #1
         
         If Err.Number = 0 Then ' The file exists!
         
            Close #1
         
            cmbProjects.AddItem sProject
            
         End If
         
         lIdx = lIdx + 1
         sProject = String$(260, 0)
         lProject = 260
         
      Loop
   
      ' Close the key
      RegCloseKey hKey
      
   End If

End Sub


Private Sub cmbProjects_Click()

   If cmbProjects.ListIndex <> -1 Then
        
      ' Execute the project
      ShellExecute hwnd, "open", cmbProjects.List(cmbProjects.ListIndex), vbNullString, vbNullString, SW_SHOW
      
      ' Clear the selection
      cmbProjects.ListIndex = -1
      
   End If
   
End Sub


Private Sub cmdUpdate_Click()
   
   ' Clear the combo
   cmbProjects.Clear

   LoadVBProjects "5.0"
   LoadVBProjects "6.0"
   
End Sub

Private Sub Form_Load()

   ' Add child style to the window
   SetWindowLong hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) Or WS_CHILD
   
   cmdUpdate_Click
   
End Sub

Private Sub Form_Resize()
   
   On Error Resume Next
   
   cmdUpdate.Left = ScaleWidth - cmdUpdate.Width
   cmbProjects.Move 0, 0, cmdUpdate.Left - 2

End Sub

