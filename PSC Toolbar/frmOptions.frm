VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toolbar Options"
   ClientHeight    =   6585
   ClientLeft      =   4830
   ClientTop       =   1830
   ClientWidth     =   5715
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   5715
   Begin VB.CommandButton cmdApply 
      Cancel          =   -1  'True
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      MouseIcon       =   "frmOptions.frx":0442
      TabCaption(0)   =   " Options"
      TabPicture(0)   =   "frmOptions.frx":045E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame1 
         Caption         =   "Search Buttons"
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   5295
         Begin VB.CheckBox chkDBSearch 
            Caption         =   "Show DB Search"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   360
            Width           =   3495
         End
         Begin VB.CheckBox chkKioskIDSearch 
            Caption         =   "Show Kiosk ID Search"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   720
            Width           =   3495
         End
         Begin VB.CheckBox chkStreetSearch 
            Caption         =   "Show Street Search"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   1080
            Width           =   3495
         End
         Begin VB.CheckBox chkNewWindow 
            Caption         =   "Open Results in a New Window"
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   1440
            Width           =   3495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search Text"
         Height          =   1215
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   5295
         Begin VB.CheckBox chkDefaultValue 
            Caption         =   "Default Value (requires restart of IE)"
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox txtDefaultValue 
            Height          =   285
            Left            =   720
            TabIndex        =   9
            Top             =   720
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ticker Options"
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   3840
         Width           =   5295
         Begin VB.CheckBox chkHDMTicker 
            Caption         =   "Show HD Management"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   1080
            Width           =   3495
         End
         Begin VB.CheckBox chkRLTicker 
            Caption         =   "Show Repair List"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   720
            Width           =   3495
         End
         Begin VB.CheckBox chkLLTicker 
            Caption         =   "Show Late List"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   3495
         End
         Begin VB.CheckBox chkHDRTicker 
            Caption         =   "Show HD Retrieval"
            Enabled         =   0   'False
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   1440
            Width           =   3495
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub chkDefaultValue_Click()
    
    If chkDefaultValue.Value = vbChecked Then
        txtDefaultValue.Visible = True
    Else
        txtDefaultValue.Visible = False
        txtDefaultValue.Text = ""
        strDefaultValue = ""
    End If
    
End Sub

Private Sub cmdApply_Click()
        
    'Text Section
    intDefault = CInt(chkDefaultValue.Value)
    strDefaultValue = Trim(txtDefaultValue.Text)
    
    'Search Buttons Section
    intDB_Search = CInt(chkDBSearch.Value)
    intKioskID_Search = CInt(chkKioskIDSearch.Value)
    intStreet_Search = CInt(chkStreetSearch.Value)
    intNewWindow = CInt(chkNewWindow.Value)
    
    'Ticker Buttons Section
    intLateList = CInt(chkLLTicker.Value)
    intRepairList = CInt(chkRLTicker.Value)
    intHDManagementList = CInt(chkHDMTicker.Value)
    intHDRetrievalList = CInt(chkHDRTicker.Value)
    
    SaveSettings
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    cmdApply_Click
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    LoadSettings
    
    If intDefault = 1 Then
        chkDefaultValue.Value = intDefault
        txtDefaultValue.Visible = True
        txtDefaultValue.Text = strDefaultValue
    Else
        chkDefaultValue.Value = intDefault
        txtDefaultValue.Visible = False
    End If
        
'    chkKioskIDSearch.Value = intKioskID_Search
'    chkDBSearch.Value = intDB_Search
'    chkStreetSearch.Value = intStreet_Search
    chkNewWindow.Value = intNewWindow
'    chkLLTicker.Value = intLateList
'    chkRLTicker.Value = intRepairList
'    chkHDMTicker.Value = intHDManagementList
'    chkHDRTicker.Value = intHDRetrievalList
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Unload Me
    
End Sub
