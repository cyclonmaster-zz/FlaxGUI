VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "FlaxGUI"
   ClientHeight    =   10335
   ClientLeft      =   3225
   ClientTop       =   660
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBrowser.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleMode       =   0  'User
   ScaleWidth      =   15240
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   9135
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   15015
      ExtentX         =   26485
      ExtentY         =   16113
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   12375
      TabIndex        =   11
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop FlaxBasic"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10680
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start FlaxBasic"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10680
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Option"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Advanced Search"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Search"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Index Collections"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Refresh"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Forward ->"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<- Back"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Rem Loads the positions last time the flaxgui was exited
Open App.Path & "/BasicHeight.ini" For Input As #1
Input #1, sText$
Me.Height = sText$
Close #1
Open App.Path & "/BasicWidth.ini" For Input As #1
Input #1, sText$
Me.Width = sText$
Close #1
Open App.Path & "/BasicTop.ini" For Input As #1
Input #1, sText$
Me.Top = sText$
Close #1
Open App.Path & "/BasicLeft.ini" For Input As #1
Input #1, sText$
Me.Left = sText$
Close #1
WB1.Navigate "http://localhost:8090/admin/collections"
End Sub
Private Sub Form_Resize()
Rem used to get the flaxgui control just right to fit the form
On Error Resume Next
If Me.Height > 11130 Then Me.Height = 11130
WB1.Width = Me.Width - 345
WB1.Height = Me.Height - 1935
PB1.Width = WB1.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Rem Saves positions at exiting
Open App.Path & "/BasicHeight.ini" For Output As #1
Print #1, Me.Height
Close #1
Open App.Path & "/BasicWidth.ini" For Output As #1
Print #1, Me.Width
Close #1
Open App.Path & "/BasicTop.ini" For Output As #1
Print #1, Me.Top
Close #1
Open App.Path & "/BasicLeft.ini" For Output As #1
Print #1, Me.Left
Close #1
Unload Me
End
End Sub
Private Sub Label1_Click()
Rem Goes back to the previous flaxgui page
On Error GoTo DieError
WB1.GoBack
Exit Sub
DieError:
Exit Sub
End Sub

Private Sub Label10_Click()
Shell App.Path & "\stopflaxservice.bat", vbMaximizedFocus
End Sub

Private Sub Label11_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Rem Goes to the next page
On Error GoTo DieError
WB1.GoForward
Exit Sub
DieError:
Exit Sub
End Sub
Private Sub Label3_Click()
Rem Refreshes the webpage
WB1.Refresh
End Sub
Private Sub Label4_Click()
Rem Stops all processes loading the page
WB1.Stop
End Sub

Private Sub Label5_Click()
WB1.Navigate "http://localhost:8090/admin/collections"
End Sub

Private Sub Label6_Click()
WB1.Navigate "http://localhost:8090/admin/search"
End Sub

Private Sub Label7_Click()
WB1.Navigate "http://localhost:8090/admin/advanced_search"
End Sub

Private Sub Label8_Click()
WB1.Navigate "http://localhost:8090/admin/options"
End Sub

Private Sub Label9_Click()
Shell App.Path & "\startflaxservice.bat", vbMaximizedFocus
End Sub

Private Sub WB1_StatusTextChange(ByVal Text As String)
Rem Description of page
Me.Caption = WB1.LocationName
End Sub

