VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSideLogo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Side Logo"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   2850
   Icon            =   "SideLogo - Updated.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      FillColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3165
      Left            =   0
      ScaleHeight     =   3165
      ScaleWidth      =   345
      TabIndex        =   7
      Top             =   0
      Width           =   345
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.Timer tmr 
         Interval        =   1
         Left            =   720
         Top             =   0
      End
      Begin MSComDlg.CommonDialog Cdlg 
         Left            =   1080
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Side Logo"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton cmdEndColor 
         Caption         =   "End Color"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdStartColor 
         Caption         =   "Start Color"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton cmdTextColor 
         Caption         =   "Text Color"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblMSG 
         Alignment       =   2  'Center
         Caption         =   "Text to show on the Side Logo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmSideLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cL As New cSideLogo 'Links to the class object

Private Sub cmdEndColor_Click() 'End Color
Cdlg.ShowColor 'Shows the Common Dialog Colors
cL.EndColor = Cdlg.Color 'Changes the Side Logo's End Color
End Sub

Private Sub cmdExit_Click() 'Ends Application
End 'Ends Side Logo
End Sub

Private Sub cmdStartColor_Click() 'Start Color
Cdlg.ShowColor 'Shows the Common Dialog Colors
cL.StartColor = Cdlg.Color 'Changes the Side Logo's Start Color
End Sub

Private Sub cmdTextColor_Click() 'Text Color
tmr.Enabled = True 'Enables the Timer
Cdlg.ShowColor 'Shows the Common Dialog Colors
picLogo.ForeColor = Cdlg.Color 'Changes the Side Logo's Text Color
End Sub

Private Sub Form_Load() 'Loads the Form
tmr.Enabled = False 'Disables the Timer
cL.DrawingObject = picLogo
    cL.Caption = txtText 'Add the Text to the Side Logo
End Sub

Private Sub Form_Resize() 'Resizes Side Logo when for is resized
    On Error Resume Next
    picLogo.Height = Me.ScaleHeight
    On Error GoTo 0
    cL.Draw
End Sub

Private Sub tmr_Timer() 'Timer to change object
cL.DrawingObject = picLogo
    cL.Draw
tmr.Enabled = False 'Disables the Timer
End Sub

Private Sub txtText_Change() 'Text for the Side Logo
tmr.Enabled = True 'Enables the Timer
    cL.Caption = txtText
End Sub

