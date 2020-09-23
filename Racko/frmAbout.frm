VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00009000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3840
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3225
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   256
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program by Paul Bahlawan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   -120
      TabIndex        =   1
      Top             =   2160
      Width           =   3405
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2085
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Racko
'Paul Bahlawan
'Jan 2011
'
'-------------
'frmAbout
'Version & Credits

Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Revision
    lblInfo = lblInfo & Chr$(13) & Chr$(13) & "Jan. 2011"
    GdiTransparentBlt frmAbout.hDC, 10, 10, 195, 120, frmMain.picBack.hDC, 0, 0, 195, 120, vbMagenta
End Sub
