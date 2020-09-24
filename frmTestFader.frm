VERSION 5.00
Object = "*\AprjFader.vbp"
Begin VB.Form frmTestFader 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Test Fader"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fra 
      BackColor       =   &H00808080&
      Height          =   765
      Left            =   592
      TabIndex        =   2
      Top             =   6180
      Width           =   4860
      Begin VB.HScrollBar scrOpacity 
         Height          =   240
         Left            =   735
         Max             =   100
         Min             =   25
         TabIndex        =   4
         Top             =   435
         Value           =   100
         Width           =   3300
      End
      Begin VB.Label lblMax 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4095
         TabIndex        =   6
         Top             =   450
         Width           =   465
      End
      Begin VB.Label lblMin 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "25%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   315
         TabIndex        =   5
         Top             =   450
         Width           =   360
      End
      Begin VB.Label lblFactor 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1305
         TabIndex        =   3
         Top             =   180
         Width           =   2130
      End
   End
   Begin VB.PictureBox picMain 
      AutoSize        =   -1  'True
      Height          =   5520
      Left            =   202
      Picture         =   "frmTestFader.frx":0000
      ScaleHeight     =   5460
      ScaleWidth      =   5580
      TabIndex        =   0
      Top             =   135
      Width           =   5640
   End
   Begin prjFader.Fader Fader1 
      Left            =   630
      Top             =   5895
      _ExtentX        =   979
      _ExtentY        =   450
      FadeInSpeed     =   1
      FadeOutSpeed    =   1
   End
   Begin VB.Label lbReady 
      Alignment       =   2  'Zentriert
      BackColor       =   &H0080FF80&
      Caption         =   " Ready "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2670
      TabIndex        =   1
      Top             =   5865
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmTestFader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Fader1_FadeInReady()

    lbReady.Visible = True       'indication that fading is complete and the form is usable

End Sub

Private Sub Fader1_FadeOutReady()

    lbReady.Visible = True       'indication that fading is complete and the form is usable

End Sub

Private Sub Form_Activate()

    lbReady.Visible = False      'hide ready label
    scrOpacity_Scroll            'show current opacity factor

    '''''''''''''
    Fader1.FadeIn                'fade in on load
    '''''''''''''

End Sub

Private Sub Form_Unload(Cancel As Integer)

    lbReady.Visible = False      'hide ready label
    lblFactor.Visible = False    'hide opacity factor

    ''''''''''''''
    Fader1.FadeOut               'fade out on unload
    ''''''''''''''

End Sub

Private Sub lblMax_Click()

    scrOpacity = scrOpacity.Max  'user clicked on label Max so set scroll value

End Sub

Private Sub lblMin_Click()

    scrOpacity = scrOpacity.Min  'user clicked on label Min so set scroll value

End Sub

Private Sub scrOpacity_Change()

    scrOpacity_Scroll            'show set opacity factor
    picMain.SetFocus             'prvent scroll thumb from blinking
    lbReady.Visible = False      'hide ready label

    '''''''''''''''''''''''''''
    Fader1.Opacity = scrOpacity  'set new opacity
    '''''''''''''''''''''''''''

End Sub

Private Sub scrOpacity_Scroll()

    lblFactor = "Opacity Factor " & scrOpacity & "%" 'show current opacity factor

End Sub

':) Ulli's VB Code Formatter V2.14.3 (21.08.2002 11:31:29) 1 + 66 = 67 Lines
