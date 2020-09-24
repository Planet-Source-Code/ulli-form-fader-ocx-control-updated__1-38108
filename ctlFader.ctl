VERSION 5.00
Begin VB.UserControl Fader 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   630
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   630
   ToolboxBitmap   =   "ctlFader.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Label lbName 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   555
   End
End
Attribute VB_Name = "Fader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'This code is based om a submission to PSC by Ed Preston

Public Enum FadingSpeed
    [Fade Slow] = 1
    [Fade Medium] = 2
    [Fade Fast] = 4
    [Fade Very Fast] = 8
End Enum

'Properties
Private Const pnEnabled         As String = "Enabled"
Private Const pnFadeIn          As String = "FadeInSpeed"
Private Const pnFadeOut         As String = "FadeOutSpeed"
Private Const pnOpacity         As String = "Opacity"
Private myEnabled               As Boolean
Private myFadeInSpeed           As FadingSpeed
Private myFadeOutSpeed          As FadingSpeed
Private myOpacity               As Long

'Private variables
Private Alpha                   As Long
Private ParhWnd                 As Long
Private Internal                As Boolean

'Events
Public Event FadeInReady()
Public Event FadeOutReady()

'Win API
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function OSVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Const RequiredVersion As Long = 5

'Win Consts
Private Const WS_EX_LAYERED     As Long = &H80000
Private Const GWL_EXSTYLE       As Long = -20
Private Const LWA_ALPHA         As Long = 2

Public Property Let Enabled(ByVal nwEnabled As Boolean)
Attribute Enabled.VB_Description = "Sets/returns whether the Control is operable."

    myEnabled = (nwEnabled <> False) And WindowsIsSuitable
    PropertyChanged pnEnabled

End Property

Public Property Get Enabled() As Boolean

    Enabled = myEnabled

End Property

Public Sub FadeIn()

    If myEnabled Then
        For Alpha = Alpha To (myOpacity / 100) * 255 Step myFadeInSpeed
            SetLayeredWindowAttributes ParhWnd, 0, Alpha, LWA_ALPHA
            DoEvents
            Sleep 1
        Next Alpha
        Alpha = Alpha - myFadeInSpeed
        If myOpacity = 100 Then
            SetWindowLong hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
        End If
        If Not Internal Then
            RaiseEvent FadeInReady
        End If
      Else 'MYENABLED = FALSE
        If Not Internal Then
            SetLayeredWindowAttributes ParhWnd, 0, 255, LWA_ALPHA
        End If
    End If

End Sub

Public Property Get FadeInSpeed() As FadingSpeed

    FadeInSpeed = myFadeInSpeed

End Property

Public Property Let FadeInSpeed(ByVal nwFadeInSpeed As FadingSpeed)

    If nwFadeInSpeed = [Fade Very Fast] Or nwFadeInSpeed = [Fade Fast] Or nwFadeInSpeed = [Fade Medium] Or nwFadeInSpeed = [Fade Slow] Then
        myFadeInSpeed = nwFadeInSpeed
        PropertyChanged pnFadeIn
      Else 'NOT NWFADEINSPEED...
        Err.Raise 380
    End If

End Property

Public Sub FadeOut()

    If myEnabled Then
        SetWindowLong ParhWnd, GWL_EXSTYLE, GetWindowLong(ParhWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        For Alpha = Alpha To IIf(Internal, (myOpacity / 100) * 255, 0) Step -myFadeOutSpeed
            SetLayeredWindowAttributes ParhWnd, 0, Alpha, LWA_ALPHA
            DoEvents
            Sleep 1
        Next Alpha
        Alpha = Alpha + myFadeOutSpeed
        If Not Internal Then
            RaiseEvent FadeOutReady
        End If
      Else 'MYENABLED = FALSE
        If Not Internal Then
            SetLayeredWindowAttributes ParhWnd, 0, 0, LWA_ALPHA
        End If
    End If

End Sub

Public Property Get FadeOutSpeed() As FadingSpeed

    FadeOutSpeed = myFadeOutSpeed

End Property

Public Property Let FadeOutSpeed(ByVal nwFadeOutSpeed As FadingSpeed)

    If nwFadeOutSpeed = [Fade Very Fast] Or nwFadeOutSpeed = [Fade Fast] Or nwFadeOutSpeed = [Fade Medium] Or nwFadeOutSpeed = [Fade Slow] Then
        myFadeOutSpeed = nwFadeOutSpeed
        PropertyChanged pnFadeOut
      Else 'NOT NWFADEOUTSPEED...
        Err.Raise 380
    End If

End Property

Public Property Get Opacity() As Long
Attribute Opacity.VB_Description = "Percent value of opacity."

    Opacity = myOpacity

End Property

Public Property Let Opacity(ByVal nwOpacity As Long)

  Dim PreviousOpacity   As Long

    PreviousOpacity = myOpacity
    If nwOpacity >= 25 And nwOpacity <= 100 Then
        myOpacity = nwOpacity
        PropertyChanged pnOpacity
        If Ambient.UserMode Then
            Internal = True
            If myOpacity > PreviousOpacity Then
                FadeIn
                RaiseEvent FadeInReady
              ElseIf myOpacity < PreviousOpacity Then 'NOT MYOPACITY...
                FadeOut
                RaiseEvent FadeOutReady
              Else 'NOT MYOPACITY...
                RaiseEvent FadeOutReady
            End If
            Internal = False
        End If
      Else 'NOT NWOPACITY...
        Err.Raise 380
    End If

End Property

Private Sub UserControl_InitProperties()

    myFadeInSpeed = [Fade Medium]
    myFadeOutSpeed = [Fade Medium]
    myEnabled = WindowsIsSuitable
    myOpacity = 100

End Sub

Private Sub UserControl_Paint()

    lbName = Ambient.DisplayName
    UserControl_Resize

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        myEnabled = .ReadProperty(pnEnabled, True) And WindowsIsSuitable
        myFadeInSpeed = .ReadProperty(pnFadeIn, [Fade Medium])
        myFadeOutSpeed = .ReadProperty(pnFadeOut, [Fade Medium])
        myOpacity = .ReadProperty(pnOpacity, 100)
    End With 'PROPBAG

    If Ambient.UserMode Then
        ParhWnd = Parent.hWnd
        If WindowsIsSuitable Then
            SetWindowLong ParhWnd, GWL_EXSTYLE, GetWindowLong(ParhWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes ParhWnd, 0, 0, LWA_ALPHA
        End If
        Alpha = 1
    End If

End Sub

Private Sub UserControl_Resize()

    Size lbName.Width, lbName.Height

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty pnEnabled, myEnabled, WindowsIsSuitable
        .WriteProperty pnFadeIn, myFadeInSpeed, [Fade Medium]
        .WriteProperty pnFadeOut, myFadeOutSpeed, [Fade Medium]
        .WriteProperty pnOpacity, myOpacity, 100
    End With 'PROPBAG

End Sub

Private Function WindowsIsSuitable() As Boolean

    WindowsIsSuitable = ((OSVersion And &HFF&) >= RequiredVersion)
    
    'uncoment next line for experiments with other Windows'es
   'WindowsIsSuitable = True
    
End Function

':) Ulli's VB Code Formatter V2.14.3 (21.08.2002 10:50:09) 42 + 185 = 227 Lines
