VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Lights"
   ClientHeight    =   3540
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   2820
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   60
      TabIndex        =   12
      Top             =   2250
      Width           =   2685
      Begin VB.CheckBox chkBlink 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Blinking corners"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   14
         Top             =   390
         Width           =   2055
      End
      Begin VB.CheckBox chkReverse 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lights reverse back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   825
   End
   Begin VB.Frame fraLights 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   2685
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   -30
         X2              =   2670
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scene 1:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   510
         Width           =   765
      End
      Begin VB.Image imgPlay1 
         Height          =   225
         Index           =   0
         Left            =   1230
         MouseIcon       =   "frmMain.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":0594
         Tag             =   "0"
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SCR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   90
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1650
         TabIndex        =   8
         Top             =   90
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1140
         TabIndex        =   7
         Top             =   90
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scene 3:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scene 2:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   855
         Width           =   765
      End
      Begin VB.Image imgPlay3 
         Height          =   225
         Index           =   2
         Left            =   2220
         MouseIcon       =   "frmMain.frx":08A6
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":09F8
         Tag             =   "0"
         Top             =   1170
         Width           =   225
      End
      Begin VB.Image imgPlay3 
         Height          =   225
         Index           =   1
         Left            =   1725
         MouseIcon       =   "frmMain.frx":0D0A
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":0E5C
         Tag             =   "0"
         Top             =   1170
         Width           =   225
      End
      Begin VB.Image imgPlay2 
         Height          =   225
         Index           =   2
         Left            =   2220
         MouseIcon       =   "frmMain.frx":116E
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":12C0
         Tag             =   "0"
         Top             =   825
         Width           =   225
      End
      Begin VB.Image imgPlay2 
         Height          =   225
         Index           =   1
         Left            =   1725
         MouseIcon       =   "frmMain.frx":15D2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1724
         Tag             =   "0"
         Top             =   825
         Width           =   225
      End
      Begin VB.Image imgPlay1 
         Height          =   225
         Index           =   2
         Left            =   2220
         MouseIcon       =   "frmMain.frx":1A36
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1B88
         Tag             =   "0"
         Top             =   480
         Width           =   225
      End
      Begin VB.Image imgPlay1 
         Height          =   225
         Index           =   1
         Left            =   1725
         MouseIcon       =   "frmMain.frx":1E9A
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1FEC
         Tag             =   "0"
         Top             =   480
         Width           =   225
      End
      Begin VB.Image imgPlay3 
         Height          =   225
         Index           =   0
         Left            =   1230
         MouseIcon       =   "frmMain.frx":22FE
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":2450
         Tag             =   "0"
         Top             =   1170
         Width           =   225
      End
      Begin VB.Image imgPlay2 
         Height          =   225
         Index           =   0
         Left            =   1230
         MouseIcon       =   "frmMain.frx":2762
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":28B4
         Tag             =   "0"
         Top             =   825
         Width           =   225
      End
   End
   Begin VB.Timer timerKeyPress 
      Left            =   2280
      Top             =   120
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   2685
      Begin VB.TextBox txtSpeed 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   870
         MaxLength       =   4
         TabIndex        =   2
         Top             =   120
         Width           =   585
      End
      Begin VB.Image imgOFF 
         Height          =   225
         Left            =   1680
         Picture         =   "frmMain.frx":2BC6
         Top             =   180
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Image imgON 
         Height          =   225
         Left            =   1950
         Picture         =   "frmMain.frx":2ED8
         Top             =   180
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1020
      TabIndex        =   0
      Top             =   3120
      Width           =   825
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Keyboard lights by Deepak Gurung <deepak_tamu@hotmail.com><deepak.gurung@yahoo.com>
Option Explicit

Dim oKeyboard As clsKeyboard

Dim blnKeyNum As Boolean
Dim blnKeyCaps As Boolean
Dim blnKeyScroll As Boolean

Dim blnFlag As Boolean

Dim cnt As Integer

Private Sub chkReverse_Click()
    chkBlink.Enabled = chkReverse.Value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSpeed_Click()
    'Set lights counter to 0
    cnt = 0
    'Put all lights off.
    AllLightsOFF
    
    If CStr(Trim(txtSpeed.Text)) = "" Then
       txtSpeed.Text = 150
    End If
    timerKeyPress.Interval = txtSpeed.Text
End Sub

Private Sub Form_Load()
    Set oKeyboard = New clsKeyboard
    
    'Lights blinking speed.
    txtSpeed.Text = 150
    'Set new shape for controls.
    MakeFlat txtSpeed.hwnd
    MakeFlat cmdSpeed.hwnd
    MakeFlat cmdClose.hwnd
    MakeFlat fraMain.hwnd
    MakeFlat fraLights.hwnd
    MakeFlat fraOptions.hwnd
    
    'Get the default state of the lights.
    oKeyboard.GetLockStatus blnKeyCaps, blnKeyNum, blnKeyScroll
    timerKeyPress.Interval = txtSpeed.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnTmpNum As Boolean
    Dim blnTmp1 As Boolean
    Dim blnTmp2 As Boolean
    Dim blnTmp3 As Boolean
    
    'Get the current lights state
    oKeyboard.GetLockStatus blnTmp1, blnTmp2, blnTmp3
    
    'Set the default state of the lights back.
    If blnTmp1 Then
       If Not blnKeyCaps Then
          oKeyboard.PressKeyVK enumKeys.keyCapsLock, False, False, False
       End If
    Else
       If blnKeyCaps Then
          oKeyboard.PressKeyVK enumKeys.keyCapsLock, False, False, False
       End If
    End If
    
    If blnTmp2 Then
       If Not blnKeyNum Then
          oKeyboard.PressKeyVK enumKeys.keyNumLock, False, False, False
       End If
    Else
       If blnKeyNum Then
          oKeyboard.PressKeyVK enumKeys.keyNumLock, False, False, False
       End If
    End If
    
    If blnTmp3 Then
       If Not blnKeyScroll Then
          oKeyboard.PressKeyVK enumKeys.keyScrollLock, False, False, False
       End If
    Else
       If blnKeyScroll Then
          oKeyboard.PressKeyVK enumKeys.keyScrollLock, False, False, False
       End If
    End If
End Sub

Private Sub imgPlay1_Click(Index As Integer)
    If CStr(imgPlay1(Index).Tag) = "" Or CStr(imgPlay1(Index).Tag) = "0" Then
       imgPlay1(Index).Tag = "1"
       imgPlay1(Index).Picture = imgON
    Else
       imgPlay1(Index).Tag = "0"
       imgPlay1(Index).Picture = imgOFF
    End If
End Sub
Private Sub imgPlay2_Click(Index As Integer)
    If CStr(imgPlay2(Index).Tag) = "" Or CStr(imgPlay2(Index).Tag) = "0" Then
       imgPlay2(Index).Tag = "1"
       imgPlay2(Index).Picture = imgON
    Else
       imgPlay2(Index).Tag = "0"
       imgPlay2(Index).Picture = imgOFF
    End If
End Sub
Private Sub imgPlay3_Click(Index As Integer)
    If CStr(imgPlay3(Index).Tag) = "" Or CStr(imgPlay3(Index).Tag) = "0" Then
       imgPlay3(Index).Tag = "1"
       imgPlay3(Index).Picture = imgON
    Else
       imgPlay3(Index).Tag = "0"
       imgPlay3(Index).Picture = imgOFF
    End If
End Sub

Private Sub timerKeyPress_Timer()
    AllLightsOFF
    Select Case cnt
      Case 0:
           If imgPlay1(0).Tag Then oKeyboard.PressKeyVK enumKeys.keyNumLock, False, False, False
           If imgPlay1(1).Tag Then oKeyboard.PressKeyVK enumKeys.keyCapsLock, False, False, False
           If imgPlay1(2).Tag Then oKeyboard.PressKeyVK enumKeys.keyScrollLock, False, False, False
      Case 1:
           If imgPlay2(0).Tag Then oKeyboard.PressKeyVK enumKeys.keyNumLock, False, False, False
           If imgPlay2(1).Tag Then oKeyboard.PressKeyVK enumKeys.keyCapsLock, False, False, False
           If imgPlay2(2).Tag Then oKeyboard.PressKeyVK enumKeys.keyScrollLock, False, False, False
      Case 2:
           If imgPlay3(0).Tag Then oKeyboard.PressKeyVK enumKeys.keyNumLock, False, False, False
           If imgPlay3(1).Tag Then oKeyboard.PressKeyVK enumKeys.keyCapsLock, False, False, False
           If imgPlay3(2).Tag Then oKeyboard.PressKeyVK enumKeys.keyScrollLock, False, False, False
    End Select
    
    If chkReverse.Value Then
       If chkBlink.Value Then
          If cnt > 2 Then blnFlag = False
          If cnt < 0 Then blnFlag = True
       Else
          If cnt >= 2 Then blnFlag = False
          If cnt <= 0 Then blnFlag = True
       End If
       
       If blnFlag Then
          cnt = cnt + 1
       Else
          cnt = cnt - 1
       End If
    Else
       cnt = cnt + 1
       If cnt > 2 Then cnt = 0
    End If
End Sub

'This function puts off all lights.
Function AllLightsOFF()
    Dim blnTmp1 As Boolean
    Dim blnTmp2 As Boolean
    Dim blnTmp3 As Boolean
    
    oKeyboard.GetLockStatus blnTmp1, blnTmp2, blnTmp3
    Debug.Print blnKeyCaps & "-" & blnKeyNum & "-" & blnKeyScroll
    
    If blnTmp1 Then
       oKeyboard.PressKeyVK enumKeys.keyCapsLock, False, False, False
    End If
    
    If blnTmp2 Then
       oKeyboard.PressKeyVK enumKeys.keyNumLock, False, False, False
    End If
    
    If blnTmp3 Then
       oKeyboard.PressKeyVK enumKeys.keyScrollLock, False, False, False
    End If
End Function

Private Sub txtSpeed_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 Then
       KeyAscii = 0
    End If
End Sub
