VERSION 5.00
Begin VB.Form StartForm 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'Kein
   Caption         =   "Shot It"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   Icon            =   "StartForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8655
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox YNameTextBox 
      Height          =   315
      Left            =   6180
      MaxLength       =   14
      TabIndex        =   9
      Top             =   5880
      Width           =   2355
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   1440
      Picture         =   "StartForm.frx":000C
      ScaleHeight     =   855
      ScaleWidth      =   5955
      TabIndex        =   7
      Top             =   360
      Width           =   5955
   End
   Begin VB.ComboBox ResCombo 
      Height          =   315
      ItemData        =   "StartForm.frx":2453
      Left            =   6180
      List            =   "StartForm.frx":2455
      Style           =   2  'Dropdown-Liste
      TabIndex        =   4
      Top             =   5460
      Width           =   2355
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      Caption         =   "Mail: mathiaskunter@yahoo.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   2640
      Width           =   3555
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      Caption         =   "Autor: Mathias Kunter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Top             =   2400
      Width           =   3075
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      Caption         =   "Built: 6th of July 2001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   3075
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      Caption         =   "Version 1.4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   1320
      Width           =   5355
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Your name (for the highscore), maximum 14 signs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   480
      TabIndex        =   8
      Top             =   5880
      Width           =   5535
   End
   Begin VB.Label RenderDevInfo 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   5460
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Resolution:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   4860
      TabIndex        =   5
      Top             =   5460
      Width           =   1155
   End
   Begin VB.Label LeaveIt 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "LEAVE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   795
      Left            =   5280
      MouseIcon       =   "StartForm.frx":2457
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   3
      Top             =   6180
      Width           =   2415
   End
   Begin VB.Label GoForIt 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   795
      Left            =   1260
      MouseIcon       =   "StartForm.frx":25A9
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   2
      Top             =   6180
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FF0000&
      Caption         =   $"StartForm.frx":26FB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "WARNING"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   3420
      TabIndex        =   0
      Top             =   3180
      Width           =   2055
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Autor: Mathias Kunter
'Mail: mathiaskunter@yahoo.de

'Time of basic development:
'28. March 2001 to 23. April 2001
'"Updates" and some corrections were done up to:
'30. July 2001

'The landscape, the bot and the MG were created by a
'CAD program, which is also written by me.

'Version 1.3 supports variable sizes of the landscape because of the creation
'of own levels in DX Cad (aviable on http://www.planetsourcecode.com).

'The game engine is developed with VB 6.0. Due to this, you
'should have at least a CPU with 400 MHZ and a
'3d graphics card for hardware rendering.
'This game also needs DirectX 7 or higher to run.



'NOTICE:
'This code is open source and freeware.
'Use it or parts of it if you need but if you do so please give me credits
'or mail me first. Thanks!
'Sorry that the code is not very much commented, but I never do this.
'...and I'm German so don't wonder about my English :-)
'I'm looking forward for any feedback, comments or critcs.
'Please mail me at

'mathiaskunter@yahoo.de




Option Explicit

Private Sub Form_Load()
    Dim i%, GetRes%, GetName$
    
    On Local Error Resume Next
    If Not Mk3d.InitDX Then
        MsgBox "There was an error while loading DirectX. DirectX 7 or higher is needed.", vbCritical
        Mk3d.ExitDX
        End
    End If
    
    'Set the resolution-combo
    ResCombo.Clear
    For i = 0 To UBound(Mk3d.VPAbleSize)
        ResCombo.AddItem Mk3d.VPAbleSize(i, 0) & " x " & Mk3d.VPAbleSize(i, 1)
    Next i
    ResCombo.ListIndex = ResCombo.ListCount - 1
    
    'load resolution and name from file, if possible
    GetRes = -1
    Open App.Path & "\Data\Setting.dat" For Input As #1
    Input #1, GetRes
    Input #1, GetName
    Close #1
    If GetRes >= 0 And GetRes < ResCombo.ListCount Then ResCombo.ListIndex = GetRes
    YNameTextBox.Text = GetName
    
    'show RenderDevInfo
    If Mk3d.RenderState = RGB Then
        RenderDevInfo.Caption = "Render-Device: Software"
    Else
        RenderDevInfo.Caption = "Render-Device: Hardware"
    End If
End Sub


Private Sub GoForIt_Click()
    Dim YName$
    
    On Local Error Resume Next
    Err = 0
    
    Mk3d.VPSize(0) = Mk3d.VPAbleSize(ResCombo.ListIndex, 0)
    Mk3d.VPSize(1) = Mk3d.VPAbleSize(ResCombo.ListIndex, 1)
    YName = Trim(YNameTextBox.Text)
    If YName = "" Then
        MsgBox "Please enter a valid name.", vbInformation
        Exit Sub
    End If
    
    'save resolution and name to file, if possible
    Open App.Path & "\Data\Setting.dat" For Output As #1
    Write #1, ResCombo.ListIndex
    Write #1, YName
    Close #1
    If Not Err = 0 Then MsgBox "The settings couldn't be saved because the files are write-protected. Disable the write-protection in the Data-folder.", vbInformation
    
    Set Game.GameFont = StartForm.Label1.Font
    Me.Hide
    DoEvents
    RenderForm.Show
    DoEvents
    
    'init DirectDraw, Direct3D, DirectInput
    If Not Mk3d.InitDDraw(Game.GameFont) Then
        MsgBox "There was an error while loading DirectDraw.", vbCritical
        Mk3d.ExitDX
        End
    End If
    If Not Mk3d.InitD3D Then
        MsgBox "There was an error while loading Direct3D.", vbCritical
        Mk3d.ExitDX
        End
    End If
    If Not Mk3d.InitDInput Then
        MsgBox "There was an error while loading DirectInput.", vbCritical
        Mk3d.ExitDX
        End
    End If
    If Not Mk3d.InitDSound Then
        MsgBox "There was an error while loading DirectSound.", vbCritical
        Mk3d.ExitDX
        End
    End If
    Mk3d.SetClipPlane 0.01, 1000
    Randomize Timer
    ShowCursor False
    
    Game.Menu YName
    Unload Me
    End
End Sub


Private Sub LeaveIt_Click()
    Mk3d.ExitDX
    End
End Sub
