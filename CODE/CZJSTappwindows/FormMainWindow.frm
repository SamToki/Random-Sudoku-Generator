VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FormMainWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Sudoku Generator v0.31eng"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   13455
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "FormMainWindow.frx":0000
   LinkTopic       =   "FormMainWindow"
   MaxButton       =   0   'False
   MouseIcon       =   "FormMainWindow.frx":23D2
   MousePointer    =   99  'Custom
   ScaleHeight     =   7575
   ScaleWidth      =   13455
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer TimerGenerator 
      Interval        =   10
      Left            =   5880
      Top             =   6405
   End
   Begin VB.Frame FrameControls 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   7035
      TabIndex        =   103
      Top             =   945
      Width           =   6210
      Begin VB.TextBox TextboxInput 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080A000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5355
         MaxLength       =   1
         MouseIcon       =   "FormMainWindow.frx":2524
         MousePointer    =   99  'Custom
         TabIndex        =   124
         Top             =   2100
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "C"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   0
         Left            =   4200
         MouseIcon       =   "FormMainWindow.frx":2676
         MousePointer    =   99  'Custom
         TabIndex        =   123
         Top             =   2100
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "9"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   9
         Left            =   3255
         MouseIcon       =   "FormMainWindow.frx":27C8
         MousePointer    =   99  'Custom
         TabIndex        =   121
         Top             =   2100
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "8"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   8
         Left            =   2310
         MouseIcon       =   "FormMainWindow.frx":291A
         MousePointer    =   99  'Custom
         TabIndex        =   119
         Top             =   2100
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "7"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   7
         Left            =   1365
         MouseIcon       =   "FormMainWindow.frx":2A6C
         MousePointer    =   99  'Custom
         TabIndex        =   117
         Top             =   2100
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   6
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":2BBE
         MousePointer    =   99  'Custom
         TabIndex        =   115
         Top             =   2100
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   5
         Left            =   4200
         MouseIcon       =   "FormMainWindow.frx":2D10
         MousePointer    =   99  'Custom
         TabIndex        =   113
         Top             =   1365
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   3255
         MouseIcon       =   "FormMainWindow.frx":2E62
         MousePointer    =   99  'Custom
         TabIndex        =   111
         Top             =   1365
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   3
         Left            =   2310
         MouseIcon       =   "FormMainWindow.frx":2FB4
         MousePointer    =   99  'Custom
         TabIndex        =   109
         Top             =   1365
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "2"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   2
         Left            =   1365
         MouseIcon       =   "FormMainWindow.frx":3106
         MousePointer    =   99  'Custom
         TabIndex        =   107
         Top             =   1365
         Width           =   540
      End
      Begin VB.CommandButton CmdNumber 
         Caption         =   "1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   1
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":3258
         MousePointer    =   99  'Custom
         TabIndex        =   105
         Top             =   1365
         Width           =   540
      End
      Begin VB.CommandButton CmdStartReset 
         Caption         =   "&START"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   420
         MouseIcon       =   "FormMainWindow.frx":33AA
         MousePointer    =   99  'Custom
         TabIndex        =   104
         Top             =   525
         Width           =   4320
      End
      Begin VB.Shape ShapeLightGameStatusIndicator 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         Height          =   330
         Left            =   5565
         Shape           =   3  'Circle
         Top             =   525
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   9
         Left            =   3780
         TabIndex        =   122
         Top             =   2310
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   8
         Left            =   2835
         TabIndex        =   120
         Top             =   2310
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   7
         Left            =   1890
         TabIndex        =   118
         Top             =   2310
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   6
         Left            =   945
         TabIndex        =   116
         Top             =   2310
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   5
         Left            =   4725
         TabIndex        =   114
         Top             =   1575
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   3780
         TabIndex        =   112
         Top             =   1575
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   2835
         TabIndex        =   110
         Top             =   1575
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   1890
         TabIndex        =   108
         Top             =   1575
         Width           =   330
      End
      Begin VB.Label LabelNumberCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   945
         TabIndex        =   106
         Top             =   1575
         Width           =   330
      End
   End
   Begin VB.Frame FrameSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   7035
      TabIndex        =   125
      Top             =   3990
      Width           =   6210
      Begin VB.CheckBox CheckboxSlowDownGeneration 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Slo&w down generation for observation"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   315
         MouseIcon       =   "FormMainWindow.frx":34FC
         MousePointer    =   99  'Custom
         TabIndex        =   132
         Top             =   1890
         Width           =   5580
      End
      Begin VB.CheckBox CheckboxUseLegacyGenerationMethod 
         BackColor       =   &H00D0D0D0&
         Caption         =   "&Use legacy generation method, which may result in an unsolvable Sudoku or even generation failure"
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   315
         MouseIcon       =   "FormMainWindow.frx":364E
         MousePointer    =   99  'Custom
         TabIndex        =   133
         Top             =   2310
         Width           =   5580
      End
      Begin VB.HScrollBar HScrollSettingsLargeBlockMaximumFixedAmount 
         Height          =   330
         LargeChange     =   6
         Left            =   4095
         Max             =   8
         Min             =   2
         MouseIcon       =   "FormMainWindow.frx":37A0
         MousePointer    =   99  'Custom
         TabIndex        =   131
         Top             =   1155
         Value           =   6
         Width           =   1800
      End
      Begin VB.HScrollBar HScrollSettingsTotalFixedAmount 
         Height          =   330
         LargeChange     =   55
         Left            =   4095
         Max             =   54
         Min             =   17
         MouseIcon       =   "FormMainWindow.frx":38F2
         MousePointer    =   99  'Custom
         TabIndex        =   128
         Top             =   525
         Value           =   36
         Width           =   1800
      End
      Begin VB.Label LabelSettingsLargeBlockMaximumFixedAmountIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3570
         TabIndex        =   130
         Top             =   1155
         Width           =   435
      End
      Begin VB.Label LabelSettingsTotalFixedAmountIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3570
         TabIndex        =   127
         Top             =   525
         Width           =   435
      End
      Begin VB.Label LabelSettingsLargeBlockMaximumFixedAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum fixed blocks in a &large block: (may not work when the value above is too large)"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   315
         TabIndex        =   129
         Top             =   945
         Width           =   3270
      End
      Begin VB.Label LabelSettingsTotalFixedAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount of all &fixed blocks:"
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   315
         TabIndex        =   126
         Top             =   525
         Width           =   3270
      End
   End
   Begin VB.Timer TimerProgressbarAnimation 
      Interval        =   1
      Left            =   13230
      Top             =   7350
   End
   Begin VB.Timer TimerSudokuBlockAnimation 
      Interval        =   20
      Left            =   6300
      Top             =   6405
   End
   Begin VB.Timer TimerClock 
      Interval        =   500
      Left            =   12810
      Top             =   420
   End
   Begin VB.Label LabelStatusbar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   105
      TabIndex        =   134
      Top             =   7245
      Width           =   13245
   End
   Begin VB.Label LabelStartTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Game not started yet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7035
      MouseIcon       =   "FormMainWindow.frx":3A44
      MousePointer    =   99  'Custom
      TabIndex        =   135
      ToolTipText     =   "Compare this time to the clock to know how long you have spent solving this Sudoku."
      Top             =   210
      Width           =   4740
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   435
      Left            =   315
      TabIndex        =   137
      Top             =   0
      Visible         =   0   'False
      Width           =   435
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   767
   End
   Begin VB.Shape ShapeProgressbar 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   330
      Left            =   0
      Top             =   7245
      Width           =   13455
   End
   Begin VB.Label LabelRowIndicator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   6270
      Width           =   330
   End
   Begin VB.Label LabelColumnIndicator 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6195
      TabIndex        =   11
      Top             =   105
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   9
      Left            =   6195
      TabIndex        =   20
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   8
      Left            =   5565
      TabIndex        =   19
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   7
      Left            =   4935
      TabIndex        =   18
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   4200
      TabIndex        =   17
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   5
      Left            =   3570
      TabIndex        =   16
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   2940
      TabIndex        =   15
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   2205
      TabIndex        =   14
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   1575
      TabIndex        =   13
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   945
      TabIndex        =   12
      Top             =   525
      Width           =   330
   End
   Begin VB.Label LabelColumn 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   6720
      TabIndex        =   21
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   9
      Left            =   420
      TabIndex        =   9
      Top             =   6300
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   8
      Left            =   420
      TabIndex        =   8
      Top             =   5670
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   7
      Left            =   420
      TabIndex        =   7
      Top             =   5040
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   6
      Left            =   420
      TabIndex        =   6
      Top             =   4305
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   5
      Left            =   420
      TabIndex        =   5
      Top             =   3675
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   4
      Left            =   420
      TabIndex        =   4
      Top             =   3045
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   3
      Left            =   420
      TabIndex        =   3
      Top             =   2310
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   2
      Left            =   420
      TabIndex        =   2
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label LabelRow 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   105
      TabIndex        =   10
      Top             =   6720
      Width           =   645
   End
   Begin VB.Label LabelRow 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   420
      TabIndex        =   1
      Top             =   1050
      Width           =   330
   End
   Begin VB.Label LabelClock 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   12180
      MouseIcon       =   "FormMainWindow.frx":3B96
      MousePointer    =   99  'Custom
      TabIndex        =   136
      ToolTipText     =   "Clock"
      Top             =   210
      Width           =   1065
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   4
      X1              =   4725
      X2              =   4725
      Y1              =   840
      Y2              =   6825
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   3
      X1              =   2730
      X2              =   2730
      Y1              =   840
      Y2              =   6825
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   2
      X1              =   735
      X2              =   6720
      Y1              =   4830
      Y2              =   4830
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      X1              =   735
      X2              =   6720
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   80
      Left            =   5460
      TabIndex        =   101
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   79
      Left            =   4830
      TabIndex        =   100
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   78
      Left            =   4095
      TabIndex        =   99
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   77
      Left            =   3465
      TabIndex        =   98
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   76
      Left            =   2835
      TabIndex        =   97
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   75
      Left            =   2100
      TabIndex        =   96
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   74
      Left            =   1470
      TabIndex        =   95
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   73
      Left            =   840
      TabIndex        =   94
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   72
      Left            =   6090
      TabIndex        =   93
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   71
      Left            =   5460
      TabIndex        =   92
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   70
      Left            =   4830
      TabIndex        =   91
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   69
      Left            =   4095
      TabIndex        =   90
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   68
      Left            =   3465
      TabIndex        =   89
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   67
      Left            =   2835
      TabIndex        =   88
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   66
      Left            =   2100
      TabIndex        =   87
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   65
      Left            =   1470
      TabIndex        =   86
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   64
      Left            =   840
      TabIndex        =   85
      Top             =   5565
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   63
      Left            =   6090
      TabIndex        =   84
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   62
      Left            =   5460
      TabIndex        =   83
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   61
      Left            =   4830
      TabIndex        =   82
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   60
      Left            =   4095
      TabIndex        =   81
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   59
      Left            =   3465
      TabIndex        =   80
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   58
      Left            =   2835
      TabIndex        =   79
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   57
      Left            =   2100
      TabIndex        =   78
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   56
      Left            =   1470
      TabIndex        =   77
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   55
      Left            =   840
      TabIndex        =   76
      Top             =   4935
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   54
      Left            =   6090
      TabIndex        =   75
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   53
      Left            =   5460
      TabIndex        =   74
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   52
      Left            =   4830
      TabIndex        =   73
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   51
      Left            =   4095
      TabIndex        =   72
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   50
      Left            =   3465
      TabIndex        =   71
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   49
      Left            =   2835
      TabIndex        =   70
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   48
      Left            =   2100
      TabIndex        =   69
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   47
      Left            =   1470
      TabIndex        =   68
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   46
      Left            =   840
      TabIndex        =   67
      Top             =   4200
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   45
      Left            =   6090
      TabIndex        =   66
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   44
      Left            =   5460
      TabIndex        =   65
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   43
      Left            =   4830
      TabIndex        =   64
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   42
      Left            =   4095
      TabIndex        =   63
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   41
      Left            =   3465
      TabIndex        =   62
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   40
      Left            =   2835
      TabIndex        =   61
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   39
      Left            =   2100
      TabIndex        =   60
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   38
      Left            =   1470
      TabIndex        =   59
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   37
      Left            =   840
      TabIndex        =   58
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   36
      Left            =   6090
      TabIndex        =   57
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   35
      Left            =   5460
      TabIndex        =   56
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   34
      Left            =   4830
      TabIndex        =   55
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   33
      Left            =   4095
      TabIndex        =   54
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   32
      Left            =   3465
      TabIndex        =   53
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   31
      Left            =   2835
      TabIndex        =   52
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   30
      Left            =   2100
      TabIndex        =   51
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   29
      Left            =   1470
      TabIndex        =   50
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   28
      Left            =   840
      TabIndex        =   49
      Top             =   2940
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   27
      Left            =   6090
      TabIndex        =   48
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   26
      Left            =   5460
      TabIndex        =   47
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   25
      Left            =   4830
      TabIndex        =   46
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   24
      Left            =   4095
      TabIndex        =   45
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   23
      Left            =   3465
      TabIndex        =   44
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   22
      Left            =   2835
      TabIndex        =   43
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   21
      Left            =   2100
      TabIndex        =   42
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   20
      Left            =   1470
      TabIndex        =   41
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   19
      Left            =   840
      TabIndex        =   40
      Top             =   2205
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   18
      Left            =   6090
      TabIndex        =   39
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   17
      Left            =   5460
      TabIndex        =   38
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   16
      Left            =   4830
      TabIndex        =   37
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   15
      Left            =   4095
      TabIndex        =   36
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   14
      Left            =   3465
      TabIndex        =   35
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   13
      Left            =   2835
      TabIndex        =   34
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   12
      Left            =   2100
      TabIndex        =   33
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   11
      Left            =   1470
      TabIndex        =   32
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   10
      Left            =   840
      TabIndex        =   31
      Top             =   1575
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   9
      Left            =   6090
      TabIndex        =   30
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   8
      Left            =   5460
      TabIndex        =   29
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   7
      Left            =   4830
      TabIndex        =   28
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   6
      Left            =   4095
      TabIndex        =   27
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   5
      Left            =   3465
      TabIndex        =   26
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   4
      Left            =   2835
      TabIndex        =   25
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   3
      Left            =   2100
      TabIndex        =   24
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   81
      Left            =   6090
      TabIndex        =   102
      Top             =   6195
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   2
      Left            =   1470
      TabIndex        =   23
      Top             =   945
      Width           =   510
   End
   Begin VB.Label LabelSudokuBlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   1
      Left            =   840
      TabIndex        =   22
      Top             =   945
      Width           =   510
   End
   Begin VB.Shape ShapeBottombar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   330
      Left            =   0
      Top             =   7245
      Width           =   13455
   End
   Begin VB.Menu Menu 
      Caption         =   "&Menu"
      Begin VB.Menu MenuSoundSwitch 
         Caption         =   "S&ound"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenuAnimationSwitch 
         Caption         =   "A&nimation"
         Checked         =   -1  'True
         Shortcut        =   {F3}
      End
      Begin VB.Menu Menu1_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuEXIT 
         Caption         =   "E&XIT"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Menu2_ 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "&About"
         Begin VB.Menu MenuAboutDownload 
            Caption         =   "&Download latest version..."
         End
         Begin VB.Menu MenuAboutUpdate 
            Caption         =   "Check for &update manually..."
         End
         Begin VB.Menu MenuAboutGitHub 
            Caption         =   "&GitHub @SamToki user page..."
         End
         Begin VB.Menu MenuAboutLicense 
            Caption         =   "Released under &license GNU GPL v3..."
         End
         Begin VB.Menu MenuAboutCopyright 
            Caption         =   "TM && (C) 2015-2022 SAM TOKI STUDIO"
            Enabled         =   0   'False
         End
         Begin VB.Menu MenuAboutDate 
            Caption         =   "2022/07/26"
            Enabled         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "FormMainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === INFORMATION ===
'
'  SAM TOKI STUDIO
'  This is a .frm source code file.
'
'  Random Sudoku Generator
'
'  Powered by Sam Toki
'  Version: v0.30eng
'  Date:    2022/07/25 (Mon)
'  History: First version v0.10 was built on 2020/03/28.
'
'  WARNING: Commercial use of this computer software is strictly prohibited.
'           Open source license:      GNU GPL v3
'           Creative Commons license: CC BY-NC 3.0
'
'  Copyright: TM & (C) 2015-2022 SAM TOKI STUDIO. All rights reserved.
'             SAM TOKI STUDIO is a trademark of CZJ Software Technologies (CZJST) Inc. in the P.R.C and other countries.
'
'  ---------------------------------------------------------------------------------------------------------------------
'
'  === NOTES FOR REFERENCE ===
'
'  ...
'
'  ---------------------------------------------------------------------------------------------------------------------

'[] DECLARATIONS []

Option Explicit

'Declare Menu...
Public setsoundswitch As Boolean
Public setanimationswitch As Boolean

'Declare Game...
Public gamestatus As Integer  '0-Initial, 1-Generating, 2-Ongoing.
Public gamegenerationstep As Integer
Public gameinputstep As Integer  '1-Row, 2-Column, 3-Number.

'Declare Sudoku...
Public sudokublockdata As Variant  '(1 To 9)(1 To 9)
Public sudokublockstatus As Variant  '(1 To 9)(1 To 9)  0-Filling, 1-Fixed, 2-FillingAgainstRule, 3-FixedAgainstRule.

Public sudokutotalfixed As Integer
Public sudokutotalfilling As Integer
Public sudokutotalfilled As Integer
Public sudokufillednumbercount As Variant  '(1 To 9)

Public sudokucurrentrow As Integer
Public sudokucurrentcolumn As Integer
Public sudokupreviousrow As Integer
Public sudokupreviouscolumn As Integer

'Declare Lottery...
Public lotterytotal As Integer
Public lotterynumber As Integer

'Declare Settings...
Public settotalfixed As Integer
Public setlargeblockmaximumfixed As Integer
Public setuselegacygenerationmethod As Boolean

'Declare Dialog...
Public answer
Public dontshowagain1 As Boolean

'Declare Animation...
Public progressbaranimationtarget As Integer  'Range: 0~13455
Public labelrowindicatoranimationtarget As Integer  'Range: 364~6270
Public labelcolumnindicatoranimationtarget As Integer  'Range: 289~6195
Public sudokucurrentrowanimation As Integer
Public sudokucurrentcolumnanimation As Integer

'Declare Temporary Variants...
Public forloop1 As Integer
Public forloop2 As Integer
Public forloop3 As Integer
Public forloop4 As Integer
Public forloop5 As Integer
Public tempvariant As Integer
Public preventinfiniteloopcounter1 As Integer
Public preventinfiniteloopcounter2 As Integer

'Declare Special...
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
     ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5

'  ---------------------------------------------------------------------------------------------------------------------

'[] LOAD []

    Public Sub Form_Load()
        setsoundswitch = True
        setanimationswitch = True
        setuselegacygenerationmethod = False
        dontshowagain1 = False

        Call Initialization
    End Sub
    Public Sub Initialization()
        'Initialize Game...
        gamestatus = 0
        gamegenerationstep = 0
        gameinputstep = 0

        'Initialize Sudoku...
        sudokublockdata = Array(Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                )
        sudokublockstatus = Array(Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                                  Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0) _
                                  )

        sudokutotalfixed = 0
        sudokutotalfilling = 81 - sudokutotalfixed
        sudokutotalfilled = 0
        settotalfixed = HScrollSettingsTotalFixedAmount.Value
        setlargeblockmaximumfixed = HScrollSettingsLargeBlockMaximumFixedAmount.Value
        sudokufillednumbercount = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        sudokucurrentrow = 0
        sudokucurrentcolumn = 0
        sudokupreviousrow = 0
        sudokupreviouscolumn = 0

        'Initialize Lottery...
        lotterytotal = 0
        lotterynumber = 0

        'Initialize Animation...
        progressbaranimationtarget = 0
        labelrowindicatoranimationtarget = 364
        labelcolumnindicatoranimationtarget = 289
        sudokucurrentrowanimation = 0
        sudokucurrentcolumnanimation = 0
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] COMMANDS []

    'CMD Menu...
    Public Sub MenuSoundSwitch_Click()
        Select Case setsoundswitch
            Case False
                setsoundswitch = True
                MenuSoundSwitch.Checked = True
            Case True
                setsoundswitch = False
                MenuSoundSwitch.Checked = False
        End Select
    End Sub
    Public Sub MenuAnimationSwitch_Click()
        Select Case setanimationswitch
            Case False
                setanimationswitch = True
                TimerSudokuBlockAnimation.Interval = 20
                MenuAnimationSwitch.Checked = True
            Case True
                setanimationswitch = False
                TimerSudokuBlockAnimation.Interval = 1
                MenuAnimationSwitch.Checked = False
        End Select
    End Sub
    Public Sub MenuEXIT_Click()
        End
    End Sub
    Public Sub MenuAboutDownload_Click()
        Call ShellExecute(Me.hWnd, "open", "https://github.com/SamToki/VB6---Random-Sudoku-Generator/raw/main/RELEASE/Random%20Sudoku%20Generator.exe", "", "", SW_SHOW)
    End Sub
    Public Sub MenuAboutUpdate_Click()
        Call ShellExecute(Me.hWnd, "open", "https://github.com/SamToki/VB6---Random-Sudoku-Generator", "", "", SW_SHOW)
    End Sub
    Public Sub MenuAboutGitHub_Click()
        Call ShellExecute(Me.hWnd, "open", "https://github.com/SamToki", "", "", SW_SHOW)
    End Sub
    Public Sub MenuAboutLicense_Click()
        Call ShellExecute(Me.hWnd, "open", "https://www.gnu.org/licenses/gpl-3.0.html", "", "", SW_SHOW)
    End Sub

    'CMD Settings...
    Public Sub HScrollSettingsTotalFixedAmount_Change()
        settotalfixed = HScrollSettingsTotalFixedAmount.Value
        LabelSettingsTotalFixedAmountIndicator.Caption = settotalfixed
    End Sub
    Public Sub HScrollSettingsTotalFixedAmount_Scroll()
        Call HScrollSettingsTotalFixedAmount_Change
    End Sub
    Public Sub HScrollSettingsLargeBlockMaximumFixedAmount_Change()
        setlargeblockmaximumfixed = HScrollSettingsLargeBlockMaximumFixedAmount.Value
        LabelSettingsLargeBlockMaximumFixedAmountIndicator.Caption = setlargeblockmaximumfixed
        HScrollSettingsTotalFixedAmount.Max = setlargeblockmaximumfixed * 9
    End Sub
    Public Sub HScrollSettingsLargeBlockMaximumFixedAmount_Scroll()
        Call HScrollSettingsLargeBlockMaximumFixedAmount_Change
    End Sub
    Public Sub CheckboxSlowDownGeneration_Click()
        Select Case CheckboxSlowDownGeneration.Value
            Case 0
                TimerGenerator.Interval = 10
            Case 1
                TimerGenerator.Interval = 500
        End Select
    End Sub
    Public Sub CheckboxUseLegacyGenerationMethod_Click()
        Select Case CheckboxUseLegacyGenerationMethod.Value
            Case 0
                setuselegacygenerationmethod = False
            Case 1
                setuselegacygenerationmethod = True
        End Select
    End Sub

    'CMD Controls...
    Public Sub CmdStartReset_Click()
        If gamestatus = 0 Then
            If setsoundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Proximity Connection.wav"
            Call Initialization: gamestatus = 1: Call Refresher
            LabelStatusbar.Caption = "Game started!"
            LabelStartTime.Caption = "Generation started at " & LabelClock.Caption
            CmdStartReset.Caption = "RE&SET"
        Else
            CmdStartReset.SetFocus
            If setsoundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Recycle.wav"
            Call Initialization: gamestatus = 0: Call Refresher
            LabelStatusbar.Caption = "Game reset! Now you can adjust settings."
            LabelStartTime.Caption = "Game not started yet"
            CmdStartReset.Caption = "&START"
        End If
    End Sub
    Public Sub CmdNumber_Click(index As Integer)
        TextboxInput.Text = index
    End Sub
    Public Sub LabelSudokuBlock_Click(index As Integer)
        If gamestatus = 1 Then
            MsgBox "Sudoku generation in progress. You cannot select the blocks during generation.", vbInformation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
            Exit Sub
        End If

        sudokucurrentrow = -Int(-index / 9)
        If (index Mod 9 = 0) Then
            sudokucurrentcolumn = 9
        Else
            sudokucurrentcolumn = index Mod 9
        End If

        If setsoundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Menu Command.wav"
        gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1
    End Sub

    Public Sub TextboxInput_Change()
        If Not (gamestatus = 2) Then Exit Sub
        TextboxInput.SetFocus
        'If the change is to clear the textbox, then do nothing...
        If TextboxInput.Text = "" Then Exit Sub

        If setsoundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Startup.wav"

        Select Case gameinputstep
            Case 1
                Select Case TextboxInput.Text
                    Case "1"
                        sudokucurrentrow = 1
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "2"
                        sudokucurrentrow = 2
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "3"
                        sudokucurrentrow = 3
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "4"
                        sudokucurrentrow = 4
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "5"
                        sudokucurrentrow = 5
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "6"
                        sudokucurrentrow = 6
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "7"
                        sudokucurrentrow = 7
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "8"
                        sudokucurrentrow = 8
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "9"
                        sudokucurrentrow = 9
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 2: sudokucurrentrowanimation = 1: Exit Sub
                    Case "W", "w"
                        If sudokucurrentrow > 1 Then sudokucurrentrow = sudokucurrentrow - 1 Else sudokucurrentrow = 1
                        If sudokucurrentcolumn = 0 Then sudokucurrentcolumn = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "S", "s"
                        If sudokucurrentrow < 9 Then sudokucurrentrow = sudokucurrentrow + 1 Else sudokucurrentrow = 9
                        If sudokucurrentcolumn = 0 Then sudokucurrentcolumn = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "A", "a"
                        If sudokucurrentcolumn > 1 Then sudokucurrentcolumn = sudokucurrentcolumn - 1 Else sudokucurrentcolumn = 1
                        If sudokucurrentrow = 0 Then sudokucurrentrow = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "D", "d"
                        If sudokucurrentcolumn < 9 Then sudokucurrentcolumn = sudokucurrentcolumn + 1 Else sudokucurrentcolumn = 9
                        If sudokucurrentrow = 0 Then sudokucurrentrow = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case ""
                        TextboxInput.Text = "": gameinputstep = 1: Exit Sub
                    Case Else
                        MsgBox "Invalid input." & vbCrLf & vbCrLf & "You have pressed a wrong key. Please ensure that your fingers are on the right keys. Acceptable keys are from ""1"" to ""9"", and ""WASD"" for selection change.", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                        TextboxInput.Text = "": gameinputstep = 1: Exit Sub
                End Select
            Case 2
                Select Case TextboxInput.Text
                    Case "1"
                        sudokucurrentcolumn = 1
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "2"
                        sudokucurrentcolumn = 2
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "3"
                        sudokucurrentcolumn = 3
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "4"
                        sudokucurrentcolumn = 4
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "5"
                        sudokucurrentcolumn = 5
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "6"
                        sudokucurrentcolumn = 6
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "7"
                        sudokucurrentcolumn = 7
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "8"
                        sudokucurrentcolumn = 8
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "9"
                        sudokucurrentcolumn = 9
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                        TextboxInput.Text = "": gameinputstep = 3: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "W", "w"
                        If sudokucurrentrow > 1 Then sudokucurrentrow = sudokucurrentrow - 1 Else sudokucurrentrow = 1
                        If sudokucurrentcolumn = 0 Then sudokucurrentcolumn = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "S", "s"
                        If sudokucurrentrow < 9 Then sudokucurrentrow = sudokucurrentrow + 1 Else sudokucurrentrow = 9
                        If sudokucurrentcolumn = 0 Then sudokucurrentcolumn = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "A", "a"
                        If sudokucurrentcolumn > 1 Then sudokucurrentcolumn = sudokucurrentcolumn - 1 Else sudokucurrentcolumn = 1
                        If sudokucurrentrow = 0 Then sudokucurrentrow = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "D", "d"
                        If sudokucurrentcolumn < 9 Then sudokucurrentcolumn = sudokucurrentcolumn + 1 Else sudokucurrentcolumn = 9
                        If sudokucurrentrow = 0 Then sudokucurrentrow = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case ""
                        TextboxInput.Text = "": gameinputstep = 2: Exit Sub
                    Case Else
                        MsgBox "Invalid input." & vbCrLf & vbCrLf & "You have pressed a wrong key. Please ensure that your fingers are on the right keys. Acceptable keys are from ""1"" to ""9"", and ""WASD"" for selection change.", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                        TextboxInput.Text = "": gameinputstep = 2: Exit Sub
                End Select
            Case 3
                Select Case TextboxInput.Text
                    Case "W", "w"
                        If sudokucurrentrow > 1 Then sudokucurrentrow = sudokucurrentrow - 1 Else sudokucurrentrow = 1
                        If sudokucurrentcolumn = 0 Then sudokucurrentcolumn = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "S", "s"
                        If sudokucurrentrow < 9 Then sudokucurrentrow = sudokucurrentrow + 1 Else sudokucurrentrow = 9
                        If sudokucurrentcolumn = 0 Then sudokucurrentcolumn = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "A", "a"
                        If sudokucurrentcolumn > 1 Then sudokucurrentcolumn = sudokucurrentcolumn - 1 Else sudokucurrentcolumn = 1
                        If sudokucurrentrow = 0 Then sudokucurrentrow = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                    Case "D", "d"
                        If sudokucurrentcolumn < 9 Then sudokucurrentcolumn = sudokucurrentcolumn + 1 Else sudokucurrentcolumn = 9
                        If sudokucurrentrow = 0 Then sudokucurrentrow = 1
                        TextboxInput.Text = "": gameinputstep = 3: Call Refresher: sudokucurrentrowanimation = 1: sudokucurrentcolumnanimation = 1: Exit Sub
                End Select
                If (((sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 1) Or (sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 3)) And (dontshowagain1 = False)) Then
                    answer = MsgBox("You are trying to change the number in a fixed block." & vbCrLf & vbCrLf & "It is acceptable only when fixing an unsolvable randomly generated Sudoku." & vbCrLf & vbCrLf & "Do you want to continue?" & vbCrLf & "[Yes] Change this number" & vbCrLf & "[No] Leave it unchanged" & vbCrLf & "[Cancel] Change it, and don't show again", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Random Sudoku Generator")
                    If answer = vbNo Then
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    End If
                    If answer = vbCancel Then dontshowagain1 = True
                End If
                Select Case TextboxInput.Text
                    Case "1"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 1
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "2"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 2
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "3"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 3
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "4"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 4
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "5"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 5
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "6"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 6
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "7"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 7
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "8"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 8
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "9"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 9
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case "0"
                        sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 0
                        TextboxInput.Text = "": gameinputstep = 1: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case ""
                        TextboxInput.Text = "": gameinputstep = 3: Call SudokuRuleChecker: Call Refresher: Exit Sub
                    Case Else
                        MsgBox "Invalid input." & vbCrLf & vbCrLf & "You have pressed a wrong key. Please ensure that your fingers are on the right keys. Acceptable keys are from ""1"" to ""9"", and ""0"" for ""Clear"", ""WASD"" for selection change, ""Enter"" for ""Reset Game"".", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                        TextboxInput.Text = "": gameinputstep = 3: Call SudokuRuleChecker: Call Refresher: Exit Sub
                End Select
        End Select
    End Sub

'[] TIMERS []

    Public Sub TimerClock_Timer()
        LabelClock.Caption = Format((Hour(Time)), "00") & ":" & Format((Minute(Time)), "00") & ":" & Format((Second(Time)), "00")
    End Sub

    Public Sub TimerGenerator_Timer()
        If Not (gamestatus = 1) Then Exit Sub

        Select Case setuselegacygenerationmethod
            Case False
                Call SudokuGenerator
            Case True
                Call SudokuGeneratorLegacy
        End Select
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ENGINE []

    'Refresh statistics, display, and commands...
    Public Sub Refresher()
        'Statistics...
        Select Case gamestatus
            Case 0
                sudokutotalfixed = 0
                sudokutotalfilling = 81 - sudokutotalfixed
                sudokutotalfilled = 0
                LabelStatusbar.Caption = "Ready"
            Case 1
                tempvariant = 0
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If (((sudokublockstatus(forloop1)(forloop2) = 1) Or (sudokublockstatus(forloop1)(forloop2) = 3)) And (Not (sudokublockdata(forloop1)(forloop2) = 0))) Then tempvariant = tempvariant + 1
                    Next
                Next
                sudokutotalfixed = tempvariant
                sudokutotalfilling = 81 - sudokutotalfixed
                tempvariant = 0
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If (Not ((sudokublockstatus(forloop1)(forloop2) = 1) Or (sudokublockstatus(forloop1)(forloop2) = 3))) And (Not (sudokublockdata(forloop1)(forloop2) = 0)) Then tempvariant = tempvariant + 1
                    Next
                Next
                sudokutotalfilled = tempvariant
            Case 2
                sudokutotalfixed = settotalfixed
                sudokutotalfilling = 81 - sudokutotalfixed
                tempvariant = 0
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If (Not ((sudokublockstatus(forloop1)(forloop2) = 1) Or (sudokublockstatus(forloop1)(forloop2) = 3))) And (Not (sudokublockdata(forloop1)(forloop2) = 0)) Then tempvariant = tempvariant + 1
                    Next
                Next
                sudokutotalfilled = tempvariant
                Select Case gameinputstep
                    Case 0
                        Exit Sub
                    Case 1
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter row number. Enter ""WASD"" for selection change"
                    Case 2
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please enter column number. Enter ""WASD"" for selection change"
                    Case 3
                        LabelStatusbar.Caption = "Filled " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Please fill the block. Enter ""0"" for ""Clear"", ""WASD"" for selection change"
                    Case Else
                        MsgBox "ERROR: Game input step is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                End Select

                'Judge Sudoku solved...
                tempvariant = 0
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If ((sudokublockstatus(forloop1)(forloop2) = 2) Or (sudokublockstatus(forloop1)(forloop2) = 3)) Then tempvariant = 444
                    Next
                Next
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If sudokublockdata(forloop1)(forloop2) = 0 Then tempvariant = 888
                    Next
                Next
                If ((sudokutotalfilled >= sudokutotalfilling) And (tempvariant = 0) And (gameinputstep = 1)) Then MsgBox "Congrats! You have solved the Sudoku.", vbInformation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select

        'Row and column number...
        For forloop1 = 0 To 9: LabelRow(forloop1).ForeColor = &HFFFFFF: Next
        If Not (sudokucurrentrow = 0) Then LabelRow(sudokucurrentrow).ForeColor = &H0&
        For forloop1 = 0 To 9: LabelColumn(forloop1).ForeColor = &HFFFFFF: Next
        If Not (sudokucurrentcolumn = 0) Then LabelColumn(sudokucurrentcolumn).ForeColor = &H0&

        'Sudoku blocks...
        For forloop1 = 1 To 81
            If (forloop1 Mod 9 = 0) Then
                tempvariant = 9
            Else
                tempvariant = forloop1 Mod 9
            End If
            Select Case sudokublockstatus(-Int(-forloop1 / 9))(tempvariant)
                Case 0
                    LabelSudokuBlock(forloop1).BackColor = &HFFFFFF: LabelSudokuBlock(forloop1).ForeColor = &H0&
                Case 1
                    LabelSudokuBlock(forloop1).BackColor = &HC0C0C0: LabelSudokuBlock(forloop1).ForeColor = &H0&
                Case 2
                    LabelSudokuBlock(forloop1).BackColor = &HFFFFFF: LabelSudokuBlock(forloop1).ForeColor = &HFF&
                Case 3
                    LabelSudokuBlock(forloop1).BackColor = &HC0C0C0: LabelSudokuBlock(forloop1).ForeColor = &HFF&
                Case Else
                    MsgBox "ERROR: Sudoku block fixed-or-not data is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
            End Select
            LabelSudokuBlock(forloop1).BorderStyle = 0
        Next
        If Not (sudokucurrentrow = 0 And sudokucurrentcolumn = 0) Then
            LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BorderStyle = 1
            LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BackColor = &HFFE0C0
        End If

        For forloop1 = 1 To 81
            If (forloop1 Mod 9 = 0) Then
                tempvariant = 9
            Else
                tempvariant = forloop1 Mod 9
            End If
            LabelSudokuBlock(forloop1).Caption = sudokublockdata(-Int(-forloop1 / 9))(tempvariant)
            If LabelSudokuBlock(forloop1).Caption = "0" Then LabelSudokuBlock(forloop1).Caption = ""
        Next

        'Number count...
        For forloop1 = 1 To 9
            tempvariant = 0

            For forloop2 = 1 To 9
                For forloop3 = 1 To 9
                    If sudokublockdata(forloop2)(forloop3) = forloop1 Then tempvariant = tempvariant + 1
                Next
            Next

            sudokufillednumbercount(forloop1) = tempvariant
            LabelNumberCount(forloop1).Caption = tempvariant
            If tempvariant > 9 Then
                LabelNumberCount(forloop1).Caption = "X"
            End If
            If tempvariant >= 9 Then
                If tempvariant = 9 Then
                    LabelNumberCount(forloop1).ForeColor = &HC000&
                Else
                    LabelNumberCount(forloop1).ForeColor = &HFF&
                End If
            Else
                LabelNumberCount(forloop1).ForeColor = &H0&
            End If
        Next

        'Game status light bulb indicator...
        Select Case gamestatus
            Case 0
                ShapeLightGameStatusIndicator.BackColor = &HC0C0C0
            Case 1
                ShapeLightGameStatusIndicator.BackColor = &HFFA000
            Case 2
                ShapeLightGameStatusIndicator.BackColor = &HE000&
            Case Else
                MsgBox "ERROR: Game status is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select

        'Progressbar...
        Select Case gamestatus
            Case 0
                ShapeProgressbar.BackColor = &HC0C0C0
                progressbaranimationtarget = 0
            Case 1
                ShapeProgressbar.BackColor = &HFFA000
                If setuselegacygenerationmethod = False Then
                    Select Case gamegenerationstep
                        Case 0 To 4
                            progressbaranimationtarget = (sudokutotalfilled / sudokutotalfilling) * 13455
                        Case 5 To 9
                            progressbaranimationtarget = (sudokutotalfixed / settotalfixed) * 13455
                        Case Else
                            MsgBox "ERROR: Game generation step is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                    End Select
                Else
                    progressbaranimationtarget = (sudokutotalfixed / settotalfixed) * 13455
                End If
            Case 2
                ShapeProgressbar.BackColor = &HE000&
                progressbaranimationtarget = (sudokutotalfilled / sudokutotalfilling) * 13455
            Case Else
                MsgBox "ERROR: Game status is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select

        'Commands...
        Select Case gamestatus
            Case 0
                HScrollSettingsTotalFixedAmount.Enabled = True: HScrollSettingsLargeBlockMaximumFixedAmount.Enabled = True: CheckboxUseLegacyGenerationMethod.Enabled = True
                TextboxInput.Enabled = False
                For forloop1 = 0 To 9: CmdNumber(forloop1).Enabled = False: Next
            Case 1
                HScrollSettingsTotalFixedAmount.Enabled = False: HScrollSettingsLargeBlockMaximumFixedAmount.Enabled = False: CheckboxUseLegacyGenerationMethod.Enabled = False
                TextboxInput.Enabled = False
                For forloop1 = 0 To 9: CmdNumber(forloop1).Enabled = False: Next
            Case 2
                HScrollSettingsTotalFixedAmount.Enabled = False: HScrollSettingsLargeBlockMaximumFixedAmount.Enabled = False: CheckboxUseLegacyGenerationMethod.Enabled = False
                TextboxInput.Enabled = True: TextboxInput.SetFocus
                For forloop1 = 0 To 9: CmdNumber(forloop1).Enabled = True: Next
            Case Else
                MsgBox "ERROR: Game status is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select
    End Sub

    'Check Sudoku...
    Public Sub LargeBlockMaximumFixedChecker()
        tempvariant = 0

        Select Case sudokucurrentrow
            Case 1 To 3
                Select Case sudokucurrentcolumn
                    Case 1 To 3
                        For forloop1 = 1 To 3
                            For forloop2 = 1 To 3
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 4 To 6
                        For forloop1 = 1 To 3
                            For forloop2 = 4 To 6
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 7 To 9
                        For forloop1 = 1 To 3
                            For forloop2 = 7 To 9
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                End Select
            Case 4 To 6
                Select Case sudokucurrentcolumn
                    Case 1 To 3
                        For forloop1 = 4 To 6
                            For forloop2 = 1 To 3
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 4 To 6
                        For forloop1 = 4 To 6
                            For forloop2 = 4 To 6
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 7 To 9
                        For forloop1 = 4 To 6
                            For forloop2 = 7 To 9
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                End Select
            Case 7 To 9
                Select Case sudokucurrentcolumn
                    Case 1 To 3
                        For forloop1 = 7 To 9
                            For forloop2 = 1 To 3
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 4 To 6
                        For forloop1 = 7 To 9
                            For forloop2 = 4 To 6
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                    Case 7 To 9
                        For forloop1 = 7 To 9
                            For forloop2 = 7 To 9
                                If sudokublockstatus(forloop1)(forloop2) = 1 Then tempvariant = tempvariant + 1
                            Next
                        Next
                        If tempvariant >= setlargeblockmaximumfixed Then
                            tempvariant = 444: Exit Sub
                        End If
                End Select
        End Select
    End Sub

    Public Sub SudokuRuleChecker()
        'Initialize sudokublockstatus...
        For forloop1 = 1 To 9
            For forloop2 = 1 To 9
                If sudokublockstatus(forloop1)(forloop2) = 2 Then sudokublockstatus(forloop1)(forloop2) = 0
                If sudokublockstatus(forloop1)(forloop2) = 3 Then sudokublockstatus(forloop1)(forloop2) = 1
            Next
        Next

        'Locate block...
        For forloop1 = 1 To 9
            For forloop2 = 1 To 9

                'Check horizontal numbers...
                For forloop3 = 1 To 9
                    If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop1)(forloop3)) And (Not (forloop2 = forloop3)) Then
                        Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                    End If
                Next

                'Check vertical numbers...
                For forloop3 = 1 To 9
                    If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop2)) And (Not (forloop1 = forloop3)) Then
                        Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                    End If
                Next

                'Check large block...
                Select Case forloop1
                    Case 1 To 3
                        Select Case forloop2
                            Case 1 To 3
                                For forloop3 = 1 To 3
                                    For forloop4 = 1 To 3
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 4 To 6
                                For forloop3 = 1 To 3
                                    For forloop4 = 4 To 6
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 7 To 9
                                For forloop3 = 1 To 3
                                    For forloop4 = 7 To 9
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                        End Select
                    Case 4 To 6
                        Select Case forloop2
                            Case 1 To 3
                                For forloop3 = 4 To 6
                                    For forloop4 = 1 To 3
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 4 To 6
                                For forloop3 = 4 To 6
                                    For forloop4 = 4 To 6
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 7 To 9
                                For forloop3 = 4 To 6
                                    For forloop4 = 7 To 9
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                        End Select
                    Case 7 To 9
                        Select Case forloop2
                            Case 1 To 3
                                For forloop3 = 7 To 9
                                    For forloop4 = 1 To 3
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 4 To 6
                                For forloop3 = 7 To 9
                                    For forloop4 = 4 To 6
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                            Case 7 To 9
                                For forloop3 = 7 To 9
                                    For forloop4 = 7 To 9
                                        If (sudokublockdata(forloop1)(forloop2) = sudokublockdata(forloop3)(forloop4)) And (Not ((forloop1 = forloop3) And (forloop2 = forloop4))) Then
                                            Select Case sudokublockstatus(forloop1)(forloop2): Case 0: sudokublockstatus(forloop1)(forloop2) = 2: Case 1: sudokublockstatus(forloop1)(forloop2) = 3: End Select
                                        End If
                                    Next
                                Next
                        End Select
                End Select

            Next
        Next
    End Sub

    'Generate Sudoku...
    Public Sub RandomNumberGenerator()
        If lotterytotal = 0 Then
            MsgBox "ERROR: Calling function ""RandomNumberGenerator"" when variant ""lotterytotal"" is 0." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End If

        lotterynumber = 0
        While lotterynumber = 0
            Randomize
            lotterynumber = Int((lotterytotal + 1) * Rnd)
        Wend
    End Sub

    Public Sub SudokuGenerator()
        If Not (gamestatus = 1) Then Exit Sub

        Select Case gamegenerationstep
            Case 0
                Call Refresher
                gamegenerationstep = 1
                LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Initialized"

                'Prevent infinite loop...
                preventinfiniteloopcounter1 = 0
            Case 1
                'Toleration of infinite loop...
                If preventinfiniteloopcounter1 > 20 Then
                    'Clear previous row...
                    answer = MsgBox("The generation has stuck." & vbCrLf & vbCrLf & "The generator will clear the previous row and try again. If this keeps popping up, please click [Yes] to abort the generation." & vbCrLf & vbCrLf & "Do you want to abort the generation?", vbQuestion + vbYesNo + vbDefaultButton2, "Random Sudoku Generator")
                    If answer = vbYes Then
                        gamegenerationstep = 444
                        LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Generation aborted"
                        Exit Sub
                    Else
                        'Move back...
                        sudokucurrentrow = sudokucurrentrow - 1
                        sudokucurrentcolumn = 0
                        'Partly clear...
                        For forloop1 = sudokucurrentrow To 9
                            For forloop2 = 1 To 9
                                sudokublockdata(forloop1)(forloop2) = 0
                                sudokublockstatus(forloop1)(forloop2) = 0
                            Next
                        Next
                        gamegenerationstep = 0
                        LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Cleared previous row"
                        Exit Sub
                    End If
                End If

                'Locate next block...
                If sudokucurrentrow = 0 Then sudokucurrentrow = 1
                sudokucurrentcolumn = sudokucurrentcolumn + 1
                If sudokucurrentcolumn > 9 Then
                    sudokucurrentrow = sudokucurrentrow + 1
                    sudokucurrentcolumn = 0
                    gamegenerationstep = 0
                    LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Located next row"
                    Exit Sub
                End If
                Call Refresher
                gamegenerationstep = 2
                LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Located next block"

                'Prevent infinite loop...
                preventinfiniteloopcounter2 = 0
            Case 2
                'Toleration of infinite loop...
                If preventinfiniteloopcounter2 > 30 Then
                    'Move back...
                    sudokucurrentcolumn = 0
                    'Partly clear...
                    For forloop1 = sudokucurrentrow To 9
                        For forloop2 = 1 To 9
                            sudokublockdata(forloop1)(forloop2) = 0
                            sudokublockstatus(forloop1)(forloop2) = 0
                        Next
                    Next
                    preventinfiniteloopcounter1 = preventinfiniteloopcounter1 + 1
                    gamegenerationstep = 1
                    LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Cleared current row"
                    Exit Sub
                End If

                'Fill but do not fix the block...
                lotterytotal = 9: lotterynumber = 0
                Do
                    Call RandomNumberGenerator
                    sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = lotterynumber
                    sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 0
                    Call Refresher
                Loop Until sudokufillednumbercount(lotterynumber) <= 9
                gamegenerationstep = 3
                LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Filled with a random number"
            Case 3
                'Check integrity...
                Call SudokuRuleChecker
                Call Refresher
                If sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 2 Or sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 3 Then
                    preventinfiniteloopcounter2 = preventinfiniteloopcounter2 + 1
                    gamegenerationstep = 2
                    LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Filled with a random number; Number invalid"
                Else
                    gamegenerationstep = 4
                    LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Filled with a random number; Number valid"
                End If
            Case 4
                'Done...
                Call Refresher
                sudokupreviousrow = sudokucurrentrow: sudokupreviouscolumn = sudokucurrentcolumn
                gamegenerationstep = 1
                LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Done"

                'Finish generation...
                If sudokutotalfilled = sudokutotalfilling Then
                    gamegenerationstep = 5
                    LabelStatusbar.Caption = "Generating " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Generation completed"
                End If
            Case 5
                'Locate random block...
                lotterytotal = 9: lotterynumber = 0
                Do
                    Call RandomNumberGenerator: sudokucurrentrow = lotterynumber
                    Call RandomNumberGenerator: sudokucurrentcolumn = lotterynumber
                Loop Until sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 0
                Call Refresher
                gamegenerationstep = 6
                LabelStatusbar.Caption = "Fixing " & sudokutotalfixed & "/" & settotalfixed & " --- Found an unfixed block"
            Case 6
                'Check integrity...
                Call Refresher
                Call LargeBlockMaximumFixedChecker
                If tempvariant = 444 Then
                    gamegenerationstep = 5
                    LabelStatusbar.Caption = "Fixing " & sudokutotalfixed & "/" & settotalfixed & " --- Found an unfixed block; Location invalid"
                Else
                    gamegenerationstep = 7
                    LabelStatusbar.Caption = "Fixing " & sudokutotalfixed & "/" & settotalfixed & " --- Found an unfixed block; Location valid"
                End If
            Case 7
                'Fix the block...
                sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 1
                Call Refresher
                gamegenerationstep = 5
                LabelStatusbar.Caption = "Fixing " & sudokutotalfixed & "/" & settotalfixed & " --- Found an unfixed block; Fixed"

                'Finish fix...
                If sudokutotalfixed = settotalfixed Then
                    gamegenerationstep = 8
                    LabelStatusbar.Caption = "Fixing " & sudokutotalfilled & "/" & sudokutotalfilling & " --- Fix completed"
                End If
            Case 8
                'Clear unfixed blocks...
                For forloop1 = 1 To 9
                    For forloop2 = 1 To 9
                        If sudokublockstatus(forloop1)(forloop2) = 0 Then sudokublockdata(forloop1)(forloop2) = 0
                    Next
                Next
                Call Refresher
                gamegenerationstep = 9
                LabelStatusbar.Caption = "Cleared unfixed blocks"
            Case 9
                'Finish all...
                If setsoundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Print Complete.wav"
                LabelStartTime.Caption = "Game started at " & LabelClock.Caption
                gamestatus = 2: gamegenerationstep = 0: gameinputstep = 1: sudokucurrentrow = 0: sudokucurrentcolumn = 0: Call Refresher
            Case 444
                'Abort generation...
                MsgBox "Unable to continue generating this Sudoku.", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                Call CmdStartReset_Click
            Case Else
                MsgBox "ERROR: Game generation step is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select
    End Sub

    Public Sub SudokuGeneratorLegacy()
        If Not (gamestatus = 1) Then Exit Sub

        Select Case gamegenerationstep
            Case 0
                Call Refresher
                gamegenerationstep = 1
                LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Initialized"

                'Prevent infinite loop...
                preventinfiniteloopcounter1 = 0
            Case 1
                'Toleration of infinite loop...
                If preventinfiniteloopcounter1 > 20 Then
                    If sudokupreviousrow = 0 And sudokupreviouscolumn = 0 Then
                        gamegenerationstep = 444
                        LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Generation aborted"
                        Exit Sub
                    Else
                        'Clear previous block...
                        answer = MsgBox("The generation has stuck." & vbCrLf & vbCrLf & "The generator will clear the previous block and try again. If still stuck, the generation will be aborted. If this keeps popping up, please click [Yes] to abort manually." & vbCrLf & vbCrLf & "Do you want to abort the generation?", vbQuestion + vbYesNo + vbDefaultButton2, "Random Sudoku Generator")
                        If answer = vbYes Then
                            gamegenerationstep = 444
                            LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Generation aborted"
                            Exit Sub
                        Else
                            sudokucurrentrow = sudokupreviousrow: sudokucurrentcolumn = sudokupreviouscolumn
                            sudokupreviousrow = 0: sudokupreviouscolumn = 0
                            sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 0
                            sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 0
                            gamegenerationstep = 0
                            LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Cleared previous block"
                            Exit Sub
                        End If
                    End If
                End If

                'Locate random block...
                lotterytotal = 9: lotterynumber = 0
                Do
                    Call RandomNumberGenerator: sudokucurrentrow = lotterynumber
                    Call RandomNumberGenerator: sudokucurrentcolumn = lotterynumber
                Loop Until sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 0
                Call Refresher
                gamegenerationstep = 2
                LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Found an empty block"
            Case 2
                'Check integrity...
                Call Refresher
                Call LargeBlockMaximumFixedChecker
                If tempvariant = 444 Then
                    gamegenerationstep = 1
                    LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Found an empty block; Location invalid"
                Else
                    gamegenerationstep = 3
                    LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Found an empty block; Location valid"
                End If

                'Prevent infinite loop...
                preventinfiniteloopcounter2 = 0
            Case 3
                'Toleration of infinite loop...
                If preventinfiniteloopcounter2 > 30 Then
                    sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = 0
                    sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 0
                    preventinfiniteloopcounter1 = preventinfiniteloopcounter1 + 1
                    gamegenerationstep = 1
                    LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Cleared current block"
                    Exit Sub
                End If

                'Fill and fix the block...
                lotterytotal = 9: lotterynumber = 0
                Do
                    Call RandomNumberGenerator
                    sudokublockdata(sudokucurrentrow)(sudokucurrentcolumn) = lotterynumber
                    sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 1
                    Call Refresher
                Loop Until sudokufillednumbercount(lotterynumber) <= 9
                gamegenerationstep = 4
                LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Filled with a random number"
            Case 4
                'Check integrity...
                Call SudokuRuleChecker
                Call Refresher
                If sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 2 Or sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn) = 3 Then
                    preventinfiniteloopcounter2 = preventinfiniteloopcounter2 + 1
                    gamegenerationstep = 3
                    LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Filled with a random number; Number invalid"
                Else
                    gamegenerationstep = 5
                    LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Filled with a random number; Number valid"
                End If
            Case 5
                'Done...
                Call Refresher
                sudokupreviousrow = sudokucurrentrow: sudokupreviouscolumn = sudokucurrentcolumn
                gamegenerationstep = 0
                LabelStatusbar.Caption = "Generating " & sudokutotalfixed & "/" & settotalfixed & " --- Done"

                'Finish generation...
                If sudokutotalfixed = settotalfixed Then
                    If setsoundswitch = True Then WindowsMediaPlayer1.URL = "C:\Windows\Media\Windows Print Complete.wav"
                    LabelStartTime.Caption = "Game started at " & LabelClock.Caption
                    gamestatus = 2: gamegenerationstep = 0: gameinputstep = 1: sudokucurrentrow = 0: sudokucurrentcolumn = 0: Call Refresher
                End If
            Case 444
                'Abort generation...
                MsgBox "Unable to continue generating this Sudoku." & vbCrLf & vbCrLf & "Please reduce the amount of fixed blocks to make it easier to generate, or uncheck ""Use legacy generation method"".", vbExclamation + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
                Call CmdStartReset_Click
            Case Else
                MsgBox "ERROR: Game generation step is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select
    End Sub

'  ---------------------------------------------------------------------------------------------------------------------

'[] ANIMATIONS []

    Public Sub TimerProgressbarAnimation_Timer()
        labelrowindicatoranimationtarget = 364 + (sudokucurrentrow / 9) * (6270 - 364)
        labelcolumnindicatoranimationtarget = 289 + (sudokucurrentcolumn / 9) * (6195 - 289)

        Select Case setanimationswitch
            Case False
                ShapeProgressbar.Width = progressbaranimationtarget
                LabelRowIndicator.Top = labelrowindicatoranimationtarget
                LabelColumnIndicator.Left = labelcolumnindicatoranimationtarget
            Case True
                If ShapeProgressbar.Width = progressbaranimationtarget Then GoTo TimerProgressbarAnimation_Skip1_
                If ShapeProgressbar.Width > progressbaranimationtarget Then ShapeProgressbar.Width = ShapeProgressbar.Width - Abs(ShapeProgressbar.Width - progressbaranimationtarget) / 4
                If ShapeProgressbar.Width < progressbaranimationtarget Then ShapeProgressbar.Width = ShapeProgressbar.Width + Abs(ShapeProgressbar.Width - progressbaranimationtarget) / 4
                If Abs(ShapeProgressbar.Width - progressbaranimationtarget) < 10 Then ShapeProgressbar.Width = progressbaranimationtarget
TimerProgressbarAnimation_Skip1_:
                If LabelRowIndicator.Top = labelrowindicatoranimationtarget Then GoTo TimerProgressbarAnimation_Skip2_
                If LabelRowIndicator.Top > labelrowindicatoranimationtarget Then LabelRowIndicator.Top = LabelRowIndicator.Top - Abs(LabelRowIndicator.Top - labelrowindicatoranimationtarget) / 4
                If LabelRowIndicator.Top < labelrowindicatoranimationtarget Then LabelRowIndicator.Top = LabelRowIndicator.Top + Abs(LabelRowIndicator.Top - labelrowindicatoranimationtarget) / 4
                If Abs(LabelRowIndicator.Top - labelrowindicatoranimationtarget) < 10 Then LabelRowIndicator.Top = labelrowindicatoranimationtarget
TimerProgressbarAnimation_Skip2_:
                If LabelColumnIndicator.Left = labelcolumnindicatoranimationtarget Then GoTo TimerProgressbarAnimation_Skip3_
                If LabelColumnIndicator.Left > labelcolumnindicatoranimationtarget Then LabelColumnIndicator.Left = LabelColumnIndicator.Left - Abs(LabelColumnIndicator.Left - labelcolumnindicatoranimationtarget) / 4
                If LabelColumnIndicator.Left < labelcolumnindicatoranimationtarget Then LabelColumnIndicator.Left = LabelColumnIndicator.Left + Abs(LabelColumnIndicator.Left - labelcolumnindicatoranimationtarget) / 4
                If Abs(LabelColumnIndicator.Left - labelcolumnindicatoranimationtarget) < 10 Then LabelColumnIndicator.Left = labelcolumnindicatoranimationtarget
TimerProgressbarAnimation_Skip3_:
        End Select
    End Sub

    Public Sub TimerSudokuBlockAnimation_Timer()
        'Highlight current row and column with light yellow color, animated. And highlight the row and column number...
        Select Case gameinputstep
            Case 0
                Exit Sub
            Case 1
                Exit Sub
            Case 2
                If sudokublockstatus(sudokucurrentrow)(sudokucurrentrowanimation) = 0 Or sudokublockstatus(sudokucurrentrow)(sudokucurrentrowanimation) = 2 Then
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentrowanimation).BackColor = &HC0FFFF
                Else
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentrowanimation).BackColor = &H80C0C0
                End If
                If sudokucurrentrowanimation < 9 Then sudokucurrentrowanimation = sudokucurrentrowanimation + 1
                LabelRow(sudokucurrentrow).ForeColor = &H0&
            Case 3
                If sudokublockstatus(sudokucurrentrow)(sudokucurrentrowanimation) = 0 Or sudokublockstatus(sudokucurrentrow)(sudokucurrentrowanimation) = 2 Then
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentrowanimation).BackColor = &HC0FFFF
                Else
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentrowanimation).BackColor = &H80C0C0
                End If
                If sudokucurrentrowanimation < 9 Then sudokucurrentrowanimation = sudokucurrentrowanimation + 1
                LabelRow(sudokucurrentrow).ForeColor = &H0&

                If sudokublockstatus(sudokucurrentcolumnanimation)(sudokucurrentcolumn) = 0 Or sudokublockstatus(sudokucurrentcolumnanimation)(sudokucurrentcolumn) = 2 Then
                    LabelSudokuBlock(sudokucurrentcolumnanimation * 9 - 9 + sudokucurrentcolumn).BackColor = &HC0FFFF
                Else
                    LabelSudokuBlock(sudokucurrentcolumnanimation * 9 - 9 + sudokucurrentcolumn).BackColor = &H80C0C0
                End If
                If sudokucurrentcolumnanimation < 9 Then sudokucurrentcolumnanimation = sudokucurrentcolumnanimation + 1
                LabelColumn(sudokucurrentcolumn).ForeColor = &H0&
            Case Else
                MsgBox "ERROR: Game input step is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
        End Select

        'Highlight current block with light green or red color...
        If gameinputstep = 3 Then
            'Show block border...
            LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BorderStyle = 1

            'Change backcolor...
            Select Case sudokublockstatus(sudokucurrentrow)(sudokucurrentcolumn)
                Case 0, 2
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BackColor = &HC0FFC0
                Case 1, 3
                    LabelSudokuBlock(sudokucurrentrow * 9 - 9 + sudokucurrentcolumn).BackColor = &HC0C0FF
                Case Else
                    MsgBox "ERROR: Sudoku block fixed-or-not data is out of range." & vbCrLf & vbCrLf & "Please send a feedback to @SamToki via GitHub so as to help solve the problem.", vbCritical + vbOKOnly + vbDefaultButton1, "Random Sudoku Generator"
            End Select
        End If
    End Sub
