VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form formMain 
   Caption         =   "RGB-HSL-Gamma Worker"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "formMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerResize 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6240
      Top             =   3720
   End
   Begin VB.Timer tmrEyeDropper 
      Left            =   6240
      Top             =   3240
   End
   Begin VB.Frame frameAbout 
      Caption         =   "By rpeterclark"
      Height          =   735
      Left            =   8280
      TabIndex        =   70
      Top             =   6120
      Width           =   1335
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   69
      Top             =   6975
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8969
            Picture         =   "formMain.frx":151A
            Key             =   "FILENAME"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Picture         =   "formMain.frx":1792
            Key             =   "IMAGETYPE"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Picture         =   "formMain.frx":1A12
            Key             =   "BITDEPTH"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Picture         =   "formMain.frx":1B80
            Key             =   "DIMENSIONS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameCode 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   60
      Top             =   6120
      Width           =   8055
      Begin VB.CommandButton cmdCopyCode 
         Caption         =   "Copy C&ode"
         Height          =   375
         Left            =   6720
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "<gammagroup id="""" />"
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame framePreview 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   59
      Top             =   2880
      Width           =   8055
      Begin VB.PictureBox picPreviewOptions 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   6720
         ScaleHeight     =   2775
         ScaleWidth      =   1215
         TabIndex        =   63
         Top             =   240
         Width           =   1215
         Begin VB.CommandButton cmdLoadPNG 
            Caption         =   "&Load PNG"
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1215
         End
         Begin VB.Frame frameOptions 
            Caption         =   "Options"
            Height          =   1335
            Left            =   0
            TabIndex        =   64
            Top             =   1440
            Width           =   1215
            Begin VB.OptionButton opGrey 
               Caption         =   "Gray Off"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton opGrey 
               Caption         =   "Gray 1"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   17
               Top             =   480
               Width           =   975
            End
            Begin VB.OptionButton opGrey 
               Caption         =   "Gray 2"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   18
               Top             =   720
               Width           =   975
            End
            Begin VB.CheckBox chkBoost 
               Caption         =   "Boost"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   960
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdPNGUpdate 
            Caption         =   "&Update"
            Default         =   -1  'True
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox chkAutoUpdate 
            Caption         =   "Auto Update"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.PictureBox picBack 
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2715
         ScaleWidth      =   6435
         TabIndex        =   61
         Top             =   240
         Width           =   6495
         Begin VB.PictureBox picDraw 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2775
            Left            =   0
            ScaleHeight     =   185
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   433
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   0
            Width           =   6495
            Begin VB.Image imgPreviewTile 
               Appearance      =   0  'Flat
               Height          =   120
               Left            =   5760
               Picture         =   "formMain.frx":1BFF
               Top             =   120
               Visible         =   0   'False
               Width           =   120
            End
         End
      End
   End
   Begin VB.Frame frameSwatch 
      Caption         =   "Swatch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   8280
      TabIndex        =   57
      Top             =   2880
      Width           =   1335
      Begin VB.CommandButton cmdEyeDropper 
         Caption         =   "Eye &Dropper"
         Height          =   375
         Left            =   120
         TabIndex        =   72
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdColorPicker 
         Caption         =   "Color &Picker"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.PictureBox picSwatch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1755
         ScaleWidth      =   1035
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frameGamma 
      Caption         =   "Gamma"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   8280
      TabIndex        =   25
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdCopyGamma 
         Caption         =   "&Copy"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Blue"
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   1095
         Begin VB.TextBox txtGammaBlue 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Green"
         Height          =   615
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   1095
         Begin VB.TextBox txtGammaGreen 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Red"
         Height          =   615
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1095
         Begin VB.TextBox txtGammaRed 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   9
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
      End
   End
   Begin VB.Frame frameHSL 
      Caption         =   "HSL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4200
      TabIndex        =   24
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton cmdCopyHSL 
         Caption         =   "Copy &HSL"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   3735
      End
      Begin VB.Frame frameLuminance 
         Caption         =   "Luminance"
         Height          =   615
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   3735
         Begin VB.TextBox txtLuminance 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
         Begin VB.PictureBox picLuminance 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1320
            ScaleHeight     =   180
            ScaleMode       =   0  'User
            ScaleWidth      =   2205
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   400
            Width           =   2265
            Begin VB.PictureBox picLuminancePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   1
               Left            =   1132
               ScaleHeight     =   135
               ScaleWidth      =   1125
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   0
               Width           =   1132
            End
            Begin VB.PictureBox picLuminancePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   0
               Left            =   0
               ScaleHeight     =   135
               ScaleWidth      =   1125
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   0
               Width           =   1132
            End
         End
         Begin MSComctlLib.Slider sliderLuminance 
            Height          =   255
            Left            =   1220
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   200
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   450
            _Version        =   393216
            Max             =   240
            TickStyle       =   3
         End
         Begin MSComCtl2.UpDown UpDownLuminance 
            Height          =   285
            Left            =   960
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtLuminance"
            BuddyDispid     =   196640
            OrigLeft        =   1200
            OrigTop         =   240
            OrigRight       =   1455
            OrigBottom      =   495
            Increment       =   15
            Max             =   239
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame frameSaturation 
         Caption         =   "Saturation"
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   3735
         Begin VB.PictureBox picSaturation 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1320
            ScaleHeight     =   180
            ScaleMode       =   0  'User
            ScaleWidth      =   2205
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   400
            Width           =   2265
         End
         Begin VB.TextBox txtSaturation 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   6
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
         Begin MSComctlLib.Slider sliderSaturation 
            Height          =   255
            Left            =   1220
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   200
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   450
            _Version        =   393216
            Max             =   240
            TickStyle       =   3
         End
         Begin MSComCtl2.UpDown UpDownSaturation 
            Height          =   285
            Left            =   960
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtSaturation"
            BuddyDispid     =   196645
            OrigLeft        =   1200
            OrigTop         =   240
            OrigRight       =   1455
            OrigBottom      =   495
            Increment       =   15
            Max             =   239
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame frameHue 
         Caption         =   "Hue"
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   3735
         Begin VB.TextBox txtHue 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   5
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
         Begin VB.PictureBox picHue 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1320
            ScaleHeight     =   180
            ScaleMode       =   0  'User
            ScaleWidth      =   2205
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   400
            Width           =   2265
            Begin VB.PictureBox picHuePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   5
               Left            =   1880
               ScaleHeight     =   135
               ScaleWidth      =   375
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   0
               Width           =   377
            End
            Begin VB.PictureBox picHuePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   4
               Left            =   1500
               ScaleHeight     =   135
               ScaleWidth      =   375
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   0
               Width           =   377
            End
            Begin VB.PictureBox picHuePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   3
               Left            =   1131
               ScaleHeight     =   135
               ScaleWidth      =   375
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   0
               Width           =   377
            End
            Begin VB.PictureBox picHuePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   2
               Left            =   754
               ScaleHeight     =   135
               ScaleWidth      =   375
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   0
               Width           =   377
            End
            Begin VB.PictureBox picHuePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   1
               Left            =   377
               ScaleHeight     =   135
               ScaleWidth      =   375
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   0
               Width           =   377
            End
            Begin VB.PictureBox picHuePiece 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   135
               Index           =   0
               Left            =   0
               ScaleHeight     =   135
               ScaleWidth      =   375
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   0
               Width           =   377
            End
         End
         Begin MSComctlLib.Slider sliderHue 
            Height          =   255
            Left            =   1220
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   200
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   450
            _Version        =   393216
            Max             =   239
            TickStyle       =   3
         End
         Begin MSComCtl2.UpDown UpDownHue 
            Height          =   285
            Left            =   960
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtHue"
            BuddyDispid     =   196647
            OrigLeft        =   1200
            OrigTop         =   240
            OrigRight       =   1455
            OrigBottom      =   495
            Increment       =   15
            Max             =   239
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
   End
   Begin VB.Frame frameRGB 
      Caption         =   "RGB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   3975
      Begin VB.Frame frameBlue 
         Caption         =   "Blue"
         Height          =   615
         Left            =   120
         TabIndex        =   65
         Top             =   1440
         Width           =   3735
         Begin VB.PictureBox picBlue 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1320
            ScaleHeight     =   180
            ScaleMode       =   0  'User
            ScaleWidth      =   2205
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   400
            Width           =   2265
         End
         Begin VB.TextBox txtBlue 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   2
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
         Begin MSComctlLib.Slider sliderBlue 
            Height          =   255
            Left            =   1220
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   200
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   450
            _Version        =   393216
            Max             =   255
            TickStyle       =   3
         End
         Begin MSComCtl2.UpDown UpDownBlue 
            Height          =   285
            Left            =   960
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtBlue"
            BuddyDispid     =   196653
            OrigLeft        =   1200
            OrigTop         =   240
            OrigRight       =   1455
            OrigBottom      =   495
            Increment       =   15
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdCopyRGB 
         Caption         =   "Copy &RGB"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Frame frameGreen 
         Caption         =   "Green"
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   3735
         Begin VB.TextBox txtGreen 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
         Begin VB.PictureBox picGreen 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1320
            ScaleHeight     =   180
            ScaleMode       =   0  'User
            ScaleWidth      =   2205
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   400
            Width           =   2265
         End
         Begin MSComctlLib.Slider sliderGreen 
            Height          =   255
            Left            =   1220
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   200
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   450
            _Version        =   393216
            Max             =   255
            TickStyle       =   3
         End
         Begin MSComCtl2.UpDown UpDownGreen 
            Height          =   285
            Left            =   960
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtGreen"
            BuddyDispid     =   196657
            OrigLeft        =   1200
            OrigTop         =   240
            OrigRight       =   1455
            OrigBottom      =   495
            Increment       =   15
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame frameRed 
         Caption         =   "Red"
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   3735
         Begin VB.PictureBox picRed 
            AutoRedraw      =   -1  'True
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1320
            ScaleHeight     =   180
            ScaleMode       =   0  'User
            ScaleWidth      =   2205
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   400
            Width           =   2265
         End
         Begin MSComctlLib.Slider sliderRed 
            Height          =   255
            Left            =   1215
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   195
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   450
            _Version        =   393216
            Max             =   255
            TickStyle       =   3
         End
         Begin MSComCtl2.UpDown UpDownRed 
            Height          =   285
            Left            =   961
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRed"
            BuddyDispid     =   196661
            OrigLeft        =   1200
            OrigTop         =   240
            OrigRight       =   1455
            OrigBottom      =   495
            Increment       =   15
            Max             =   255
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtRed 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "0"
            Top             =   240
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PNG As New clsPNG
Private FileDlg As New clsFileDialog
Private Tiler As New clsTiler

Private Type PointAPI
    x As Long
    y As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()
    Dim lngRGBLastSelected As Long
    Dim strRGBLastSelected As String
    Dim intRed As Integer
    Dim intGreen As Integer
    Dim intBlue As Integer
    
    Me.Width = Val(VBA.GetSetting(App.EXEName, "Dimensions", "LastWidth"))
    Me.Height = Val(VBA.GetSetting(App.EXEName, "Dimensions", "LastHeight"))
    Me.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    strRGBLastSelected = VBA.GetSetting(App.EXEName, "Colors", "LastSelectedColor")
    If strRGBLastSelected <> "" Then
        lngRGBLastSelected = CLng(strRGBLastSelected)
    Else
        lngRGBLastSelected = RGB(128, 128, 128)
    End If
    
    intRed = RGBRed(lngRGBLastSelected)
    intGreen = RGBGreen(lngRGBLastSelected)
    intBlue = RGBBlue(lngRGBLastSelected)
    
    opGrey.Item(0).Value = True
    picDraw.Visible = False
    
    Tiler.SetBackdrop imgPreviewTile.Picture
    Tiler.Attach picBack.hwnd
            
    Call InitHueScale
    Call SyncAll(intRed, intGreen, intBlue)
    Call SyncAllScales(intRed, intGreen, intBlue)
End Sub

Private Sub Form_Resize()
    Dim intFormMainMinWidth As Integer
    Dim intFormMainMinHeight As Integer
    Dim intBorderBufferTwipsX As Integer
    Dim intBorderBufferTwipsY As Integer
    Dim intInnerBufferTwipsX As Integer
    Dim intInnerBufferTwipsY As Integer
    Dim intInnerFrameBufferTwipsX As Integer
    Dim intInnerFrameBufferTwipsY As Integer
    
    timerResize.Enabled = False
    
    intFormMainMinWidth = 9870
    intFormMainMinHeight = 7800
    
    intBorderBufferTwipsX = 255
    intBorderBufferTwipsY = 570
    
    intInnerBufferTwipsX = 110
    intInnerBufferTwipsY = 105
    
    intInnerFrameBufferTwipsX = 120
    intInnerFrameBufferTwipsY = 120
    
    If Me.WindowState <> 1 Then
        If Me.Width < intFormMainMinWidth Then Me.Width = intFormMainMinWidth
        If Me.Height < intFormMainMinHeight Then Me.Height = intFormMainMinHeight
        
        'Adjust Height
        frameCode.tOp = Me.Height - (frameCode.Height + sbMain.Height + intBorderBufferTwipsY)
        frameAbout.tOp = Me.Height - (frameAbout.Height + sbMain.Height + intBorderBufferTwipsY)
        
        frameSwatch.Height = Me.Height - frameSwatch.tOp - intInnerBufferTwipsY - frameAbout.Height - sbMain.Height - intBorderBufferTwipsY
        cmdEyeDropper.tOp = frameSwatch.Height - (cmdEyeDropper.Height + intInnerFrameBufferTwipsY)
        cmdColorPicker.tOp = cmdEyeDropper.tOp - (cmdColorPicker.Height + intInnerFrameBufferTwipsY)
        picSwatch.Height = frameSwatch.Height - picSwatch.tOp - (cmdColorPicker.Height + cmdEyeDropper.Height + (intInnerFrameBufferTwipsY * 3))
        
        framePreview.Height = Me.Height - framePreview.tOp - intInnerBufferTwipsY - frameCode.Height - sbMain.Height - intBorderBufferTwipsY
        picBack.Height = framePreview.Height - picBack.tOp - intInnerFrameBufferTwipsY
        picDraw.Height = picBack.Height
        
        'Adjust Width
        frameGamma.left = Me.Width - (frameGamma.Width + intBorderBufferTwipsX)
        frameSwatch.left = Me.Width - (frameSwatch.Width + intBorderBufferTwipsX)
        frameAbout.left = Me.Width - (frameAbout.Width + intBorderBufferTwipsX)
        
        frameCode.Width = Me.Width - (frameAbout.Width + (intInnerBufferTwipsX * 2) + intBorderBufferTwipsX)
        txtCode.Width = frameCode.Width - (txtCode.left + intInnerFrameBufferTwipsX + cmdCopyCode.Width + intInnerFrameBufferTwipsX)
        cmdCopyCode.left = txtCode.left + txtCode.Width + intInnerFrameBufferTwipsX
        
        framePreview.Width = Me.Width - (frameSwatch.Width + (intInnerBufferTwipsX * 2) + intBorderBufferTwipsX)
        picBack.Width = framePreview.Width - (picBack.left + picPreviewOptions.Width + (intInnerFrameBufferTwipsX * 2))
        picDraw.Width = picBack.Width
        picPreviewOptions.left = picBack.left + picBack.Width + intInnerFrameBufferTwipsX
        
        frameRGB.Width = ((Me.Width - (frameRGB.left + intBorderBufferTwipsX + frameGamma.Width + intInnerBufferTwipsX)) / 2) - (intInnerBufferTwipsX / 2)
        frameRed.Width = frameRGB.Width - (frameRed.left + intInnerFrameBufferTwipsX)
        frameGreen.Width = frameRGB.Width - (frameGreen.left + intInnerFrameBufferTwipsX)
        frameBlue.Width = frameRGB.Width - (frameBlue.left + intInnerFrameBufferTwipsX)
        sliderRed.Width = frameRed.Width - (sliderRed.left + (intInnerFrameBufferTwipsX - 60))
        sliderGreen.Width = frameGreen.Width - (sliderGreen.left + (intInnerFrameBufferTwipsX - 60))
        sliderBlue.Width = frameBlue.Width - (sliderBlue.left + (intInnerFrameBufferTwipsX - 60))
        picRed.Width = frameRed.Width - (picRed.left + intInnerFrameBufferTwipsX)
        picGreen.Width = frameGreen.Width - (picGreen.left + intInnerFrameBufferTwipsX)
        picBlue.Width = frameBlue.Width - (picBlue.left + intInnerFrameBufferTwipsX)
        cmdCopyRGB.Width = frameRGB.Width - (cmdCopyRGB.left + intInnerFrameBufferTwipsX + cmdReset.Width + intInnerFrameBufferTwipsX)
        cmdReset.left = cmdCopyRGB.left + cmdCopyRGB.Width + intInnerFrameBufferTwipsX
        
        frameHSL.left = frameRGB.left + frameRGB.Width + intInnerBufferTwipsX
        frameHSL.Width = Me.Width - (frameHSL.left + intBorderBufferTwipsX + frameGamma.Width + intInnerBufferTwipsX)
        frameHue.Width = frameHSL.Width - (frameHue.left + intInnerFrameBufferTwipsX)
        frameSaturation.Width = frameHSL.Width - (frameSaturation.left + intInnerFrameBufferTwipsX)
        frameLuminance.Width = frameHSL.Width - (frameLuminance.left + intInnerFrameBufferTwipsX)
        sliderHue.Width = frameHue.Width - (sliderHue.left + (intInnerFrameBufferTwipsX - 60))
        sliderSaturation.Width = frameSaturation.Width - (sliderSaturation.left + (intInnerFrameBufferTwipsX - 60))
        sliderLuminance.Width = frameLuminance.Width - (sliderLuminance.left + (intInnerFrameBufferTwipsX - 60))
        picHue.Width = frameHue.Width - (picHue.left + intInnerFrameBufferTwipsX)
        picSaturation.Width = frameSaturation.Width - (picSaturation.left + intInnerFrameBufferTwipsX)
        picLuminance.Width = frameLuminance.Width - (picLuminance.left + intInnerFrameBufferTwipsX)
        
        picHuePiece(0).Width = (picHue.Width / 6) + 15
        picHuePiece(1).Width = (picHue.Width / 6) + 15
        picHuePiece(2).Width = (picHue.Width / 6) + 15
        picHuePiece(3).Width = (picHue.Width / 6) + 15
        picHuePiece(4).Width = (picHue.Width / 6) + 15
        picHuePiece(5).Width = (picHue.Width / 6) + 15
        
        picHuePiece(1).left = (picHue.Width / 6)
        picHuePiece(2).left = ((picHue.Width / 6) * 2)
        picHuePiece(3).left = ((picHue.Width / 6) * 3)
        picHuePiece(4).left = ((picHue.Width / 6) * 4)
        picHuePiece(5).left = ((picHue.Width / 6) * 5)
        
        picLuminancePiece(0).Width = (picLuminance.Width / 2)
        picLuminancePiece(1).Width = (picLuminance.Width / 2)
        picLuminancePiece(1).left = (picLuminance.Width / 2)
        cmdCopyHSL.Width = frameHSL.Width - (cmdCopyHSL.left + intInnerBufferTwipsX)
        
        timerResize.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VBA.SaveSetting App.EXEName, "Colors", "LastSelectedColor", CStr(RGB(txtRed.Text, txtGreen.Text, txtBlue.Text))
    VBA.SaveSetting App.EXEName, "Dimensions", "LastWidth", Me.Width
    VBA.SaveSetting App.EXEName, "Dimensions", "LastHeight", Me.Height
End Sub

Private Sub InitHueScale()
    Dim objGradient As New clsGradient
    
    With objGradient
        .Color1 = RGB(255, 0, 0)
        .Color2 = RGB(255, 255, 0)
        .Angle = 0
        .Draw picHuePiece(0)
        
        .Color1 = RGB(255, 255, 0)
        .Color2 = RGB(0, 255, 0)
        .Angle = 0
        .Draw picHuePiece(1)
        
        .Color1 = RGB(0, 255, 0)
        .Color2 = RGB(0, 255, 255)
        .Angle = 0
        .Draw picHuePiece(2)
        
        .Color1 = RGB(0, 255, 255)
        .Color2 = RGB(0, 0, 255)
        .Angle = 0
        .Draw picHuePiece(3)
        
        .Color1 = RGB(0, 0, 255)
        .Color2 = RGB(255, 0, 255)
        .Angle = 0
        .Draw picHuePiece(4)
        
        .Color1 = RGB(255, 0, 255)
        .Color2 = RGB(255, 0, 0)
        .Angle = 0
        .Draw picHuePiece(5)
    End With
    
    picHuePiece(0).Refresh
    picHuePiece(1).Refresh
    picHuePiece(2).Refresh
    picHuePiece(3).Refresh
    picHuePiece(4).Refresh
    picHuePiece(5).Refresh
End Sub

Private Sub SyncAll(ByVal intRed, ByVal intGreen, ByVal intBlue)
    Call SyncRGB(intRed, intGreen, intBlue)
    Call SyncHLS(intRed, intGreen, intBlue)
    Call SyncGamma(intRed, intGreen, intBlue)
    Call SyncSwatch(intRed, intGreen, intBlue)
    Call SyncPreview(False)
    Call SyncCode
End Sub

Private Sub SyncSB()
    If PNG.CurrentFilename <> "" Then
        sbMain.Panels("FILENAME").Text = CompactedPathSh(PNG.CurrentFilename, Me.ScaleX((sbMain.Panels("FILENAME").Width - sbMain.Panels("FILENAME").Picture.Width), vbTwips, vbPixels), Me.hDC)
        sbMain.Panels("IMAGETYPE").Text = Choose(PNG.ColorType + 1, "Grayscale", , "RGB", "Palette", "Grayscale+Alpha", , "RGB+Alpha")
        sbMain.Panels("BITDEPTH").Text = "Bit Depth: " & PNG.BitDepth
        sbMain.Panels("DIMENSIONS").Text = PNG.Width & "x" & PNG.Height & " px"
    End If
End Sub

Private Sub SyncAllScales(ByVal intRed, ByVal intGreen, ByVal intBlue)
    Dim objGradient As New clsGradient
    Dim lngRGBSaturationLow As Long
    Dim lngRGBSaturationHigh As Long
    Dim lngRGBLuminanceLow As Long
    Dim lngRGBLuminanceMid As Long
    Dim lngRGBLuminanceHigh As Long
    Dim intHue As Integer
    Dim intSaturation As Integer
    Dim intLuminance As Integer
    
    Call ColorRGBToHLS(RGB(intRed, intGreen, intBlue), intHue, intLuminance, intSaturation)
    lngRGBSaturationLow = ColorHLSToRGB(intHue, intLuminance, 1)
    lngRGBSaturationHigh = ColorHLSToRGB(intHue, intLuminance, 240)
    lngRGBLuminanceLow = ColorHLSToRGB(intHue, 1, intSaturation)
    lngRGBLuminanceMid = ColorHLSToRGB(intHue, 120, intSaturation)
    lngRGBLuminanceHigh = ColorHLSToRGB(intHue, 240, intSaturation)
    
    With objGradient
        .Color1 = RGB(0, intGreen, intBlue)
        .Color2 = RGB(255, intGreen, intBlue)
        .Angle = 0
        .Draw picRed

        .Color1 = RGB(intRed, 0, intBlue)
        .Color2 = RGB(intRed, 255, intBlue)
        .Angle = 0
        .Draw picGreen
        
        .Color1 = RGB(intRed, intGreen, 0)
        .Color2 = RGB(intRed, intGreen, 255)
        .Angle = 0
        .Draw picBlue
        
        .Color1 = RGB(RGBRed(lngRGBSaturationLow), RGBGreen(lngRGBSaturationLow), RGBBlue(lngRGBSaturationLow))
        .Color2 = RGB(RGBRed(lngRGBSaturationHigh), RGBGreen(lngRGBSaturationHigh), RGBBlue(lngRGBSaturationHigh))
        .Angle = 0
        .Draw picSaturation
        
        .Color1 = RGB(RGBRed(lngRGBLuminanceLow), RGBGreen(lngRGBLuminanceLow), RGBBlue(lngRGBLuminanceLow))
        .Color2 = RGB(RGBRed(lngRGBLuminanceMid), RGBGreen(lngRGBLuminanceMid), RGBBlue(lngRGBLuminanceMid))
        .Angle = 0
        .Draw picLuminancePiece(0)
        
        .Color1 = RGB(RGBRed(lngRGBLuminanceMid), RGBGreen(lngRGBLuminanceMid), RGBBlue(lngRGBLuminanceMid))
        .Color2 = RGB(RGBRed(lngRGBLuminanceHigh), RGBGreen(lngRGBLuminanceHigh), RGBBlue(lngRGBLuminanceHigh))
        .Angle = 0
        .Draw picLuminancePiece(1)
    End With
    
    picRed.Refresh
    picGreen.Refresh
    picBlue.Refresh
    picSaturation.Refresh
    picLuminancePiece(0).Refresh
    picLuminancePiece(1).Refresh

    sliderRed.Value = intRed
    sliderGreen.Value = intGreen
    sliderBlue.Value = intBlue
    sliderSaturation.Value = intSaturation
    sliderLuminance.Value = intLuminance
End Sub

Private Sub SyncRGB(ByVal intRed, ByVal intGreen, ByVal intBlue)
    txtRed.Text = intRed
    txtBlue.Text = intBlue
    txtGreen.Text = intGreen
End Sub

Private Sub SyncHLS(ByVal intRed, ByVal intGreen, ByVal intBlue)
    Dim intHue As Integer
    Dim intSaturation As Integer
    Dim intLuminance As Integer
    
    Call ColorRGBToHLS(RGB(intRed, intGreen, intBlue), intHue, intLuminance, intSaturation)
    
    txtHue.Text = intHue
    txtLuminance.Text = intLuminance
    txtSaturation.Text = intSaturation
End Sub

Private Sub SyncGamma(ByVal intRed, ByVal intGreen, ByVal intBlue)
    Dim intGammaRed As Integer
    Dim intGammaGreen As Integer
    Dim intGammaBlue As Integer
    
    Call ColorRGBToGamma(RGB(intRed, intGreen, intBlue), intGammaRed, intGammaGreen, intGammaBlue)
    txtGammaRed.Text = intGammaRed
    txtGammaGreen.Text = intGammaGreen
    txtGammaBlue.Text = intGammaBlue
End Sub

Private Sub SyncSwatch(ByVal intRed, ByVal intGreen, ByVal intBlue)
    picSwatch.BackColor = RGB(intRed, intGreen, intBlue)
End Sub

Private Sub SyncPreview(boolForce As Boolean)
    If (boolForce Or chkAutoUpdate) And PNG.CurrentFilename <> "" Then
        picDraw.Cls
        PNG.GammaEnabled = True
        PNG.GammaRed = txtGammaRed.Text
        PNG.GammaGreen = txtGammaGreen.Text
        PNG.GammaBlue = txtGammaBlue.Text
        If opGrey.Item(1) Then
            PNG.Gray = 1
        ElseIf opGrey.Item(2) Then
            PNG.Gray = 2
        Else
            PNG.Gray = 0
        End If
        If chkBoost Then
            PNG.Boost = 1
        Else
            PNG.Boost = 0
        End If
        
        Tiler.TileArea picBack.hDC, 0, 0, picBack.ScaleWidth \ Screen.TwipsPerPixelX, picBack.ScaleHeight \ Screen.TwipsPerPixelY
        
        PNG.DrawToDC picDraw.hDC, picDraw.ScaleWidth \ 2 - PNG.Width \ 2, picDraw.ScaleHeight \ 2 - PNG.Height \ 2
        BitBlt picDraw.hDC, 0, 0, picDraw.ScaleWidth, picDraw.ScaleHeight, picBack.hDC, 0, 0, vbSrcCopy
        PNG.DrawToDC picDraw.hDC, picDraw.ScaleWidth \ 2 - PNG.Width \ 2, picDraw.ScaleHeight \ 2 - PNG.Height \ 2
        
        picDraw.Refresh
    End If
End Sub

Private Sub SyncCode()
    Dim strCode As String
    strCode = "<gammagroup id="""" value="""" gray="""" boost="""" />"
    strCode = Replace(strCode, "value=""""", "value=""" & txtGammaRed.Text & ", " & txtGammaGreen.Text & ", " & txtGammaBlue.Text & """")
    If opGrey.Item(1) Then
        strCode = Replace(strCode, "gray=""""", "gray=""1""")
    ElseIf opGrey.Item(2) Then
        strCode = Replace(strCode, "gray=""""", "gray=""2""")
    Else
        strCode = Replace(strCode, "gray=""""", "gray=""0""")
    End If
    If chkBoost Then
        strCode = Replace(strCode, "boost=""""", "boost=""1""")
    Else
        strCode = Replace(strCode, "boost=""""", "boost=""0""")
    End If
    txtCode.Text = strCode
End Sub

Private Sub opGrey_Click(Index As Integer)
    Call SyncPreview(False)
    Call SyncCode
End Sub

Private Sub picSwatch_Click()
    Call ShowColorPicker
End Sub

Private Sub picSwatch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngReturn As Long
    If tmrEyeDropper.Interval > 0 Then
        lngReturn = ReleaseCapture
        tmrEyeDropper.Interval = 0
    End If
End Sub

Private Sub sliderRed_Scroll()
    txtRed.Text = sliderRed.Value
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub sliderGreen_Scroll()
    txtGreen.Text = sliderGreen.Value
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub sliderBlue_Scroll()
    txtBlue.Text = sliderBlue.Value
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub sliderHue_Scroll()
    Dim lngRGB As Long
    
    txtHue.Text = sliderHue.Value
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub sliderSaturation_Scroll()
    Dim lngRGB As Long
    
    txtSaturation.Text = sliderSaturation.Value
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub sliderLuminance_Scroll()
    Dim lngRGB As Long
    
    txtLuminance.Text = sliderLuminance.Value
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub timerResize_Timer()
    Call SyncPreview(True)
    Call InitHueScale
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncSB
    timerResize.Enabled = False
End Sub

Private Sub txtGammaRed_Change()
    Call SyncPreview(False)
    Call SyncCode
End Sub

Private Sub txtGammaGreen_Change()
    Call SyncPreview(False)
    Call SyncCode
End Sub

Private Sub txtGammaBlue_Change()
    Call SyncPreview(False)
    Call SyncCode
End Sub

Private Sub txtRed_GotFocus()
    txtRed.SelStart = 0
    txtRed.SelLength = Len(txtRed.Text)
End Sub

Private Sub txtGreen_GotFocus()
    txtGreen.SelStart = 0
    txtGreen.SelLength = Len(txtGreen.Text)
End Sub

Private Sub txtBlue_GotFocus()
    txtBlue.SelStart = 0
    txtBlue.SelLength = Len(txtBlue.Text)
End Sub

Private Sub txtHue_GotFocus()
    txtHue.SelStart = 0
    txtHue.SelLength = Len(txtHue.Text)
End Sub

Private Sub txtSaturation_GotFocus()
    txtSaturation.SelStart = 0
    txtSaturation.SelLength = Len(txtSaturation.Text)
End Sub

Private Sub txtLuminance_GotFocus()
    txtLuminance.SelStart = 0
    txtLuminance.SelLength = Len(txtLuminance.Text)
End Sub

Private Sub txtGammaRed_GotFocus()
    txtGammaRed.SelStart = 0
    txtGammaRed.SelLength = Len(txtGammaRed.Text)
End Sub

Private Sub txtGammaGreen_GotFocus()
    txtGammaGreen.SelStart = 0
    txtGammaGreen.SelLength = Len(txtGammaGreen.Text)
End Sub

Private Sub txtGammaBlue_GotFocus()
    txtGammaBlue.SelStart = 0
    txtGammaBlue.SelLength = Len(txtGammaBlue.Text)
End Sub

Private Sub txtCode_GotFocus()
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
End Sub

Private Sub txtHue_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lngRGB As Long
    
    Call Validate(txtHue, 239, 0)
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub txtSaturation_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lngRGB As Long
    
    Call Validate(txtSaturation, 240, 0)
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub txtLuminance_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lngRGB As Long
    
    Call Validate(txtLuminance, 240, 0)
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub txtRed_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Validate(txtRed, 255, 0)
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub txtGreen_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Validate(txtGreen, 255, 0)
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub txtBlue_KeyUp(KeyCode As Integer, Shift As Integer)
    Call Validate(txtBlue, 255, 0)
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub txtGammaRed_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lngRGB As Long
    
    Call Validate(txtGammaRed, 4096, -4096)
    lngRGB = ColorGammaToRGB(txtGammaRed.Text, txtGammaGreen.Text, txtGammaBlue.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncHLS(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub txtGammaGreen_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lngRGB As Long
    
    Call Validate(txtGammaGreen, 4096, -4096)
    lngRGB = ColorGammaToRGB(txtGammaRed.Text, txtGammaGreen.Text, txtGammaBlue.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncHLS(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub txtGammaBlue_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim lngRGB As Long
    
    Call Validate(txtGammaBlue, 4096, -4096)
    lngRGB = ColorGammaToRGB(txtGammaRed.Text, txtGammaGreen.Text, txtGammaBlue.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncHLS(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub UpDownRed_Change()
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub UpDownGreen_Change()
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub UpDownBlue_Change()
    Call SyncAll(txtRed.Text, txtGreen.Text, txtBlue.Text)
    Call SyncAllScales(txtRed.Text, txtGreen.Text, txtBlue.Text)
End Sub

Private Sub UpDownHue_Change()
    Dim lngRGB As Long
    
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub UpDownSaturation_Change()
    Dim lngRGB As Long
    
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub UpDownLuminance_Change()
    Dim lngRGB As Long
    
    lngRGB = ColorHLSToRGB(txtHue.Text, txtLuminance.Text, txtSaturation.Text)
    Call SyncRGB(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncGamma(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncSwatch(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
    Call SyncAllScales(RGBRed(lngRGB), RGBGreen(lngRGB), RGBBlue(lngRGB))
End Sub

Private Sub cmdColorPicker_Click()
    Call ShowColorPicker
End Sub

Private Sub cmdCopyGamma_Click()
    Clipboard.Clear
    Clipboard.SetText (txtGammaRed.Text & ", " & txtGammaGreen.Text & ", " & txtGammaBlue.Text)
End Sub

Private Sub cmdCopyRGB_Click()
    Clipboard.Clear
    Clipboard.SetText (txtRed.Text & ", " & txtGreen.Text & ", " & txtBlue.Text)
End Sub

Private Sub cmdCopyHSL_Click()
    Clipboard.Clear
    Clipboard.SetText (txtHue.Text & ", " & txtSaturation.Text & ", " & txtGreen.Text)
End Sub

Private Sub Validate(ByRef txtField As TextBox, ByVal intMax As Integer, ByVal intMin As Integer)
    If Val(txtField.Text) > intMax Then
        txtField.Text = intMax
    ElseIf Val(txtField.Text) < intMin Then
        txtField.Text = intMin
    Else
        txtField.Text = Val(txtField.Text)
    End If
End Sub

Private Sub ShowColorPicker()
    Dim sColor As SelectedColor
    sColor = ShowColor(Me.hwnd, True, RGB(txtRed.Text, txtGreen.Text, txtBlue.Text))
    If Not sColor.bCanceled Then
        Call SyncAll(RGBRed(sColor.oSelectedColor), RGBGreen(sColor.oSelectedColor), RGBBlue(sColor.oSelectedColor))
        Call SyncAllScales(RGBRed(sColor.oSelectedColor), RGBGreen(sColor.oSelectedColor), RGBBlue(sColor.oSelectedColor))
    End If
End Sub

Private Sub cmdLoadPNG_Click()
    Dim strPreviousFile As String
    
    If PNG.CurrentFilename <> "" And PNG.Interlaced = False Then
        strPreviousFile = PNG.CurrentFilename
    End If
    
    FileDlg.Owner = Me
    FileDlg.Flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST
    FileDlg.Filter = "Portable Network Graphics (PNG)|*.png|All Files|*.*"
    FileDlg.ShowOpen
    If FileDlg.FileName <> "" Then
        Select Case PNG.LoadPNGFile(FileDlg.FileName)
            Case pngeFileNotFound
                MsgBox "The specified file could not be found.", vbCritical
            Case pngeOpenError
                MsgBox "Error opening the file.", vbCritical
            Case pngeInvalidFile
                MsgBox "The specified file is no valid PNG file.", vbCritical
            Case pngeSucceeded
                If Not PNG.Interlaced Then
                    picDraw.Visible = True
                    Call SyncPreview(True)
                    Call SyncSB
                    formMain.Caption = Right(PNG.CurrentFilename, Len(PNG.CurrentFilename) - InStrRev(PNG.CurrentFilename, "\")) & " - " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
                    If PNG.ColorType = pngeRGB Then
                        MsgBox "This PNG file does not have an alpha channel." & vbCrLf & "Currently this program cannot preview gamma values on PNG files without an alpha channel." & vbCrLf & "I'm working on adding this functionality.", vbInformation, "Sorry"
                    End If
                Else
                    MsgBox "This program does not currently support interlaced PNG files. Resave the PNG as non-interlaced if you wish to open it in this program.", vbInformation, "Interlaced PNG file not supported"
                    If strPreviousFile <> "" Then
                        PNG.LoadPNGFile (strPreviousFile)
                    End If
                End If
        End Select
    End If
End Sub

Private Sub cmdPNGUpdate_Click()
    Call SyncPreview(True)
End Sub

Private Sub cmdCopyCode_Click()
    Clipboard.Clear
    Clipboard.SetText (txtCode.Text)
End Sub

Private Sub chkAutoUpdate_Click()
    If chkAutoUpdate Then
        Call SyncPreview(False)
        Call SyncCode
    End If
End Sub

Private Sub chkBoost_Click()
    Call SyncPreview(False)
    Call SyncCode
End Sub

Private Sub cmdReset_Click()
    txtRed.Text = "128"
    txtGreen.Text = "128"
    txtBlue.Text = "128"
    Call SyncAll(128, 128, 128)
    Call SyncAllScales(128, 128, 128)
End Sub

Private Sub cmdAbout_Click()
    formAbout.Show vbModal, Me
End Sub

Private Sub cmdEyedropper_Click()
    Dim lngReturn As Long
    lngReturn = SetCapture(picSwatch.hwnd)
    tmrEyeDropper.Interval = 50
End Sub

Private Sub tmrEyeDropper_Timer()
    Static lX As Long
    Static lY As Long
    
    On Local Error Resume Next
    
    Dim P As PointAPI
    Dim H As Long
    Dim hD As Long
    Dim R As Long
    
    GetCursorPos P
    If P.x = lX And P.y = lY Then Exit Sub
    lX = P.x: lY = P.y
    H = WindowFromPoint(lX, lY)
    hD = GetDC(H)
    ScreenToClient H, P
    R = GetPixel(hD, P.x, P.y)
    If R = -1 Then
        BitBlt picSwatch.hDC, 0, 0, 1, 1, hD, P.x, P.y, vbSrcCopy
        R = picSwatch.Point(0, 0)
    Else
        picSwatch.PSet (0, 0), R
    End If
    ReleaseDC H, hD
    picSwatch.BackColor = R
    
    Call SyncAll(RGBRed(R), RGBGreen(R), RGBBlue(R))
    Call SyncAllScales(RGBRed(R), RGBGreen(R), RGBBlue(R))
End Sub
