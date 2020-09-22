VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "GeoStar"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8625
   DrawWidth       =   2
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgList 
      Left            =   6720
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49E2
            Key             =   ""
            Object.Tag             =   "&New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F7E
            Key             =   ""
            Object.Tag             =   "&Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":551A
            Key             =   ""
            Object.Tag             =   "&Create Star"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AB6
            Key             =   ""
            Object.Tag             =   "C&lear"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6052
            Key             =   ""
            Object.Tag             =   "&Random Star"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65EE
            Key             =   ""
            Object.Tag             =   "&Save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstTabStrip 
      Left            =   6120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CE26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB7A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comdiag 
      Left            =   5040
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbr1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglstImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.Tag             =   "NEW"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.Tag             =   "SAVE"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CREATE"
            Object.Tag             =   "CREATE"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CLEAR"
            Object.Tag             =   "CLEAR"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RANDOM"
            Object.Tag             =   "RANDOM"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXIT"
            Object.Tag             =   "EXIT"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCon 
      BorderStyle     =   0  'None
      Height          =   5025
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   3375
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   615
         Left            =   2280
         Picture         =   "frmMain.frx":10F5E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4320
         Width           =   975
      End
      Begin VB.CommandButton cmdRandom 
         Caption         =   "&Random"
         Height          =   615
         Left            =   1200
         Picture         =   "frmMain.frx":114E8
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Frame fraAttributes 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   3135
         Begin VB.TextBox txtLineSize 
            Height          =   285
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   61
            Text            =   "1"
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox txtSideLength 
            Height          =   285
            Left            =   1440
            TabIndex        =   60
            Text            =   "1200"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtNumber 
            Height          =   285
            Left            =   1440
            TabIndex        =   59
            Text            =   "4"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtWidth 
            Height          =   285
            Left            =   1440
            TabIndex        =   58
            Text            =   "540"
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtDensity 
            Height          =   285
            Left            =   1440
            TabIndex        =   57
            Text            =   "13"
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label lblLine_Size 
            Alignment       =   1  'Right Justify
            Caption         =   "Line Size:"
            Height          =   255
            Left            =   600
            TabIndex        =   66
            Top             =   2760
            Width           =   735
         End
         Begin VB.Line lnBorderR 
            BorderColor     =   &H80000010&
            Index           =   4
            X1              =   3120
            X2              =   3120
            Y1              =   3360
            Y2              =   0
         End
         Begin VB.Line lnBorderT 
            BorderColor     =   &H00E0E0E0&
            Index           =   4
            X1              =   3120
            X2              =   0
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lnBorderL 
            BorderColor     =   &H00E0E0E0&
            Index           =   4
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   3360
         End
         Begin VB.Line lnBorderB 
            BorderColor     =   &H80000010&
            Index           =   4
            X1              =   0
            X2              =   3120
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Image imgStyle 
            Height          =   480
            Index           =   3
            Left            =   2520
            Picture         =   "frmMain.frx":11A72
            Top             =   2040
            Width           =   480
         End
         Begin VB.Image imgStyle 
            Height          =   480
            Index           =   2
            Left            =   2520
            Picture         =   "frmMain.frx":1273C
            Top             =   1440
            Width           =   480
         End
         Begin VB.Image imgStyle 
            Height          =   480
            Index           =   1
            Left            =   2520
            Picture         =   "frmMain.frx":13406
            Top             =   840
            Width           =   480
         End
         Begin VB.Image imgStyle 
            Height          =   480
            Index           =   0
            Left            =   2520
            Picture         =   "frmMain.frx":140D0
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblSideLength 
            Alignment       =   1  'Right Justify
            Caption         =   "Side length:"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblNumber 
            Alignment       =   1  'Right Justify
            Caption         =   "# of Sides:"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   64
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblWidth 
            Alignment       =   1  'Right Justify
            Caption         =   "Width between:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label lblDensity 
            Alignment       =   1  'Right Justify
            Caption         =   "Density:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   2160
            Width           =   1215
         End
      End
      Begin VB.Frame fraColours 
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         Begin VB.Frame fraRGB 
            BorderStyle     =   0  'None
            Height          =   2535
            Index           =   0
            Left            =   0
            TabIndex        =   44
            Top             =   840
            Width           =   3135
            Begin VB.TextBox txtForeR 
               Height          =   285
               Left            =   2520
               TabIndex        =   49
               Text            =   "lr"
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox txtForeG 
               Height          =   285
               Left            =   2520
               TabIndex        =   48
               Text            =   "lr"
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox txtForeB 
               Height          =   285
               Left            =   2520
               TabIndex        =   47
               Text            =   "lr"
               Top             =   1320
               Width           =   375
            End
            Begin VB.PictureBox picForePreview 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   1635
               TabIndex        =   46
               Top             =   2040
               Width           =   1695
            End
            Begin VB.CommandButton cmdCustomFore 
               Caption         =   "Custom.."
               Height          =   375
               Left            =   1920
               TabIndex        =   45
               Top             =   2040
               Width           =   1095
            End
            Begin MSComctlLib.Slider sldrForeR 
               Height          =   255
               Left            =   600
               TabIndex        =   50
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   20
               TickFrequency   =   10
               Value           =   20
            End
            Begin MSComctlLib.Slider sldrForeG 
               Height          =   255
               Left            =   600
               TabIndex        =   51
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   100
               TickFrequency   =   10
               Value           =   100
            End
            Begin MSComctlLib.Slider sldrForeB 
               Height          =   255
               Left            =   600
               TabIndex        =   52
               Top             =   1320
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   230
               TickFrequency   =   10
               Value           =   230
            End
            Begin VB.Line lnBorderR 
               BorderColor     =   &H80000010&
               Index           =   3
               X1              =   3120
               X2              =   3120
               Y1              =   2520
               Y2              =   0
            End
            Begin VB.Line lnBorderT 
               BorderColor     =   &H00E0E0E0&
               Index           =   3
               X1              =   3120
               X2              =   0
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line lnBorderL 
               BorderColor     =   &H00E0E0E0&
               Index           =   3
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   2520
            End
            Begin VB.Line lnBorderB 
               BorderColor     =   &H80000010&
               Index           =   3
               X1              =   0
               X2              =   3120
               Y1              =   2520
               Y2              =   2520
            End
            Begin VB.Label lblForeB 
               Alignment       =   1  'Right Justify
               Caption         =   "Blue:"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label lblForeG 
               Alignment       =   1  'Right Justify
               Caption         =   "Green:"
               Height          =   255
               Left            =   60
               TabIndex        =   54
               Top             =   720
               Width           =   555
            End
            Begin VB.Label lblForeR 
               Alignment       =   1  'Right Justify
               Caption         =   "Red:"
               Height          =   255
               Left            =   120
               TabIndex        =   53
               Top             =   120
               Width           =   495
            End
            Begin VB.Line lnSeparator2 
               BorderColor     =   &H80000010&
               Index           =   1
               X1              =   120
               X2              =   3000
               Y1              =   1920
               Y2              =   1920
            End
         End
         Begin VB.Frame fraRGB 
            BorderStyle     =   0  'None
            Height          =   2535
            Index           =   1
            Left            =   0
            TabIndex        =   32
            Top             =   840
            Width           =   3135
            Begin VB.CommandButton cmdCustomBack 
               Caption         =   "Custom.."
               Height          =   375
               Left            =   1920
               TabIndex        =   37
               Top             =   2040
               Width           =   1095
            End
            Begin VB.PictureBox picBackPreview 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   1635
               TabIndex        =   36
               Top             =   2040
               Width           =   1695
            End
            Begin VB.TextBox txtBackB 
               Height          =   285
               Left            =   2520
               TabIndex        =   35
               Text            =   "255"
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox txtBackG 
               Height          =   285
               Left            =   2520
               TabIndex        =   34
               Text            =   "255"
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox txtBackR 
               Height          =   285
               Left            =   2520
               TabIndex        =   33
               Text            =   "255"
               Top             =   120
               Width           =   375
            End
            Begin MSComctlLib.Slider sldrBackR 
               Height          =   255
               Left            =   600
               TabIndex        =   38
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   255
               TickFrequency   =   10
               Value           =   255
            End
            Begin MSComctlLib.Slider sldrBackG 
               Height          =   255
               Left            =   600
               TabIndex        =   39
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   255
               TickFrequency   =   10
               Value           =   255
            End
            Begin MSComctlLib.Slider sldrBackB 
               Height          =   255
               Left            =   600
               TabIndex        =   40
               Top             =   1320
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   255
               TickFrequency   =   10
               Value           =   255
            End
            Begin VB.Line lnSeparator2 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   120
               X2              =   3000
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Label lblBackR 
               Alignment       =   1  'Right Justify
               Caption         =   "Red:"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   120
               Width           =   495
            End
            Begin VB.Label lblBackG 
               Alignment       =   1  'Right Justify
               Caption         =   "Green:"
               Height          =   255
               Left            =   60
               TabIndex        =   42
               Top             =   720
               Width           =   555
            End
            Begin VB.Label lblBackB 
               Alignment       =   1  'Right Justify
               Caption         =   "Blue:"
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   1320
               Width           =   495
            End
            Begin VB.Line lnBorderB 
               BorderColor     =   &H80000010&
               Index           =   1
               X1              =   0
               X2              =   3120
               Y1              =   2520
               Y2              =   2520
            End
            Begin VB.Line lnBorderL 
               BorderColor     =   &H00E0E0E0&
               Index           =   1
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   2520
            End
            Begin VB.Line lnBorderT 
               BorderColor     =   &H00E0E0E0&
               Index           =   1
               X1              =   3120
               X2              =   0
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line lnBorderR 
               BorderColor     =   &H80000010&
               Index           =   1
               X1              =   3120
               X2              =   3120
               Y1              =   2520
               Y2              =   -480
            End
         End
         Begin VB.Frame fraRGB 
            BorderStyle     =   0  'None
            Height          =   2535
            Index           =   2
            Left            =   0
            TabIndex        =   20
            Top             =   840
            Width           =   3135
            Begin VB.CommandButton cmdCustomShadow 
               Caption         =   "Custom.."
               Height          =   375
               Left            =   1920
               TabIndex        =   25
               Top             =   2040
               Width           =   1095
            End
            Begin VB.PictureBox picShadowPreview 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   120
               ScaleHeight     =   315
               ScaleWidth      =   1635
               TabIndex        =   24
               Top             =   2040
               Width           =   1695
            End
            Begin VB.TextBox txtShadowB 
               Height          =   285
               Left            =   2520
               TabIndex        =   23
               Text            =   "150"
               Top             =   1320
               Width           =   375
            End
            Begin VB.TextBox txtShadowG 
               Height          =   285
               Left            =   2520
               TabIndex        =   22
               Text            =   "150"
               Top             =   720
               Width           =   375
            End
            Begin VB.TextBox txtShadowR 
               Height          =   285
               Left            =   2520
               TabIndex        =   21
               Text            =   "150"
               Top             =   120
               Width           =   375
            End
            Begin MSComctlLib.Slider sldrShadowR 
               Height          =   255
               Left            =   600
               TabIndex        =   26
               Top             =   120
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   255
               TickFrequency   =   10
               Value           =   150
            End
            Begin MSComctlLib.Slider sldrShadowG 
               Height          =   255
               Left            =   600
               TabIndex        =   27
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   255
               TickFrequency   =   10
               Value           =   150
            End
            Begin MSComctlLib.Slider sldrShadowB 
               Height          =   255
               Left            =   600
               TabIndex        =   28
               Top             =   1320
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   450
               _Version        =   393216
               Max             =   255
               SelStart        =   150
               TickFrequency   =   10
               Value           =   150
            End
            Begin VB.Line lnSeparator2 
               BorderColor     =   &H80000010&
               Index           =   2
               X1              =   120
               X2              =   3000
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Label lblShadowR 
               Alignment       =   1  'Right Justify
               Caption         =   "Red:"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   120
               Width           =   495
            End
            Begin VB.Label lblShadowG 
               Alignment       =   1  'Right Justify
               Caption         =   "Green:"
               Height          =   255
               Left            =   60
               TabIndex        =   30
               Top             =   720
               Width           =   555
            End
            Begin VB.Label lblShadowB 
               Alignment       =   1  'Right Justify
               Caption         =   "Blue:"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   1320
               Width           =   495
            End
            Begin VB.Line lnBorderB 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   0
               X2              =   3120
               Y1              =   2520
               Y2              =   2520
            End
            Begin VB.Line lnBorderL 
               BorderColor     =   &H00E0E0E0&
               Index           =   0
               X1              =   0
               X2              =   0
               Y1              =   0
               Y2              =   2520
            End
            Begin VB.Line lnBorderT 
               BorderColor     =   &H00E0E0E0&
               Index           =   0
               X1              =   3120
               X2              =   0
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line lnBorderR 
               BorderColor     =   &H80000010&
               Index           =   0
               X1              =   3120
               X2              =   3120
               Y1              =   2520
               Y2              =   0
            End
         End
         Begin VB.ComboBox cmbColourSelect 
            Height          =   315
            ItemData        =   "frmMain.frx":14D9A
            Left            =   120
            List            =   "frmMain.frx":14DA7
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   2895
         End
         Begin VB.Line lnBorderR 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   3120
            X2              =   3120
            Y1              =   720
            Y2              =   0
         End
         Begin VB.Line lnBorderT 
            BorderColor     =   &H00E0E0E0&
            Index           =   5
            X1              =   3120
            X2              =   0
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lnBorderL 
            BorderColor     =   &H00E0E0E0&
            Index           =   5
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   720
         End
         Begin VB.Line lnBorderB 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   0
            X2              =   3120
            Y1              =   720
            Y2              =   720
         End
      End
      Begin VB.Frame fraShadow 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox txtShadowSize 
            Height          =   285
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   16
            Text            =   "1"
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optShadow 
            Height          =   255
            Index           =   0
            Left            =   1395
            TabIndex        =   15
            Top             =   3000
            Value           =   -1  'True
            Width           =   215
         End
         Begin VB.CheckBox chkShadow 
            Alignment       =   1  'Right Justify
            Caption         =   "      Shadow:"
            Height          =   255
            Left            =   180
            TabIndex        =   14
            Top             =   120
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.OptionButton optShadow 
            Height          =   255
            Index           =   1
            Left            =   2100
            TabIndex        =   13
            Top             =   2760
            Width           =   215
         End
         Begin VB.OptionButton optShadow 
            Height          =   255
            Index           =   2
            Left            =   2340
            TabIndex        =   12
            Top             =   2040
            Width           =   215
         End
         Begin VB.OptionButton optShadow 
            Height          =   255
            Index           =   3
            Left            =   2100
            TabIndex        =   11
            Top             =   1320
            Width           =   215
         End
         Begin VB.OptionButton optShadow 
            Height          =   255
            Index           =   4
            Left            =   1395
            TabIndex        =   10
            Top             =   1080
            Width           =   215
         End
         Begin VB.OptionButton optShadow 
            Alignment       =   1  'Right Justify
            Height          =   255
            Index           =   5
            Left            =   660
            TabIndex        =   9
            Top             =   1320
            Width           =   215
         End
         Begin VB.OptionButton optShadow 
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   8
            Top             =   2040
            Width           =   215
         End
         Begin VB.OptionButton optShadow 
            Height          =   255
            Index           =   7
            Left            =   660
            TabIndex        =   7
            Top             =   2760
            Width           =   215
         End
         Begin VB.Shape shpCircle 
            FillColor       =   &H00FF8080&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   1140
            Shape           =   3  'Circle
            Tag             =   "b"
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label lblShadowSize 
            Alignment       =   1  'Right Justify
            Caption         =   "Shadow Size:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
         Begin VB.Shape shpShadow 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   1140
            Shape           =   3  'Circle
            Tag             =   "b"
            Top             =   1920
            Width           =   735
         End
         Begin VB.Line lblSeparator 
            BorderColor     =   &H80000010&
            X1              =   180
            X2              =   3000
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line lnShadow 
            Index           =   5
            X1              =   1020
            X2              =   900
            Y1              =   1680
            Y2              =   1560
         End
         Begin VB.Line lnShadow 
            Index           =   1
            X1              =   1980
            X2              =   2100
            Y1              =   2640
            Y2              =   2760
         End
         Begin VB.Line lnShadow 
            Index           =   3
            X1              =   1980
            X2              =   2100
            Y1              =   1680
            Y2              =   1560
         End
         Begin VB.Line lnShadow 
            Index           =   7
            X1              =   1020
            X2              =   900
            Y1              =   2640
            Y2              =   2760
         End
         Begin VB.Line lnShadow 
            Index           =   6
            X1              =   900
            X2              =   780
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line lnShadow 
            Index           =   2
            X1              =   2100
            X2              =   2220
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line lnShadow 
            Index           =   4
            X1              =   1500
            X2              =   1500
            Y1              =   1560
            Y2              =   1440
         End
         Begin VB.Line lnShadow 
            BorderColor     =   &H00FF0000&
            Index           =   0
            X1              =   1500
            X2              =   1500
            Y1              =   2880
            Y2              =   2760
         End
         Begin VB.Line lnBorderB 
            BorderColor     =   &H80000010&
            Index           =   2
            X1              =   0
            X2              =   3120
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line lnBorderL 
            BorderColor     =   &H00E0E0E0&
            Index           =   2
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   3360
         End
         Begin VB.Line lnBorderT 
            BorderColor     =   &H00E0E0E0&
            Index           =   2
            X1              =   3120
            X2              =   0
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line lnBorderR 
            BorderColor     =   &H80000010&
            Index           =   2
            X1              =   3120
            X2              =   3120
            Y1              =   3360
            Y2              =   0
         End
      End
      Begin MSComctlLib.TabStrip tabProperties 
         Height          =   3975
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   7011
         ImageList       =   "imglstTabStrip"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Attributes"
               Key             =   "Atrributes"
               Object.Tag             =   "Atrributes"
               Object.ToolTipText     =   "Atrributes"
               ImageVarType    =   2
               ImageIndex      =   1
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Colours"
               Key             =   "Colours"
               Object.Tag             =   "Colours"
               Object.ToolTipText     =   "Colours"
               ImageVarType    =   2
               ImageIndex      =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Shadow"
               Key             =   "Shadow"
               Object.Tag             =   "Shadow"
               Object.ToolTipText     =   "Shadow"
               ImageVarType    =   2
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCreate 
         Appearance      =   0  'Flat
         Caption         =   "&Create"
         Height          =   615
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmMain.frx":14DCB
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4320
         Width           =   1095
      End
   End
   Begin VB.PictureBox pbx1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5025
      Left            =   0
      ScaleHeight     =   4965
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin MSComctlLib.ImageList imglstImages 
      Left            =   5520
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15355
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AF79
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20415
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25C09
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27915
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27EB5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuStar 
      Caption         =   "&Star"
      Begin VB.Menu mnuStarCreate 
         Caption         =   "&Create Star"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuStarRandom 
         Caption         =   "&Random Star"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuStarSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStarClear 
         Caption         =   "C&lear"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Name          : frmMain
' Description   : The main Form
' Lines of Code : 1395
'
' Modified      : 10/10/2001
'
' --------------------------------------------------
Option Explicit

    'Declare PI
    Const m_cPI = 3.141592654

    'Declare public variables
    Dim m_blnGstar As Boolean
    Dim m_blnOverride As Boolean

Public Function FileExist(p_strSfname As String) As Boolean
    
    'Declare Variables
    Dim lngTempbx1      As Long
    
    On Error Resume Next
    
    'See if file exists
    lngTempbx1 = GetAttr(p_strSfname)
    
    'If there was an error then the file doesn't exist
    If Err Then
        FileExist = False
    Else
        FileExist = True
    End If
    
End Function


Public Function ResetControls()
 
    'Error Handling
    On Error GoTo PROC_ERR

    'Reset all controls and clear screen
    txtForeR.Text = "56"
    txtForeG.Text = "112"
    txtForeB.Text = "168"
    
    UpdatefromText "ForeColour"
    
    txtBackR.Text = "255"
    txtBackG.Text = "255"
    txtBackB.Text = "255"
    
    UpdatefromText "BackColour"
    
    txtShadowR.Text = "200"
    txtShadowG.Text = "200"
    txtShadowB.Text = "255"
    
    UpdatefromText "Shadow"
    
    txtDensity.Text = "25"
    txtWidth.Text = "500"
    txtLineSize.Text = "1"
    txtNumber.Text = "4"
    txtSideLength.Text = "1200"
    
    txtShadowSize.Text = "1"
    chkShadow.Value = vbChecked
    
    pbx1.Cls
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Function

Public Function UpdatefromText(strColour As String)

    'Declare Variables
    Dim lngR        As Long
    Dim lngG        As Long
    Dim lngB        As Long

    'Error Handling
    On Error GoTo PROC_ERR
    
Select Case strColour

    Case "ForeColour"   'Change ForeColour Values
        If txtForeR.Text <> "sr" Or txtForeR.Text <> "lr" Then lngR = CLng(Val(txtForeR.Text))
        If txtForeG.Text <> "sr" Or txtForeG.Text <> "lr" Then lngG = CLng(Val(txtForeG.Text))
        If txtForeB.Text <> "sr" Or txtForeB.Text <> "lr" Then lngB = CLng(Val(txtForeB.Text))
        
        sldrForeR.Value = lngR
        sldrForeG.Value = lngG
        sldrForeB.Value = lngB
        
        picForePreview.BackColor = RGB(lngR, lngG, lngB)
        
        lngR = 0
        lngG = 0
        lngB = 0
    
    
    Case "BackColour"   'Change BackColour Values
        lngR = CLng(Val(txtBackR.Text))
        lngG = CLng(Val(txtBackG.Text))
        lngB = CLng(Val(txtBackB.Text))
        
        sldrBackR.Value = lngR
        sldrBackG.Value = lngG
        sldrBackB.Value = lngB
        
        picBackPreview.BackColor = RGB(lngR, lngG, lngB)
        
        lngR = 0
        lngG = 0
        lngB = 0
    
    
    Case "Shadow"   'Change Shadow Values
        lngR = CLng(Val(txtShadowR.Text))
        lngG = CLng(Val(txtShadowG.Text))
        lngB = CLng(Val(txtShadowB.Text))
        
        sldrShadowR.Value = lngR
        sldrShadowG.Value = lngG
        sldrShadowB.Value = lngB
        
        picShadowPreview.BackColor = RGB(lngR, lngG, lngB)
        
        lngR = 0
        lngG = 0
        lngB = 0
    
End Select
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    

End Function

Public Function UpdatefromSldr(strColour As String)

    'Error Handling
    On Error GoTo PROC_ERR
    
    'Declare Variables
    Dim lngR        As Long
    Dim lngG        As Long
    Dim lngB        As Long
    
    'If textboxes contain a random text value then exit
    If txtForeR.Text = "sr" Or txtForeR.Text = "lr" Then Exit Function
    If txtForeG.Text = "sr" Or txtForeG.Text = "lr" Then Exit Function
    If txtForeB.Text = "sr" Or txtForeB.Text = "lr" Then Exit Function
    
Select Case strColour
    
    Case "ForeColour"   'Change ForeColour Values
        lngR = sldrForeR.Value
        lngG = sldrForeG.Value
        lngB = sldrForeB.Value
        
        txtForeR.Text = lngR
        txtForeG.Text = lngG
        txtForeB.Text = lngB
        
        picForePreview.BackColor = RGB(lngR, lngG, lngB)
        
        lngR = 0
        lngG = 0
        lngB = 0
    
    
    Case "BackColour"   'Change BackColour Values
        lngR = sldrBackR.Value
        lngG = sldrBackG.Value
        lngB = sldrBackB.Value
        
        txtBackR.Text = lngR
        txtBackG.Text = lngG
        txtBackB.Text = lngB
        
        picBackPreview.BackColor = RGB(lngR, lngG, lngB)
        
        lngR = 0
        lngG = 0
        lngB = 0
    
    
    Case "Shadow"   'Change Shadow Values
        lngR = sldrShadowR.Value
        lngG = sldrShadowG.Value
        lngB = sldrShadowB.Value
        
        txtShadowR.Text = lngR
        txtShadowG.Text = lngG
        txtShadowB.Text = lngB
        
        picShadowPreview.BackColor = RGB(lngR, lngG, lngB)
        
        lngR = 0
        lngG = 0
        lngB = 0
    
End Select
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    

End Function



Public Function SaveStar()
    
    'Error Handling
    On Error GoTo PROC_ERR

    'Delcare Variables
    Dim strSfname As String

    'Show Common Dialog
    comdiag.Filter = "Bitmap|*.bmp"
    comdiag.DialogTitle = "Save as..."
    comdiag.InitDir = App.Path
    comdiag.ShowSave
    strSfname = comdiag.filename
        
    'See if user cancelled
    If strSfname = "" Then
        Exit Function
    End If
            
    'See if file exists
    If FileExist(strSfname) = True Then
        MsgBox "The File already exist!", vbCritical, "Error"
        Exit Function
    End If
            
    'Save file
    If (LCase$(Right$(strSfname, 4)) = ".bmp") Then
        SavePicture pbx1.Image, strSfname
    Else
        strSfname = strSfname & ".bmp"
        SavePicture pbx1.Image, strSfname
    End If

    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    

End Function

Public Function CreateStar()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Declare Variables
    Dim lngLength       As Long
    Dim lngWidth        As Long
    Dim dblSides        As Double
    Dim lngDensity      As Long
    Dim strFTR          As String
    Dim strFTG          As String
    Dim strFTB          As String
    Dim blnShadow       As Boolean
    Dim strShadowPos    As String
    Dim intLineSize     As Integer
    Dim lngSR           As Long
    Dim lngSG           As Long
    Dim lngSB           As Long
    Dim lngBR           As Long
    Dim lngBG           As Long
    Dim lngBB           As Long
    Dim intShadowSize   As Integer
    
    'Get attribute variables
    lngLength = CLng(Val(txtSideLength.Text))
    lngWidth = CLng(Val(txtWidth.Text))
    lngDensity = CLng(Val(txtDensity.Text))
    intLineSize = CInt(Val(txtLineSize.Text))
    dblSides = Val(txtNumber.Text)
    
    'Get shadow variables
    blnShadow = False
    If chkShadow.Value = vbChecked Then blnShadow = True
    
    strShadowPos = shpCircle.Tag
    intShadowSize = CInt(Val(txtShadowSize.Text))
    
    
    'Get colour variables
    strFTR = txtForeR.Text
    strFTG = txtForeG.Text
    strFTB = txtForeB.Text
    
    lngSR = CLng(Val(txtShadowR.Text))
    lngSG = CLng(Val(txtShadowG.Text))
    lngSB = CLng(Val(txtShadowB.Text))
    
    lngBR = CLng(Val(txtBackR.Text))
    lngBG = CLng(Val(txtBackG.Text))
    lngBB = CLng(Val(txtBackB.Text))
    
    'Change Picturebox BackColor
    pbx1.BackColor = RGB(lngBR, lngBG, lngBB)
    
    'Draw star
    GeoStar Me, pbx1, lngLength, lngWidth, dblSides, lngDensity, strFTR, strFTG, strFTB, blnShadow, strShadowPos, intLineSize, intShadowSize, lngSR, lngSG, lngSB
    
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    

End Function

Public Function SelectAll(txtbx As TextBox)

    'Error Handling
    On Error GoTo PROC_ERR

    'Select all Text
    txtbx.SelStart = 0
    txtbx.SelLength = Len(txtbx.Text)

PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    

End Function

Public Function LimitTxt(txtbx As TextBox, lngMax As Long, Optional lngLow As Long)

    'Error Handling
    On Error GoTo PROC_ERR
    
    'See if user typed in lngLow, if not, make it Zero
    If lngLow = Null Then lngLow = 0

    'Check to see if text is greater or smaller than allowable values
    If CLng(Val(txtbx.Text)) > lngMax Then
        txtbx.Text = lngMax
    End If
    If CLng(Val(txtbx.Text)) < lngLow Then
        txtbx.Text = lngLow
    End If
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT


End Function

Private Sub cmdRandom_Click()

    'Error Handling
    On Error GoTo PROC_ERR
    
    'Make Random values then create star
    RandomValues
    CreateStar
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub cmdCreate_Click()

    'Error Handling
    On Error GoTo PROC_ERR
    
    'Create Star
    CreateStar
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub cmdExit_Click()

    'Error Handling
    On Error GoTo PROC_ERR
    
    'Exit
    Unload Me

    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub Form_Load()

    'Error Handling
    On Error GoTo PROC_ERR

    Set CoolMenuObj = New CoolMenu
    Call CoolMenuObj.Install(Me.hwnd, imgList, True, True)

    'Reset controls
    ResetControls
    
    'Set Combo Text
    cmbColourSelect.Text = "ForeColour"
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub Form_Resize()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    ' Can't move or resize a minimized  form
    If Not Me.WindowState = vbMinimized Then
        
        'If form is smaller than original height then stop it!
        If Me.Height < 6075 Then
            Me.Height = 6075
        End If
        If Me.Width < 8745 Then
            Me.Width = 8745
        End If
        
    End If
    
    ' Can't move or resize a minimized  form
    If Not Me.WindowState = vbMinimized Then
        
        'Resize all controls to follow form
        With Me
            .pbx1.Height = .Height - 400 - 370 - 285
            .fraCon.Left = .Width - fraCon.Width - 195
            .pbx1.Width = .fraCon.Left - 120
            .fraCon.Height = .Height - 465 - 370 - 285
            .cmdRandom.Top = .fraCon.Height - cmdRandom.Height - 60
            .cmdCreate.Top = .fraCon.Height - cmdCreate.Height - 60
            .cmdExit.Top = .fraCon.Height - cmdExit.Height - 60
        
        End With
        
    End If
    
    'Clear screen
    pbx1.Cls
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub Form_Unload(p_intCancel As Integer)
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    CoolMenuObj.Install &O0
    
    ' Explicit Clean Up
    Set frmMain = Nothing
    Set CoolMenuObj = Nothing
    
    ' Yield to processes
    DoEvents
    
    'End
    End
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub mnuFileExit_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub mnuFileNew_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    ResetControls
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub mnuFileSave_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    SaveStar
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub mnuStarClear_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    pbx1.Cls
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub mnuStarCreate_Click()

    'Error Handling
    On Error GoTo PROC_ERR
    
    CreateStar
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    

End Sub

Private Sub mnuStarRandom_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    RandomValues
    CreateStar
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrBackB_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all BackColor Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrBackG_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all BackColor Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrBackR_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all BackColor Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrForeB_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all ForeColor Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrForeG_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all ForeColor Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrForeR_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all ForeColor Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrShadowB_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all Shadow Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrShadowG_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all Shadow Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub sldrShadowR_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all Shadow Controls
    UpdatefromSldr cmbColourSelect.Text
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub tbr1_ButtonClick(ByVal p_Button As MSComctlLib.Button)
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'See what button was pressed
    Select Case p_Button.Tag
            
        Case "NEW"
            
            'Create a new star
            ResetControls
            
        Case "CREATE"
            
            'Create the Star
            CreateStar
            
        Case "CLEAR"
        
            'Clear the Screen
            pbx1.Cls
            
        Case "SAVE"
            
            'Save the star
            SaveStar
            
        Case "EXIT"
            
            'Exit
            Unload Me
            
        Case "RANDOM"
            
            'Create random values then make star
            RandomValues
            CreateStar
            
    End Select
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub cmbColourSelect_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Declare Variables
    Dim x  As Integer

    'Make all frames invisible
    For x = 0 To 2 Step 1
        fraRGB(x).Visible = False
    Next x
    
    Select Case cmbColourSelect.Text
    
        Case "ForeColour"   'ForeColour was selected
            fraRGB(0).Visible = True
        
        Case "BackColour"   'BackColour was selected
            fraRGB(1).Visible = True
        
        Case "Shadow"       'ShadowColour was selected
            fraRGB(2).Visible = True
        
    End Select
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub optShadow_Click(intShadowIndex As Integer)
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Declare Variables
    Dim x As Integer

    For x = 0 To 7 Step 1
        'Clear all lines
        lnShadow(x).BorderColor = RGB(0, 0, 0)
    Next x
    
    'Change the line colour
    lnShadow(intShadowIndex).BorderColor = RGB(0, 0, 255)
    
    'Make the ciricle tag equal the code of the shadow position, E.g. For bottom-right the code is br
    Select Case intShadowIndex
    
        Case 0
            shpCircle.Tag = "b"
            shpShadow.Top = shpCircle.Top + 120
            shpShadow.Left = shpCircle.Left
            
        Case 1
            shpCircle.Tag = "br"
            shpShadow.Top = shpCircle.Top + 120
            shpShadow.Left = shpCircle.Left + 120
            
        Case 2
            shpCircle.Tag = "r"
            shpShadow.Top = shpCircle.Top
            shpShadow.Left = shpCircle.Left + 120
            
        Case 3
            shpCircle.Tag = "tr"
            shpShadow.Top = shpCircle.Top - 120
            shpShadow.Left = shpCircle.Left + 120
            
        Case 4
            shpCircle.Tag = "t"
            shpShadow.Top = shpCircle.Top - 120
            shpShadow.Left = shpCircle.Left
            
        Case 5
            shpCircle.Tag = "tl"
            shpShadow.Top = shpCircle.Top - 120
            shpShadow.Left = shpCircle.Left - 120
            
        Case 6
            shpCircle.Tag = "l"
            shpShadow.Top = shpCircle.Top
            shpShadow.Left = shpCircle.Left - 120
            
        Case 7
            shpCircle.Tag = "bl"
            shpShadow.Top = shpCircle.Top + 120
            shpShadow.Left = shpCircle.Left - 120
    
    End Select
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub tabProperties_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Make all frames invisible
    fraShadow.Visible = False
    fraColours.Visible = False
    fraAttributes.Visible = False
    
    Select Case tabProperties.SelectedItem.Caption
    
        Case "&Attributes"  'Attributes was selected
            fraAttributes.Visible = True
        
        Case "&Colours"     'Colours was selected
            fraColours.Visible = True
        
        Case "&Shadow"      'Shadow was selected
            fraShadow.Visible = True
        
    End Select
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub cmdCustomFore_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Declare Variables
    Dim lngCol  As Long
    Dim lngR    As Long
    Dim lngG    As Long
    Dim lngB    As Long
    Dim lngX    As Long
    
    'Set values
    lngR = 0
    lngG = 0
    lngB = 0
    
    'Show common dialog
    comdiag.ShowColor
    
    'Get RGB Values
    lngCol = comdiag.Color
    
    If lngCol = 0 And pbx1.BackColor = -2147483643 Then
        lngCol = 255
        lngR = 255
        lngG = 255
        lngB = 255
    End If
    
    For lngX = 1 To 513 Step 1
        
        If lngCol >= 65536 Then
            lngCol = lngCol - 65536
            lngB = lngB + 1
        ElseIf lngCol >= 256 Then
            lngCol = lngCol - 256
            lngG = lngG + 1
        Else
            lngR = lngCol
        End If
        
    Next lngX
    
    
    'Write values
    txtForeR.Text = lngR
    txtForeG.Text = lngG
    txtForeB.Text = lngB
    
    sldrForeR.Value = lngR
    sldrForeG.Value = lngG
    sldrForeB.Value = lngB
    
    picForePreview.BackColor = comdiag.Color
    
    comdiag.Color = 0
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub cmdCustomBack_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Declare Variables
    Dim lngCol  As Long
    Dim lngR    As Long
    Dim lngG    As Long
    Dim lngB    As Long
    Dim lngX    As Long
    
    'Set values
    lngR = 0
    lngG = 0
    lngB = 0
    
    'Show common dialog
    comdiag.ShowColor
    
    'Get RGB Values
    lngCol = comdiag.Color
    
    If lngCol = 0 And pbx1.BackColor = -2147483643 Then
        lngCol = 255
        lngR = 255
        lngG = 255
        lngB = 255
    End If
    
    For lngX = 1 To 513 Step 1
        
        If lngCol >= 65536 Then
            lngCol = lngCol - 65536
            lngB = lngB + 1
        ElseIf lngCol >= 256 Then
            lngCol = lngCol - 256
            lngG = lngG + 1
        Else
            lngR = lngCol
        End If
        
    Next lngX
    
    
    'Write values
    txtBackR.Text = lngR
    txtBackG.Text = lngG
    txtBackB.Text = lngB
    
    sldrBackR.Value = lngR
    sldrBackG.Value = lngG
    sldrBackB.Value = lngB
    
    picBackPreview.BackColor = comdiag.Color
    
    comdiag.Color = 0
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub cmdCustomShadow_Click()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Declare Variables
    Dim lngCol  As Long
    Dim lngR    As Long
    Dim lngG    As Long
    Dim lngB    As Long
    Dim lngX    As Long
    
    'Set values
    lngR = 0
    lngG = 0
    lngB = 0
    
    'Show common dialog
    comdiag.ShowColor
    
    'Get RGB Values
    lngCol = comdiag.Color
    
    If lngCol = 0 And pbx1.BackColor = -2147483643 Then
        lngCol = 255
        lngR = 255
        lngG = 255
        lngB = 255
    End If
    
    For lngX = 1 To 513 Step 1
        
        If lngCol >= 65536 Then
            lngCol = lngCol - 65536
            lngB = lngB + 1
        ElseIf lngCol >= 256 Then
            lngCol = lngCol - 256
            lngG = lngG + 1
        Else
            lngR = lngCol
        End If
        
    Next lngX
    
    
    'Write values
    txtShadowR.Text = lngR
    txtShadowG.Text = lngG
    txtShadowB.Text = lngB
    
    sldrShadowR.Value = lngR
    sldrShadowG.Value = lngG
    sldrShadowB.Value = lngB
    
    picShadowPreview.BackColor = comdiag.Color
    
    comdiag.Color = 0
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtBackB_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all BackColour controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtBackG_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all BackColour controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtBackR_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all BackColour controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtForeB_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all ForeColour controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtForeG_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all ForeColour controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtForeR_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all ForeColour controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtShadowB_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all shadow controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtShadowG_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all shadow controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Private Sub txtShadowR_Change()
    
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Update all shadow controls
    UpdatefromText cmbColourSelect
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Sub

Public Function RandomValues()
            
    'Error Handling
    On Error GoTo PROC_ERR
    
    'Randomize all values
    Randomize
    txtSideLength.Text = Int(Rnd * 800) + 400
    
    Randomize
    If pbx1.Height > pbx1.Width Then
        txtWidth.Text = Int(Rnd * Int(pbx1.Width / 2 - CLng(Val(txtSideLength.Text)) - 500) + 400)
    Else
        txtWidth.Text = Int(Rnd * Int(pbx1.Height / 2 - CLng(Val(txtSideLength.Text)) - 500) + 400)
    End If
    
    Randomize
    txtDensity.Text = Int(Rnd * 27) + 3
    
    Randomize
    txtNumber.Text = Int(Rnd * 12) + 3
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Error"
    Resume PROC_EXIT
    
    
End Function
