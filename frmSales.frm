VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmSales 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7200
   ClientLeft      =   1575
   ClientTop       =   870
   ClientWidth     =   8595
   ControlBox      =   0   'False
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00A27E66&
      Caption         =   "&Member"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00A27E66&
      Caption         =   "&EFT (Cash Card)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4680
      Picture         =   "frmSales.frx":27A2
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   25
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   22
      Top             =   960
      Width           =   2505
   End
   Begin HoverButton.Button Button1 
      Height          =   250
      Left            =   8280
      TabIndex        =   20
      Top             =   110
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483632
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483632
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483630
      Caption         =   "r"
      CaptionDown     =   "r"
      CaptionOver     =   "r"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "WYEIWYG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   19
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Print Invoice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sales Catagory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A27E66&
      Height          =   615
      Left            =   5880
      TabIndex        =   14
      Top             =   1200
      Width           =   2535
      Begin VB.OptionButton BSales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "B Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1440
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton ASales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "A Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox picCombo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   5160
      ScaleHeight     =   780
      ScaleWidth      =   3060
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   3060
      Begin MSForms.ComboBox ComboBox 
         DataField       =   "StdName"
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   2955
         VariousPropertyBits=   1822443547
         BackColor       =   16777215
         DisplayStyle    =   3
         Size            =   "5212;1085"
         ListRows        =   10
         DropButtonStyle =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00A27E66&
      Caption         =   "C&omplimentary"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00A27E66&
      Caption         =   "&Cash Sales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6630
      TabIndex        =   2
      Top             =   6180
      Width           =   1710
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   6840
      Width           =   1245
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFInvoice 
      Height          =   2550
      Left            =   240
      TabIndex        =   1
      Top             =   2820
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   4498
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   14268829
      BackColorSel    =   14268829
      BackColorBkg    =   16777215
      GridColorFixed  =   14268829
      GridColorUnpopulated=   14268829
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin HoverButton.Button Command1 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   4210752
      Caption         =   "Hold [F4]"
      CaptionDown     =   "Hold [F4]"
      CaptionOver     =   "Hold [F4]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Command4 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   4210752
      Caption         =   "Close [Esc]"
      CaptionDown     =   "Close [Esc]"
      CaptionOver     =   "Close [Esc]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Button2 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   4210752
      Caption         =   "Discount [F5]"
      CaptionDown     =   "Discount [F5]"
      CaptionOver     =   "Discount [F5]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Button3 
      Height          =   375
      Left            =   480
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   4210752
      Caption         =   "New [F3]"
      CaptionDown     =   "New [F3]"
      CaptionOver     =   "New [F3]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   5640
      Picture         =   "frmSales.frx":4099
      ScaleHeight     =   975
      ScaleWidth      =   2895
      TabIndex        =   26
      Top             =   480
      Width           =   2895
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Standered Edition"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   29
         Top             =   480
         Width           =   1275
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "[F6] Quick Search"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5880
      TabIndex        =   31
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "[Mohamed Firag]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   30
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Sales "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   24
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8295
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00800000&
      X1              =   8400
      X2              =   8400
      Y1              =   6960
      Y2              =   7080
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00800000&
      X1              =   8400
      X2              =   8400
      Y1              =   6000
      Y2              =   6120
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00800000&
      X1              =   90
      X2              =   8400
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00800000&
      X1              =   2040
      X2              =   8400
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00800000&
      X1              =   2040
      X2              =   8400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00800000&
      X1              =   2040
      X2              =   8400
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00800000&
      X1              =   2040
      X2              =   2040
      Y1              =   6120
      Y2              =   6960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   2040
      X2              =   2040
      Y1              =   840
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   2040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   90
      X2              =   90
      Y1              =   840
      Y2              =   7080
   End
   Begin VB.Label bal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A27E66&
      Height          =   375
      Left            =   7485
      TabIndex        =   18
      Top             =   6510
      Width           =   885
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1920
      Top             =   6000
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   195
      Top             =   2760
      Width           =   8205
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   11
      Top             =   6600
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total For This Bill:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   5640
      Width           =   1260
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxx.xxx MRf"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A27E66&
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   5490
      Width           =   2370
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount Tandered:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4800
      TabIndex        =   6
      Top             =   6240
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   6255
      Left            =   105
      Top             =   840
      Width           =   1950
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1920
      Top             =   6960
      Width           =   6495
   End
   Begin VB.Menu optDeletes 
      Caption         =   "&Deletes"
      Visible         =   0   'False
      Begin VB.Menu optDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu optdiscount 
         Caption         =   "&Set Discount"
         Enabled         =   0   'False
      End
      Begin VB.Menu prp 
         Caption         =   "&Properties"
      End
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ===============================================================================//
'// Program: OpenPOS(Point of Sales)                                               //
'// Developed by: Sofwathullah Mohamed                                             //
'// Sofwath@Hotmail.Com                                                            //
'// You are free to use and modify this program as long as you give credit to the  //
'// original developer. Any comments or bugs report to sofwath@hotmail.com         //
'// Ver: 0.1                                                                       //
'// This Program is Still Under Development and Some of the Modules are Missing    //
'// ===============================================================================//

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hnd As Long, ByVal clval As Long, ByVal alph As Byte, ByVal flago As Long) As Long
Dim EditinG As Boolean
Dim TempHold As String
Public LastRow As Integer
Private Sub Button1_Click()
    Unload Me
End Sub
Private Sub Button2_Click()
    'frmDiscount.Show 1
    CalculateDiscount
End Sub
Private Sub Button3_Click()
'new
    SetTextBox2Grid
    Form_Load
    MSHFInvoice.Row = 1
    MSHFInvoice.Col = 0
    MSHFInvoice_Click
    DoEvents
End Sub
Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Text2.Enabled = True
        Text2.SetFocus
    Else
        Text2.Enabled = False
        Text2.Text = Empty
        Label8.Visible = False
        MSHFInvoice.Col = 0
        MSHFInvoice_Click
        DoEvents
    End If
End Sub
Private Sub ComboBox_GotFocus()
    ComboBox.Visible = True
    picCombo.Visible = True
End Sub
Private Sub ComboBox_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            '// done with desc
            If MSHFInvoice.Row <> 1 Then
                MSHFInvoice.Col = 0
                cr = MSHFInvoice.Row - 1
                For r = cr To 1 Step -1
                    MSHFInvoice.Row = r
                    If Not MSHFInvoice.Text = Empty Then '// move up
                        MSHFInvoice.Row = r + 1
                        Exit For
                    End If
                Next
            End If
            If Not ComboBox.Text = Empty Then
                DataEnvironment1.Recordsets(1).Filter = "description = '" & Trim(ComboBox.Text) & "'"
                If DataEnvironment1.Recordsets(1).RecordCount = 1 Then
                    MSHFInvoice.Col = 0
                    MSHFInvoice.Text = DataEnvironment1.Recordsets(1)!code
                    MSHFInvoice.Col = 1
                    MSHFInvoice.Text = DataEnvironment1.Recordsets(1)!Description
                    MSHFInvoice.Col = 3
                    If BSales = True Then
                        MSHFInvoice.Text = DataEnvironment1.Recordsets(1)!B_Price
                    Else
                        MSHFInvoice.Text = DataEnvironment1.Recordsets(1)!A_price
                    End If
                    MSHFInvoice.Col = 2
                    Grid_EnterCell
                End If
            End If
        Case vbKeyUp
            ComboBox.DropDown
        Case vbKeyDown
            ComboBox.DropDown
        Case vbKeyEscape
            MSHFInvoice.SetFocus
        Case vbKeyLeft
            MSHFInvoice.Col = 0
            Call SetTextBox2Grid
        Case vbKeyRight
            MSHFInvoice.Col = 2
            Call SetTextBox2Grid
    End Select
    'Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub ComboBox_KeyPress(KeyAscii As MSForms.ReturnInteger)
'// key press
    Screen.MousePointer = vbHourglass
    ComboBox.Clear
    ComboBox.DropDown
    With DataEnvironment1
        If Not .Recordsets(1).State = 1 Then
            .Recordsets(1).Open
        End If
        If Not (Chr(KeyAscii)) = Empty Then
            .Recordsets(1).Filter = "description LIKE '*" & Trim((ComboBox.Text) & (Chr$(KeyAscii))) & "*'"
        Else
            .Recordsets(1).Filter = "description LIKE '*" & Trim(ComboBox.Text) & "*'"
        End If
        'search for matching items
        If .Recordsets(1).RecordCount > 0 Then
            .Recordsets(1).MoveFirst
            While Not .Recordsets(1).EOF
                ComboBox.AddItem (.Recordsets(1)!Description)
                .Recordsets(1).MoveNext
            Wend
        End If
    End With
    Screen.MousePointer = vbDefault
End Sub
Private Sub ComboBox_LostFocus()
    ComboBox.Visible = False
    picCombo.Visible = False
End Sub
Private Sub Command1_Click()
    Dim newSales As New frmSales
    indexSales = indexSales + 1
    'newSales.Caption = "Sales (Hold #:" & Forms.Count & ")"
    newSales.Show
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Call SetTextBox2Grid
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Button3_Click
        Case vbKeyF4
            Command1_Click
        Case vbKeyF5
            Button2_Click
        Case vbKeyF6
            frmItemSearch.Show 1
        Case vbKeyEscape
            Command4_Click
    End Select
End Sub

Private Sub Form_Load()
'// initilize form and do the setup
    With MSHFInvoice
        .Clear
        .Cols = 5
        .RowHeightMin = 320
        .Row = 0
        .Col = 0: .Text = "Code"
        .Col = 1: .Text = "Description"
        .Col = 2: .Text = "Qty"
        .Col = 3: .Text = "Rate"
        .Col = 4: .Text = "Total"
        .ColWidth(1) = 3900
        .ColWidth(3) = 1000
        .Rows = 200
    End With
    Text9.Text = Empty
    Text9.Visible = False
    picCombo.Visible = False
    Total.Caption = "00.00"
    bal.Caption = "00.00"
'    DTPicker1.Value = Now
    ASales.Value = False
    BSales.Value = True
    Check3.Value = 0
    Check2.Enabled = True
    Text2.Text = Empty
    Text1.Text = Empty
'    Combo1.Enabled = False
    Text2.Enabled = False
    'SelfService.Value = True
    Option1.Value = True ' make default cash sales
    '--
    Label8.Visible = False
    MSHFInvoice.Col = 0
    MSHFInvoice_Click
    DoEvents
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Private Sub MSHFInvoice_Click()
'// row selected so take action based on it
    Select Case MSHFInvoice.Col
        Case 0 'item code
            Call SetTextBox2Grid
        Case 1 'item desc
            Call SetComboBox2Grid
        Case 2 'qty
            Call SetTextBox2Grid
    End Select
    LastRow = MSHFInvoice.Row
    '-- tool tip
    Dim c As Integer, r As Integer
    c = MSHFInvoice.Col
    r = MSHFInvoice.Row
    DoEvents
    MSHFInvoice.ToolTipText = Empty
    MSHFInvoice.Col = 0
    If Not MSHFInvoice.Text = Empty Then
        MSHFInvoice.ToolTipText = "Max Possible Discount For Item " & MSHFInvoice.Text & " Is MRf " & ReturnMaxDiscountP(MSHFInvoice.Text) & " Per Unit"
    End If
    MSHFInvoice.Col = c
    MSHFInvoice.Row = r
End Sub
Private Sub SetComboBox2Grid()
'// set the combo box to the grid
    With picCombo
        .Move (MSHFInvoice.CellLeft + MSHFInvoice.Left), _
        MSHFInvoice.CellTop + MSHFInvoice.Top, MSHFInvoice.CellWidth - 25, MSHFInvoice.CellHeight - 25
        ComboBox.Height = 300
        ComboBox.Width = MSHFInvoice.CellWidth - 25
        ComboBox.Text = MSHFInvoice.Text
        .ZOrder 0
        .BackColor = MSHFInvoice.CellBackColor
        ComboBox.BackColor = MSHFInvoice.CellBackColor
        ComboBox.Visible = True
        .Visible = True
        ComboBox.SetFocus
    End With
End Sub
Private Sub SetTextBox2Grid()
'// set text box to grid
On Error Resume Next
    With Text9
        .Move MSHFInvoice.CellLeft + MSHFInvoice.Left, _
        MSHFInvoice.CellTop + MSHFInvoice.Top, MSHFInvoice.CellWidth - 25, MSHFInvoice.CellHeight - 25
        .Text = MSHFInvoice.Text
        If Not .Text = Empty Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
        .ZOrder 0
        .BackColor = MSHFInvoice.CellBackColor
        .Visible = True
        .SetFocus
    End With
End Sub
Private Sub MSHFInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call Grid_EnterCell
    End Select
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub MSHFInvoice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'to do mouse right click, to delete selected row
    If Button = vbRightButton Then ' right click
        PopupMenu optDeletes
    End If
End Sub
Private Sub optDelete_Click()
    'delete
    Dim r As Integer, rr As Integer
    Dim LastRowPlusOne As Integer
    Dim TStr As String
    
    With MSHFInvoice
        .Row = LastRow
        LastRowPlusOne = LastRow + 1
        'first delete
        For r = 0 To 4
            .Col = r
            .Text = Empty
        Next
        For rr = LastRowPlusOne To .Rows
            .Col = 0
            .Row = rr
            If .Text <> Empty Then
                For r = 0 To 4
                    .Row = rr
                    .Col = r
                    TStr = .Text
                    .Text = Empty
                    .Row = rr - 1
                    .Text = TStr
                Next
            Else
                Exit For
            End If
        Next
        .Row = LastRow - 1
    End With
    Call DoTotal
End Sub
Private Sub Option1_Click()
'    Combo1.Enabled = False
    MSHFInvoice.Col = 0
    MSHFInvoice_Click
    DoEvents
End Sub
Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Option2_Click()
    TempHold = InputBox("Details", "Complementry Sales Details")
'    Combo1.Enabled = False
End Sub
Private Sub Option4_Click()
    If Option4.Value = True Then
        Text2.Enabled = True
        Text2.SetFocus
    End If
End Sub
Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Option3_Click()
    MSHFInvoice.Col = 0
    MSHFInvoice_Click
    DoEvents
End Sub
Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub picCombo_LostFocus()
    picCombo.Visible = False
End Sub
Private Sub prp_Click()
    '--
    With MSHFInvoice
        .Col = 0
        If Not .Text = Empty Then
            Call MakeChildTrans(230, frmItm)
            frmItm.Label18.Caption = "View Item"
            frmItm.EditItem (.Text)
            frmItm.Label9.Visible = False
            frmItm.Text8(0).Visible = False
            frmItm.Show 1
        End If
    End With
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Tot As Currency
    Dim Paid As Currency
    Dim Balance As Currency
    If KeyAscii = vbKeyReturn Then
        Tot = CDbl(Total.Caption)
        If Tot > 0 And (Val(Text1.Text) > 0) Then
            Paid = CDbl(Text1.Text)
            Text1.Text = Format(Paid, "##,##0.00")
            Balance = Paid - Total
            If Balance >= 0 Then
                bal.Caption = Format(Balance, "##,##0.00")
                'ok done now need to do the db stuff
            Else
                'Text4.Text = Empty
                Beep
                Text1.Text = "00.00"
                Text1.SelStart = 0
                Text1.SelLength = Len(Text1.Text)
                Text1.SetFocus
            End If
        End If
    End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim mNam As String
        If Not Text2.Text = Empty Then
            mNam = GetMemName(Text2.Text)
            If Not mNam = Empty Then
                Label8.Caption = "Name :[ " & mNam & " ]"
                Label8.Visible = True
                MSHFInvoice.Col = 0
                MSHFInvoice_Click
            Else
                Text2.Text = Empty
                Label8.Visible = False
                Check3.Value = 0
            End If
        End If
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text2_LostFocus()
'   chk mem
    Dim mNam As String
    If Not Text2.Text = Empty Then
        mNam = GetMemName(Text2.Text)
        If Not mNam = Empty Then
            Label8.Caption = "Name :[ " & mNam & " ]"
            Label8.Visible = True
            MSHFInvoice.Col = 0
            MSHFInvoice_Click
        Else
            Text2.Text = Empty
            Label8.Visible = False
            Check3.Value = 0
        End If
    End If
End Sub
Public Function GetMemName(memNo As String) As String
    On Error Resume Next
    GetMemName = Empty
    With DataEnvironment1.rsmem
        If Not .State = adStateClosed Then .Close
        .Open "Select * From Memberz Where code = '" & memNo & "'"
        If Not .AbsolutePosition = adPosEOF Then
            GetMemName = !Name
        End If
    End With
End Function
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    'On Error Resume Next
    Dim qty As Double
    Dim Rate As Double
    Dim Total As Currency
    Dim tRow As Integer, tCol As Integer, ttRow As Integer
    Dim r As Integer, cr As Integer, rr As Integer
    Dim qtyOK As Boolean
    Select Case KeyCode
        Case vbKeyReturn
            Select Case MSHFInvoice.Col
                Case 0
                    If Not DataEnvironment1.Recordsets(1).State = 1 Then
                        DataEnvironment1.Recordsets(1).Open
                    End If
                    DataEnvironment1.Recordsets(1).Filter = "code ='" & Text9.Text & "'"
                    If DataEnvironment1.Recordsets(1).RecordCount = 1 Then
                        If MSHFInvoice.Row <> 1 Then '// move up part
                            MSHFInvoice.Col = 0
                            cr = MSHFInvoice.Row - 1
                            For r = cr To 1 Step -1
                                MSHFInvoice.Row = r
                                If Not MSHFInvoice.Text = Empty Then
                                    MSHFInvoice.Row = r + 1
                                    Exit For
                                End If
                            Next
                        End If
                        MSHFInvoice.Col = 1
                        MSHFInvoice.Text = DataEnvironment1.Recordsets(1)!Description
                        MSHFInvoice.Col = 3
                        If BSales = True Then
                            MSHFInvoice.Text = DataEnvironment1.Recordsets(1)!B_Price
                        Else
                            MSHFInvoice.Text = DataEnvironment1.Recordsets(1)!A_price
                        End If
                        MSHFInvoice.Col = 0
                        MSHFInvoice.Text = Text9.Text
                        MSHFInvoice.Col = 2
                        Grid_EnterCell
                    Else
                        'MsgBox "Invalid Code, Use Description to Locate or Enter Code Again", vbInformation
                        tRow = MSHFInvoice.Row
                        tCol = MSHFInvoice.Col
                        frmMessage.Show 1
                        frmMessage.Timer1.Enabled = True
                        Grid_EnterCell
                        Text9.SelStart = 0
                        Text9.SelLength = Len(Text9.Text)
                    End If
                Case 2
                    If Not Text9.Text = Empty Then ' some validation checks
                        qty = Val(Text9.Text)
                    Else
                        qty = 1
                    End If
                    If qty <= 1 Then qty = 1
                    With MSHFInvoice
                        .Text = qty
                        .Col = 3
                        If Not .Text = Empty Then ' validate (do nothing if no items)
                            Rate = CDbl(.Text)
                            Total = qty * Rate
                            .Col = 4
                            .Text = Format(Total, "##,##0.00")
                            .Col = 0
                            qtyOK = IsAvailable(.Text, qty)
                            If qtyOK = False Then
                                ttRow = .Row
                                DoMsg ("Quantity Not In Hand")
                                For rr = 0 To 4
                                    .Col = rr
                                    .Text = Empty
                                Next
                                Call optDelete_Click
                                MSHFInvoice.Col = 0
                                MSHFInvoice.Row = ttRow
                                MSHFInvoice_Click
                                Exit Sub
                            End If
                        Else
                            .Col = 2
                            .Text = Empty
                        End If
                        If Rate > 0 Then
                            'MSHFInvoice.Row = MSHFInvoice.Row + 1 'change row (if only data in grid)
                            Call DoTotal
                        End If
                        .Col = 0
                        'If Check2.Value = 0 Then Call LoopForRepAndIncrement(.Text, .Row)
                        Call LoopForRepAndIncrement(.Text, .Row)
                        .Row = .Row + 1
                        .Col = 0
                    End With
                    Call Grid_EnterCell
                Case Else
            End Select
        Case vbKeyLeft
            'left
            With MSHFInvoice
                Select Case .Col
                    Case 0
                        If .Row > 1 Then
                            .Text = Text9.Text
                            .Col = 2
                            .Row = .Row - 1
                            KeyCode = 0
                            Call SetTextBox2Grid
                        End If
                    Case 2
                        With Text9
                            If Not .Text = Empty Then
                                .Text = .Text
                            End If
                            .Visible = False
                            .Text = Empty
                        End With
                        .Col = 1
                        Call SetComboBox2Grid
                    Case Else
                End Select
            End With
        Case vbKeyRight
            Select Case MSHFInvoice.Col
                Case 0
                    MSHFInvoice.Col = 1
                    Call SetComboBox2Grid
                Case 2
                    MSHFInvoice.Col = 0
                    MSHFInvoice.Row = MSHFInvoice.Row + 1
                    Call SetTextBox2Grid
            End Select
        Case vbKeyUp
            If Not MSHFInvoice.Row = 1 Then
                MSHFInvoice.Row = MSHFInvoice.Row - 1
                Call SetTextBox2Grid
            End If
        Case vbKeyDown
            If Not MSHFInvoice.Row = MSHFInvoice.Rows - 1 Then
                MSHFInvoice.Row = MSHFInvoice.Row + 1
                Call SetTextBox2Grid
            End If
        Case Else
    End Select
    Call Form_KeyDown(KeyCode, Shift)
    KeyCode = 0
End Sub
Private Sub Grid_EnterCell()
    '// when click on cell
    Dim tCol As Integer
    tCol = MSHFInvoice.Col
    Select Case tCol
        Case 0, 2
            With Text9
                .Move MSHFInvoice.CellLeft + MSHFInvoice.Left, _
                MSHFInvoice.CellTop + MSHFInvoice.Top, MSHFInvoice.CellWidth - 25, MSHFInvoice.CellHeight - 25
                .Text = MSHFInvoice.Text
                If Len(.Text) > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End If
                .Visible = True
                .ZOrder 0
                .SetFocus
            End With
    End Select
End Sub
Private Sub DoTotal()
'// calculate total
    Dim BillTotal As Currency
    Dim CurrentRow As Integer
    Dim CurrentCol As Integer
    Dim r As Integer
    Dim rCount As Integer
    If EditinG = False Then 'check if workin then disable change of price and sales type
        EditinG = True
        Frame1.Enabled = False
        Check2.Enabled = False
        'Frame2.Enabled = False
    End If
    With MSHFInvoice
        CurrentRow = .Row
        CurrentCol = .Col
        rCount = .Rows - 1
        .Col = 4
        For r = 1 To rCount
            .Row = r
            If Not .Text = Empty Then
                BillTotal = BillTotal + CDbl(.Text)
            End If
        Next r
        Total.Caption = Format(BillTotal, "##,##0.00") ' currency format
        .Col = CurrentCol
        .Row = CurrentRow
    End With
End Sub
Private Sub Text9_LostFocus()
    Text9.Visible = False
End Sub
Private Sub LoopForRepAndIncrement(icode As String, crow As Integer)
    'check rapatition
    If crow = 1 Then Exit Sub
    Dim r, rr, rrr As Integer
    Dim qty As Double
    Dim Rate As Double
    Dim Total As Currency
    Dim founD As Boolean
    Dim qtyOK As Boolean
    founD = False
    crow = crow - 1
    With MSHFInvoice
        For r = crow To 1 Step -1
            .Row = r
            .Col = 0
            If .Text = icode Then
                .Col = 2
                .Row = crow + 1
                qty = CDbl(.Text)
                .Row = r
                .Text = CDbl(.Text) + qty
                '---total up (new)
                qty = CDbl(.Text)
                '--
                qtyOK = IsAvailable(icode, qty)
                If qtyOK = False Then
                    '// daa now wot
                    DoMsg ("Quantity Not In Hand")
                    For rrr = 0 To 4
                        .Col = rrr
                        .Text = Empty
                    Next
                    'Call optDelete_Click
                    'Grid_EnterCell
                    'Text9.SelStart = 0
                    'Text9.SelLength = Len(Text9.Text)
                '--
                Else
                    .Col = 3
                    Rate = CDbl(.Text)
                    Total = qty * Rate
                    .Col = 4
                    .Text = Format(Total, "##,##0.00")
                End If
                '---
                .Row = crow + 1
                For rr = 0 To 4
                    .Col = rr
                    .Text = Empty
                Next
                .Row = crow
                Call DoTotal
                founD = True
                Exit For
            End If
        Next
        If Not founD Then
            .Row = crow + 1
        End If
    End With
End Sub
Private Sub Timer1_Timer()
    Label2.Caption = Now
End Sub
Private Sub CalculateDiscount()
    '--maxDiscountP
    Dim BillTotal As Currency
    Dim CurrentRow As Integer
    Dim CurrentCol As Integer
    Dim r As Integer
    Dim rCount As Integer
    Dim cc As Double
    Dim rr As Double
    Dim pp As Double
    With MSHFInvoice
        CurrentRow = .Row
        CurrentCol = .Col
        
        rCount = .Rows - 1
        .Col = 0
        For r = 1 To rCount
            .Row = r
            If Not .Text = Empty Then
                'BillTotal = BillTotal + CDbl(ReturnMaxDiscountP(.Text))
                pp = IIf(Check3.Value = 0, CDbl(ReturnMaxDiscountP(.Text)), CDbl(ReturnMemDiscountP(.Text)))
                .Col = 2
                BillTotal = BillTotal + (CDbl(.Text) * pp)
                .Col = 0
            End If
        Next r
        If BillTotal <= 0 Then Exit Sub
        cc = (CDbl(Total.Caption) / 100)
        rr = BillTotal / cc
        'Total.Caption = Format(BillTotal, "##,##0.00") ' currency format
        .Col = CurrentCol
        .Row = CurrentRow
    End With
    frmDiscount.Label7.Caption = "(Max. % is " & Format(rr, "##.##") & "%)"
    frmDiscount.Label5.Caption = Format(BillTotal, "##,##0.00")
    frmDiscount.Tag = rr
    frmDiscount.bal.Caption = Total.Caption
    If Check3.Value = 1 Then
        frmDiscount.Label6.Visible = False
        frmDiscount.Label7.Visible = False
        frmDiscount.Label8.Visible = False
        frmDiscount.Text1.Visible = False
        frmDiscount.Text2.Visible = False
        frmDiscount.Label2.Caption = "Member Discount:"
        frmDiscount.Label9.Caption = Format((CDbl(frmDiscount.bal.Caption) - CDbl(frmDiscount.Label5.Caption)), "##,##0.00")
    End If
    frmDiscount.Show 1
    '--
End Sub
Private Function ReturnMaxDiscountP(icode As String) As Double
    '-
    On Error Resume Next
    Dim peC As Double
    Dim prC As Double
    ReturnMaxDiscountP = 0
    DataEnvironment1.Recordsets(1).Filter = "code ='" & icode & "'"
    If DataEnvironment1.Recordsets(1).RecordCount = 1 Then
        peC = DataEnvironment1.Recordsets(1)!maxDiscountP
        prC = DataEnvironment1.Recordsets(1)!B_Price
        ReturnMaxDiscountP = (prC / 100) * peC
    End If
    Debug.Print icode, ReturnMaxDiscountP
End Function
Private Function ReturnMemDiscountP(icode As String) As Double
    '-memberz discount (yeah i know could have implemented this in returnmaxdiscountP 2, but.., i have other future plans)
    Dim peC As Double
    Dim prC As Double
    ReturnMemDiscountP = 0
    DataEnvironment1.Recordsets(1).Filter = "code ='" & icode & "'"
    If DataEnvironment1.Recordsets(1).RecordCount = 1 Then
        peC = DataEnvironment1.Recordsets(1)!memDiscountP
        prC = DataEnvironment1.Recordsets(1)!B_Price
        ReturnMemDiscountP = (prC / 100) * peC
    End If
End Function
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not Text1.Text = Empty Then
            If Not bal.Caption = "00.00" Then
                DoSalesDBNoDis
                Button3_Click
            End If
        End If
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Public Sub DoSalesDBNoDis()
    'sales write
    Dim rCount As Double
    Dim bnO As String
    Dim r As Integer, rr As Integer
    Dim itm As String, qty As Integer
    With DataEnvironment1.rsdailysales
        If Not .State = adStateClosed Then .Close
        .Open "Select * from daily_sales"
        .AddNew
            bnO = getNextBillNo
            !bill = bnO
            !BDate = Date
            !Time = Time
            !Total = CDbl(Total.Caption)
            !Complementry = False
            If Option1.Value = True Then !CashSales = True
            !DiscountedSales = False
            !BSales = IIf(BSales.Value = True, True, False)
            !memberSales = IIf(Check3.Value = 1, True, False)
            !member = Text2.Text
            !EFT = IIf(Option3.Value = True, True, False)
        .Update
    End With
    '--
    With DataEnvironment1.rssalesitems
        If Not .State = adStateClosed Then .Close
        .Open "Select * From sales_items"
    End With
    With MSHFInvoice
        rCount = .Rows - 1
        For r = 1 To rCount
            .Row = r
            .Col = 0
            If Not .Text = Empty Then
                DataEnvironment1.rssalesitems.AddNew
                DataEnvironment1.rssalesitems!Item = .Text
                itm = .Text
                .Col = 2
                DataEnvironment1.rssalesitems!qty = CDbl(.Text)
                qty = .Text
                .Col = 3
                DataEnvironment1.rssalesitems!price = CDbl(.Text)
                DataEnvironment1.rssalesitems!billNo = bnO
                DataEnvironment1.rssalesitems.Update
                Call minusStock(itm, qty)
            End If
        Next r
    End With
End Sub
Public Function getNextBillNo() As String
    On Error Resume Next
    Dim billNos As Double
    'generate next bill
    With DataEnvironment1.rsbillno
        If Not .State = adStateClosed Then .Close
        .Open "Select * From billno"
        billNos = (!lastbillno + 1)
        getNextBillNo = !BillSafix & Str(billNos) & !Yearz
        !lastbillno = billNos
        .Update
    End With
End Function
Public Sub DoSalesDBWithDis()
    'sales write
    Dim rCount As Double
    Dim bnO As String
    Dim r As Integer, rr As Integer
    Dim itm As String, qty As Integer
    With DataEnvironment1.rsdailysales
        If Not .State = adStateClosed Then .Close
        .Open "Select * from daily_sales"
        .AddNew
            bnO = getNextBillNo
            !bill = bnO
            !BDate = Date
            !Time = Time
            !Total = CDbl(frmDiscount.Label9.Caption)
            !Complementry = False
            If Option1.Value = True Then !CashSales = True
            !DiscountedSales = True
            !BSales = IIf(BSales.Value = True, True, False)
            !memberSales = IIf(Check3.Value = 1, True, False)
            !member = Text2.Text
            !EFT = IIf(Option3.Value = True, True, False)
            !DiscountGiven = (CDbl(frmDiscount.bal.Caption)) - (CDbl(frmDiscount.Label9.Caption))
            !Totalb4dis = CDbl(frmDiscount.bal.Caption)
        .Update
    End With
    '--
    With DataEnvironment1.rssalesitems
        If Not .State = adStateClosed Then .Close
        .Open "Select * From sales_items"
    End With
    With MSHFInvoice
        rCount = .Rows - 1
        For r = 1 To rCount
            .Row = r
            .Col = 0
            If Not .Text = Empty Then
                DataEnvironment1.rssalesitems.AddNew
                DataEnvironment1.rssalesitems!Item = .Text
                itm = .Text
                .Col = 2
                DataEnvironment1.rssalesitems!qty = CDbl(.Text)
                qty = .Text
                .Col = 3
                DataEnvironment1.rssalesitems!price = CDbl(.Text)
                DataEnvironment1.rssalesitems!billNo = bnO
                DataEnvironment1.rssalesitems.Update
                Call minusStock(itm, qty)
            End If
        Next r
    End With
    Button3_Click
End Sub
Private Function IsAvailable(icode As String, iqty As Double) As Boolean
    Dim qq As Double
    IsAvailable = False
    DataEnvironment1.Recordsets(1).Filter = "code ='" & icode & "'"
    If DataEnvironment1.Recordsets(1).RecordCount = 1 Then
        qq = DataEnvironment1.Recordsets(1)!Inhand
        If qq >= iqty Then
            IsAvailable = True
        End If
    End If
End Function
Private Sub minusStock(icode As String, iqty As Integer)
    Dim qq As Integer
    Dim newQty As Integer
    DataEnvironment1.Recordsets(1).Filter = "code ='" & icode & "'"
    If DataEnvironment1.Recordsets(1).RecordCount = 1 Then
        qq = DataEnvironment1.Recordsets(1)!Inhand
        newQty = qq - iqty
        DataEnvironment1.Recordsets(1)!Inhand = newQty
        DataEnvironment1.Recordsets(1).Update
    End If
End Sub
