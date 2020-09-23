VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmItm 
   BorderStyle     =   0  'None
   Caption         =   "Items "
   ClientHeight    =   5775
   ClientLeft      =   990
   ClientTop       =   855
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   7
      Left            =   7320
      TabIndex        =   14
      Text            =   "0"
      ToolTipText     =   "Total Cost For The Item"
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   6
      Left            =   7320
      TabIndex        =   9
      ToolTipText     =   "Total Cost For The Item"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   5
      Left            =   5160
      TabIndex        =   8
      Text            =   "20"
      ToolTipText     =   "Total Cost For The Item"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   4
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "Total Cost For The Item"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   3
      Left            =   5160
      TabIndex        =   6
      ToolTipText     =   "Total Cost For The Item"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   2
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Total Cost For The Item"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Total Cost For The Item"
      Top             =   2160
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmItm.frx":0000
      DataField       =   "Type"
      DataMember      =   "types"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Type"
      Text            =   "DataCombo1"
      Object.DataMember      =   "types"
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
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Inactive"
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
      Height          =   255
      Left            =   7200
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text11 
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
      Height          =   1095
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Text            =   "frmItm.frx":0025
      Top             =   4440
      Width           =   5895
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      DataField       =   "A_price"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Index           =   0
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Total Cost For The Item"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "description"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   1200
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin HoverButton.Button Button1 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4800
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
      Caption         =   "Done [F3]"
      CaptionDown     =   "Done [F3]"
      CaptionOver     =   "Done [F3]"
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
      Left            =   120
      TabIndex        =   18
      Top             =   5280
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
      Caption         =   "Cancel [Esc]"
      CaptionDown     =   "Cancel [Esc]"
      CaptionOver     =   "Cancel [Esc]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Bindings        =   "frmItm.frx":002C
      DataField       =   "Catogory"
      DataMember      =   "catogory"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "Catogory"
      Text            =   "DataCombo1"
      Object.DataMember      =   "catogory"
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
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frmItm.frx":0054
      DataField       =   "brand"
      DataMember      =   "brands"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   4800
      TabIndex        =   12
      Top             =   3480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "brand"
      Text            =   "DataCombo1"
      Object.DataMember      =   "brands"
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "frmItm.frx":007A
      DataField       =   "color"
      DataMember      =   "colors"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "color"
      Text            =   "DataCombo1"
      Object.DataMember      =   "colors"
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
   Begin MSDataListLib.DataCombo DataCombo5 
      Bindings        =   "frmItm.frx":00A0
      DataField       =   "origin"
      DataMember      =   "origins"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   4800
      TabIndex        =   13
      Top             =   3960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "origin"
      Text            =   "DataCombo1"
      Object.DataMember      =   "origins"
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
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   8400
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   8400
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00A27E66&
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1440
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   8535
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   37
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "New Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Memo"
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
      Left            =   1680
      TabIndex        =   35
      Top             =   4440
      Width           =   420
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Origin"
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
      Left            =   4080
      TabIndex        =   34
      Top             =   3960
      Width           =   420
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Color"
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
      Left            =   1680
      TabIndex        =   33
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Size"
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
      Left            =   6720
      TabIndex        =   32
      Top             =   3480
      Width           =   285
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Brand"
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
      Left            =   4080
      TabIndex        =   31
      Top             =   3480
      Width           =   420
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Type"
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
      Left            =   1680
      TabIndex        =   30
      Top             =   3480
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mem. Dis. %"
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
      Left            =   3960
      TabIndex        =   29
      Top             =   2640
      Width           =   915
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cost (Unit)"
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
      Left            =   1680
      TabIndex        =   28
      Top             =   1680
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max. Dis. %"
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
      Left            =   3960
      TabIndex        =   27
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "RO Level"
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
      Left            =   6240
      TabIndex        =   25
      Top             =   1680
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "On Hand"
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
      Left            =   3960
      TabIndex        =   24
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category"
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
      Left            =   4920
      TabIndex        =   23
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Price(B)"
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
      Left            =   1680
      TabIndex        =   22
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Min. Price(A)"
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
      Left            =   1680
      TabIndex        =   21
      Top             =   2640
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description"
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
      Left            =   1680
      TabIndex        =   20
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Code"
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
      Left            =   1680
      TabIndex        =   19
      Top             =   720
      Width           =   750
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   1440
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   72
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   15
      TabIndex        =   26
      Top             =   2280
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmItm"
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

Private Sub Button1_Click()
'--Save if valid
'MsgBox "ok"
    Dim c As Boolean
    If Label18.Caption = "New Item" Then
        c = IsValidItm(False)
    ElseIf Label18.Caption = "Edit Item" Then
        c = IsValidItm(True)
    End If
    If c = True Then
        'add it
        With DataEnvironment1.rsCommand3
            !code = Text1.Text
            !Description = Text2.Text
            !CostPrice = Text8(0).Text
            !B_Price = Text8(1).Text
            !A_price = Text8(2).Text
            !Inhand = Text8(3).Text
            !maxDiscountP = Text8(4).Text
            !memDiscountP = Text8(5).Text
            !RLevel = Text8(6).Text
            !Size = Text8(7).Text
            !Memo = Text11.Text
            !CatCode = DataCombo2.Text
            !Type = DataCombo1.Text
            !Color = DataCombo4.Text
            !brand = DataCombo3.Text
            !origin = DataCombo5.Text
            !inactive = Check1.Value
            If Not Label18.Caption = "View Item" Then 'update only on add and edit
                .Update
                Call BatchUpdateThingy
            End If
        End With
    End If
    Unload Me
End Sub
Private Sub Command4_Click()
    If Label18.Caption = "New Item" Then
        DataEnvironment1.rsCommand3.CancelUpdate ' cancel and pending trans
    End If
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Button1_Click
        Case vbKeyEscape
            Command4_Click
    End Select
End Sub
Private Sub Form_Load()
    'AddNewItem
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text11_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Public Sub AddNewItem()
    With DataEnvironment1.rsCommand3
        If Not .State = adStateClosed Then .Close
        .Open "Select * From items"
        .AddNew
    End With
End Sub
Public Sub EditItem(itmCode As String)
    On Error Resume Next
    If Not DataEnvironment1.rsCommand3.State = adStateClosed Then DataEnvironment1.rsCommand3.Close
    DataEnvironment1.rsCommand3.Open "Select * From items Where code='" & itmCode & "'"
    '--
    With DataEnvironment1.rsCommand3
            Text1.Text = !code
            Text2.Text = !Description
            Text8(0).Text = !CostPrice
            Text8(1).Text = !B_Price
            Text8(2).Text = !A_price
            Text8(3).Text = !Inhand
            Text8(4).Text = !maxDiscountP
            Text8(5).Text = !memDiscountP
            Text8(6).Text = !RLevel
            Text8(7).Text = !Size
            Text11.Text = !Memo
            DataCombo2.Text = !CatCode
            DataCombo1.Text = !Type
            DataCombo4.Text = !Color
            DataCombo3.Text = !brand
            DataCombo5.Text = !origin
            Check1.Value = !inactive
    End With
    Text8(3).Enabled = False
End Sub
Private Function IsValidItm(EditinG As Boolean) As Boolean
    IsValidItm = True ' be positive ;)
    Dim i As Integer
    If Len(Text1.Text) > 1 Then       ' chk if itm is already in list
        DataEnvironment1.rsCommand1.Filter = "code='" & Text1.Text & "'"
        If DataEnvironment1.rsCommand1.RecordCount > 0 And Not EditinG Then ' yeah in
            Call DoMsg("Item Code Already Registered")
            IsValidItm = False
            Exit Function
        End If
        For i = 0 To Me.Count - 1 ' this is a slow method but who cares abt this kinda speed in MS stuff
            If TypeOf Me(i) Is TextBox Or TypeOf Me(i) Is DataCombo Then
                If Len(Me(i).Text) < 1 Then
                    Call DoMsg("No Fields Should Be Empty")
                    IsValidItm = False
                    Exit Function
                End If
            End If
        Next
    Else
        Call DoMsg("Item Code Missing")
        IsValidItm = False
    End If
End Function
Private Sub Text8_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
'--validation chk
On Error Resume Next
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(".")
            KeyAscii = IIf(Index = 3 Or Index = 6, 0, KeyAscii)
        Case vbKeyBack
            KeyAscii = vbKeyBack
        Case KeyAscii = vbKeyDelete
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub BatchUpdateThingy()
    With DataEnvironment1.rsbatch
        If Not .State = adStateClosed Then .Close
        .Open "Select * From batch"
        .AddNew
        '--to do : add user
        !icode = Text1.Text
        !qty = Text8(3).Text
        !batch = 1
        !Date = Now
        !prevqty = 0
        !totalqty = Text8(3).Text
        !Add = True
        !transtype = "Openning Stock"
        !cost = Text8(0).Text
        .Update
    End With
End Sub
