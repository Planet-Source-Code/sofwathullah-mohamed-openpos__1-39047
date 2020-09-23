VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmDailySales 
   Appearance      =   0  'Flat
   BackColor       =   &H00A27E66&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1800
      ScaleHeight     =   345
      ScaleWidth      =   1905
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -15
         TabIndex        =   2
         Top             =   -15
         Width           =   1945
         _ExtentX        =   3440
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57737217
         CurrentDate     =   37483
      End
   End
   Begin HoverButton.Button Button1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Caption         =   "Ok [F3]"
      CaptionDown     =   "Ok [F3]"
      CaptionOver     =   "Ok [F3]"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Date Of Sales"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1470
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmDailySales"
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
    With DataEnvironment1.rsrptSalesSummry
        If Not .State = adStateClosed Then .Close
        .Open "SELECT daily_sales.bill, daily_sales.Bdate, daily_sales.total, daily_sales.CashSales " & _
        "From daily_sales " & _
        "WHERE (((daily_sales.Bdate)=#" & DTPicker1.Value & "#));"
    End With
    rptSales.Sections(1).Controls.Item(3).Caption = DTPicker1.Value
    rptSales.Show 1
    Unload Me
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    If KeyCode = vbKeyF3 Then
        Button1_Click
    ElseIf KeyCode = vbKeyEscape Then
        Command4_Click
    End If
End Sub
