VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainUI 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   Caption         =   "OpenPOS [V-Neck]"
   ClientHeight    =   7695
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2566
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   0
      Picture         =   "frmMainUI.frx":0000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9600
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainUI.frx":0B17
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainUI.frx":12BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainUI.frx":1AC6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "©2002. Designed and developed by Sofwath@Hotmail.Com"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5340
      TabIndex        =   1
      Top             =   1440
      Width           =   4950
   End
   Begin VB.Menu sys 
      Caption         =   "&System"
      Begin VB.Menu sys_abt 
         Caption         =   "&About"
      End
      Begin VB.Menu sys_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu inv 
      Caption         =   "&Inventory"
      Begin VB.Menu itm_mang 
         Caption         =   "Item &Management"
         Shortcut        =   {F4}
      End
      Begin VB.Menu inv_cat 
         Caption         =   "Add &Categories"
         Enabled         =   0   'False
      End
      Begin VB.Menu inv_prop 
         Caption         =   "Item &Property management"
      End
   End
   Begin VB.Menu mem 
      Caption         =   "&Members"
      Begin VB.Menu mem_manag 
         Caption         =   "&Members management"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu pos 
      Caption         =   "&POS"
      Begin VB.Menu pos_sales 
         Caption         =   "&Sales"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu rep 
      Caption         =   "&Reports"
      Begin VB.Menu rep_dailysales 
         Caption         =   "&Daily Sales [Today]"
      End
      Begin VB.Menu repDailySales 
         Caption         =   "Daily &Sales"
      End
      Begin VB.Menu rep_sreorep 
         Caption         =   "Stock &Reorder Report"
      End
   End
End
Attribute VB_Name = "frmMainUI"
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

Dim xx As Double, yy As Double
Private Sub Form_Load()
    OpenDB
    ListView1.Width = Me.Width - 110
    Dim itmX As ListItem
    Set itmX = ListView1.ListItems.Add(1, "Sal", "Sales", 1, 1)
    Set itmX = ListView1.ListItems.Add(2, "Itm", "Item Management", 3, 1)
    Set itmX = ListView1.ListItems.Add(3, "Mem", "Members Management", 2, 1)
End Sub
Private Sub Form_Resize()
    ListView1.Width = Me.Width - 110
    Label1.Left = Me.Width - ((Label1.Width) + 150)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub inv_cat_Click()
    frmCatogory.Show 1
End Sub
Private Sub inv_prop_Click()
    frmMangeItemProperties.Show 1
End Sub
Private Sub itm_mang_Click()
    frmItmList.Show 1
End Sub
Private Sub ListView1_Click()
    '--menu stuff
    On Error Resume Next
    Select Case ListView1.HitTest(xx, yy).Key
        Case "Sal"
            frmSales.Show 1
        Case "Itm"
            frmItmList.Show 1
        Case "Mem"
            frmMemList.Show 1
    End Select
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Select Case ListView1.SelectedItem.Key
            Case "Sal"
            frmSales.Show 1
        Case "Itm"
            frmItmList.Show 1
        Case "Mem"
            frmMemList.Show 1
        End Select
    End If
End Sub
Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xx = X
    yy = Y
End Sub

Private Sub mem_manag_Click()
    frmMemList.Show 1
End Sub

Private Sub pos_sales_Click()
    frmSales.Show 1
End Sub

Private Sub rep_dailysales_Click()
'    SELECT daily_sales.bill, daily_sales.Bdate, daily_sales.total, daily_sales.CashSales
'    From daily_sales
'    WHERE (((daily_sales.Bdate)=#8/14/2002#));
    With DataEnvironment1.rsrptSalesSummry
        If Not .State = adStateClosed Then .Close
        .Open "SELECT daily_sales.bill, daily_sales.Bdate, daily_sales.total, daily_sales.CashSales " & _
        "From daily_sales " & _
        "WHERE (((daily_sales.Bdate)=#" & Date & "#));"
    End With
    rptSales.Sections(1).Controls.Item(3).Caption = Now
    rptSales.Show
End Sub
Private Sub repDailySales_Click()
    frmDailySales.Show 1
End Sub

Private Sub sys_abt_Click()
    frmMsgBox.Label1.Caption = "OpenPOS is designed for small shops. Everything is in DB, just generate the reports. Main modules are here. User management can easily be implemented." & vbCrLf & " ©2002 Sofwath@Hotmail.COM"
    frmMsgBox.Command4.Visible = False
    frmMsgBox.Show
End Sub

Private Sub sys_exit_Click()
    End
End Sub
