VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmStockUpdate 
   Appearance      =   0  'Flat
   BackColor       =   &H00A27E66&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
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
   ScaleHeight     =   3615
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
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
      Index           =   6
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text8 
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
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Total Cost For The Item"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text8 
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
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text8 
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
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text8 
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
      Index           =   3
      Left            =   5880
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text8 
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
      Index           =   4
      Left            =   5880
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text8 
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
      Index           =   5
      Left            =   5880
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin HoverButton.Button Button1 
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   3120
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
      Left            =   5760
      TabIndex        =   9
      Top             =   3120
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
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     "
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Stock"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape2 
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   7095
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
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   750
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
      Left            =   240
      TabIndex        =   16
      Top             =   1080
      Width           =   795
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
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   915
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
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Qty"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   1560
      Width           =   630
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
      Left            =   4680
      TabIndex        =   12
      Top             =   2040
      Width           =   885
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
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   780
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
      Left            =   4680
      TabIndex        =   10
      Top             =   2520
      Width           =   915
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   120
      Top             =   450
      Width           =   6855
   End
End
Attribute VB_Name = "frmStockUpdate"
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

Public qty As Double
Public founD As Boolean
Private Sub Button1_Click()
    'action part
    Dim newqTyTotal As Double
    Dim newQ As Double
    'If founD Then ' usless but.. can be implemented later for validation
        With DataEnvironment1.rsCommand3
            newQ = CDbl(Text8(3).Text)
            newqTyTotal = qty + newQ
            'update rec
            !code = Text8(6).Text
            !Description = Label12.Caption
            !CostPrice = Text8(0).Text
            !B_Price = Text8(1).Text
            !A_price = Text8(2).Text
            !maxDiscountP = Text8(4).Text
            !memDiscountP = Text8(5).Text
            !Inhand = newqTyTotal
            .Update
            Call BatchUpdateThingy(Text8(6).Text, newqTyTotal)
        End With
        If frmItmList.Visible Then
            frmItmList.FillGridDB
            frmItmList.Refresh
        End If
        Unload Me
    'End If
End Sub
Private Sub Command4_Click()
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
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Public Sub GetRecord(itmCode As String)
    founD = False
    With DataEnvironment1.rsCommand3
        If Not .State = adStateClosed Then .Close
        .Open "Select * From items Where code ='" & itmCode & "'"
        If Not .EOF Or Not .BOF Then
            Text8(6).Text = !code
            Label12.Caption = !Description
            Text8(0).Text = !CostPrice
            Text8(1).Text = !B_Price
            Text8(2).Text = !A_price
            Text8(4).Text = !maxDiscountP
            Text8(5).Text = !memDiscountP
            qty = !Inhand
        Else
            DoMsg ("Record Not Found")
            Text8(6).Text = Empty
            Label12.Caption = Empty
            Text8(0).Text = Empty
            Text8(1).Text = Empty
            Text8(2).Text = Empty
            Text8(4).Text = Empty
            Text8(5).Text = Empty
            qty = Empty
        End If
    End With
End Sub
Private Sub Text8_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Index = 6 Then
        GetRecord (Text8(6).Text)
    Else
        Call Form_KeyDown(KeyCode, Shift)
    End If
End Sub
Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case Asc(".")
            KeyAscii = IIf(Index = 3, 0, KeyAscii)
        Case vbKeyBack
            KeyAscii = vbKeyBack
        Case KeyAscii = vbKeyDelete
        Case Else
            KeyAscii = IIf(Index = 6, KeyAscii, 0)
    End Select
End Sub
Private Sub BatchUpdateThingy(codeZ As String, newTot As Double)
    On Error Resume Next
    Dim oldB As Integer
    With DataEnvironment1.rsbatch
        If Not .State = adStateClosed Then .Close
        .Open "Select * From batch where icode='" & codeZ & "' order by batch"
        If .RecordCount > 0 Then
            If Not .EOF Then .MoveLast
            oldB = !batch
            .AddNew
            '--to do : add user
            !icode = Text8(6).Text
            !qty = CDbl(Text8(3).Text)
            !batch = (oldB + 1)
            !Date = Now
            !prevqty = qty
            !totalqty = newTot
            !Add = True
            !transtype = "New Stock"
            !cost = CDbl(Text8(0).Text)
            !salesPA = CDbl(Text8(2).Text)
            !salesPB = CDbl(Text8(1).Text)
            .Update
        End If
    End With
End Sub
