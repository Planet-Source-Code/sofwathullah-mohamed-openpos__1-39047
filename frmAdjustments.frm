VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmAdjustments 
   Appearance      =   0  'Flat
   BackColor       =   &H00A27E66&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
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
   ScaleHeight     =   4095
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   795
      Index           =   4
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2520
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton DeductStock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Deduct From Stock"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   1650
      Value           =   -1  'True
      Width           =   1665
   End
   Begin VB.OptionButton AddStock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Add To Stock"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2760
      TabIndex        =   1
      Top             =   1650
      Width           =   1260
   End
   Begin HoverButton.Button Button1 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3600
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
      TabIndex        =   6
      Top             =   3600
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
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     "
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     "
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      Height          =   4095
      Left            =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Stock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Remarks"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Balance"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Adjustment"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   1590
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Description"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1110
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Code"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   630
      Width           =   750
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   120
      Top             =   480
      Width           =   6855
   End
End
Attribute VB_Name = "frmAdjustments"
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
Public newQty
Private Sub Button1_Click()
    Dim minuS As Boolean
    minuS = IIf(AddStock.Value = True, False, True)
    If (CDbl(Label8.Caption) = qty) And (Not Text1(2).Text = Empty) Then
        If AddStock.Value = True Then
            Label8.Caption = AdjStockPlus(CDbl(Text1(2).Text))
        Else
            Label8.Caption = AdjStockMinus(CDbl(Text1(2).Text))
        End If
    End If
    If Not (CDbl(Label8.Caption) = qty) Then ' changes need and log
        DataEnvironment1.rsCommand3!Inhand = CDbl(Label8.Caption)
        DataEnvironment1.rsCommand3.Update
        If minuS = True Then
            Call BatchUpdate(False, Text1(4).Text)
        Else
            Call BatchUpdate(True, Text1(4).Text)
        End If
    End If
    If frmItmList.Visible Then
        frmItmList.FillGridDB
        frmItmList.Refresh
    End If
    Unload Me
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Public Sub GetRecord(itmCode As String)
    founD = False
    With DataEnvironment1.rsCommand3
        If Not .State = adStateClosed Then .Close
        .Open "Select * From items Where code ='" & itmCode & "'"
        Text8(1).Text = !code
        Label10.Caption = !Description
        qty = !Inhand
        Label8.Caption = qty
    End With
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
        Case vbKeyBack
            KeyAscii = vbKeyBack
        Case KeyAscii = vbKeyDelete
        Case vbKeyReturn
            '--
            If Not Text1(2).Text = Empty Then
                If AddStock.Value = True Then
                    Label8.Caption = AdjStockPlus(CDbl(Text1(2).Text))
                Else
                    Label8.Caption = AdjStockMinus(CDbl(Text1(2).Text))
                End If
            End If
        Case Else
            KeyAscii = IIf(Index = 4, KeyAscii, 0)
    End Select
End Sub
Private Function AdjStockMinus(adj As Double) As Double
    Dim newQty As Double
    newQty = qty - adj
    If Not newQty < 0 Then
        Label8.Caption = newQty
        AdjStockMinus = newQty
    Else
       DoMsg ("Cannot Be Nagative")
       AdjStockMinus = qty
    End If
End Function
Public Sub DoMsg(msg As String)
    frmMessage.Label1.Caption = msg
    frmMessage.Show 1
    frmMessage.Timer1.Enabled = True
End Sub
Private Sub BatchUpdate(Addin As Boolean, remKs As String)
    'Store all adjustment details
    With DataEnvironment1.rsadj
        If Not .State = adStateClosed Then .Close
        .Open "Select * From adj"
        .AddNew
        '--to do : add user
        !icode = Text8(1).Text
        !qtyb4 = qty
        !adj = Text1(2).Text
        !added = Addin
        !Date = Now
        ' !user = Cuser
        !rem = remKs
        .Update
    End With
End Sub
Private Function AdjStockPlus(adj As Double) As Double
    Dim newQty As Double
    newQty = qty + adj
    Label8.Caption = newQty
    AdjStockPlus = newQty
End Function
Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 And KeyAscii = vbKeyReturn Then
        GetRecord (Text8(1).Text)
    End If
End Sub
