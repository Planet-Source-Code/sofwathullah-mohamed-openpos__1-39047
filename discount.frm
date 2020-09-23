VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmDiscount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin HoverButton.Button Command4 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2760
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
   Begin HoverButton.Button Button1 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2280
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
   Begin VB.Line Line8 
      X1              =   6840
      X2              =   6840
      Y1              =   120
      Y2              =   3120
   End
   Begin VB.Line Line7 
      X1              =   1440
      X2              =   6840
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line6 
      X1              =   1440
      X2              =   6840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line5 
      X1              =   1440
      X2              =   1440
      Y1              =   120
      Y2              =   3120
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set Dis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   19
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "‘"
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
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   1320
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   6960
      X2              =   6960
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Amount Tendered:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   16
      Top             =   2040
      Width           =   1605
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   15
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label Label10 
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
      Left            =   5640
      TabIndex        =   14
      Top             =   2640
      Width           =   885
   End
   Begin VB.Label Label9 
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
      Left            =   5640
      TabIndex        =   13
      Top             =   1680
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total After Discount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   1800
      Width           =   1770
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discount Amount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Max. % is 100%)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   9
      Top             =   960
      Width           =   1545
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Discount %:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label5 
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
      Left            =   5640
      TabIndex        =   7
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max Possible Discount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Width           =   1965
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total For Bill:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   1140
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
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   885
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "?Ð???"
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
      Left            =   -840
      TabIndex        =   3
      Top             =   960
      Width           =   720
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   6840
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "frmDiscount"
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
    frmSales.DoSalesDBWithDis
    Unload Me
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
Private Sub Form_Load()
    '--win2k/winme
    Dim lOldStyle As Long
    Dim bTrans As Byte ' The level of transparency (0 - 255)

    bTrans = 220
    lOldStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, bTrans, LWA_ALPHA
    '--
    Label14.Caption = "Set " & vbCrLf & "Discount"
End Sub
Private Sub Label14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 FormMove Me
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    '-- % bai
    Dim bb As Double
    Dim pp As Double
    If KeyCode = vbKeyReturn Then
        If CDbl(Text1.Text) > CDbl(Me.Tag) Then
            DoMsg ("Cannot Exceed Max. Pecent")
            Exit Sub
        Else
            bb = CDbl(bal.Caption)
            pp = (bal / 100) * CDbl(Text1.Text)
            Text2.Text = Format(pp, "##,##0.00")
            Label9.Caption = Format((bb - (CDbl(Text2.Text))), "##,##0.00")
            Text3.SetFocus
        End If
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bb As Double
    Dim pp As Double
    If KeyCode = vbKeyReturn Then
        If CDbl(Text2.Text) > CDbl(Label5.Caption) Then
            DoMsg ("Cannot Exceed Max. Amount")
            Exit Sub
        Else
          bb = CDbl(bal.Caption)
          pp = (CDbl(Text2.Text) / bb) * 100
          Text1.Text = Format(pp, "##,##0.00")
          Label9.Caption = Format((bb - (CDbl(Text2.Text))), "##,##0.00")
          Text3.SetFocus
        End If
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not Label10.Caption = "00.00" Then
            Button1_Click
        End If
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim Tot As Currency
    Dim Paid As Currency
    Dim Balance As Currency
    If KeyAscii = vbKeyReturn Then
        If Not Text3.Text = Empty Then
            Tot = CDbl(Label9.Caption)
            If Tot > 0 And (CDbl(Text3.Text) > 0) Then
                Paid = CDbl(Text3.Text)
                Text3.Text = Format(Paid, "##,##0.00")
                Balance = Paid - Tot
                If Balance >= 0 Then
                    Label10.Caption = Format(Balance, "##,##0.00")
                    'ok done now need to do the db stuff
                Else
                    'Text4.Text = Empty
                    Beep
                    Label10.Caption = "00.00"
                    Text3.SelStart = 0
                    Text3.SelLength = Len(Text3.Text)
                    Text3.SetFocus
                End If
            End If
        End If
    End If
End Sub
