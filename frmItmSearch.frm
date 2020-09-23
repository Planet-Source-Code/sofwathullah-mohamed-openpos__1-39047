VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmItemSearch 
   BorderStyle     =   0  'None
   Caption         =   "Item Search"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   5415
   End
   Begin VB.TextBox Text1 
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
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin HoverButton.Button Button6 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "View Info [F3]"
      CaptionDown     =   "View Info [F3]"
      CaptionOver     =   "View Info [F3]"
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
      Left            =   5400
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Shape Shape6 
      Height          =   2895
      Left            =   1200
      Top             =   120
      Width           =   5655
   End
   Begin VB.Shape Shape5 
      Height          =   3120
      Left            =   0
      Top             =   0
      Width           =   6960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Search::"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   615
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
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   3000
      Width           =   6975
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "L"
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
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ë"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6480
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search  By Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   -360
      TabIndex        =   1
      Top             =   600
      Width           =   1980
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmItemSearch"
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
    Unload Me
End Sub
Private Sub Button6_Click()
    frmItm.Label18.Caption = "View Item"
    frmItm.EditItem (GetItemCode(List1.Text))
    frmItm.Label9.Visible = False
    frmItm.Text8(0).Visible = False
    frmItm.Show 1
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Button6_Click
        Case vbKeyEscape
            Button1_Click
    End Select
End Sub
Private Sub Form_Load()
    Label1.Caption = "Search" & vbCrLf & "By" & vbCrLf & "Description"
    Text1_Change
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Button6_Click
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text1_Change()
    With DataEnvironment1.rsitm
        If Not .State = adStateClosed Then
            .Close
        End If
        If Not (Text1.Text = Empty Or Len(Text1.Text) < 1) Then
            .Open "Select * From items where description like '%" & Text1.Text & "%' order by description"
        Else
            .Open "Select * From items Where description Like '%' Order By description"
        End If
        List1.Clear
        Do While Not .EOF
            List1.AddItem !Description
            .MoveNext
        Loop
    End With
End Sub
Public Function GetItemCode(des As String) As String
    With DataEnvironment1.rsitm
        If Not .State = adStateClosed Then .Close
        .Open "Select * From items where description = '" & des & "'"
        If Not .EOF Or Not .BOF Then
            GetItemCode = !code
        End If
    End With
End Function
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
