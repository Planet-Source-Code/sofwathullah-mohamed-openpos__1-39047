VERSION 5.00
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmMemList 
   Appearance      =   0  'Flat
   BackColor       =   &H00A27E66&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
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
   ScaleHeight     =   3120
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   5655
   End
   Begin HoverButton.Button Button6 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
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
      Caption         =   "Edit [F4]"
      CaptionDown     =   "Edit [F4]"
      CaptionOver     =   "Edit [F4]"
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
      Left            =   5640
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
      Left            =   1320
      TabIndex        =   7
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
   Begin HoverButton.Button Button3 
      Height          =   375
      Left            =   4200
      TabIndex        =   8
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
      Caption         =   "Delete [F5]"
      CaptionDown     =   "Delete [F5]"
      CaptionOver     =   "Delete [F5]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   3015
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "Members"
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
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "m"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Search::"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape5 
      Height          =   3120
      Left            =   0
      Top             =   0
      Width           =   7200
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   1200
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmMemList"
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
Private Sub Button2_Click()
    frmMemberz.Label1.Caption = "Add Member"
    frmMemberz.AddMember
    frmMemberz.Show 1
End Sub
Private Sub Button3_Click()
    Dim Mcode As String
    If Not List1.Text = Empty Then
        Mcode = GetMemNo(List1.Text)
        With DataEnvironment1.rsmem
            If Not .State = adStateClosed Then .Close
            .Open "Select * From Memberz Where code='" & Mcode & "'"
            If .RecordCount = 1 Then
                frmMsgBox.Label1.Caption = "Are You Sure You Want To Delete This Record?"
                frmMsgBox.Show 1
                If frmMsgBox.Tag = "yes" Then
                    .Delete ' daa da magic command
                    Text1_Change
                    Me.Refresh
                End If
            End If
        End With
    End If
End Sub
Private Sub Button6_Click()
    If Not List1.Text = Empty Then
        frmMemberz.GetMember GetMemNo(List1.Text)
        frmMemberz.Label1.Caption = "Edit Member"
        frmMemberz.Show 1
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Button2_Click
        Case vbKeyF4
            Button6_Click
        Case vbKeyF5
            Button3_Click
        Case vbKeyEscape
            Button1_Click
    End Select
End Sub
Private Sub Form_Load()
    Label1.Caption = "Members" & vbCrLf & "Management"
    Text1_Change
End Sub
Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Private Sub List1_DblClick()
    If Not List1.Text = Empty Then
        frmMemberz.GetMember GetMemNo(List1.Text)
        frmMemberz.Label1.Caption = "View Member"
        frmMemberz.Show 1
    End If
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        List1_DblClick
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Public Sub Text1_Change()
    On Error Resume Next
    With DataEnvironment1.rsmem
        If Not .State = adStateClosed Then .Close
        If Not (Text1.Text = Empty Or Len(Text1.Text) < 1) Then
            .Open "Select * From Memberz where name like '%" & Text1.Text & "%' order by name"
        Else
            .Open "Select * From Memberz Where name Like '%' Order By name"
        End If
        List1.Clear
        Do While Not .EOF
            List1.AddItem !Name
            .MoveNext
        Loop
    End With
End Sub
Public Function GetMemNo(memName As String) As String
    On Error Resume Next
    With DataEnvironment1.rsmem
        If Not .State = adStateClosed Then .Close
        .Open "Select * From Memberz Where name = '" & memName & "'"
        If Not .AbsolutePosition = adPosEOF Then
            GetMemNo = !code
        End If
    End With
End Function
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
