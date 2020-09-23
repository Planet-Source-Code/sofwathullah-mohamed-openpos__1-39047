VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmMemberz 
   Appearance      =   0  'Flat
   BackColor       =   &H00A27E66&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
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
   ScaleHeight     =   4575
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
      Height          =   1035
      Index           =   3
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sex"
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   7560
      TabIndex        =   30
      Top             =   1470
      Width           =   1335
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Female"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Male"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00A27E66&
         BackStyle       =   0  'Transparent
         Caption         =   ""
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A27E66&
         Height          =   390
         Left            =   0
         TabIndex        =   32
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00A27E66&
         BackStyle       =   0  'Transparent
         Caption         =   "€"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A27E66&
         Height          =   390
         Left            =   0
         TabIndex        =   31
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Marital Status"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   7320
      TabIndex        =   28
      Top             =   3180
      Width           =   1575
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Devoiced"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Married"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Single"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Index           =   6
      Left            =   3240
      TabIndex        =   8
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Index           =   5
      Left            =   3240
      TabIndex        =   7
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Index           =   4
      Left            =   6600
      TabIndex        =   5
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Index           =   1
      Left            =   6720
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "code"
      DataMember      =   "items"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin HoverButton.Button Button1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3600
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
      Left            =   240
      TabIndex        =   14
      Top             =   4080
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      ScaleHeight     =   345
      ScaleWidth      =   1785
      TabIndex        =   24
      Top             =   2760
      Width           =   1815
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -15
         TabIndex        =   6
         Top             =   -15
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   661
         _Version        =   393216
         Format          =   57671681
         CurrentDate     =   37471
      End
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Occupation :"
      Height          =   195
      Left            =   1920
      TabIndex        =   27
      Top             =   3840
      Width           =   915
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Email :"
      Height          =   195
      Left            =   1920
      TabIndex        =   26
      Top             =   3360
      Width           =   465
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact Numbers :"
      Height          =   195
      Left            =   5160
      TabIndex        =   25
      Top             =   2760
      Width           =   1350
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DOB :"
      Height          =   195
      Left            =   1920
      TabIndex        =   23
      Top             =   2760
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Address :"
      Height          =   195
      Left            =   1920
      TabIndex        =   22
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Name (Full) :"
      Height          =   195
      Left            =   1920
      TabIndex        =   21
      Top             =   1080
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIC/PP Numbar :"
      Height          =   195
      Left            =   5400
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Member Code :"
      Height          =   195
      Left            =   1920
      TabIndex        =   19
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00A27E66&
      BackStyle       =   0  'Transparent
      Caption         =   "ƒ"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   80.25
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
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   4575
      Left            =   0
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "Add Members"
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
      TabIndex        =   16
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
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
      Left            =   1680
      TabIndex        =   15
      Top             =   120
      Width           =   7350
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   1680
      Top             =   360
      Width           =   7335
   End
End
Attribute VB_Name = "frmMemberz"
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
    Dim v As Boolean
    With DataEnvironment1.rsmem
        If Not Label1.Caption = "View Member" Then
            If Label1.Caption = "New Mameber" Then
                .Find ("code='" & Text1(0).Text & "'")
            End If
            If Not .AbsolutePosition = adPosEOF And Label1.Caption = "New Member" Then
                DoMsg ("Code Already Registered")
                Exit Sub
            Else
                v = validTha
                If v = True Then
                    If Label1.Caption = "Add Member" Then
                        .AddNew
                    End If
                Else
                    DoMsg ("Important Data Missing")
                    Exit Sub
                End If
            End If
        End If
        !code = Text1(0).Text
        !nidpp = Text1(1).Text
        !Name = Text1(2).Text
        !address = Text1(3).Text
        !tel = Text1(4).Text
        !email = Text1(5).Text
        !occup = Text1(6).Text
        !dob = DTPicker1.Value
        !sex = IIf(Option1.Value = True, "M", "F")
        If Option3.Value = True Then
            !mstate = 1
        ElseIf Option4.Value = True Then
            !mstate = 2
        Else
            !mstate = 3
        End If
        If Not Label1.Caption = "View Member" Then
            .Update
            If frmMemList.Visible = True Then
                frmMemList.Text1_Change
            End If
        End If
    End With
    Unload Me
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
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
    Option1.Value = True
    Option3.Value = True
End Sub
Private Sub Label20_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Public Sub AddMember()
    With DataEnvironment1.rsmem
        If Not .State = adStateClosed Then .Close
        .Open "Select * From Memberz"
    End With
End Sub
Public Function GetMember(memberCode As String) As Boolean
    On Error Resume Next
    With DataEnvironment1.rsmem
        If Not .State = adStateClosed Then .Close
        .Open "Select * From Memberz where code='" & memberCode & "'"
        If .RecordCount > 0 Then
            Text1(0).Text = !code
            Text1(2).Text = !Name
            Text1(1).Text = !nidpp
            Text1(3).Text = !address
            Text1(4).Text = !tel
            Text1(5).Text = !email
            Text1(6).Text = !occup
            DTPicker1.Value = !dob
            If !sex = "M" Then Option1.Value = True Else Option2.Value = True
            Select Case !mstate
                Case 1
                    Option3.Value = True
                Case 2
                    Option4.Value = True
                Case 3
                    Option5.Value = True
            End Select
            GetMember = True
        Else
            DoMsg ("No Record(s) Found")
            GetMember = False
        End If
    End With
End Function
Public Function validTha() As Boolean
    If (Not Text1(0).Text = Empty) And (Not Text1(2).Text = Empty) And (Not Text1(3).Text = Empty) Then
        validTha = True
    Else
        validTha = False
    End If
End Function
Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Option4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Option5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

