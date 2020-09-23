VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmMangeItemProperties 
   BorderStyle     =   0  'None
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmItmProperties.frx":0000
      Height          =   2655
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      DefColWidth     =   287
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "types"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmItmProperties.frx":001F
      DataField       =   "property"
      DataMember      =   "property"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   1560
      TabIndex        =   8
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "property"
      Text            =   "DataCombo1"
      Object.DataMember      =   "property"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HoverButton.Button Button1 
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "&Done (F6)"
      CaptionDown     =   "&Done (F6)"
      CaptionOver     =   "&Done (F6)"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Add (F3)"
      CaptionDown     =   "Add (F3)"
      CaptionOver     =   "Add (F3)"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Update (F4)"
      CaptionDown     =   "Update (F4)"
      CaptionOver     =   "Update (F4)"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Button4 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   16777215
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Caption         =   "Delete (F5)"
      CaptionDown     =   "Delete (F5)"
      CaptionOver     =   "Delete (F5)"
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
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select Property"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   120
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "Mange Item Properties ::::::::::::::::::::::::::::::::::::::::::"
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
      TabIndex        =   0
      Top             =   120
      Width           =   5160
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmMangeItemProperties"
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

Sub CancelTran()
On Error Resume Next
    Select Case LCase(DataCombo1.Text)
        Case "types"
            DataEnvironment1.rstypes.Cancel
        Case "colors"
            DataEnvironment1.rscolors.Cancel
        Case "brands"
            DataEnvironment1.rsbrands.Cancel
        Case "catogory"
            DataEnvironment1.rscatogory.Cancel
        Case "origins"
            DataEnvironment1.rsorigins.Cancel
    End Select
End Sub

Private Sub Button1_Click()
    Unload Me
End Sub

Private Sub Button2_Click()
On Error Resume Next
    If DataGrid1.Text = Empty Then
            CancelTran
            Exit Sub
    End If
    Select Case LCase(DataCombo1.Text)
        Case "types"
            DataEnvironment1.rstypes.AddNew
        Case "colors"
            DataEnvironment1.rscolors.AddNew
        Case "brands"
            DataEnvironment1.rsbrands.AddNew
        Case "catogory"
            DataEnvironment1.rscatogory.AddNew
        Case "origins"
            DataEnvironment1.rsorigins.AddNew
    End Select
End Sub

Private Sub Button3_Click()
On Error Resume Next
    If DataGrid1.Text = Empty Then Exit Sub
    Select Case LCase(DataCombo1.Text)
        Case "types"
            DataEnvironment1.rstypes.Update
        Case "colors"
            DataEnvironment1.rscolors.Update
        Case "brands"
            DataEnvironment1.rsbrands.Update
        Case "catogory"
            DataEnvironment1.rscatogory.Update
        Case "origins"
            DataEnvironment1.rsorigins.Update
    End Select
End Sub

Private Sub Button4_Click()
Dim res As Integer
On Error Resume Next
    frmMsgBox.Label1.Caption = "Are You Sure You Want To Delete This Record?"
    frmMsgBox.Show 1
    If frmMsgBox.Tag = "yes" Then
        Select Case LCase(DataCombo1.Text)
            Case "types"
                DataEnvironment1.rstypes.Delete
            Case "colors"
                DataEnvironment1.rscolors.Delete
            Case "brands"
                DataEnvironment1.rsbrands.Delete
            Case "catogory"
                DataEnvironment1.rscatogory.Delete
            Case "origins"
                DataEnvironment1.rsorigins.Delete
        End Select
    End If
End Sub

Private Sub DataCombo1_Change()
'--
    DataGrid1.DataMember = DataCombo1.Text
    DataGrid1.Refresh
End Sub
Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    DataGrid1.DataMember = DataCombo1.Text
    DataGrid1.Refresh
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
        Case vbKeyF3
            Button2_Click
            DataGrid1.Text = ""
        Case vbKeyF4
            Button3_Click
            DataGrid1.SetFocus
        Case vbKeyF5
            Button4_Click
        Case vbKeyF6
            Button3_Click
        Case vbKeyEscape
            CancelTran
        Case Else
            'do noth'n
    End Select
End Sub

Private Sub Form_Load()
    '--
    DataGrid1.DataMember = DataCombo1.Text
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub
