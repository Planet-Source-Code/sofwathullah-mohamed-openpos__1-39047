VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form frmItmList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   1725
   ClientTop       =   660
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AND"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmItmList.frx":0000
      DataField       =   "property"
      DataMember      =   "property"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ListField       =   "property"
      BoundColumn     =   "property"
      Text            =   "Field"
      Object.DataMember      =   "property"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin HoverButton.Button Button1 
      Height          =   250
      Left            =   8280
      TabIndex        =   10
      Top             =   110
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BackColor       =   -2147483633
      HoverBackColor  =   -2147483632
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   -2147483632
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   -2147483632
      HoverShadowColor=   -2147483632
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483630
      Caption         =   "r"
      CaptionDown     =   "r"
      CaptionOver     =   "r"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFInvoice 
      Height          =   4935
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   8705
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   14268829
      BackColorSel    =   14268829
      BackColorBkg    =   16777215
      GridColorFixed  =   14268829
      GridColorUnpopulated=   14268829
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin HoverButton.Button Button3 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   6960
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
      Caption         =   "New [F3]"
      CaptionDown     =   "New [F3]"
      CaptionOver     =   "New [F3]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Picture         =   "frmItmList.frx":0028
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Button2 
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   6960
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
   Begin HoverButton.Button Button4 
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   6960
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
   Begin HoverButton.Button Button5 
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   6960
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
   Begin HoverButton.Button Button6 
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   6960
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
      Caption         =   "Update [F6]"
      CaptionDown     =   "Update [F6]"
      CaptionOver     =   "Update [F6]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin HoverButton.Button Button7 
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   6960
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
      Caption         =   "Adjust [F7]"
      CaptionDown     =   "Adjust [F7]"
      CaptionOver     =   "Adjust [F7]"
      ShowFocusRect   =   0   'False
      Sink            =   -1  'True
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Item Management "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3480
      TabIndex        =   15
      Top             =   120
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search  By Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   14
      Top             =   480
      Width           =   1980
   End
   Begin MSForms.ComboBox ComboBox1 
      Bindings        =   "frmItmList.frx":107A
      DataField       =   "type"
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   1815
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3201;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ARE"
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
      Height          =   195
      Left            =   4320
      TabIndex        =   13
      Top             =   960
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00A27E66&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   27.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   190
      TabIndex        =   12
      Top             =   480
      Width           =   435
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1320
      Top             =   6735
      Width           =   375
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   5100
      Left            =   285
      Top             =   1470
      Width           =   8145
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00A27E66&
      FillColor       =   &H00A27E66&
      FillStyle       =   0  'Solid
      Height          =   7215
      Left            =   120
      Top             =   360
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Left            =   0
      Top             =   0
      Width           =   8655
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   8160
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8295
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00A27E66&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00A27E66&
      Height          =   855
      Left            =   1080
      Top             =   6720
      Width           =   7455
   End
End
Attribute VB_Name = "frmItmList"
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
    With MSHFInvoice
        .Col = 0
        If Not .Text = Empty Then
            Call MakeChildTrans(230, frmItm)
            frmItm.Label18.Caption = "Edit Item"
            frmItm.EditItem (.Text)
            frmItm.Show 1
        End If
    End With
End Sub
Private Sub Button3_Click()
    frmItm.Label18.Caption = "New Item"
    frmItm.AddNewItem
    frmItm.Show 1
End Sub
Private Sub Button4_Click()
    'del thingy ( i hate to del. stuff )
    MSHFInvoice.Col = 0
    If Not MSHFInvoice.Text = Empty Then
        With DataEnvironment1.rsCommand3
            If Not .State = adStateClosed Then .Close
            .Open "Select * From items Where code='" & MSHFInvoice.Text & "' AND Inhand=0"
            If .RecordCount = 1 Then
                frmMsgBox.Label1.Caption = "Are You Sure You Want To Delete This Record?"
                frmMsgBox.Show 1
                If frmMsgBox.Tag = "yes" Then
                    .Delete ' daa da magic command
                    FillGridDB
                    Me.Refresh
                End If
            Else
                DoMsg ("First Adjust Stock To Deleted")
            End If
        End With
    End If
End Sub
Private Sub Button5_Click()
    'done with this form
    Unload Me
End Sub
Private Sub Button6_Click()
    With MSHFInvoice
        .Col = 0
        If Not .Text = Empty Then
            Call MakeChildTrans(230, frmStockUpdate)
            frmStockUpdate.GetRecord (.Text)
            frmStockUpdate.Show 1
        End If
    End With
End Sub
Private Sub Button7_Click()
    With MSHFInvoice
        .Col = 0
        If Not .Text = Empty Then
            Call MakeChildTrans(230, frmAdjustments)
            frmAdjustments.GetRecord (.Text)
            frmAdjustments.Show 1
        End If
    End With
End Sub
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub ComboBox1_Change()
    FillGridDB
End Sub
Private Sub ComboBox1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Call Form_KeyDown(CDbl(KeyCode), Shift)
End Sub
Private Sub DataCombo1_Change()
    '--
    On Error Resume Next
    If Not DataEnvironment1.rsCommand3.State = adStateClosed Then DataEnvironment1.rsCommand3.Close
    DataEnvironment1.rsCommand3.Open "Select * From " & LCase(DataCombo1.Text)
    ComboBox1.Clear
    Do While Not DataEnvironment1.rsCommand3.EOF
        ComboBox1.AddItem DataEnvironment1.rsCommand3.Fields(0)
        DataEnvironment1.rsCommand3.MoveNext
    Loop
    ComboBox1.ListIndex = 0
End Sub
Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Form_GotFocus()
    FillGridDB
    Me.Refresh
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Button3_Click
        Case vbKeyF4
            Button2_Click
        Case vbKeyF5
            Button4_Click
        Case vbKeyF6
            Button6_Click
        Case vbKeyF7
            Button7_Click
        Case vbKeyEscape
            Button5_Click
    End Select
End Sub
Private Sub Form_Load()
    With MSHFInvoice
        .Cols = 4
        .RowHeightMin = 320
        .Row = 0
        .Col = 0: .Text = "Code"
        .Col = 1: .Text = "Description"
        .Col = 2: .Text = "In Hand"
        .Col = 3: .Text = "Price"
        '.Col = 4: .Text = "Cat."
        .ColWidth(1) = 4100
        .ColWidth(3) = 1000
        .ColWidth(0) = 1600
        '.Rows = 200
    End With
    Label1.Caption = "Search" & vbCrLf & "By" & vbCrLf & "Description"
    '--
    If Not DataEnvironment1.Recordsets(1).State = 1 Then
        DataEnvironment1.Recordsets(1).Open
    End If
    DataEnvironment1.Recordsets(1).Filter = "code > '*'" ' reset crazy thingy
    '--
    FillGridDB
    
End Sub
Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormMove Me
End Sub
Public Sub FillGridDB()
    On Error Resume Next
    Dim flD As String
    Dim r As Integer
    With DataEnvironment1.rsCommand1
        If Not .State = adStateClosed Then .Close
        If (Text1.Text = Empty Or Len(Text1.Text) < 1) And Check1.Value = 0 Then
            .Open "Select * From items where description like '%' order by description"
        Else
            If Check1.Value = 0 Then
                .Open "Select * From items where description like '%" & _
                    Text1.Text & "%' order by description"
                Debug.Print .RecordCount
            Else
                flD = IIf(DataCombo1.Text = "Catogory", "CatCode", _
                    Mid(DataCombo1.Text, 1, Len(DataCombo1.Text) - 1))
                .Open "Select * From items where description like '%" & _
                    Text1.Text & "%' AND " & flD & " ='" & ComboBox1.Text & "' order by description"
            End If
        End If
        If .RecordCount > 0 Then
            MSHFInvoice.Rows = .RecordCount + 1
            r = 0
            Do While Not .EOF
                r = r + 1
                MSHFInvoice.Row = r
                MSHFInvoice.Col = 0: MSHFInvoice.Text = !code
                MSHFInvoice.Col = 1: MSHFInvoice.Text = !Description
                MSHFInvoice.Col = 2: MSHFInvoice.Text = !Inhand
                MSHFInvoice.Col = 3: MSHFInvoice.Text = !B_Price
                .MoveNext
                Fancy
            Loop
            MSHFInvoice.Row = 1: MSHFInvoice.Col = 0
        End If
    End With
End Sub
Public Sub Fancy()
    Dim CurrentCell As Integer
    Dim r As Integer
    With MSHFInvoice
        If .Row Mod 2 = 0 And .Row <> 0 Then
            '// trying to make this row diff col
            CurrentCell = .Col
            For r = 0 To 3
                .Col = r
                .CellBackColor = &HFAEDDE       'RGB(174, 245, 214)
            Next
            .Col = CurrentCell
        End If
    End With
End Sub
Private Sub MSHFInvoice_Click()
    'Button2_Click
End Sub
Private Sub MSHFInvoice_DblClick()
    ViewItm
End Sub
Private Sub MSHFInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ViewItm
    End If
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text1_Change()
    FillGridDB
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then FillGridDB
End Sub
Private Sub ViewItm()
    With MSHFInvoice
        .Col = 0
        If Not .Text = Empty Then
            Call MakeChildTrans(230, frmItm)
            frmItm.Label18.Caption = "View Item"
            frmItm.EditItem (.Text)
            frmItm.Show 1
        End If
    End With
End Sub
