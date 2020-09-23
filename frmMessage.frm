VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   570
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3240
      Top             =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   105
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid Item Code, Please Retry..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   1
      Top             =   165
      Width           =   3300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00A27E66&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   570
      Left            =   0
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   360
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00A27E66&
      BorderWidth     =   2
      Height          =   540
      Left            =   495
      Top             =   15
      Width           =   3210
   End
End
Attribute VB_Name = "frmMessage"
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

Private Sub Form_Load()
    Dim lOldStyle As Long
    Dim bTrans As Byte ' The level of transparency (0 - 255)
    bTrans = 200
    lOldStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    ''// SetWindowLong Me.hwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    SetLayeredWindowAttributes Me.hwnd, 0, bTrans, LWA_ALPHA
End Sub

Private Sub Timer1_Timer()
    '--
    Timer1.Enabled = False
    
    '--
    Unload Me
    'Me.Hide
End Sub
