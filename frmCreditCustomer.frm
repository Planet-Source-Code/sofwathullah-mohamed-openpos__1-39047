VERSION 5.00
Begin VB.Form frmCustomers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credit Customers"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Settle "
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Account Active"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1080
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Amnt. Due:"
      Height          =   195
      Left            =   4200
      TabIndex        =   19
      Top             =   1560
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Credit Value:"
      Height          =   195
      Left            =   1080
      TabIndex        =   17
      Top             =   1560
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contact"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
