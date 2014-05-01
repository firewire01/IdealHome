VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   15720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   14400
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustomerFile 
      Caption         =   "CUSTOMER FILE"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11880
      TabIndex        =   7
      Top             =   3720
      Width           =   3855
   End
   Begin VB.CommandButton cmdSales 
      Caption         =   "SALES"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8520
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CommandButton cmdInventory 
      Caption         =   "INVENTORY"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4320
      TabIndex        =   5
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CommandButton cmdDelivery 
      Caption         =   "DELIVERY REPORT"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      Picture         =   "frmMenu.frx":0000
      ScaleHeight     =   1935
      ScaleWidth      =   3255
      TabIndex        =   3
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "888 Mabugat Road, Tabok, Mandaue City"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Trading and Manufacturing Inc."
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2040
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "IDEAL HOME"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   11655
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelivery_Click()
    frmDelivery.Show
End Sub

Private Sub cmdExit_Click()
    DB.Close
    End
End Sub

Private Sub Form_Load()
    Call setDBaseTable
End Sub
