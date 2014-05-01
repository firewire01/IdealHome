VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDelivery 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delivery Report ENTRY"
   ClientHeight    =   10455
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   16020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   12
      Top             =   9720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   11
      Top             =   9120
      Width           =   1455
   End
   Begin VB.TextBox txtEmployeeNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   9240
      Width           =   2535
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   8520
      Width           =   14175
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   3360
      Picture         =   "frmDelivery.frx":0000
      ScaleHeight     =   1125
      ScaleWidth      =   1125
      TabIndex        =   31
      Top             =   120
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delivery Information Details"
      Height          =   4335
      Left            =   120
      TabIndex        =   30
      Top             =   4080
      Width           =   15735
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13560
         TabIndex        =   8
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtQuantity 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtItemCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin MSFlexGridLib.MSFlexGrid grd 
         Height          =   2295
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   7
         BackColorBkg    =   0
         FormatString    =   $"frmDelivery.frx":0893
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   14400
         TabIndex        =   48
         Top             =   3600
         Width           =   180
      End
      Begin VB.Label lblTotalAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Php 0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13560
         TabIndex        =   41
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lblGrossAmount 
         Alignment       =   1  'Right Justify
         Caption         =   "Php 0.00"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13560
         TabIndex        =   40
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Total Amount:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12360
         TabIndex        =   39
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Discount:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12720
         TabIndex        =   38
         Top             =   3600
         Width           =   825
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Gross Amount:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   12240
         TabIndex        =   37
         Top             =   3120
         Width           =   1320
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4200
         TabIndex        =   34
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Item Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.TextBox txtPONo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   5
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox txtSalesman 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtTerms 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox txtCustomerNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtDeliveryNo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblEmpLastName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   47
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label lblEmpFirstName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   46
      Top             =   9720
      Width           =   3015
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4200
      TabIndex        =   45
      Top             =   9840
      Width           =   60
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "last name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   44
      Top             =   10200
      Width           =   690
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "first name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5400
      TabIndex        =   43
      Top             =   10200
      Width           =   720
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Employee No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   42
      Top             =   9360
      Width           =   1410
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Prepared By:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   36
      Top             =   9000
      Width           =   1140
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   35
      Top             =   8640
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   0
      X2              =   16080
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "PO / SO No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10560
      TabIndex        =   29
      Top             =   3000
      Width           =   1290
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Salesman"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10560
      TabIndex        =   28
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Terms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10560
      TabIndex        =   27
      Top             =   2040
      Width           =   660
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10560
      TabIndex        =   26
      Top             =   1560
      Width           =   510
   End
   Begin VB.Label lblAddress 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   3600
      Width           =   12975
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "first name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6000
      TabIndex        =   24
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "last name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3360
      TabIndex        =   23
      Top             =   3360
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4800
      TabIndex        =   22
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label lblFirstName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label lblLastName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Address  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   19
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label lblDeliveredTo 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   2400
      Width           =   7815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Delivered To "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   17
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Customer No. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delivery Receipt  No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080C0FF&
      Caption         =   "DELIVERY REPORT (ENTRY)"
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   14
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "IDEAL HOME TRADING AND MFG. INC."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   240
      Width           =   7815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "FILE"
      Begin VB.Menu mnuUpdate 
         Caption         =   "UPDATE records"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "Back to MENU"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "EXIT"
      End
   End
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Dim DeliveryNumber As String
Dim CustomerNumber As String
Dim r As Integer
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim Amount As Currency
Dim TAmount As Currency
Dim found As Boolean
Dim Discount As Double

Private Sub cmdClear_Click()
    txtDeliveryNo.Text = ""
    txtCustomerNo.Text = ""
    lblDeliveredTo.Caption = ""
    lblLastName.Caption = ""
    lblFirstName.Caption = ""
    lblAddress.Caption = ""
    txtDate.Text = ""
    txtTerms.Text = ""
    txtSalesman.Text = ""
    txtPONo.Text = ""
    txtItemCode.Text = ""
    txtQuantity.Text = ""
    lblGrossAmount.Caption = "Php 0.00"
    txtDiscount.Text = ""
    lblTotalAmount.Caption = "Php 0.00"
    txtDiscount.Text = 0
    txtRemarks.Text = ""
    txtEmployeeNo.Text = ""
    lblEmpLastName.Caption = ""
    lblEmpFirstName.Caption = ""
    grd.Rows = 1
    grd.Rows = grd.Rows + 1
    r = 1
End Sub

Private Sub Form_Load()
    Call setDBaseTable
    Call cmdClear_Click
    ShowCursor False 'hide cursor
End Sub

Private Sub mnuBack_Click()
    frmMenu.Show
    Call cmdClear_Click
End Sub

Private Sub mnuExit_Click()
    Call cmdClear_Click
    DB.Close
    End
End Sub

Private Sub txtCustomerNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CustomerNumber = UCase(Trim(txtCustomerNo.Text))
        If txtCustomerNo.Text = "" Then
            MsgBox "Customer Number is empty." & vbCrLf & "You need to fill-up the customer number field", vbCritical, "Customer No Field"
            txtCustomerNo.SetFocus
        Else
            If checkCUSTOMERFILE(CustomerNumber) = True Then
                MsgBox "Customer number doesn't exists" & vbCrLf & "Please check the Customer File for your customer number reference", vbInformation, "Customer No Field"
                txtCustomerNo.SetFocus
            Else
                txtCustomerNo.Text = CustomerNumber
                With rsCF
                    If IsNull(!CUSTOMERSTATUS) = True Then
                        MsgBox "There are no customer file available", vbCritical, "Customer File"
                        txtCustomerNo.SetFocus
                    ElseIf UCase(Trim(!CUSTOMERSTATUS)) = "NA" Then
                        MsgBox "Customer is not active", vbInformation, "Customer File"
                        txtCustomerNo.SetFocus
                    Else
                        If IsNull(!COMPNAME) = True Then
                            MsgBox "No customer's company name found" & vbCrLf & "Please check Customer File", vbCritical, "Delivered To Field"
                            txtCustomerNo.SetFocus
                        Else
                            lblDeliveredTo.Caption = UCase(!COMPNAME)
                        End If
                        If IsNull(!CUSTOMERLASTNAME) = True Then
                            lblLastName.Caption = ""
                        Else
                            lblLastName.Caption = UCase(Trim(!CUSTOMERLASTNAME))
                        End If
                        If IsNull(!CUSTOMERFIRSTNAME) = True Then
                            lblFirstName.Caption = ""
                        Else
                            lblFirstName.Caption = UCase(Trim(!CUSTOMERFIRSTNAME))
                        End If
                        If IsNull(!CUSTOMERADDRESS) = True Then
                            lblAddress.Caption = ""
                        Else
                            lblAddress.Caption = UCase(!CUSTOMERADDRESS)
                        End If
                    End If
                End With
                txtDate.Text = ""
                txtDate.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtDate.Text) = "" Then 'checks if date is empty'
            MsgBox "Date is required to continue delivery report.", vbCritical, "Delivery Report Date Field"
            txtDate.SetFocus
        ElseIf IsDate(Trim(txtDate.Text)) = False Then 'checks if date is a valid date'
            MsgBox "Date is invalid" & vbCrLf & "Please enter a valid date", vbCritical, "Delivery Report Date Field"
            txtDate.SetFocus
        ElseIf DateValue(Format(Trim(CDate(txtDate.Text)))) > DateValue(Now) Then 'checks if date is tomorrow'
            MsgBox "Error:" & vbCrLf & "Date is invalid" & vbCrLf & "Must be on or before current date", vbCritical, "Delivery Report Date Field"
            txtDate.SetFocus
        Else
            txtDate.Text = Format(CDate(Trim(txtDate.Text)), "mm/dd/yyyy")
            txtTerms.Text = ""
            txtTerms.SetFocus
        End If
    End If
End Sub

Private Sub txtDeliveryNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DeliveryNumber = UCase(Trim(txtDeliveryNo.Text))
        If DeliveryNumber = "" Then
            MsgBox "Delivery Receipt Number is empty." & vbCrLf & "You need to fill-up the delivery receipt number field", vbCritical, "Delivery Receipt No Field"
            txtDeliveryNo.SetFocus
        Else
            If checkDRHEADERFILE(DeliveryNumber) = False Then
                MsgBox "Delivery Receipt Number already exists." & vbCrLf & "Please enter a new Delivery Receipt Number", vbInformation, "Delivery Receipt No Field"
                txtDeliveryNo.SetFocus
            Else
                txtDeliveryNo.Text = DeliveryNumber
                txtCustomerNo.Text = ""
                txtCustomerNo.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(txtDiscount.Text) = False Then
            MsgBox "Please enter a valid discount", vbCritical, "Discount Field"
            txtDiscount.SetFocus
        Else
            Discount = txtDiscount.Text / 100
            TAmount = Discount * Amount
            lblTotalAmount.Caption = "Php " + FormatNumber(Amount - TAmount, 2, True, True, True)
            txtRemarks.Text = ""
            txtRemarks.SetFocus
        End If
    End If
End Sub

Private Sub txtItemCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim ITEMCODE As String
        ITEMCODE = UCase(Trim(txtItemCode.Text))
        If ITEMCODE = "" Then
            MsgBox "Item code is empty." & vbCrLf & "Please enter an item code", vbInformation, "Item Code Field"
            txtItemCode.SetFocus
        Else
            If checkINVENTORYFILE(ITEMCODE) = True Then
                MsgBox "Item Code / Product code doesn't exists" & vbCrLf & "Please refer to the Inventory File", vbCritical, "Item Code Field"
                txtItemCode.SetFocus
            Else
                With rsINV
                    If IsNull(!INVSTATUS) = True Then
                        MsgBox "Product doesn't exists" & vbCrLf & "Please review Inventory File", vbInformation, "Inventory File"
                        txtItemCode.SetFocus
                    ElseIf UCase(Trim(!INVSTATUS)) = "NA" Then
                        MsgBox "Product is already out of stock", vbInformation, "Inventory File"
                        txtItemCode.SetFocus
                    Else
                        If IsNull(!INVITEMCODE) = True Then
                            MsgBox "Item Code information is missing / empty", vbInformation, "Inventory File"
                            txtItemCode.SetFocus
                        ElseIf IsNull(!INVDSC) = True Then
                            MsgBox "Product Description information is missing / empty", vbInformation, "Inventory File"
                            txtItemCode.SetFocus
                        ElseIf IsNull(!INVUNITOFMEASURE) = True Then
                            MsgBox "Product unit of measure information is missing / empty", vbInformation, "Inventory File"
                            txtItemCode.SetFocus
                        ElseIf IsNull(!INVUNITPRICE) = True Then
                            MsgBox "Product unit price information is missing / empty", vbInformation, "Inventory File"
                            txtItemCode.SetFocus
                        Else
                            
                            k = r
                            If grd.TextMatrix(r, 1) = "" Then
                                For i = 1 To k
                                    If found = True Then
                                        Exit For
                                    Else
                                        If ITEMCODE = Trim(grd.TextMatrix(i, 1)) Then
                                            found = True
                                            r = i
                                            j = k - 1
                                            grd.Rows = grd.Rows - 1
                                            txtItemCode.Text = ""
                                            txtQuantity.SetFocus
                                            Exit Sub
                                        End If
                                    End If
                                Next i
                            End If
                            
                            
                            grd.TextMatrix(r, 1) = UCase(Trim(!INVITEMCODE))
                            grd.TextMatrix(r, 2) = UCase(Trim(!INVDSC))
                            grd.TextMatrix(r, 4) = UCase(Trim(!INVUNITOFMEASURE))
                            grd.TextMatrix(r, 5) = FormatNumber(!INVUNITPRICE, 2, True, True, True)
                            If (CInt(!INVQUANTITY) <= CInt(!INVREORDER)) = True Then
                                MsgBox "Notice:" & vbCrLf & "Product is below reorder level", vbInformation, "REORDER LEVEL"
                            End If
                            txtItemCode.Text = ""
                            txtQuantity.SetFocus
                        End If
                    End If
                End With
            End If
        End If
    End If
End Sub

Private Sub txtPONo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPONo.Text = UCase(Trim(txtPONo.Text))
        txtItemCode.Text = ""
        txtItemCode.SetFocus
    End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsEmpty(grd.TextMatrix(r, 3)) = True Then
            MsgBox "Quantity is empty" & vbCrLf & "Please enter a quantity", vbCritical, "Quantity Field"
            txtQuantity.SetFocus
        ElseIf isIntegral(Trim(txtQuantity.Text)) = False And grd.TextMatrix(r, 3) = "" Then
            MsgBox "Please enter a valid quantity", vbInformation, "Quantity Field"
            txtQuantity.SetFocus
        Else
            If grd.TextMatrix(r, 1) = "" Or grd.TextMatrix(r, 2) = "" Or grd.TextMatrix(r, 4) = "" Or grd.TextMatrix(r, 5) = "" Then
                MsgBox "Please enter an item code", vbInformation, "Item Code Field"
                txtItemCode.SetFocus
            Else
                If checkINVENTORYFILE(UCase(Trim(grd.TextMatrix(r, 1)))) = False Then
                    With rsINV
                        If txtQuantity.Text <> "" Then
                            If CInt(txtQuantity.Text) > CInt(!INVQUANTITY) Then
                                MsgBox "There are only " & Trim(!INVQUANTITY) & " " & LCase(Trim(!INVUNITOFMEASURE)) & " available for delivery", vbInformation, "Inventory File"
                                txtQuantity.SetFocus
                            ElseIf CInt(txtQuantity.Text) <= 0 Then
                                MsgBox "Quantity must be greater than zero(0)", vbInformation, "Quantity Field"
                                txtQuantity.SetFocus
                            Else
                                grd.TextMatrix(r, 3) = CInt(txtQuantity.Text)
                                grd.TextMatrix(r, 6) = FormatNumber(grd.TextMatrix(r, 3) * Trim(!INVUNITPRICE), 2, True, True, True)
                            End If
                        End If
                        txtQuantity.Text = ""
                        Amount = 0
                        For i = 1 To r
                            Amount = Amount + grd.TextMatrix(i, 6)
                        Next i
                        lblGrossAmount.Caption = "Php" + FormatNumber(Amount, 2, True, True, True)
                        lblTotalAmount.Caption = "Php" + FormatNumber(Amount, 2, True, True, True)
                        If txtDiscount.Text <> 0 Then
                            Discount = txtDiscount.Text / 100
                            TAmount = Discount * Amount
                            lblTotalAmount.Caption = "Php " + FormatNumber(Amount - TAmount, 2, True, True, True)
                        End If
                        found = False
                        r = r + 1
                        grd.Rows = grd.Rows + 1
                        txtItemCode.SetFocus
                    End With
                End If
            End If
        End If
    End If
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRemarks.Text = UCase(Trim(txtRemarks.Text))
        txtEmployeeNo.Text = ""
        txtEmployeeNo.SetFocus
    End If
End Sub

Private Sub txtSalesman_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(Trim(txtSalesman.Text)) = True Then
            MsgBox "Error:" & vbCrLf & "Salesman name is not valid. It contains numeric values", vbCritical, "Salesman Field"
            txtSalesman.Text = ""
            txtSalesman.SetFocus
        ElseIf checkLetter(UCase(Trim(txtSalesman.Text))) = False Then
            MsgBox "Incorrect Salesman Name" & vbCrLf & "Please indicate the correct name", vbCritical, "Salesman Field"
            txtSalesman.SetFocus
        Else
            txtSalesman.Text = UCase(Trim(txtSalesman.Text))
            txtPONo.Text = ""
            txtPONo.SetFocus
        End If
    End If
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTerms.Text = UCase(txtTerms.Text)
        txtSalesman.Text = ""
        txtSalesman.SetFocus
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ShowCursor True 'show cursor
End Sub
