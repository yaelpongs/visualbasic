VERSION 5.00
Begin VB.Form cmdPayroll 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   15615
   ScaleWidth      =   28560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "PRINT"
      Height          =   975
      Left            =   16440
      Picture         =   "cmdLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "COMPUTE GROSS"
      Height          =   495
      Left            =   360
      TabIndex        =   55
      Top             =   10200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NET INCOME"
      Height          =   495
      Left            =   9120
      TabIndex        =   54
      Top             =   10680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COMPUTE DEDUCTION"
      Height          =   495
      Left            =   9120
      TabIndex        =   53
      Top             =   9960
      Width           =   1455
   End
   Begin VB.TextBox txtnetincome 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   52
      Top             =   10680
      Width           =   1695
   End
   Begin VB.TextBox txttotdeduction 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   51
      Top             =   9960
      Width           =   1695
   End
   Begin VB.TextBox txtpag 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   50
      Top             =   9240
      Width           =   3735
   End
   Begin VB.TextBox txtphil 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   48
      Top             =   8280
      Width           =   3735
   End
   Begin VB.TextBox txttax 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   46
      Top             =   7320
      Width           =   3735
   End
   Begin VB.TextBox txtsss1 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   44
      Top             =   6360
      Width           =   3735
   End
   Begin VB.TextBox txtGrossPay 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   41
      Top             =   9480
      Width           =   3735
   End
   Begin VB.TextBox txtmeal 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   39
      Top             =   8520
      Width           =   3735
   End
   Begin VB.TextBox txtperhour 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   37
      Top             =   7560
      Width           =   3735
   End
   Begin VB.TextBox txtperday 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   35
      Top             =   6600
      Width           =   3735
   End
   Begin VB.TextBox txt15th 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   33
      Top             =   5640
      Width           =   3735
   End
   Begin VB.TextBox txtpagibig 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   30
      Top             =   4680
      Width           =   3735
   End
   Begin VB.TextBox txtphilhealth 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   28
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox txttin 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   26
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox txtSSS 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   24
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtdateTo 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   3720
      Width           =   1695
   End
   Begin VB.TextBox txtdatefrom 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   20
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txttranno 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3720
      Width           =   3735
   End
   Begin VB.TextBox txtdatehired 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   15
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtmonthlysalary 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   13
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtposition 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   3735
   End
   Begin VB.TextBox txtempID 
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton closebttn 
      Caption         =   "CLOSE"
      Height          =   975
      Left            =   16440
      Picture         =   "cmdLogin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton findbttn 
      Caption         =   "FIND"
      Height          =   975
      Left            =   16440
      Picture         =   "cmdLogin.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton deletebttn 
      Caption         =   "DELETE"
      Height          =   975
      Left            =   16440
      Picture         =   "cmdLogin.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton savebttn 
      Caption         =   "SAVE"
      Height          =   975
      Left            =   16440
      Picture         =   "cmdLogin.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton addbttn 
      Caption         =   "ADD"
      Height          =   975
      Left            =   16440
      Picture         =   "cmdLogin.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label22 
      Caption         =   "PAG-IBIG:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   49
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label21 
      Caption         =   "PHILHEALTH:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   47
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label20 
      Caption         =   "TAX WITH HELD:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   45
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label Label19 
      Caption         =   "SSS#:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   43
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "LIST OF DEDUCTIONS"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   42
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label17 
      Caption         =   "GROSS PAY:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "MEAL/TRAVEL ALLOWANCE:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   38
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Label Label15 
      Caption         =   "RATE PER HOUR:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "RATE PER DAY:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "RATE PER 15TH DAY OF THE MONTH:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label12 
      Caption         =   "BREAK DOWN OF WAGES"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "PAG-IBIG#:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   29
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "PHILHEALTH#:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   27
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "TIN#:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   25
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "SSS#:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   23
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "DATE COVERED TO:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "DATE COVERED FROM:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   19
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   " TRANSACTION NO:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "DATE HIRED:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10320
      TabIndex        =   14
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "MONTHLY SALARY:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "POSITION:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "ADDRESS:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label dfsz 
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label txtid 
      Caption         =   "EMPLOYEE ID:"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "cmdPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub addbttn_Click()
On Error Resume Next

txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus

txttranno.Text = ""
txtempID.Text = ""
txtname.Text = ""
txtmonthlysalary.Text = ""
txtdatefrom.Text = ""
txtdateTo.Text = ""
txtdatehired.Text = ""
txtSSS.Text = ""
txttin.Text = ""
txtphilhealth.Text = ""
txtpagibig.Text = ""
txt15th.Text = ""
txtperday.Text = ""
txtperhour.Text = ""
txtmeal.Text = ""
txtGrossPay.Text = ""
txtsss1.Text = ""
txttax.Text = ""
txtphil.Text = ""
txtpag.Text = ""
txttotdeduction.Text = ""
txtnetincome.Text = ""
End Sub

Private Sub closebttn_Click()
 Unload Me
End Sub

Private Sub Command1_Click()
Dim xsss As Single
Dim xtax As Single
Dim xpag As Single
Dim xphil As Single
Dim xtOTD As Double
xsss = 300
txtsss1.Text = xsss

xtax = 500
txttax.Text = xtax

xpag = 100
txtpag.Text = xpag

xphil = 100
txtphil.Text = xphil

xtOTD = xsss + xtax + xphil + xpag

txttotdeduction.Text = xtOTD
End Sub

Private Sub Command2_Click()
Dim xNet As Double
Dim xG As Double
Dim xD As Double


xG = txtGrossPay.Text
xD = txttotdeduction.Text

xNet = xG - xD
txtnetincome.Text = xNet
End Sub

Private Sub Command3_Click()
Dim xrate15 As Double
Dim xSalary As Double
Dim xrateperday As Double
Dim xrateperhour As Double
Dim xmeal As Double
Dim xGross As Double



xSalary = txtmonthlysalary.Text
xrate15 = xSalary / 2
txt15th.Text = xrate15

xrateperday = txtmonthlysalary.Text / 26
txtperday.Text = xrateperday

xrateperhour = txtperday.Text / 8
txtperhour.Text = xrateperhour

xmeal = 500

txtmeal.Text = xmeal

xGross = xmeal + xrate15

txtGrossPay.Text = xGross
End Sub

Private Sub Command4_Click()

End Sub

Private Sub deletebttn_Click()
conPayroll.Execute "Delete * from payroll where tranno='" & Trim(txttranno.Text) & "'"
MsgBox "Record has been deleted.."
End Sub

Private Sub cmdFind_Click()
txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus

End Sub

Private Sub cmdGross_Click()
Dim xrate15 As Double
Dim xSalary As Double
Dim xrateperday As Double
Dim xrateperhour As Double
Dim xmeal As Double
Dim xGross As Double



xSalary = txtmonthlysalary.Text
xrate15 = xSalary / 2
txt15th.Text = xrate15

xrateperday = txtmonthlysalary.Text / 26
txtperday.Text = xrateperday

xrateperhour = txtperday.Text / 8
txtperhour.Text = xrateperhour

xmeal = 500

txtmeal.Text = xmeal

xGross = xmeal + xrate15

txtGrossPay.Text = xGross

End Sub

Private Sub findbttn_Click()
txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus
End Sub

Private Sub Form_Load()
openWORKSPACEODBC
openconPayroll
End Sub

Private Sub savebttn_Click()
openrstPayroll "SELECT * FROM payroll where tranno='" & Trim(txttranno.Text) & "'"
If Not rstPayroll.EOF Then
'if not found
    With rstPayroll
        .Edit
            .Fields("tranno").Value = Trim(txttranno.Text)
            .Fields("employeeid").Value = Trim(txtempID.Text)
            .Fields("datefrom").Value = Trim(txtdatefrom.Text)
            .Fields("dateto").Value = Trim(txtdateTo.Text)
            .Fields("rate15").Value = Trim(txt15th.Text)
            .Fields("rateperday").Value = Trim(txtperday.Text)
            .Fields("rateperhour").Value = Trim(txtperhour.Text)
            .Fields("meal").Value = Trim(txtmeal.Text)
            .Fields("grosspay").Value = Trim(txtGrossPay.Text)
            .Fields("datehired").Value = Trim(txtdatehired.Text)
            .Fields("sssno").Value = Trim(txtSSS.Text)
            .Fields("tinno").Value = Trim(txttin.Text)
            .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
            .Fields("pagibigno").Value = Trim(txtpagibig.Text)
            .Fields("sss").Value = Trim(txtsss1.Text)
            .Fields("tax").Value = Trim(txttax.Text)
            .Fields("pagibig").Value = Trim(txtpag.Text)
            .Fields("philhealth").Value = Trim(txtphil.Text)
            .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
            .Fields("netincome").Value = Trim(txtnetincome.Text)
            
        .Update
        
    End With
Else
    'not found
        With rstPayroll
            .AddNew
                .Fields("tranno").Value = Trim(txttranno.Text)
                .Fields("employeeid").Value = Trim(txtempID.Text)
                .Fields("datefrom").Value = Trim(txtdatefrom.Text)
                .Fields("dateto").Value = Trim(txtdateTo.Text)
                .Fields("rate15").Value = Trim(txt15th.Text)
                .Fields("rateperday").Value = Trim(txtperday.Text)
                .Fields("rateperhour").Value = Trim(txtperhour.Text)
                .Fields("meal").Value = Trim(txtmeal.Text)
                .Fields("grosspay").Value = Trim(txtGrossPay.Text)
                .Fields("datehired").Value = Trim(txtdatehired.Text)
                .Fields("sssno").Value = Trim(txtSSS.Text)
                .Fields("tinno").Value = Trim(txttin.Text)
                .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
                .Fields("pagibigno").Value = Trim(txtpagibig.Text)
                .Fields("sss").Value = Trim(txtsss1.Text)
                .Fields("tax").Value = Trim(txttax.Text)
                .Fields("pagibig").Value = Trim(txtpagibig.Text)
                .Fields("philhealth").Value = Trim(txtphil.Text)
                .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
                .Fields("netincome").Value = Trim(txtnetincome.Text)
            .Update
            
        End With
End If

End Sub

Private Sub Text3_Change()

End Sub

Private Sub Text16_Change()

End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtposition.SetFocus
    End If
End Sub

Private Sub txtEmployeeID_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    openrstEmployee "Select * from employee where employeeid ='" & Trim(txtEmployeeID.Text) & "'"
     If Not rstEmployee.EOF Then
        With rstEmployee
            txtEmployeeID.Text = .Fields("employeeid").Value
            txtname.Text = .Fields("employeename").Value
            txtaddress.Text = .Fields("address").Value
            txtposition.Text = .Fields("position").Value
            txtsalary.Text = .Fields("salary").Value
            txtdatehired.Text = .Fields("datehired").Value
        End With
    End If
    
End If
End Sub

Private Sub txtempID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    openrstEmployee "SELECT * FROM EMPLOYEE WHERE EMPLOYEEID='" & Trim(txtempID.Text) & "'"
    If Not rstEmployee.EOF Then
        With rstEmployee
            txtposition.Text = .Fields("position").Value
            txtaddress.Text = .Fields("address").Value
            txtempID.Text = .Fields("employeeid").Value
            txtname.Text = .Fields("employeename").Value
            txtmonthlysalary.Text = .Fields("salary").Value
            txtdatehired.Text = .Fields("datehired").Value
        End With
    End If
End If


End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress.SetFocus
    End If
End Sub

Private Sub txtposition_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsalary.SetFocus
    End If
End Sub

Private Sub txtsalary_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdatehired.SetFocus
    End If
End Sub

Private Sub txttranno_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    openrstPayroll "Select * from payroll where tranno='" & Trim(txttranno.Text) & "'"
     If Not rstPayroll.EOF Then
        With rstPayroll
            txttranno.Text = .Fields("tranno").Value
            txtempID.Text = .Fields("employeeid").Value
            txtdatefrom.Text = .Fields("datefrom").Value
            txtdateTo.Text = .Fields("dateto").Value
            txt15th.Text = .Fields("rate15").Value
            txtperday.Text = .Fields("rateperday").Value
            txtperhour.Text = .Fields("rateperhour").Value
            txtmeal.Text = .Fields("meal").Value
            txtGrossPay.Text = .Fields("grosspay").Value
            txtdatehired.Text = .Fields("datehired").Value
            txtSSS.Text = .Fields("sssno").Value
            txttin.Text = .Fields("tinno").Value
            txtphilhealth.Text = .Fields("philhealthno").Value
            txtpagibig.Text = .Fields("pagibigno").Value
            txtsss1.Text = .Fields("sss").Value
            txttax.Text = .Fields("tax").Value
            txtpag.Text = .Fields("pagibig").Value
            txtphil.Text = .Fields("philhealth").Value
            txttotdeduction.Text = .Fields("totaldeduction").Value
            txtnetincome.Text = .Fields("netincome").Value
            
        End With
    End If
    
End If
End Sub
