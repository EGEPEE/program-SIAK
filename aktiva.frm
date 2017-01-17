VERSION 5.00
Begin VB.Form aktivapilih 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AKTIVA"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "KEMBALI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "AKUMULASI PENYUSUTAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton sewa 
      Caption         =   "SEWA DIBAYAR MUKA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PERALATAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PERLENGKAPAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PIUTANG DAGANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton kas 
      Caption         =   "KAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label euraian 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label ereferensi 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "aktivapilih"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
aktivapilih.Hide
inputan.editref = "12"
inputan.editurai = "Piutang Dagang"
inputan.Show
End Sub

Private Sub Command2_Click()
aktivapilih.Hide
inputan.editref = "14"
inputan.editurai = "Perlengkapan"
inputan.Show
End Sub

Private Sub Command3_Click()
aktivapilih.Hide
inputan.editref = "18"
inputan.editurai = "Peralatan"
inputan.Show
End Sub

Private Sub Command4_Click()
aktivapilih.Hide
inputan.editref = "19"
inputan.editurai = "Akumulasi Penyusutan"
inputan.Show
End Sub

Private Sub Command5_Click()
aktivapilih.Hide
pilihan.Show
End Sub

Private Sub kas_Click()
aktivapilih.Hide
inputan.editref = "11"
inputan.editurai = "Kas"
inputan.Show
End Sub

Private Sub sewa_Click()
aktivapilih.Hide
inputan.editref = "15"
inputan.editurai = "Sewa dibayar muka"
inputan.Show
End Sub
