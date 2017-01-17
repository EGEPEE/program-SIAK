VERSION 5.00
Begin VB.Form kasinput 
   Caption         =   "KAS"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "KREDIT"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1680
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DEBET"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox etanggal 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Uraian"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "KREDIT / DEBET ="
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label editurai 
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Uraian"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label editref 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Referensi"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "kasinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
aktivapilih.ereferensi = editref
aktivapilih.euraian = editurai
End Sub

