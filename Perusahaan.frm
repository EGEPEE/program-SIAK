VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form perusahaan 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7350
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Profil Perusahaan"
      TabPicture(0)   =   "Perusahaan.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(1)=   "npwp"
      Tab(0).Control(2)=   "izin"
      Tab(0).Control(3)=   "kode"
      Tab(0).Control(4)=   "bidang"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(9)=   "Label1"
      Tab(0).Control(10)=   "nama"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Ubah Profil"
      TabPicture(1)   =   "Perusahaan.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label10"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label12"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Text1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Text2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Text5"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Command2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin VB.CommandButton Command3 
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
         Height          =   495
         Left            =   -74760
         TabIndex        =   23
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ubah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SIMPAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   21
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   20
         Top             =   2520
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   17
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   16
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label12 
         Caption         =   "NPWP                     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Izin Usaha               :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Kode Saham           :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Bidang Usaha        :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Nama Perusahaan :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label npwp 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72600
         TabIndex        =   10
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label izin 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72600
         TabIndex        =   9
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label kode 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72600
         TabIndex        =   8
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label bidang 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72600
         TabIndex        =   7
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "NPWP                     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Izin Usaha               :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Kode Saham           :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Bidang Usaha        :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Perusahaan :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label nama 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72600
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
   End
End
Attribute VB_Name = "perusahaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
nama = Text1
bidang = Text2
kode = Text3
izin = Text4
npwp = Text5
End Sub

Private Sub Command2_Click()
pass = InputBox("Masukkan Password : ", "Admin", "Enter Password here")
If Not (pass = "12345") Then
    MsgBox "Anda Bukan Admin", vbCritical, "Masukkan password yang tepat!"""
Else
    MsgBox "Silahkan ubah data"
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
End If
End Sub

Private Sub Command3_Click()
perusahaan.Hide
main.Show
End Sub

