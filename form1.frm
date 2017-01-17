VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form main 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SIAK"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1320
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form1.frx":7AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form1.frx":10148
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form1.frx":165D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "form1.frx":1EA39
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   9384
            MinWidth        =   7585
            Text            =   "Created by Athiatul Bary, Ega Prasetianti dan Isyania Farahani - 2IA01"
            TextSave        =   "Created by Athiatul Bary, Ega Prasetianti dan Isyania Farahani - 2IA01"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            Object.Width           =   1834
            MinWidth        =   35
            TextSave        =   "10/22/2016"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   1808
            MinWidth        =   9
            TextSave        =   "4:40 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1140
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   2011
      ButtonWidth     =   2619
      ButtonHeight    =   1852
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buku Besar"
            Key             =   "ddaftar"
            Description     =   "buku besar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Profil Perusahaan"
            Key             =   "dprofil"
            Description     =   "Profil Perusahaan"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Transaksi"
            Key             =   "dtransaksi"
            Description     =   "Transaksi"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buku Jurnal"
            Key             =   "dbuku"
            Description     =   "Buku Jurnal"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Key             =   "dclose"
            Description     =   "close"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAM INI TERDIRI DARI :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELAMAT DATANG DI PROGRAM AKUNTANSI "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Toolbar1_ButtonClick(ByVal Toolbar1 As Button)
Select Case Toolbar1.Key
Case Is = "ddaftar"
main.Hide
bukubesar.Command1.Visible = False
bukubesar.Command2.Visible = True
bukubesar.Show
Case Is = "dbuku"
main.Hide
buku.Show
Case Is = "dprofil"
main.Hide
perusahaan.Show
Case Is = "dtransaksi"
main.Hide
pilihan.Show
Case Is = "dclose"
Unload Me
'Case Is = "laporan"
'pil = MsgBox("Apakah anda ingin membuat Laporan?", vbYesNo + vbInformation, "Buat Laporan")
'If pil = 6 Then
'laporan1.Show
'End If
Case Is = "Delete"
End Select
End Sub
