VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bukubesar 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUKU BESAR"
   ClientHeight    =   6825
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   8145
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   240
      TabIndex        =   12
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9763
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Jurnal"
         Caption         =   "Referensi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Tanggal"
         Caption         =   "Tanggal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Uraian"
         Caption         =   "Uraian"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Debet"
         Caption         =   "Debet"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Kredit"
         Caption         =   "Kredit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2400
      Top             =   6360
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Kampus\SEMESTER 3\SIAK\TRANSAKSI VB\akuntansi.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Kampus\SEMESTER 3\SIAK\TRANSAKSI VB\akuntansi.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "bukubesar"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
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
      Left            =   240
      TabIndex        =   1
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label totsaldo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Kredit :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label totdebet 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label totkredit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Referensi :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Debet : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Menu aktiva 
      Caption         =   "1-0000 AKTIVA"
      Begin VB.Menu kas 
         Caption         =   "1-1000 KAS"
      End
      Begin VB.Menu piudag 
         Caption         =   "1-2000 PIUTANG DAGANG"
      End
      Begin VB.Menu pelengkap 
         Caption         =   "1-4000 PERLENGKAPAN"
      End
      Begin VB.Menu sewamuka 
         Caption         =   "1-5000 SEWA BAYAR DIMUKA"
      End
      Begin VB.Menu peralat 
         Caption         =   "1-8000 PERALATAN"
      End
      Begin VB.Menu nyusut 
         Caption         =   "1-9000 AKUMULASI PENYUSUTAN"
      End
   End
   Begin VB.Menu wajiban 
      Caption         =   "2-0000 KEWAJIBAN"
      Begin VB.Menu hutgang 
         Caption         =   "2-1000 HUTANG DAGANG"
      End
      Begin VB.Menu hutgaj 
         Caption         =   "2-2000 HUTANG GAJI"
      End
   End
   Begin VB.Menu modalmain 
      Caption         =   "3-0000 MODAL"
      Begin VB.Menu modalsub 
         Caption         =   "3-1000 MODAL"
      End
      Begin VB.Menu prive 
         Caption         =   "3-2000 PRIVE"
      End
      Begin VB.Menu rugilaba 
         Caption         =   "3-3000 IKHTISAR RUGI LABA"
      End
   End
   Begin VB.Menu dapat 
      Caption         =   "4-0000 PENDAPATAN"
      Begin VB.Menu Jual 
         Caption         =   "4-1000 PENJUALAN"
      End
   End
   Begin VB.Menu bebanmain 
      Caption         =   "5-0000 BEBAN"
      Begin VB.Menu beper 
         Caption         =   "5-1000 BEBAN PERLENGKAPAN"
      End
      Begin VB.Menu bega 
         Caption         =   "5-2000 BEBAN GAJI"
      End
      Begin VB.Menu bewa 
         Caption         =   "5-3000 BEBAN SEWA"
      End
      Begin VB.Menu bepen 
         Caption         =   "5-4000 BEBAN PENYUSUTAN"
      End
      Begin VB.Menu berup 
         Caption         =   "5-5000 BEBAN RUPA-RUPA"
      End
   End
End
Attribute VB_Name = "bukubesar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bega_Click()
Label4 = "5-2000"
Text3 = "52"
Call itung
End Sub

Private Sub bepen_Click()
Label4 = "5-4000"
Text3 = "54"
Call itung
End Sub

Private Sub beper_Click()
Label4 = "5-1000"
Text3 = "51"
Call itung
End Sub

Private Sub berup_Click()
Label4 = "5-5000"
Text3 = 55
Call itung
End Sub

Private Sub bewa_Click()
Label4 = "5-3000"
Text3 = 53
Call itung
End Sub

Private Sub Command1_Click()
jawab = MsgBox("APAKAH ANDA AKAN KEMBALI KE FORM TRANSAKSI?", vbYesNo, "KONFIRMASI")
If jawab = vbYes Then
bukubesar.Hide
inputan.Show
inputan.Frame1.Visible = True
inputan.butsimpan.Visible = True
inputan.Command1.Visible = False
inputan.DataGrid1.Visible = False
inputan.btnjurnal.Top = 4320
inputan.posting.Visible = False
inputan.btnubah.Visible = False
inputan.btnjurnal.Caption = "JURNAL UMUM"
Else
bukubesar.Hide
main.Show
End If
End Sub


Private Sub Command2_Click()
jawab = MsgBox("APAKAH ANDA AKAN MENGINPUT DATA?", vbYesNo, "KONFIRMASI")
If jawab = vbYes Then
bukubesar.Hide
pilihan.Show
Else
bukubesar.Hide
main.Show
End If
End Sub

Private Sub Form_Load()
DataGrid1.Refresh
Adodc1.Refresh
End Sub

Private Sub hutgaj_Click()
Label4 = "2-2000"
Text3 = 22
Call itung
End Sub

Private Sub hutgang_Click()
Label4 = "2-1000"
Text3 = "21"
Call itung
End Sub

Private Sub Jual_Click()
Label4 = "4-1000"
Text3 = 41
Call itung
End Sub

Private Sub kas_Click()
Label4 = "1-1000"
Text3 = 11
Call itung
End Sub

Private Sub Label1_Click()

End Sub

Private Sub modalsub_Click()
Label4 = "3-1000"
Text3 = "31"
Call itung
End Sub

Private Sub nyusut_Click()
Label4 = "1-9000"
Text3 = "19"
Call itung
End Sub

Private Sub pelengkap_Click()
Label4 = "1-4000"
Text3 = "14"
Call itung
End Sub

Private Sub peralat_Click()
Label4 = "1-8000"
Text3 = "18"
Call itung
End Sub

Private Sub piudag_Click()
Label4 = "1-2000"
Text3 = "12"
Call itung
End Sub

Private Sub itung()
totdebet = ""
totkredit = ""
Adodc1.Recordset.Filter = "Jurnal = '" & Text3 & "'"
If Adodc1.Recordset.RecordCount > 0 Then
    Text5 = Adodc1.Recordset!Debet
    Text6 = Adodc1.Recordset!Kredit
    Adodc1.Recordset.MoveFirst
    totaldebet = 0
    totalkredit = 0
    Adodc1.Recordset.MoveFirst
    Do Until Adodc1.Recordset.EOF
        totaldebet = Val(totaldebet) + Val(Adodc1.Recordset!Debet)
        totalkredit = Val(totalkredit) + Val(Adodc1.Recordset!Kredit)
        Adodc1.Recordset.MoveNext
    Loop
    totdebet = totaldebet
    totkredit = totalkredit
Else
End If
End Sub

Private Sub prive_Click()
Label4 = "3-2000"
Text3 = "32"
Call itung
End Sub

Private Sub rugilaba_Click()
Label4 = "3-3000"
Text3 = "33"
Call itung
End Sub

Private Sub sewamuka_Click()
Label4 = "1-5000"
Text3 = "15"
Call itung
End Sub

