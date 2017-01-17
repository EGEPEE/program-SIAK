VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form inputan 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9195
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   9165
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   5640
      TabIndex        =   46
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   7800
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton btnjurnal 
      Caption         =   "JURNAL UMUM"
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
      Left            =   6000
      TabIndex        =   27
      Top             =   4320
      Width           =   1575
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
      Height          =   615
      Left            =   7680
      TabIndex        =   26
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mutasi 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   4815
      Begin VB.TextBox editkredit 
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
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox editdebet 
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
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.OptionButton opdebet 
         Caption         =   "Debet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton opkredit 
         Caption         =   "Kredit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox etanggal 
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
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah uang: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Uraian "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label editurai 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ref :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label editref 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8160
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "akuntan"
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
   Begin VB.CommandButton butkeluar 
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
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton butsimpan 
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
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mutasi 2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.ComboBox comjurnal 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   9
         Text            =   "-----PILIH SATU------"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox eurai 
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
         Left            =   1200
         TabIndex        =   2
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label ltanggal 
         Caption         =   "Label6"
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
         Left            =   1200
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label labelkredit 
         Caption         =   "labelkredit"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label labeldebet 
         Caption         =   "labeldebet"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "No. Ref:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Uraian : "
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
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Kredit : "
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
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form2.frx":0342
      Height          =   1095
      Left            =   120
      TabIndex        =   40
      Top             =   8040
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "Jurnal"
         Caption         =   "Jurnal"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "UBAH BUKU JURNAL UMUM"
      Height          =   3135
      Left            =   120
      TabIndex        =   29
      Top             =   4920
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox ubahref 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   47
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox ubahkredit 
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
         Left            =   1200
         TabIndex        =   38
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox ubahdebet 
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
         Left            =   1200
         TabIndex        =   36
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox ubahurai 
         DataField       =   "Jurnal"
         DataSource      =   "Adodc2"
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
         Left            =   1200
         TabIndex        =   34
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox ubahtanggal 
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
         Left            =   1200
         TabIndex        =   31
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Kredit : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Debet : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Uraian : "
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
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ref :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton posting 
      Caption         =   "POSTING"
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
      Left            =   4440
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton btnubah 
      Caption         =   "UBAH"
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
      Left            =   7680
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0357
      Height          =   4215
      Left            =   120
      TabIndex        =   39
      Top             =   600
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7435
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
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
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
   Begin VB.Label bbkredit 
      Caption         =   "bbkredit"
      Height          =   255
      Left            =   5160
      TabIndex        =   45
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label bbdebet 
      Caption         =   "bbdebet"
      Height          =   255
      Left            =   5160
      TabIndex        =   44
      Top             =   8640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label bburai 
      Caption         =   "urai"
      Height          =   255
      Left            =   5160
      TabIndex        =   43
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label bbjurnal 
      Caption         =   "ref"
      Height          =   255
      Left            =   5160
      TabIndex        =   42
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label bbtanggal 
      Caption         =   "tanggal"
      Height          =   255
      Left            =   5160
      TabIndex        =   41
      Top             =   7440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSAKSI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "inputan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnjurnal_Click()
Select Case btnjurnal.Caption
Case "JURNAL UMUM"
    Label1.Caption = "PEMBUKUAN JURNAL UMUM"
    Frame1.Visible = False
    Frame2.Visible = False
    DataGrid1.Visible = True
    btnubah.Visible = True
    Command1.Visible = False
    butsimpan.Visible = False
    btnjurnal.Caption = "TRANSAKSI"
    btnjurnal.Top = 6840
    posting.Visible = True
    butkeluar.Visible = False
Case "TRANSAKSI"
    inputan.Hide
    pilihan.Show
    btnjurnal.Top = 4320
    posting.Visible = False
    Frame1.Visible = True
    DataGrid1.Visible = False
    butsimpan.Visible = True
    btnjurnal.Caption = "JURNAL UMUM"
End Select
End Sub

Private Sub btnubah_Click()
    Select Case btnubah.Caption
    Case "UBAH"
        pass = InputBox("Masukkan Password : ", "Admin", "Enter Password here")
        If Not (pass = "12345") Then
        MsgBox "Anda Bukan Admin", vbCritical, "Masukkan password yang tepat!"
        Else
        Frame3.Visible = True
        ubahtanggal = Adodc1.Recordset!Tanggal
        ubahref = Adodc1.Recordset!Jurnal
        ubahurai = Adodc1.Recordset!Uraian
        ubahdebet = Adodc1.Recordset!Debet
        ubahkredit = Adodc1.Recordset!Kredit
        btnubah.Caption = "UPDATE"
        End If
    Case "UPDATE"
       Adodc1.Recordset!Tanggal = ubahtanggal
       Adodc1.Recordset!Jurnal = ubahref
       Adodc1.Recordset!Uraian = ubahurai
       Adodc1.Recordset!Debet = ubahdebet
       Adodc1.Recordset!Kredit = ubahkredit
       Adodc1.Recordset.Update
       DataGrid1.Refresh
       ubahtanggal = ""
       ubahref = ""
       ubahurai = ""
       ubahdebet = ""
       ubahkredit = ""
       btnubah.Caption = "UBAH"
    End Select
End Sub

Private Sub butkeluar_Click()
inputan.Hide
main.Show
End Sub

Private Sub butsimpan_Click()
ltanggal = etanggal
labelkredit = editdebet
labeldebet = editkredit
Frame1.Visible = False
Frame2.Visible = True
List1.Visible = True
Adodc1.Recordset.AddNew
Adodc1.Recordset!Tanggal = etanggal.Text
Adodc1.Recordset!Jurnal = editref.Caption
Adodc1.Recordset!Uraian = editurai.Caption
Adodc1.Recordset!Debet = editdebet.Text
Adodc1.Recordset!Kredit = editkredit.Text
Adodc1.Recordset.Update
DataGrid1.Refresh
butsimpan.Visible = False
Command1.Visible = True
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset!Tanggal = etanggal.Text
Adodc1.Recordset!Jurnal = comjurnal.Text
Adodc1.Recordset!Uraian = eurai.Text
Adodc1.Recordset!Debet = labeldebet.Caption
Adodc1.Recordset!Kredit = labelkredit.Caption
Adodc1.Recordset.Update
DataGrid1.Refresh
Call Form_Load
etanggal = ""
ltanggal = ""
editref = ""
editurai = ""
editdebet = ""
opdebet.Value = False
opkredit.Value = False
editkredit = ""
comjurnal = ""
eurai = ""
Frame2.Visible = False
List1.Visible = False
Frame1.Visible = True
Command1.Visible = False
butsimpan.Visible = True
End Sub


Private Sub Form_Load()
comjurnal.AddItem ("11")
comjurnal.AddItem ("12")
comjurnal.AddItem ("14")
comjurnal.AddItem ("15")
comjurnal.AddItem ("18")
comjurnal.AddItem ("19")
comjurnal.AddItem ("------------------------------------------------")
comjurnal.AddItem ("21")
comjurnal.AddItem ("22")
comjurnal.AddItem ("------------------------------------------------")
comjurnal.AddItem ("31")
comjurnal.AddItem ("32")
comjurnal.AddItem ("33")
comjurnal.AddItem ("------------------------------------------------")
comjurnal.AddItem ("41")
comjurnal.AddItem ("------------------------------------------------")
comjurnal.AddItem ("51")
comjurnal.AddItem ("52")
comjurnal.AddItem ("53")
comjurnal.AddItem ("54")
comjurnal.AddItem ("55")
List1.AddItem ("11. KAS")
List1.AddItem ("12. PIUTANG DAGANG")
List1.AddItem ("14. PERLENGKAPAN")
List1.AddItem ("15. SEWA BAYAR DIMUKA")
List1.AddItem ("18. PERALATAN")
List1.AddItem ("19. AKUMULASI PENYUSUTAN")
List1.AddItem ("21. HUTANG DAGANG")
List1.AddItem ("12. HUTANG GAJI")
List1.AddItem ("12. PIUTANG DAGANG")
List1.AddItem ("14. PERLENGKAPAN")
List1.AddItem ("15. SEWA BAYAR DIMUKA")
List1.AddItem ("18. PERALATAN")
List1.AddItem ("19. AKUMULASI PENYUSUTAN")
List1.AddItem ("---------------------------------------------------------------------------------")
List1.AddItem ("21. HUTANG DAGANG")
List1.AddItem ("22. HUTANG GAJI")
List1.AddItem ("---------------------------------------------------------------------------------")
List1.AddItem ("31. MODAL")
List1.AddItem ("32. PRIVE")
List1.AddItem ("33. IKHTISAR RUGI-LABA")
List1.AddItem ("---------------------------------------------------------------------------------")
List1.AddItem ("41. PENJUALAN")
List1.AddItem ("---------------------------------------------------------------------------------")
List1.AddItem ("51. BEBAN PERLENGKAPAN")
List1.AddItem ("52. BEBAN GAJI")
List1.AddItem ("53. BEBAN SEWA")
List1.AddItem ("54. BEBAN PENYUSUTAN")
List1.AddItem ("55. BEBAN RUPA-RUPA")
ubahref.AddItem ("11")
ubahref.AddItem ("12")
ubahref.AddItem ("14")
ubahref.AddItem ("15")
ubahref.AddItem ("18")
ubahref.AddItem ("19")
ubahref.AddItem ("------------------------------------------------")
ubahref.AddItem ("21")
ubahref.AddItem ("22")
ubahref.AddItem ("------------------------------------------------")
ubahref.AddItem ("31")
ubahref.AddItem ("32")
ubahref.AddItem ("33")
ubahref.AddItem ("------------------------------------------------")
ubahref.AddItem ("41")
ubahref.AddItem ("------------------------------------------------")
ubahref.AddItem ("51")
ubahref.AddItem ("52")
ubahref.AddItem ("53")
ubahref.AddItem ("54")
ubahref.AddItem ("55")
End Sub

Private Sub opdebet_Click()
Label4.Caption = "Kredit :"
editdebet.Visible = True
labelkredit.Visible = True
editkredit.Visible = False
labeldebet.Visible = False
editkredit.Text = 0
End Sub

Private Sub opkredit_Click()
Label4.Caption = "Debet :"
editkredit.Visible = True
editdebet.Visible = False
labeldebet.Visible = True
labelkredit.Visible = False
editdebet.Text = 0
End Sub

Private Sub posting_Click()
jawab = MsgBox("DATA YANG TERINPUT TIDAK DAPAT DIUBAH, LANJUTKAN? ", vbYesNo + vbExclamation + vbCritical, "KONFIRMASI")
If jawab = vbYes Then
    Adodc1.Recordset.Update
    bbtanggal = Adodc1.Recordset!Tanggal
    bbjurnal = Adodc1.Recordset!Jurnal
    bburai = Adodc1.Recordset!Uraian
    bbdebet = Adodc1.Recordset!Debet
    bbkredit = Adodc1.Recordset!Kredit
    Call inputbb
Else
End If

End Sub

Private Sub inputbb()
        Adodc2.Recordset.AddNew
        Adodc2.Recordset!Jurnal = bbjurnal
        Adodc2.Recordset!Tanggal = bbtanggal
        Adodc2.Recordset!Uraian = bburai
        Adodc2.Recordset!Debet = bbdebet
        Adodc2.Recordset!Kredit = bbkredit
        Adodc2.Recordset.Update
        Adodc1.Recordset.Delete
        DataGrid2.Refresh
        bukubesar.DataGrid1.Refresh
        Call buku
End Sub

Private Sub buku()
jawab = MsgBox("DATA SUDAH TERINPUT KE BUKU BESAR, APAKAH ANDA AKAN MELIHAT BUKU BESAR?", vbYesNo + vbInformation, "KONFIRMASI")
If jawab = vbYes Then
bukubesar.Adodc1.Refresh
bukubesar.DataGrid1.Refresh
bukubesar.Show
inputan.Hide
Else
inputan.Show
End If
End Sub
Private Sub ubahref_Change()
Call Form_Load
End Sub
