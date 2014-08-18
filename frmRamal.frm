VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmRamal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERAMALAN"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   15390
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15840
      Top             =   6720
      Visible         =   0   'False
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
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin TabDlg.SSTab tabPenjualan 
      Height          =   8415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   14843
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "DATA PENJUALAN"
      TabPicture(0)   =   "frmRamal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAmbil"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtPenjualan"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListViewTampil"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDouble"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSingle"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdKeluar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "EXPONENSIAL SMOOTING SINGLE"
      TabPicture(1)   =   "frmRamal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "Command5"
      Tab(1).Control(2)=   "ListViewSingle"
      Tab(1).Control(3)=   "ListViewHasilSingle"
      Tab(1).Control(4)=   "ListViewDataSingle"
      Tab(1).Control(5)=   "txtalfa"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "EXPONENSIAL SMOOTING DOUBLE"
      TabPicture(2)   =   "frmRamal.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command2"
      Tab(2).Control(1)=   "Command6"
      Tab(2).Control(2)=   "ListViewDouble"
      Tab(2).Control(3)=   "ListViewDataDouble"
      Tab(2).Control(4)=   "ListViewHasilDouble"
      Tab(2).Control(5)=   "Text1"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "GRAFIK"
      TabPicture(3)   =   "frmRamal.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command3"
      Tab(3).Control(1)=   "mscSingle"
      Tab(3).Control(2)=   "ListViewGrafik"
      Tab(3).ControlCount=   3
      Begin VB.CommandButton Command3 
         BackColor       =   &H00000000&
         Caption         =   "Data Penjualan"
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
         Left            =   -61920
         TabIndex        =   17
         Top             =   7800
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00000000&
         Caption         =   "Data Penjualan"
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
         Left            =   -63240
         TabIndex        =   16
         Top             =   7800
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000000&
         Caption         =   "Data Penjualan"
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
         Left            =   -63240
         TabIndex        =   15
         Top             =   7800
         Width           =   1575
      End
      Begin MSChart20Lib.MSChart mscSingle 
         Height          =   7335
         Left            =   -71280
         OleObjectBlob   =   "frmRamal.frx":0070
         TabIndex        =   12
         Top             =   600
         Width           =   11055
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00000000&
         Caption         =   "Lihat Grafik"
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
         Left            =   -61560
         TabIndex        =   11
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00000000&
         Caption         =   "Lihat Grafik"
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
         Left            =   -61560
         TabIndex        =   9
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdKeluar 
         BackColor       =   &H00000000&
         Caption         =   "Keluar"
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
         Left            =   13440
         TabIndex        =   7
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdSingle 
         BackColor       =   &H00000000&
         Caption         =   "Exponensial Smooting Single"
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
         Left            =   7200
         TabIndex        =   6
         Top             =   7800
         Width           =   3015
      End
      Begin VB.CommandButton cmdDouble 
         BackColor       =   &H00000000&
         Caption         =   "Exponensial Smooting Double"
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
         Left            =   10320
         TabIndex        =   5
         Top             =   7800
         Width           =   3015
      End
      Begin MSComctlLib.ListView ListViewTampil 
         Height          =   6135
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   10821
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox txtPenjualan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         TabIndex        =   1
         Top             =   660
         Width           =   2295
      End
      Begin VB.CommandButton cmdAmbil 
         BackColor       =   &H00000000&
         Caption         =   "Ambil Data"
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
         Left            =   4080
         TabIndex        =   2
         Top             =   660
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListViewSingle 
         Height          =   7095
         Left            =   -74760
         TabIndex        =   8
         Top             =   600
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   12515
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListViewDouble 
         Height          =   7095
         Left            =   -74760
         TabIndex        =   10
         Top             =   600
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   12515
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListViewHasilSingle 
         Height          =   3255
         Left            =   -64680
         TabIndex        =   14
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListViewGrafik 
         Height          =   7095
         Left            =   -74520
         TabIndex        =   18
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   12515
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListViewDataSingle 
         Height          =   3255
         Left            =   -64680
         TabIndex        =   20
         Top             =   3960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListViewDataDouble 
         Height          =   2535
         Left            =   -63960
         TabIndex        =   21
         Top             =   4680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   4471
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListViewHasilDouble 
         Height          =   4095
         Left            =   -63960
         TabIndex        =   19
         Top             =   600
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -63840
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   7320
         Width           =   1335
      End
      Begin VB.TextBox txtalfa 
         Height          =   375
         Left            =   -64560
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun Penjualan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   780
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView ListViewCadangan 
      Height          =   6135
      Left            =   15360
      TabIndex        =   13
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10821
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmRamal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CITY_CHART(ByVal chart As MSChart)
'Set chart.DataSource = Me.ListViewDataDouble.ListItems.Count

chart.AllowSelections = False
chart.DoSetCursor = True
chart.MousePointer = VtMousePointerArrowQuestion
chart.chartType = VtChChartType2dLine

chart.ShowLegend = True

With chart.Plot.SeriesCollection(1)
    .LegendText = "saya"
End With
With chart.Plot.SeriesCollection(2)
    .LegendText = "dua"
End With
With chart.Legend
    .Location.LocationType = VtChLocationTypeTop
    .TextLayout.HorzAlignment = VtHorizontalAlignmentCenter
    .VtFont.VtColor.Set 255, 0, 0
    .Backdrop.Fill.Style = VtFillStyleBrush
    .Backdrop.Fill.Brush.Style = VtBrushStyleSolid
    .Backdrop.Fill.Brush.FillColor.Set 219, 230, 255
End With

'mengatur judul grafik
chart.Title = "Metode Single"
With chart.Title.VtFont
    .Name = "Calibri"
    .Size = 20
    .Effect = VtFontEffectUnderline
End With

'mengatur title untuk sumbu x dan y
With chart.Plot.Axis(1, 1)
.AxisTitle.VtFont.Size = 9
.AxisTitle.VtFont.Name = "Calibri"
.AxisTitle.VtFont.Effect = Bold
.AxisTitle.Visible = True
.AxisTitle.Text = "Penjualan"
End With

With chart.Plot.Axis(0, 1)
    .AxisTitle.VtFont.Size = 9
    .AxisTitle.VtFont.Name = "Calibri"
    .AxisTitle.VtFont.Effect = Bold
    .AxisTitle.Visible = True
    .AxisTitle.Text = "Bulan"
End With

chart.Footnote = "Sumber"

'mengatur warna grafik
With chart.Plot.SeriesCollection(1)
        .DataPoints(-1).Brush.FillColor.Set 45, 44, 78
    End With

    ' mengatur warna background grafik
    With chart.Backdrop.Fill
        .Style = VtFillStyleBrush
        .Brush.FillColor.Set 255, 255, 255
    End With
End Sub

Private Sub Grafik()
Dim X(1 To 12, 1 To 3) As Variant

'X(1, 1) = "Jagung"
X(1, 2) = "Penjualan"
X(1, 3) = "Peramalan"
'X(1, 4) = "kedelai"
'X(1, 5) = "Singkong"
'X(1, 6) = "Tebu"

X(2, 1) = "Januari"
X(2, 2) = 2
X(2, 3) = 5
'X(2, 4) = 9
'X(2, 5) = 10
'X(2, 6) = 14

X(3, 1) = "Februari"
X(3, 2) = 4
X(3, 3) = 6
'X(3, 4) = 10
'X(3, 5) = 8
'X(3, 6) = 19

X(4, 1) = "Maret"
X(4, 2) = 9
X(4, 3) = 10

X(5, 1) = "April"
X(5, 2) = 18
X(5, 3) = 11

X(6, 1) = "Mei"
X(6, 2) = 10
X(6, 3) = 8

X(7, 1) = "Juni"
X(7, 2) = 10
X(7, 3) = 8

X(8, 1) = "Juli"
X(8, 2) = 10
X(8, 3) = 8

X(9, 1) = "Agustus"
X(9, 2) = 10
X(9, 3) = 8

X(10, 1) = "September"
X(10, 2) = 10
X(10, 3) = 8

X(11, 1) = "Oktober"
X(11, 2) = 10
X(11, 3) = 8

X(11, 1) = "Nopember"
X(11, 2) = 10
X(11, 3) = 8

X(12, 1) = "Desember"
X(12, 2) = 10
X(12, 3) = 8

mscSingle.ChartData = X
mscSingle.ShowLegend = True 'menampilkan legend
mscSingle.chartType = VtChChartType2dLine 'menampilkan tipe grafik
'mscSingle.Footnote = "Sumber dari "
mscSingle.Title = "Metode Single"

With mscSingle.Plot.Axis(1, 1)
    .AxisTitle.Text = "Penjualan"
End With

With mscSingle.Plot.Axis(0, 1)
    .AxisTitle.Text = "Bulan"
End With

End Sub
Private Sub TabMati()
Me.tabPenjualan.Tab = 0
Me.tabPenjualan.TabEnabled(1) = False
Me.tabPenjualan.TabEnabled(2) = False
Me.tabPenjualan.TabEnabled(3) = False
End Sub
Private Sub metodeSingle()
Set rsPenjualan = koneksi.Execute("select * from penjualan where year(penjualan.bulan)='" & Me.txtPenjualan.Text & "'")

Me.ListViewSingle.ListItems.Clear
Me.ListViewSingle.ColumnHeaders.Clear
Me.ListViewHasilSingle.ListItems.Clear
Me.ListViewHasilSingle.ColumnHeaders.Clear

'nama = Me.ListViewSingle.ColumnHeaders.Add(1, , "Bulan", 1000, 0)
'nama = Me.ListViewCadangan.ColumnHeaders.Add(2, , "Penjualan", 1000, 0)


'For a = 1 To 12
    nama = Me.ListViewSingle.ColumnHeaders.Add(1, , "NO", 500, 0)
    nama = Me.ListViewSingle.ColumnHeaders.Add(2, , "BULAN", 1500, 0)
    nama = Me.ListViewSingle.ColumnHeaders.Add(3, , "PENJUALAN ", 1500, 1)
    nama = Me.ListViewSingle.ColumnHeaders.Add(4, , "Xi-1", 1500, 1)
    nama = Me.ListViewSingle.ColumnHeaders.Add(5, , "(Xi-Xi-1)", 1500, 1)
    nama = Me.ListViewSingle.ColumnHeaders.Add(6, , "|Xi-Xi-1|", 1500, 1)
    nama = Me.ListViewSingle.ColumnHeaders.Add(7, , "Kesalahan Relatif", 1500, 1)
    
    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(1, , "Alfa", 700, 0)
    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(2, , "Ramalan", 1500, 0)
    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(3, , "Jumlah", 1500, 0)
    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(4, , "NF", 1500, 0)
    
'Next
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(1).SubItems(1) = "Alfa 0.1"

urut = 1
While Not rsPenjualan.EOF
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , urut)
    Me.ListViewSingle.ListItems(urut + 1).SubItems(2) = rsPenjualan.Fields(2)
    rsPenjualan.MoveNext
    urut = urut + 1
Wend
For a = 1 To 9
    Set data = Me.ListViewHasilSingle.ListItems.Add(, , "0." & a)

Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(2).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(3).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(4).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(5).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(6).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(7).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(8).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(9).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(10).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(11).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(12).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(13).SubItems(1) = "DESEMBER"


Me.ListViewSingle.ListItems(2).SubItems(3) = 0
Me.ListViewSingle.ListItems(2).SubItems(4) = 0
Me.ListViewSingle.ListItems(2).SubItems(5) = 0
Me.ListViewSingle.ListItems(2).SubItems(6) = 0

Me.ListViewSingle.ListItems(3).SubItems(3) = Me.ListViewSingle.ListItems(2).SubItems(2)
Me.ListViewSingle.ListItems(3).SubItems(4) = Me.ListViewSingle.ListItems(3).SubItems(2) - Me.ListViewSingle.ListItems(3).SubItems(3)
Me.ListViewSingle.ListItems(3).SubItems(5) = Abs(Me.ListViewSingle.ListItems(3).SubItems(4))
Me.ListViewSingle.ListItems(3).SubItems(6) = Me.ListViewSingle.ListItems(3).SubItems(5) / Me.ListViewSingle.ListItems(3).SubItems(2)

For a = 4 To 13
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.1 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.9 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next


For a = 4 To 13
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 4 To 13
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next
hasil = 0
For a = 4 To 13
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(1).SubItems(1) = (0.1 * Me.ListViewSingle.ListItems(13).SubItems(2)) + (0.9 * Me.ListViewSingle.ListItems(13).SubItems(3))
Me.ListViewHasilSingle.ListItems(1).SubItems(2) = hasil + Me.ListViewSingle.ListItems(3).SubItems(6)
Me.ListViewHasilSingle.ListItems(1).SubItems(3) = (Me.ListViewHasilSingle.ListItems(1).SubItems(2) / 11) * 100

Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(15).SubItems(1) = "Alfa 0.2"

'urut = 1
'jml = 16
For a = 16 To 27
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 15)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 15).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(16).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(17).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(18).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(19).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(20).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(21).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(22).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(23).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(24).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(25).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(26).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(27).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(16).SubItems(3) = 0
Me.ListViewSingle.ListItems(16).SubItems(4) = 0
Me.ListViewSingle.ListItems(16).SubItems(5) = 0
Me.ListViewSingle.ListItems(16).SubItems(6) = 0

Me.ListViewSingle.ListItems(17).SubItems(3) = Me.ListViewSingle.ListItems(16).SubItems(2)
Me.ListViewSingle.ListItems(17).SubItems(4) = Me.ListViewSingle.ListItems(17).SubItems(2) - Me.ListViewSingle.ListItems(17).SubItems(3)
Me.ListViewSingle.ListItems(17).SubItems(5) = Abs(Me.ListViewSingle.ListItems(17).SubItems(4))
Me.ListViewSingle.ListItems(17).SubItems(6) = Me.ListViewSingle.ListItems(17).SubItems(5) / Me.ListViewSingle.ListItems(17).SubItems(2)

For a = 18 To 27
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.2 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.8 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 18 To 27
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 18 To 27
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next
hasil = 0
For a = 18 To 27
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(2).SubItems(1) = (0.2 * Me.ListViewSingle.ListItems(27).SubItems(2)) + (0.8 * Me.ListViewSingle.ListItems(27).SubItems(3))
Me.ListViewHasilSingle.ListItems(2).SubItems(2) = hasil + Me.ListViewSingle.ListItems(17).SubItems(6)
Me.ListViewHasilSingle.ListItems(2).SubItems(3) = (Me.ListViewHasilSingle.ListItems(2).SubItems(2) / 11) * 100


Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(29).SubItems(1) = "Alfa 0.3"

'urut = 1
'jml = 16
For a = 30 To 41
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 29)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 29).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(30).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(31).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(32).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(33).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(34).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(35).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(36).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(37).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(38).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(39).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(40).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(41).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(30).SubItems(3) = 0
Me.ListViewSingle.ListItems(30).SubItems(4) = 0
Me.ListViewSingle.ListItems(30).SubItems(5) = 0
Me.ListViewSingle.ListItems(30).SubItems(6) = 0

Me.ListViewSingle.ListItems(31).SubItems(3) = Me.ListViewSingle.ListItems(30).SubItems(2)
Me.ListViewSingle.ListItems(31).SubItems(4) = Me.ListViewSingle.ListItems(31).SubItems(2) - Me.ListViewSingle.ListItems(31).SubItems(3)
Me.ListViewSingle.ListItems(31).SubItems(5) = Abs(Me.ListViewSingle.ListItems(31).SubItems(4))
Me.ListViewSingle.ListItems(31).SubItems(6) = Me.ListViewSingle.ListItems(31).SubItems(5) / Me.ListViewSingle.ListItems(31).SubItems(2)

For a = 32 To 41
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.3 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.7 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 32 To 41
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 32 To 41
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next

hasil = 0
For a = 32 To 41
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(3).SubItems(1) = (0.3 * Me.ListViewSingle.ListItems(41).SubItems(2)) + (0.7 * Me.ListViewSingle.ListItems(41).SubItems(3))
Me.ListViewHasilSingle.ListItems(3).SubItems(2) = hasil + Me.ListViewSingle.ListItems(31).SubItems(6)
Me.ListViewHasilSingle.ListItems(3).SubItems(3) = (Me.ListViewHasilSingle.ListItems(3).SubItems(2) / 11) * 100


Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(43).SubItems(1) = "Alfa 0.4"

'urut = 1
'jml = 16
For a = 44 To 55
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 43)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 43).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(44).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(45).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(46).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(47).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(48).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(49).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(50).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(51).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(52).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(53).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(54).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(55).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(44).SubItems(3) = 0
Me.ListViewSingle.ListItems(44).SubItems(4) = 0
Me.ListViewSingle.ListItems(44).SubItems(5) = 0
Me.ListViewSingle.ListItems(44).SubItems(6) = 0

Me.ListViewSingle.ListItems(45).SubItems(3) = Me.ListViewSingle.ListItems(44).SubItems(2)
Me.ListViewSingle.ListItems(45).SubItems(4) = Me.ListViewSingle.ListItems(45).SubItems(2) - Me.ListViewSingle.ListItems(45).SubItems(3)
Me.ListViewSingle.ListItems(45).SubItems(5) = Abs(Me.ListViewSingle.ListItems(45).SubItems(4))
Me.ListViewSingle.ListItems(45).SubItems(6) = Me.ListViewSingle.ListItems(45).SubItems(5) / Me.ListViewSingle.ListItems(45).SubItems(2)

For a = 46 To 55
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.4 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.6 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 46 To 55
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 46 To 55
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next

hasil = 0
For a = 46 To 55
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(4).SubItems(1) = (0.4 * Me.ListViewSingle.ListItems(55).SubItems(2)) + (0.6 * Me.ListViewSingle.ListItems(13).SubItems(5))
Me.ListViewHasilSingle.ListItems(4).SubItems(2) = hasil + Me.ListViewSingle.ListItems(45).SubItems(6)
Me.ListViewHasilSingle.ListItems(4).SubItems(3) = (Me.ListViewHasilSingle.ListItems(4).SubItems(2) / 11) * 100


Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(57).SubItems(1) = "Alfa 0.5"

'urut = 1
'jml = 16
For a = 58 To 69
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 57)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 57).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(58).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(59).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(60).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(61).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(62).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(63).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(64).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(65).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(66).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(67).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(68).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(69).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(58).SubItems(3) = 0
Me.ListViewSingle.ListItems(58).SubItems(4) = 0
Me.ListViewSingle.ListItems(58).SubItems(5) = 0
Me.ListViewSingle.ListItems(58).SubItems(6) = 0

Me.ListViewSingle.ListItems(59).SubItems(3) = Me.ListViewSingle.ListItems(58).SubItems(2)
Me.ListViewSingle.ListItems(59).SubItems(4) = Me.ListViewSingle.ListItems(59).SubItems(2) - Me.ListViewSingle.ListItems(59).SubItems(3)
Me.ListViewSingle.ListItems(59).SubItems(5) = Abs(Me.ListViewSingle.ListItems(59).SubItems(4))
Me.ListViewSingle.ListItems(59).SubItems(6) = Me.ListViewSingle.ListItems(59).SubItems(5) / Me.ListViewSingle.ListItems(59).SubItems(2)

For a = 60 To 69
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.5 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.5 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 60 To 69
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 60 To 69
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next

hasil = 0
For a = 60 To 69
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(5).SubItems(1) = (0.5 * Me.ListViewSingle.ListItems(69).SubItems(2)) + (0.5 * Me.ListViewSingle.ListItems(69).SubItems(3))
Me.ListViewHasilSingle.ListItems(5).SubItems(2) = hasil + Me.ListViewSingle.ListItems(59).SubItems(6)
Me.ListViewHasilSingle.ListItems(5).SubItems(3) = (Me.ListViewHasilSingle.ListItems(5).SubItems(2) / 11) * 100


Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(71).SubItems(1) = "Alfa 0.6"

'urut = 1
'jml = 16
For a = 72 To 83
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 71)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 71).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(72).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(73).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(74).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(75).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(76).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(77).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(78).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(79).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(80).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(81).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(82).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(83).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(72).SubItems(3) = 0
Me.ListViewSingle.ListItems(72).SubItems(4) = 0
Me.ListViewSingle.ListItems(72).SubItems(5) = 0
Me.ListViewSingle.ListItems(72).SubItems(6) = 0

Me.ListViewSingle.ListItems(73).SubItems(3) = Me.ListViewSingle.ListItems(73).SubItems(2)
Me.ListViewSingle.ListItems(73).SubItems(4) = Me.ListViewSingle.ListItems(73).SubItems(2) - Me.ListViewSingle.ListItems(73).SubItems(3)
Me.ListViewSingle.ListItems(73).SubItems(5) = Abs(Me.ListViewSingle.ListItems(73).SubItems(4))
Me.ListViewSingle.ListItems(73).SubItems(6) = Me.ListViewSingle.ListItems(73).SubItems(5) / Me.ListViewSingle.ListItems(73).SubItems(2)

For a = 74 To 83
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.6 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.4 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 74 To 83
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 74 To 83
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next

hasil = 0
For a = 74 To 83
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(6).SubItems(1) = (0.6 * Me.ListViewSingle.ListItems(83).SubItems(2)) + (0.4 * Me.ListViewSingle.ListItems(83).SubItems(3))
Me.ListViewHasilSingle.ListItems(6).SubItems(2) = hasil + Me.ListViewSingle.ListItems(74).SubItems(6)
Me.ListViewHasilSingle.ListItems(6).SubItems(3) = (Me.ListViewHasilSingle.ListItems(6).SubItems(2) / 11) * 100



Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(85).SubItems(1) = "Alfa 0.7"

'urut = 1
'jml = 16
For a = 86 To 97
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 85)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 85).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(86).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(87).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(88).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(89).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(90).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(91).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(92).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(93).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(94).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(95).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(96).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(97).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(86).SubItems(3) = 0
Me.ListViewSingle.ListItems(86).SubItems(4) = 0
Me.ListViewSingle.ListItems(86).SubItems(5) = 0
Me.ListViewSingle.ListItems(86).SubItems(6) = 0

Me.ListViewSingle.ListItems(87).SubItems(3) = Me.ListViewSingle.ListItems(86).SubItems(2)
Me.ListViewSingle.ListItems(87).SubItems(4) = Me.ListViewSingle.ListItems(87).SubItems(2) - Me.ListViewSingle.ListItems(87).SubItems(3)
Me.ListViewSingle.ListItems(87).SubItems(5) = Abs(Me.ListViewSingle.ListItems(87).SubItems(4))
Me.ListViewSingle.ListItems(87).SubItems(6) = Me.ListViewSingle.ListItems(87).SubItems(5) / Me.ListViewSingle.ListItems(87).SubItems(2)

For a = 88 To 97
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.7 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.3 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 88 To 97
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 88 To 97
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next

hasil = 0
For a = 88 To 97
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(7).SubItems(1) = (0.7 * Me.ListViewSingle.ListItems(97).SubItems(2)) + (0.3 * Me.ListViewSingle.ListItems(97).SubItems(3))
Me.ListViewHasilSingle.ListItems(7).SubItems(2) = hasil + Me.ListViewSingle.ListItems(87).SubItems(6)
Me.ListViewHasilSingle.ListItems(7).SubItems(3) = (Me.ListViewHasilSingle.ListItems(7).SubItems(2) / 11) * 100


Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(99).SubItems(1) = "Alfa 0.8"

'urut = 1
'jml = 16
For a = 100 To 111
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 99)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 99).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(100).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(101).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(102).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(103).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(104).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(105).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(106).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(107).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(108).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(109).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(110).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(111).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(100).SubItems(3) = 0
Me.ListViewSingle.ListItems(100).SubItems(4) = 0
Me.ListViewSingle.ListItems(100).SubItems(5) = 0
Me.ListViewSingle.ListItems(100).SubItems(6) = 0

Me.ListViewSingle.ListItems(101).SubItems(3) = Me.ListViewSingle.ListItems(100).SubItems(2)
Me.ListViewSingle.ListItems(101).SubItems(4) = Me.ListViewSingle.ListItems(101).SubItems(2) - Me.ListViewSingle.ListItems(101).SubItems(3)
Me.ListViewSingle.ListItems(101).SubItems(5) = Abs(Me.ListViewSingle.ListItems(101).SubItems(4))
Me.ListViewSingle.ListItems(101).SubItems(6) = Me.ListViewSingle.ListItems(101).SubItems(5) / Me.ListViewSingle.ListItems(101).SubItems(2)

For a = 102 To 111
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.8 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.2 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 102 To 111
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 102 To 111
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next

hasil = 0
For a = 102 To 111
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(8).SubItems(1) = (0.8 * Me.ListViewSingle.ListItems(111).SubItems(2)) + (0.1 * Me.ListViewSingle.ListItems(111).SubItems(3))
Me.ListViewHasilSingle.ListItems(8).SubItems(2) = hasil + Me.ListViewSingle.ListItems(101).SubItems(6)
Me.ListViewHasilSingle.ListItems(8).SubItems(3) = (Me.ListViewHasilSingle.ListItems(8).SubItems(2) / 11) * 100


Set data = Me.ListViewSingle.ListItems.Add(, , "")
Set data = Me.ListViewSingle.ListItems.Add(, , "")
Me.ListViewSingle.ListItems(113).SubItems(1) = "Alfa 0.9"

'urut = 1
'jml = 16
For a = 114 To 125
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewSingle.ListItems.Add(, , a - 113)
'    Me.ListViewSingle.ListItems(a).SubItems(2) = Me.ListViewCadangan.ListItems(16).Text
'    rsPenjualan.MoveNext
'    urut = urut + 1
'    jml = jml + 1
Next
For a = 1 To 12
Me.ListViewSingle.ListItems(a + 113).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next
'Set data = Me.ListViewSingle.ListItems.Add(, , "")
'data.SubItems(2) = 0

Me.ListViewSingle.ListItems(114).SubItems(1) = "JANUARI"
Me.ListViewSingle.ListItems(115).SubItems(1) = "FEBRUARI"
Me.ListViewSingle.ListItems(116).SubItems(1) = "MARET"
Me.ListViewSingle.ListItems(117).SubItems(1) = "APRIL"
Me.ListViewSingle.ListItems(118).SubItems(1) = "MEI"
Me.ListViewSingle.ListItems(119).SubItems(1) = "JUNI"
Me.ListViewSingle.ListItems(120).SubItems(1) = "JULI"
Me.ListViewSingle.ListItems(121).SubItems(1) = "AGUSTUS"
Me.ListViewSingle.ListItems(122).SubItems(1) = "SEPTEMBER"
Me.ListViewSingle.ListItems(123).SubItems(1) = "OKTOBER"
Me.ListViewSingle.ListItems(124).SubItems(1) = "NOPEMBER"
Me.ListViewSingle.ListItems(125).SubItems(1) = "DESEMBER"

Me.ListViewSingle.ListItems(114).SubItems(3) = 0
Me.ListViewSingle.ListItems(114).SubItems(4) = 0
Me.ListViewSingle.ListItems(114).SubItems(5) = 0
Me.ListViewSingle.ListItems(114).SubItems(6) = 0

Me.ListViewSingle.ListItems(115).SubItems(3) = Me.ListViewSingle.ListItems(114).SubItems(2)
Me.ListViewSingle.ListItems(115).SubItems(4) = Me.ListViewSingle.ListItems(115).SubItems(2) - Me.ListViewSingle.ListItems(115).SubItems(3)
Me.ListViewSingle.ListItems(115).SubItems(5) = Abs(Me.ListViewSingle.ListItems(115).SubItems(4))
Me.ListViewSingle.ListItems(115).SubItems(6) = Me.ListViewSingle.ListItems(115).SubItems(5) / Me.ListViewSingle.ListItems(115).SubItems(2)

For a = 116 To 125
    Me.ListViewSingle.ListItems(a).SubItems(3) = (0.9 * Me.ListViewSingle.ListItems(a - 1).SubItems(2)) + (0.1 * Me.ListViewSingle.ListItems(a - 1).SubItems(3))
Next

For a = 116 To 125
    Me.ListViewSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a).SubItems(2) - Me.ListViewSingle.ListItems(a).SubItems(3)
Next

For a = 116 To 125
    Me.ListViewSingle.ListItems(a).SubItems(5) = Abs(ListViewSingle.ListItems(a).SubItems(4))
Next

hasil = 0
For a = 116 To 125
    Me.ListViewSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a).SubItems(5) / Me.ListViewSingle.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewSingle.ListItems(a).SubItems(6)
Next

Me.ListViewHasilSingle.ListItems(9).SubItems(1) = (0.9 * Me.ListViewSingle.ListItems(125).SubItems(2)) + (0.1 * Me.ListViewSingle.ListItems(125).SubItems(3))
Me.ListViewHasilSingle.ListItems(9).SubItems(2) = hasil + Me.ListViewSingle.ListItems(115).SubItems(6)
Me.ListViewHasilSingle.ListItems(9).SubItems(3) = (Me.ListViewHasilSingle.ListItems(9).SubItems(2) / 11) * 100


Me.ListViewDataSingle.ListItems.Clear
Me.ListViewDataSingle.ColumnHeaders.Clear

nama = Me.ListViewDataSingle.ColumnHeaders.Add(1, , "Bulan", 800, 0)
nama = Me.ListViewDataSingle.ColumnHeaders.Add(2, , "Penjualan", 1500, 0)
For a = 1 To 9
    nama = Me.ListViewDataSingle.ColumnHeaders.Add(a + 2, , "alfa" & a, 1500, 0)
Next
For a = 1 To 12
    Set data = Me.ListViewDataSingle.ListItems.Add(, , a)
Next

For a = 1 To 12
    Me.ListViewDataSingle.ListItems(a).SubItems(1) = Me.ListViewSingle.ListItems(a + 1).SubItems(2)
Next

For a = 1 To 12 '1
    Me.ListViewDataSingle.ListItems(a).SubItems(2) = Me.ListViewSingle.ListItems(a + 1).SubItems(3)
Next

For a = 1 To 12 '2
    Me.ListViewDataSingle.ListItems(a).SubItems(3) = Me.ListViewSingle.ListItems(a + 15).SubItems(3)
Next

For a = 1 To 12 '3
    Me.ListViewDataSingle.ListItems(a).SubItems(4) = Me.ListViewSingle.ListItems(a + 29).SubItems(3)
Next

For a = 1 To 12 '4
    Me.ListViewDataSingle.ListItems(a).SubItems(5) = Me.ListViewSingle.ListItems(a + 43).SubItems(3)
Next

For a = 1 To 12 '5
    Me.ListViewDataSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a + 57).SubItems(3)
Next

For a = 1 To 12 '6
    Me.ListViewDataSingle.ListItems(a).SubItems(7) = Me.ListViewSingle.ListItems(a + 71).SubItems(3)
Next

For a = 1 To 12 '7
    Me.ListViewDataSingle.ListItems(a).SubItems(8) = Me.ListViewSingle.ListItems(a + 85).SubItems(3)
Next

For a = 1 To 12 '8
    Me.ListViewDataSingle.ListItems(a).SubItems(9) = Me.ListViewSingle.ListItems(a + 99).SubItems(3)
Next

For a = 1 To 12 '9
    Me.ListViewDataSingle.ListItems(a).SubItems(10) = Me.ListViewSingle.ListItems(a + 113).SubItems(3)
Next

alfaa = 1
nf = Me.ListViewHasilSingle.ListItems(1).SubItems(3)
For a = 2 To 8
If nf < Me.ListViewHasilSingle.ListItems(a).SubItems(3) Then
    alfaa = alfaa
    nf = nf
Else
    nf = Me.ListViewHasilSingle.ListItems(a).SubItems(3)
    alfaa = a
End If
Next
Me.txtalfa.Text = alfaa
End Sub

Private Sub MetodeDouble()
Set rsPenjualan = koneksi.Execute("select * from penjualan where year(penjualan.bulan)='" & Me.txtPenjualan.Text & "'")

Me.ListViewDouble.ListItems.Clear
Me.ListViewDouble.ColumnHeaders.Clear
Me.ListViewHasilDouble.ListItems.Clear
Me.ListViewHasilDouble.ColumnHeaders.Clear
'Me.ListViewHasilSingle.ListItems.Clear
'Me.ListViewHasilSingle.ColumnHeaders.Clear

'nama = Me.ListViewSingle.ColumnHeaders.Add(1, , "Bulan", 1000, 0)
'nama = Me.ListViewCadangan.ColumnHeaders.Add(2, , "Penjualan", 1000, 0)


'For a = 1 To 12
    nama = Me.ListViewDouble.ColumnHeaders.Add(1, , "NO", 500, 0)
    nama = Me.ListViewDouble.ColumnHeaders.Add(2, , "BULAN", 1500, 0)
    nama = Me.ListViewDouble.ColumnHeaders.Add(3, , "PENJUALAN ", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(4, , "s'", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(5, , "s''", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(6, , "a", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(7, , "b", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(8, , "f", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(9, , "e", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(10, , "e absolute", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(11, , "Kesalahan Relatif", 1500, 1)
    nama = Me.ListViewDouble.ColumnHeaders.Add(12, , "Ramalan", 1500, 1)
    
    nama = Me.ListViewHasilDouble.ColumnHeaders.Add(1, , "Alfa", 500, 0)
    nama = Me.ListViewHasilDouble.ColumnHeaders.Add(2, , "Total", 1500, 0)
    nama = Me.ListViewHasilDouble.ColumnHeaders.Add(3, , "NF", 1500, 0)
'    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(1, , "Alfa", 700, 0)
'    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(2, , "Ramalan", 1500, 0)
'    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(3, , "Jumlah", 1500, 0)
'    nama = Me.ListViewHasilSingle.ColumnHeaders.Add(4, , "NF", 1500, 0)


For a = 1 To 9
    Set data = Me.ListViewHasilDouble.ListItems.Add(, , "0." & a)

Next


Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(1).SubItems(1) = "Alfa 0.1"



urut = 1
While Not rsPenjualan.EOF
'    Set data = Me.ListViewSingle.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewDouble.ListItems.Add(, , urut)
    Me.ListViewDouble.ListItems(urut + 1).SubItems(2) = rsPenjualan.Fields(2)
    rsPenjualan.MoveNext
    urut = urut + 1
Wend

Me.ListViewDouble.ListItems(2).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(3).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(4).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(5).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(6).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(7).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(8).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(9).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(10).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(11).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(12).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(13).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(2).SubItems(3) = Me.ListViewDouble.ListItems(2).SubItems(2)
Me.ListViewDouble.ListItems(2).SubItems(4) = Me.ListViewDouble.ListItems(2).SubItems(2)
Me.ListViewDouble.ListItems(2).SubItems(5) = 0
Me.ListViewDouble.ListItems(2).SubItems(6) = 0
Me.ListViewDouble.ListItems(2).SubItems(7) = 0
Me.ListViewDouble.ListItems(2).SubItems(8) = 0
Me.ListViewDouble.ListItems(2).SubItems(9) = 0
Me.ListViewDouble.ListItems(2).SubItems(10) = 0

Me.ListViewDouble.ListItems(3).SubItems(3) = (0.1 * Me.ListViewDouble.ListItems(3).SubItems(2)) + (0.9 * Me.ListViewDouble.ListItems(2).SubItems(3))
Me.ListViewDouble.ListItems(3).SubItems(4) = (0.1 * Me.ListViewDouble.ListItems(3).SubItems(3)) + (0.9 * Me.ListViewDouble.ListItems(2).SubItems(4))
Me.ListViewDouble.ListItems(3).SubItems(5) = (2 * Me.ListViewDouble.ListItems(3).SubItems(3)) - Me.ListViewDouble.ListItems(3).SubItems(4)
Me.ListViewDouble.ListItems(3).SubItems(6) = (0.1 / 0.9) * (Me.ListViewDouble.ListItems(3).SubItems(3) - Me.ListViewDouble.ListItems(3).SubItems(4))
Me.ListViewDouble.ListItems(3).SubItems(7) = 0
Me.ListViewDouble.ListItems(3).SubItems(8) = 0
Me.ListViewDouble.ListItems(3).SubItems(9) = 0
Me.ListViewDouble.ListItems(3).SubItems(10) = 0

For a = 4 To 13
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.1 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.9 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 4 To 13
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.1 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.9 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 4 To 13
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 4 To 13
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.1 / 0.9) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 4 To 13
b = Me.ListViewDouble.ListItems(a - 1).SubItems(5)
c = Me.ListViewDouble.ListItems(a - 1).SubItems(6)
Me.ListViewDouble.ListItems(a).SubItems(7) = b * 1 + c * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 4 To 13
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next
hasil = 0
For a = 4 To 13
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next
Me.ListViewHasilDouble.ListItems(1).SubItems(1) = hasil + Me.ListViewDouble.ListItems(2).SubItems(10)
Me.ListViewHasilDouble.ListItems(1).SubItems(2) = (Me.ListViewHasilDouble.ListItems(1).SubItems(1) / 11) * 100

Me.ListViewDouble.ListItems(2).SubItems(11) = Me.ListViewDouble.ListItems(13).SubItems(5) * 1 + ListViewDouble.ListItems(13).SubItems(6) * 1
For a = 3 To 13
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(13).SubItems(6) + Me.ListViewDouble.ListItems(13).SubItems(7) * (a - 1)
Next



'alfa 2
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(15).SubItems(1) = "Alfa 0.1"

For a = 16 To 27
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 15)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 15).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(16).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(17).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(18).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(19).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(20).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(21).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(22).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(23).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(24).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(25).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(26).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(27).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(16).SubItems(3) = Me.ListViewDouble.ListItems(16).SubItems(2)
Me.ListViewDouble.ListItems(16).SubItems(4) = Me.ListViewDouble.ListItems(16).SubItems(2)
Me.ListViewDouble.ListItems(16).SubItems(5) = 0
Me.ListViewDouble.ListItems(16).SubItems(6) = 0
Me.ListViewDouble.ListItems(16).SubItems(7) = 0
Me.ListViewDouble.ListItems(16).SubItems(8) = 0
Me.ListViewDouble.ListItems(16).SubItems(9) = 0
Me.ListViewDouble.ListItems(16).SubItems(10) = 0

Me.ListViewDouble.ListItems(17).SubItems(3) = (0.2 * Me.ListViewDouble.ListItems(17).SubItems(2)) + (0.8 * Me.ListViewDouble.ListItems(16).SubItems(3))
Me.ListViewDouble.ListItems(17).SubItems(4) = (0.2 * Me.ListViewDouble.ListItems(17).SubItems(3)) + (0.8 * Me.ListViewDouble.ListItems(16).SubItems(4))
Me.ListViewDouble.ListItems(17).SubItems(5) = (2 * Me.ListViewDouble.ListItems(17).SubItems(3)) - Me.ListViewDouble.ListItems(17).SubItems(4)
Me.ListViewDouble.ListItems(17).SubItems(6) = (0.2 / 0.8) * (Me.ListViewDouble.ListItems(17).SubItems(3) - Me.ListViewDouble.ListItems(17).SubItems(4))
Me.ListViewDouble.ListItems(17).SubItems(7) = 0
Me.ListViewDouble.ListItems(17).SubItems(8) = 0
Me.ListViewDouble.ListItems(17).SubItems(9) = 0
Me.ListViewDouble.ListItems(17).SubItems(10) = 0

For a = 18 To 27
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.2 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.8 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 18 To 27
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.2 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.8 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 18 To 27
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 18 To 27
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.2 / 0.8) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 18 To 27
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 18 To 27
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 18 To 27
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next

Me.ListViewHasilDouble.ListItems(2).SubItems(1) = hasil + Me.ListViewDouble.ListItems(17).SubItems(10)
Me.ListViewHasilDouble.ListItems(2).SubItems(2) = (Me.ListViewHasilDouble.ListItems(2).SubItems(1) / 11) * 100

Me.ListViewDouble.ListItems(16).SubItems(11) = Me.ListViewDouble.ListItems(27).SubItems(5) * 1 + ListViewDouble.ListItems(27).SubItems(6) * 1
For a = 17 To 27
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(27).SubItems(6) + Me.ListViewDouble.ListItems(27).SubItems(7) * (a - 15)
Next


'alfa 3
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(29).SubItems(1) = "Alfa 0.3"

For a = 29 To 40
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 28)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 29).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(30).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(31).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(32).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(33).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(34).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(35).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(36).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(37).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(38).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(39).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(40).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(41).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(30).SubItems(3) = Me.ListViewDouble.ListItems(30).SubItems(2)
Me.ListViewDouble.ListItems(30).SubItems(4) = Me.ListViewDouble.ListItems(30).SubItems(2)
Me.ListViewDouble.ListItems(30).SubItems(5) = 0
Me.ListViewDouble.ListItems(30).SubItems(6) = 0
Me.ListViewDouble.ListItems(30).SubItems(7) = 0
Me.ListViewDouble.ListItems(30).SubItems(8) = 0
Me.ListViewDouble.ListItems(30).SubItems(9) = 0
Me.ListViewDouble.ListItems(30).SubItems(10) = 0

Me.ListViewDouble.ListItems(31).SubItems(3) = (0.3 * Me.ListViewDouble.ListItems(31).SubItems(2)) + (0.7 * Me.ListViewDouble.ListItems(30).SubItems(3))
Me.ListViewDouble.ListItems(31).SubItems(4) = (0.3 * Me.ListViewDouble.ListItems(31).SubItems(3)) + (0.7 * Me.ListViewDouble.ListItems(30).SubItems(4))
Me.ListViewDouble.ListItems(31).SubItems(5) = (2 * Me.ListViewDouble.ListItems(31).SubItems(3)) - Me.ListViewDouble.ListItems(31).SubItems(4)
Me.ListViewDouble.ListItems(31).SubItems(6) = (0.3 / 0.7) * (Me.ListViewDouble.ListItems(31).SubItems(3) - Me.ListViewDouble.ListItems(31).SubItems(4))
Me.ListViewDouble.ListItems(31).SubItems(7) = 0
Me.ListViewDouble.ListItems(31).SubItems(8) = 0
Me.ListViewDouble.ListItems(31).SubItems(9) = 0
Me.ListViewDouble.ListItems(31).SubItems(10) = 0

For a = 32 To 41
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.3 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.7 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 32 To 41
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.3 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.7 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 32 To 41
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 32 To 41
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.2 / 0.8) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 32 To 41
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 32 To 41
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 32 To 41
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next
Me.ListViewHasilDouble.ListItems(3).SubItems(1) = hasil + Me.ListViewDouble.ListItems(31).SubItems(10)
Me.ListViewHasilDouble.ListItems(3).SubItems(2) = (Me.ListViewHasilDouble.ListItems(3).SubItems(1) / 11) * 100

'Me.ListViewDouble.ListItems(2).SubItems(11) = Me.ListViewDouble.ListItems(41).SubItems(5) * 1 + ListViewDouble.ListItems(41).SubItems(6) * 1
'For a = 3 To 13
'    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(41).SubItems(6) + Me.ListViewDouble.ListItems(41).SubItems(7) * (a - 1)
'Next

Me.ListViewDouble.ListItems(30).SubItems(11) = Me.ListViewDouble.ListItems(41).SubItems(5) * 1 + ListViewDouble.ListItems(41).SubItems(6) * 1
For a = 31 To 41
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(41).SubItems(6) + Me.ListViewDouble.ListItems(41).SubItems(7) * (a - 30)
Next



'alfa 4
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(43).SubItems(1) = "Alfa 0.4"

For a = 44 To 55
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 43)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 43).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(44).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(45).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(46).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(47).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(48).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(49).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(50).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(51).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(52).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(53).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(54).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(55).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(44).SubItems(3) = Me.ListViewDouble.ListItems(44).SubItems(2)
Me.ListViewDouble.ListItems(44).SubItems(4) = Me.ListViewDouble.ListItems(44).SubItems(2)
Me.ListViewDouble.ListItems(44).SubItems(5) = 0
Me.ListViewDouble.ListItems(44).SubItems(6) = 0
Me.ListViewDouble.ListItems(44).SubItems(7) = 0
Me.ListViewDouble.ListItems(44).SubItems(8) = 0
Me.ListViewDouble.ListItems(44).SubItems(9) = 0
Me.ListViewDouble.ListItems(44).SubItems(10) = 0

Me.ListViewDouble.ListItems(45).SubItems(3) = (0.4 * Me.ListViewDouble.ListItems(45).SubItems(2)) + (0.6 * Me.ListViewDouble.ListItems(44).SubItems(3))
Me.ListViewDouble.ListItems(45).SubItems(4) = (0.4 * Me.ListViewDouble.ListItems(45).SubItems(3)) + (0.6 * Me.ListViewDouble.ListItems(44).SubItems(4))
Me.ListViewDouble.ListItems(45).SubItems(5) = (2 * Me.ListViewDouble.ListItems(45).SubItems(3)) - Me.ListViewDouble.ListItems(45).SubItems(4)
Me.ListViewDouble.ListItems(45).SubItems(6) = (0.4 / 0.6) * (Me.ListViewDouble.ListItems(45).SubItems(3) - Me.ListViewDouble.ListItems(45).SubItems(4))
Me.ListViewDouble.ListItems(45).SubItems(7) = 0
Me.ListViewDouble.ListItems(45).SubItems(8) = 0
Me.ListViewDouble.ListItems(45).SubItems(9) = 0
Me.ListViewDouble.ListItems(45).SubItems(10) = 0

For a = 46 To 55
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.4 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.6 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 46 To 55
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.4 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.6 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 46 To 55
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 46 To 55
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.4 / 0.6) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 46 To 55
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 46 To 55
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 46 To 55
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next

Me.ListViewHasilDouble.ListItems(4).SubItems(1) = hasil + Me.ListViewDouble.ListItems(45).SubItems(10)
Me.ListViewHasilDouble.ListItems(4).SubItems(2) = (Me.ListViewHasilDouble.ListItems(4).SubItems(1) / 11) * 100

Me.ListViewDouble.ListItems(44).SubItems(11) = Me.ListViewDouble.ListItems(55).SubItems(5) * 1 + ListViewDouble.ListItems(55).SubItems(6) * 1
For a = 45 To 55
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(55).SubItems(6) + Me.ListViewDouble.ListItems(55).SubItems(7) * (a - 43)
Next



'alfa 5
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(57).SubItems(1) = "Alfa 0.5"

For a = 58 To 69
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 43)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 57).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(58).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(59).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(60).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(61).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(62).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(63).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(64).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(65).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(66).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(67).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(68).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(69).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(58).SubItems(3) = Me.ListViewDouble.ListItems(58).SubItems(2)
Me.ListViewDouble.ListItems(58).SubItems(4) = Me.ListViewDouble.ListItems(58).SubItems(2)
Me.ListViewDouble.ListItems(58).SubItems(5) = 0
Me.ListViewDouble.ListItems(58).SubItems(6) = 0
Me.ListViewDouble.ListItems(58).SubItems(7) = 0
Me.ListViewDouble.ListItems(58).SubItems(8) = 0
Me.ListViewDouble.ListItems(58).SubItems(9) = 0
Me.ListViewDouble.ListItems(58).SubItems(10) = 0

Me.ListViewDouble.ListItems(59).SubItems(3) = (0.5 * Me.ListViewDouble.ListItems(59).SubItems(2)) + (0.6 * Me.ListViewDouble.ListItems(58).SubItems(3))
Me.ListViewDouble.ListItems(59).SubItems(4) = (0.5 * Me.ListViewDouble.ListItems(59).SubItems(3)) + (0.6 * Me.ListViewDouble.ListItems(58).SubItems(4))
Me.ListViewDouble.ListItems(59).SubItems(5) = (2 * Me.ListViewDouble.ListItems(59).SubItems(3)) - Me.ListViewDouble.ListItems(59).SubItems(4)
Me.ListViewDouble.ListItems(59).SubItems(6) = (0.5 / 0.5) * (Me.ListViewDouble.ListItems(59).SubItems(3) - Me.ListViewDouble.ListItems(59).SubItems(4))
Me.ListViewDouble.ListItems(59).SubItems(7) = 0
Me.ListViewDouble.ListItems(59).SubItems(8) = 0
Me.ListViewDouble.ListItems(59).SubItems(9) = 0
Me.ListViewDouble.ListItems(59).SubItems(10) = 0

For a = 60 To 69
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.5 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.5 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 60 To 69
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.5 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.5 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 60 To 69
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 60 To 69
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.5 / 0.5) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 60 To 69
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 60 To 69
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 60 To 69
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next

Me.ListViewHasilDouble.ListItems(5).SubItems(1) = hasil + Me.ListViewDouble.ListItems(59).SubItems(10)
Me.ListViewHasilDouble.ListItems(5).SubItems(2) = (Me.ListViewHasilDouble.ListItems(5).SubItems(1) / 11) * 100

Me.ListViewDouble.ListItems(58).SubItems(11) = Me.ListViewDouble.ListItems(69).SubItems(5) * 1 + ListViewDouble.ListItems(69).SubItems(6) * 1
For a = 59 To 69
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(69).SubItems(6) + Me.ListViewDouble.ListItems(69).SubItems(7) * (a - 57)
Next



'alfa 6
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(71).SubItems(1) = "Alfa 0.6"

For a = 72 To 83
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 71)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 71).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(72).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(73).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(74).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(75).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(76).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(77).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(78).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(79).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(80).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(81).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(82).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(83).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(72).SubItems(3) = Me.ListViewDouble.ListItems(72).SubItems(2)
Me.ListViewDouble.ListItems(72).SubItems(4) = Me.ListViewDouble.ListItems(72).SubItems(2)
Me.ListViewDouble.ListItems(72).SubItems(5) = 0
Me.ListViewDouble.ListItems(72).SubItems(6) = 0
Me.ListViewDouble.ListItems(72).SubItems(7) = 0
Me.ListViewDouble.ListItems(72).SubItems(8) = 0
Me.ListViewDouble.ListItems(72).SubItems(9) = 0
Me.ListViewDouble.ListItems(72).SubItems(10) = 0

Me.ListViewDouble.ListItems(73).SubItems(3) = (0.6 * Me.ListViewDouble.ListItems(73).SubItems(2)) + (0.4 * Me.ListViewDouble.ListItems(72).SubItems(3))
Me.ListViewDouble.ListItems(73).SubItems(4) = (0.6 * Me.ListViewDouble.ListItems(73).SubItems(3)) + (0.4 * Me.ListViewDouble.ListItems(72).SubItems(4))
Me.ListViewDouble.ListItems(73).SubItems(5) = (2 * Me.ListViewDouble.ListItems(73).SubItems(3)) - Me.ListViewDouble.ListItems(73).SubItems(4)
Me.ListViewDouble.ListItems(73).SubItems(6) = (0.6 / 0.4) * (Me.ListViewDouble.ListItems(73).SubItems(3) - Me.ListViewDouble.ListItems(73).SubItems(4))
Me.ListViewDouble.ListItems(73).SubItems(7) = 0
Me.ListViewDouble.ListItems(73).SubItems(8) = 0
Me.ListViewDouble.ListItems(73).SubItems(9) = 0
Me.ListViewDouble.ListItems(73).SubItems(10) = 0

For a = 74 To 83
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.6 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.4 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 74 To 83
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.6 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.4 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 74 To 83
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 74 To 83
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.6 / 0.4) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 74 To 83
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 74 To 83
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 74 To 83
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next

Me.ListViewHasilDouble.ListItems(6).SubItems(1) = hasil + Me.ListViewDouble.ListItems(73).SubItems(10)
Me.ListViewHasilDouble.ListItems(6).SubItems(2) = (Me.ListViewHasilDouble.ListItems(6).SubItems(1) / 11) * 100

Me.ListViewDouble.ListItems(72).SubItems(11) = Me.ListViewDouble.ListItems(83).SubItems(5) * 1 + ListViewDouble.ListItems(83).SubItems(6) * 1
For a = 73 To 83
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(83).SubItems(6) + Me.ListViewDouble.ListItems(83).SubItems(7) * (a - 71)
Next



'alfa 7
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(85).SubItems(1) = "Alfa 0.7"

For a = 86 To 97
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 85)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 85).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(86).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(87).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(88).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(89).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(90).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(91).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(92).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(93).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(94).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(95).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(96).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(97).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(86).SubItems(3) = Me.ListViewDouble.ListItems(86).SubItems(2)
Me.ListViewDouble.ListItems(86).SubItems(4) = Me.ListViewDouble.ListItems(86).SubItems(2)
Me.ListViewDouble.ListItems(86).SubItems(5) = 0
Me.ListViewDouble.ListItems(86).SubItems(6) = 0
Me.ListViewDouble.ListItems(86).SubItems(7) = 0
Me.ListViewDouble.ListItems(86).SubItems(8) = 0
Me.ListViewDouble.ListItems(86).SubItems(9) = 0
Me.ListViewDouble.ListItems(86).SubItems(10) = 0

Me.ListViewDouble.ListItems(87).SubItems(3) = (0.7 * Me.ListViewDouble.ListItems(87).SubItems(2)) + (0.3 * Me.ListViewDouble.ListItems(86).SubItems(3))
Me.ListViewDouble.ListItems(87).SubItems(4) = (0.7 * Me.ListViewDouble.ListItems(87).SubItems(3)) + (0.3 * Me.ListViewDouble.ListItems(86).SubItems(4))
Me.ListViewDouble.ListItems(87).SubItems(5) = (2 * Me.ListViewDouble.ListItems(87).SubItems(3)) - Me.ListViewDouble.ListItems(87).SubItems(4)
Me.ListViewDouble.ListItems(87).SubItems(6) = (0.7 / 0.3) * (Me.ListViewDouble.ListItems(87).SubItems(3) - Me.ListViewDouble.ListItems(87).SubItems(4))
Me.ListViewDouble.ListItems(87).SubItems(7) = 0
Me.ListViewDouble.ListItems(87).SubItems(8) = 0
Me.ListViewDouble.ListItems(87).SubItems(9) = 0
Me.ListViewDouble.ListItems(87).SubItems(10) = 0

For a = 88 To 97
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.7 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.3 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 88 To 97
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.7 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.3 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 88 To 97
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 88 To 97
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.7 / 0.3) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 88 To 97
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 88 To 97
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 88 To 97
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next

Me.ListViewHasilDouble.ListItems(7).SubItems(1) = hasil + Me.ListViewDouble.ListItems(87).SubItems(10)
Me.ListViewHasilDouble.ListItems(7).SubItems(2) = (Me.ListViewHasilDouble.ListItems(7).SubItems(1) / 11) * 100

Me.ListViewDouble.ListItems(86).SubItems(11) = Me.ListViewDouble.ListItems(97).SubItems(5) * 1 + ListViewDouble.ListItems(97).SubItems(6) * 1
For a = 87 To 97
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(97).SubItems(6) + Me.ListViewDouble.ListItems(97).SubItems(7) * (a - 85)
Next



'alfa 8
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(99).SubItems(1) = "Alfa 0.8"

For a = 100 To 111
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 99)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 99).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(100).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(101).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(102).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(103).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(104).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(105).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(106).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(107).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(108).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(109).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(110).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(111).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(100).SubItems(3) = Me.ListViewDouble.ListItems(100).SubItems(2)
Me.ListViewDouble.ListItems(100).SubItems(4) = Me.ListViewDouble.ListItems(100).SubItems(2)
Me.ListViewDouble.ListItems(100).SubItems(5) = 0
Me.ListViewDouble.ListItems(100).SubItems(6) = 0
Me.ListViewDouble.ListItems(100).SubItems(7) = 0
Me.ListViewDouble.ListItems(100).SubItems(8) = 0
Me.ListViewDouble.ListItems(100).SubItems(9) = 0
Me.ListViewDouble.ListItems(100).SubItems(10) = 0

Me.ListViewDouble.ListItems(101).SubItems(3) = (0.8 * Me.ListViewDouble.ListItems(101).SubItems(2)) + (0.2 * Me.ListViewDouble.ListItems(100).SubItems(3))
Me.ListViewDouble.ListItems(101).SubItems(4) = (0.8 * Me.ListViewDouble.ListItems(101).SubItems(3)) + (0.2 * Me.ListViewDouble.ListItems(100).SubItems(4))
Me.ListViewDouble.ListItems(101).SubItems(5) = (2 * Me.ListViewDouble.ListItems(101).SubItems(3)) - Me.ListViewDouble.ListItems(101).SubItems(4)
Me.ListViewDouble.ListItems(101).SubItems(6) = (0.8 / 0.2) * (Me.ListViewDouble.ListItems(101).SubItems(3) - Me.ListViewDouble.ListItems(101).SubItems(4))
Me.ListViewDouble.ListItems(101).SubItems(7) = 0
Me.ListViewDouble.ListItems(101).SubItems(8) = 0
Me.ListViewDouble.ListItems(101).SubItems(9) = 0
Me.ListViewDouble.ListItems(101).SubItems(10) = 0

For a = 102 To 111
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.8 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.2 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 102 To 111
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.8 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.2 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 102 To 111
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 102 To 111
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.8 / 0.2) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 102 To 111
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 102 To 111
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 102 To 111
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next

Me.ListViewHasilDouble.ListItems(8).SubItems(1) = hasil + Me.ListViewDouble.ListItems(101).SubItems(10)
Me.ListViewHasilDouble.ListItems(8).SubItems(2) = (Me.ListViewHasilDouble.ListItems(8).SubItems(1) / 11) * 100

Me.ListViewDouble.ListItems(100).SubItems(11) = Me.ListViewDouble.ListItems(111).SubItems(5) * 1 + ListViewDouble.ListItems(111).SubItems(6) * 1
For a = 101 To 111
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(111).SubItems(6) + Me.ListViewDouble.ListItems(111).SubItems(7) * (a - 99)
Next


'alfa 9
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Set data = Me.ListViewDouble.ListItems.Add(, , "")
Me.ListViewDouble.ListItems(113).SubItems(1) = "Alfa 0.9"

For a = 114 To 125
    Set data = Me.ListViewDouble.ListItems.Add(, , a - 113)
Next

For a = 1 To 12
Me.ListViewDouble.ListItems(a + 113).SubItems(2) = Me.ListViewCadangan.ListItems(a).Text
Next

Me.ListViewDouble.ListItems(114).SubItems(1) = "JANUARI"
Me.ListViewDouble.ListItems(115).SubItems(1) = "FEBRUARI"
Me.ListViewDouble.ListItems(116).SubItems(1) = "MARET"
Me.ListViewDouble.ListItems(117).SubItems(1) = "APRIL"
Me.ListViewDouble.ListItems(118).SubItems(1) = "MEI"
Me.ListViewDouble.ListItems(119).SubItems(1) = "JUNI"
Me.ListViewDouble.ListItems(120).SubItems(1) = "JULI"
Me.ListViewDouble.ListItems(121).SubItems(1) = "AGUSTUS"
Me.ListViewDouble.ListItems(122).SubItems(1) = "SEPTEMBER"
Me.ListViewDouble.ListItems(123).SubItems(1) = "OKTOBER"
Me.ListViewDouble.ListItems(124).SubItems(1) = "NOPEMBER"
Me.ListViewDouble.ListItems(125).SubItems(1) = "DESEMBER"

Me.ListViewDouble.ListItems(114).SubItems(3) = Me.ListViewDouble.ListItems(114).SubItems(2)
Me.ListViewDouble.ListItems(114).SubItems(4) = Me.ListViewDouble.ListItems(114).SubItems(2)
Me.ListViewDouble.ListItems(114).SubItems(5) = 0
Me.ListViewDouble.ListItems(114).SubItems(6) = 0
Me.ListViewDouble.ListItems(114).SubItems(7) = 0
Me.ListViewDouble.ListItems(114).SubItems(8) = 0
Me.ListViewDouble.ListItems(114).SubItems(9) = 0
Me.ListViewDouble.ListItems(114).SubItems(10) = 0

Me.ListViewDouble.ListItems(115).SubItems(3) = (0.9 * Me.ListViewDouble.ListItems(115).SubItems(2)) + (0.1 * Me.ListViewDouble.ListItems(114).SubItems(3))
Me.ListViewDouble.ListItems(115).SubItems(4) = (0.9 * Me.ListViewDouble.ListItems(115).SubItems(3)) + (0.1 * Me.ListViewDouble.ListItems(114).SubItems(4))
Me.ListViewDouble.ListItems(115).SubItems(5) = (2 * Me.ListViewDouble.ListItems(115).SubItems(3)) - Me.ListViewDouble.ListItems(115).SubItems(4)
Me.ListViewDouble.ListItems(115).SubItems(6) = (0.9 / 0.1) * (Me.ListViewDouble.ListItems(115).SubItems(3) - Me.ListViewDouble.ListItems(115).SubItems(4))
Me.ListViewDouble.ListItems(115).SubItems(7) = 0
Me.ListViewDouble.ListItems(115).SubItems(8) = 0
Me.ListViewDouble.ListItems(115).SubItems(9) = 0
Me.ListViewDouble.ListItems(115).SubItems(10) = 0

For a = 116 To 125
    Me.ListViewDouble.ListItems(a).SubItems(3) = (0.9 * Me.ListViewDouble.ListItems(a).SubItems(2)) + (0.1 * Me.ListViewDouble.ListItems(a - 1).SubItems(3))
Next

For a = 116 To 125
    Me.ListViewDouble.ListItems(a).SubItems(4) = (0.9 * Me.ListViewDouble.ListItems(a).SubItems(3)) + (0.1 * Me.ListViewDouble.ListItems(a - 1).SubItems(4))
Next

For a = 116 To 125
    Me.ListViewDouble.ListItems(a).SubItems(5) = (2 * Me.ListViewDouble.ListItems(a).SubItems(3)) - Me.ListViewDouble.ListItems(a).SubItems(4)
Next

For a = 116 To 125
    Me.ListViewDouble.ListItems(a).SubItems(6) = (0.9 / 0.1) * (Me.ListViewDouble.ListItems(a).SubItems(3) - Me.ListViewDouble.ListItems(a).SubItems(4))
Next

For a = 116 To 125
Me.ListViewDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a - 1).SubItems(5) * 1 + Me.ListViewDouble.ListItems(a - 1).SubItems(6) * 1
 '   Me.ListViewDouble.ListItems(a).SubItems(7) = Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(5), "")) + Val(Format(Me.ListViewDouble.ListItems(a - 1).SubItems(6)))
Next

For a = 116 To 125
    Me.ListViewDouble.ListItems(a).SubItems(8) = Val(Me.ListViewDouble.ListItems(a).SubItems(2)) - Val(Me.ListViewDouble.ListItems(a).SubItems(7))
    Me.ListViewDouble.ListItems(a).SubItems(9) = Abs(Me.ListViewDouble.ListItems(a).SubItems(8))
Next

hasil = 0
For a = 116 To 125
    Me.ListViewDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a).SubItems(9) / Me.ListViewDouble.ListItems(a).SubItems(2)
    hasil = hasil + Me.ListViewDouble.ListItems(a).SubItems(10)
Next

Me.ListViewDouble.ListItems(114).SubItems(11) = Me.ListViewDouble.ListItems(125).SubItems(5) * 1 + ListViewDouble.ListItems(125).SubItems(6) * 1
For a = 115 To 125
    Me.ListViewDouble.ListItems(a).SubItems(11) = Me.ListViewDouble.ListItems(125).SubItems(6) + Me.ListViewDouble.ListItems(125).SubItems(7) * (a - 113)
Next



Me.ListViewHasilDouble.ListItems(9).SubItems(1) = hasil + Me.ListViewDouble.ListItems(115).SubItems(10)
Me.ListViewHasilDouble.ListItems(9).SubItems(2) = (Me.ListViewHasilDouble.ListItems(9).SubItems(1) / 11) * 100


Me.ListViewDataDouble.ListItems.Clear
Me.ListViewDataDouble.ColumnHeaders.Clear

nama = Me.ListViewDataDouble.ColumnHeaders.Add(1, , "Bulan", 800, 0)
nama = Me.ListViewDataDouble.ColumnHeaders.Add(2, , "Penjualan", 1500, 0)
For a = 1 To 9
    nama = Me.ListViewDataDouble.ColumnHeaders.Add(a + 2, , "alfa" & a, 1500, 0)
Next
For a = 1 To 12
    Set data = Me.ListViewDataDouble.ListItems.Add(, , a)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(1) = Me.ListViewDouble.ListItems(a + 1).SubItems(2)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(2) = Me.ListViewDouble.ListItems(a + 1).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(3) = Me.ListViewDouble.ListItems(a + 15).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(4) = Me.ListViewDouble.ListItems(a + 29).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(5) = Me.ListViewDouble.ListItems(a + 43).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(6) = Me.ListViewDouble.ListItems(a + 57).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(7) = Me.ListViewDouble.ListItems(a + 71).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(8) = Me.ListViewDouble.ListItems(a + 85).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(9) = Me.ListViewDouble.ListItems(a + 99).SubItems(11)
Next

For a = 1 To 12
    Me.ListViewDataDouble.ListItems(a).SubItems(10) = Me.ListViewDouble.ListItems(a + 113).SubItems(11)
Next

'Me.ListViewDataDouble.ListItems(1).SubItems(2) = Me.ListViewDouble.ListItems(13).SubItems(5) * 1 + Me.ListViewDouble.ListItems(13).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(1).SubItems(3) = Me.ListViewDouble.ListItems(27).SubItems(6) * 1 + Me.ListViewDouble.ListItems(27).SubItems(7) * 2
'Me.ListViewDataDouble.ListItems(1).SubItems(4) = Me.ListViewDouble.ListItems(41).SubItems(6) * 1 + Me.ListViewDouble.ListItems(41).SubItems(7) * 3
'Me.ListViewDataDouble.ListItems(1).SubItems(5) = Me.ListViewDouble.ListItems(55).SubItems(6) * 1 + Me.ListViewDouble.ListItems(55).SubItems(7) * 4
'Me.ListViewDataDouble.ListItems(1).SubItems(6) = Me.ListViewDouble.ListItems(69).SubItems(6) * 1 + Me.ListViewDouble.ListItems(69).SubItems(7) * 5
'Me.ListViewDataDouble.ListItems(1).SubItems(7) = Me.ListViewDouble.ListItems(83).SubItems(6) * 1 + Me.ListViewDouble.ListItems(83).SubItems(7) * 6
'Me.ListViewDataDouble.ListItems(1).SubItems(8) = Me.ListViewDouble.ListItems(97).SubItems(6) * 1 + Me.ListViewDouble.ListItems(97).SubItems(7) * 7
'Me.ListViewDataDouble.ListItems(1).SubItems(9) = Me.ListViewDouble.ListItems(111).SubItems(6) * 1 + Me.ListViewDouble.ListItems(111).SubItems(7) * 8
'Me.ListViewDataDouble.ListItems(1).SubItems(10) = Me.ListViewDouble.ListItems(125).SubItems(6) * 1 + Me.ListViewDouble.ListItems(125).SubItems(7) * 9

'Me.ListViewDataDouble.ListItems(1).SubItems(2) = Me.ListViewDouble.ListItems(13).SubItems(5) * 1 + Me.ListViewDouble.ListItems(13).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(2).SubItems(2) = Me.ListViewDouble.ListItems(27).SubItems(5) * 1 + Me.ListViewDouble.ListItems(27).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(3).SubItems(2) = Me.ListViewDouble.ListItems(41).SubItems(5) * 1 + Me.ListViewDouble.ListItems(41).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(4).SubItems(2) = Me.ListViewDouble.ListItems(55).SubItems(5) * 1 + Me.ListViewDouble.ListItems(55).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(5).SubItems(2) = Me.ListViewDouble.ListItems(69).SubItems(5) * 1 + Me.ListViewDouble.ListItems(69).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(6).SubItems(2) = Me.ListViewDouble.ListItems(83).SubItems(5) * 1 + Me.ListViewDouble.ListItems(83).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(7).SubItems(2) = Me.ListViewDouble.ListItems(97).SubItems(5) * 1 + Me.ListViewDouble.ListItems(97).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(8).SubItems(2) = Me.ListViewDouble.ListItems(111).SubItems(5) * 1 + Me.ListViewDouble.ListItems(111).SubItems(6) * 1
'Me.ListViewDataDouble.ListItems(9).SubItems(2) = Me.ListViewDouble.ListItems(125).SubItems(5) * 1 + Me.ListViewDouble.ListItems(125).SubItems(6) * 1

'For a = 1 To 7 '1
'    Me.ListViewDataDouble.ListItems(2).SubItems(a + 1) = Me.ListViewDouble.ListItems((a - 1) + (a * 12) + (14 + a)).SubItems(5) * 1 + Me.ListViewDouble.ListItems((a - 1) + (a * 12) + (14 + a)).SubItems(5) * 2
'Next
'
'For a = 1 To 12 '2
'    Me.ListViewDataDouble.ListItems(a).SubItems(3) = Me.ListViewDouble.ListItems(a + 15).SubItems(3)
'Next
'
'For a = 1 To 12 '3
'    Me.ListViewDataDouble.ListItems(a).SubItems(4) = Me.ListViewDouble.ListItems(a + 29).SubItems(3)
'Next
'
'For a = 1 To 12 '4
'    Me.ListViewDataDouble.ListItems(a).SubItems(5) = Me.ListViewDouble.ListItems(a + 43).SubItems(3)
'Next
'
'For a = 1 To 12 '5
'    Me.ListViewDataSingle.ListItems(a).SubItems(6) = Me.ListViewSingle.ListItems(a + 57).SubItems(3)
'Next
'
'For a = 1 To 12 '6
'    Me.ListViewDataSingle.ListItems(a).SubItems(7) = Me.ListViewSingle.ListItems(a + 71).SubItems(3)
'Next
'
'For a = 1 To 12 '7
'    Me.ListViewDataSingle.ListItems(a).SubItems(8) = Me.ListViewSingle.ListItems(a + 85).SubItems(3)
'Next
'
'For a = 1 To 12 '8
'    Me.ListViewDataSingle.ListItems(a).SubItems(9) = Me.ListViewSingle.ListItems(a + 99).SubItems(3)
'Next
'
'For a = 1 To 12 '9
'    Me.ListViewDataSingle.ListItems(a).SubItems(10) = Me.ListViewSingle.ListItems(a + 113).SubItems(3)
'Next

alfaa = 1
nf = Me.ListViewHasilDouble.ListItems(1).SubItems(2)
For a = 2 To 8
If nf < Me.ListViewHasilDouble.ListItems(a).SubItems(2) Then
    alfaa = alfaa
    nf = nf
Else
    nf = Me.ListViewHasilDouble.ListItems(a).SubItems(2)
    alfaa = a
End If
Next
Me.Text1.Text = alfaa
End Sub

Private Sub LoadPenjualan()
On Error GoTo AdaError
Set rsPenjualan = koneksi.Execute("select * from penjualan where year(penjualan.bulan)='" & Me.txtPenjualan.Text & "'")

Me.ListViewCadangan.ListItems.Clear
Me.ListViewCadangan.ColumnHeaders.Clear
Me.ListViewTampil.ListItems.Clear
Me.ListViewTampil.ColumnHeaders.Clear

nama = Me.ListViewCadangan.ColumnHeaders.Add(1, , "Bulan", 1000, 0)
'nama = Me.ListViewCadangan.ColumnHeaders.Add(2, , "Penjualan", 1000, 0)


'For a = 1 To 12
    nama = Me.ListViewTampil.ColumnHeaders.Add(1, , "NO", 500, 0)
    nama = Me.ListViewTampil.ColumnHeaders.Add(2, , "BULAN", 1500, 0)
    nama = Me.ListViewTampil.ColumnHeaders.Add(3, , "PENJUALAN ", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(4, , "ALFA 0.1", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(5, , "E1", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(6, , "ALFA 0.2", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(7, , "E2", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(8, , "ALFA 0.3", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(9, , "E3", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(10, , "ALFA 0.4", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(11, , "E4", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(12, , "ALFA 0.5", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(13, , "E5", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(14, , "ALFA 0.6", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(15, , "E6", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(16, , "ALFA 0.7", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(17, , "E7", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(18, , "ALFA 0.8", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(19, , "E8", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(20, , "ALFA 0.9", 1500, 1)
'    nama = Me.ListViewTampil.ColumnHeaders.Add(21, , "E9", 1500, 1)
'Next

urut = 1
While Not rsPenjualan.EOF
    Set data = Me.ListViewCadangan.ListItems.Add(, , rsPenjualan.Fields(2))
    Set data = Me.ListViewTampil.ListItems.Add(, , urut)
    Me.ListViewTampil.ListItems(urut).SubItems(2) = rsPenjualan.Fields(2)
    rsPenjualan.MoveNext
    urut = urut + 1
Wend
Set data = Me.ListViewTampil.ListItems.Add(, , "")
data.SubItems(2) = 0
Me.ListViewTampil.ListItems(1).SubItems(1) = "JANUARI"
Me.ListViewTampil.ListItems(2).SubItems(1) = "FEBRUARI"
Me.ListViewTampil.ListItems(3).SubItems(1) = "MARET"
Me.ListViewTampil.ListItems(4).SubItems(1) = "APRIL"
Me.ListViewTampil.ListItems(5).SubItems(1) = "MEI"
Me.ListViewTampil.ListItems(6).SubItems(1) = "JUNI"
Me.ListViewTampil.ListItems(7).SubItems(1) = "JULI"
Me.ListViewTampil.ListItems(8).SubItems(1) = "AGUSTUS"
Me.ListViewTampil.ListItems(9).SubItems(1) = "SEPTEMBER"
Me.ListViewTampil.ListItems(10).SubItems(1) = "OKTOBER"
Me.ListViewTampil.ListItems(11).SubItems(1) = "NOPEMBER"
Me.ListViewTampil.ListItems(12).SubItems(1) = "DESEMBER"

'Me.ListViewTampil.ListItems(1).SubItems(3) = 0
'Me.ListViewTampil.ListItems(1).SubItems(4) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(5) = 0
'Me.ListViewTampil.ListItems(1).SubItems(6) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(7) = 0
'Me.ListViewTampil.ListItems(1).SubItems(8) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(9) = 0
'Me.ListViewTampil.ListItems(1).SubItems(10) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(11) = 0
'Me.ListViewTampil.ListItems(1).SubItems(12) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(13) = 0
'Me.ListViewTampil.ListItems(1).SubItems(14) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(15) = 0
'Me.ListViewTampil.ListItems(1).SubItems(16) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(17) = 0
'Me.ListViewTampil.ListItems(1).SubItems(18) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(1).SubItems(19) = 0
'Me.ListViewTampil.ListItems(1).SubItems(20) = Me.ListViewTampil.ListItems(1).SubItems(2)
''Me.ListViewTampil.ListItems(1).SubItems(21) = 0

'Me.ListViewTampil.ListItems(2).SubItems(3) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(4) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(5) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(6) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(7) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(8) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(9) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(10) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(11) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(12) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(13) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(14) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(15) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(16) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(17) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(18) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(19) = Me.ListViewTampil.ListItems(1).SubItems(2)
'Me.ListViewTampil.ListItems(2).SubItems(20) = Me.ListViewTampil.ListItems(2).SubItems(2) - Me.ListViewTampil.ListItems(1).SubItems(2)

'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(3) = (0.1 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.9 * Me.ListViewTampil.ListItems(a - 1).SubItems(3))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(4) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(3)
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(5) = (0.2 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.8 * Me.ListViewTampil.ListItems(a - 1).SubItems(5))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(6) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(5)
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(7) = (0.3 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.7 * Me.ListViewTampil.ListItems(a - 1).SubItems(7))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(8) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(7)
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(9) = (0.4 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.6 * Me.ListViewTampil.ListItems(a - 1).SubItems(9))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(10) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(9)
'Next

'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(11) = (0.5 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.5 * Me.ListViewTampil.ListItems(a - 1).SubItems(11))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(12) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(11)
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(13) = (0.6 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.4 * Me.ListViewTampil.ListItems(a - 1).SubItems(13))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(14) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(13)
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(15) = (0.7 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.4 * Me.ListViewTampil.ListItems(a - 1).SubItems(15))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(16) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(15)
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(17) = (0.8 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.2 * Me.ListViewTampil.ListItems(a - 1).SubItems(17))
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(18) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(17)
'Next
'
'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(19) = (0.9 * Me.ListViewTampil.ListItems(a - 1).SubItems(2)) + (0.1 * Me.ListViewTampil.ListItems(a - 1).SubItems(19))
'Next

'For a = 3 To 13
'    Me.ListViewTampil.ListItems(a).SubItems(20) = Me.ListViewTampil.ListItems(a).SubItems(2) - Me.ListViewTampil.ListItems(a).SubItems(19)
'Next





'For a = 1 To Me.ListViewCadangan.ListItems.Count
'    Me.ListViewTampil.ListItems(a).SubItems(1) = rsPenjualan.Fields(2)
'Next
'    Set rsBuyer = koneksi.Execute("select * from buyer where buyer_id ='" & Me.Text1.Text & "'")
'Set rsBaru = koneksi.Execute("select * from temp where kd_penyakit='" & Me.ListViewLangkah2.ListItems.Item(a) & "'")

'    sqlCari = ""
'    sqlCari = "select * from buyer where buyer." & Kriteria & " like '%" & Me.Text1.Text & "%'"
'
'    rsBuyer.Close
'    rsBuyer.Open sqlCari, koneksi, adOpenDynamic, adLockBatchOptimistic
'Call Buka_Database(14)
'Me.ListViewClient.ListItems.Clear
'Me.ListViewClient.ColumnHeaders.Clear
'
'nama = Me.ListViewClient.ColumnHeaders.Add(1, , "Client ID", 1500, 0)
'    nama = Me.ListViewClient.ColumnHeaders.Add(2, , "Client Finish", Val(Me.ListViewClient.Width) - 1500, 0)
'
'    While Not rsClientFinish.EOF
'        Set Data = ListViewClient.ListItems.Add(, , rsClientFinish.Fields(0))
'        Data.SubItems(1) = rsClientFinish.Fields(1)
'        rsClientFinish.MoveNext
'    Wend

Exit Sub
AdaError:
MsgBox "data kurang valid", vbInformation, "Pemberitahuan"
Me.txtPenjualan.SetFocus
End Sub

Private Sub cmdAmbil_Click()
Call LoadPenjualan
End Sub

Private Sub cmdDouble_Click()
Call MetodeDouble
Me.tabPenjualan.Tab = 2
Me.tabPenjualan.TabEnabled(0) = False
Me.tabPenjualan.TabEnabled(1) = False
Me.tabPenjualan.TabEnabled(3) = False
End Sub

Private Sub cmdSingle_Click()
Call metodeSingle
Me.tabPenjualan.Tab = 1
Me.tabPenjualan.TabEnabled(0) = False
Me.tabPenjualan.TabEnabled(2) = False
Me.tabPenjualan.TabEnabled(3) = False
End Sub

Private Sub Command1_Click()
Call TabMati
End Sub

Private Sub Command2_Click()
Call TabMati
End Sub

Private Sub Command3_Click()
Call TabMati
End Sub

Private Sub Command5_Click()
Me.ListViewGrafik.ListItems.Clear
Me.ListViewGrafik.ColumnHeaders.Clear

Set data = Me.ListViewGrafik.ColumnHeaders.Add(1, , "Hasil Ramalan", Me.ListViewGrafik.Width, 0)
Set data = Me.ListViewGrafik.ListItems.Add(1, , Me.ListViewHasilSingle.ListItems(Val(Me.txtalfa.Text)).SubItems(1))

Me.tabPenjualan.Tab = 3
Me.tabPenjualan.TabEnabled(0) = False
Me.tabPenjualan.TabEnabled(2) = False
Me.tabPenjualan.TabEnabled(1) = False
'Call CITY_CHART(mscSingle)

Dim X(1 To 13, 1 To 3) As Variant

'X(1, 1) = "Jagung"
X(1, 2) = "Penjualan"
X(1, 3) = "Peramalan"
'X(1, 4) = "kedelai"
'X(1, 5) = "Singkong"
'X(1, 6) = "Tebu"

X(2, 1) = "Januari"
X(2, 2) = Me.ListViewDataSingle.ListItems(1).SubItems(1)
X(2, 3) = Me.ListViewDataSingle.ListItems(1).SubItems(Val(Me.txtalfa.Text) + 1)
'X(2, 4) = 9
'X(2, 5) = 10
'X(2, 6) = 14

X(3, 1) = "Februari"
X(3, 2) = Me.ListViewDataSingle.ListItems(2).SubItems(1)
X(3, 3) = Me.ListViewDataSingle.ListItems(2).SubItems(Val(Me.txtalfa.Text) + 1)
'X(3, 4) = 10
'X(3, 5) = 8
'X(3, 6) = 19

X(4, 1) = "Maret"
X(4, 2) = Me.ListViewDataSingle.ListItems(3).SubItems(1)
X(4, 3) = Me.ListViewDataSingle.ListItems(3).SubItems(Val(Me.txtalfa.Text) + 1)

X(5, 1) = "April"
X(5, 2) = Me.ListViewDataSingle.ListItems(4).SubItems(1)
X(5, 3) = Me.ListViewDataSingle.ListItems(4).SubItems(Val(Me.txtalfa.Text) + 1)

X(6, 1) = "Mei"
X(6, 2) = Me.ListViewDataSingle.ListItems(5).SubItems(1)
X(6, 3) = Me.ListViewDataSingle.ListItems(5).SubItems(Val(Me.txtalfa.Text) + 1)

X(7, 1) = "Juni"
X(7, 2) = Me.ListViewDataSingle.ListItems(6).SubItems(1)
X(7, 3) = Me.ListViewDataSingle.ListItems(6).SubItems(Val(Me.txtalfa.Text) + 1)

X(8, 1) = "Juli"
X(8, 2) = Me.ListViewDataSingle.ListItems(7).SubItems(1)
X(8, 3) = Me.ListViewDataSingle.ListItems(7).SubItems(Val(Me.txtalfa.Text) + 1)

X(9, 1) = "Agustus"
X(9, 2) = Me.ListViewDataSingle.ListItems(8).SubItems(1)
X(9, 3) = Me.ListViewDataSingle.ListItems(8).SubItems(Val(Me.txtalfa.Text) + 1)

X(10, 1) = "September"
X(10, 2) = Me.ListViewDataSingle.ListItems(9).SubItems(1)
X(10, 3) = Me.ListViewDataSingle.ListItems(9).SubItems(Val(Me.txtalfa.Text) + 1)

X(11, 1) = "Oktober"
X(11, 2) = Me.ListViewDataSingle.ListItems(10).SubItems(1)
X(11, 3) = Me.ListViewDataSingle.ListItems(10).SubItems(Val(Me.txtalfa.Text) + 1)

X(12, 1) = "Nopember"
X(12, 2) = Me.ListViewDataSingle.ListItems(11).SubItems(1)
X(12, 3) = Me.ListViewDataSingle.ListItems(11).SubItems(Val(Me.txtalfa.Text) + 1)

X(13, 1) = "Desember"
X(13, 2) = Me.ListViewDataSingle.ListItems(12).SubItems(1)
X(13, 3) = Me.ListViewDataSingle.ListItems(12).SubItems(Val(Me.txtalfa.Text) + 1)

mscSingle.ChartData = X
mscSingle.ShowLegend = True
mscSingle.chartType = VtChChartType2dLine
mscSingle.Footnote = "Sumber dari data yang valid"
mscSingle.Title = "Metode Single"

With mscSingle.Plot.Axis(1, 1)
    .AxisTitle.Text = "Penjualan"
End With

With mscSingle.Plot.Axis(0, 1)
    .AxisTitle.Text = "Bulan"
End With
End Sub

Private Sub Command6_Click()
Me.ListViewGrafik.ListItems.Clear
Me.ListViewGrafik.ColumnHeaders.Clear

Set data = Me.ListViewGrafik.ColumnHeaders.Add(1, , "Bulan", 1000, 0)
Set data = Me.ListViewGrafik.ColumnHeaders.Add(2, , "Ramalan", 2000, 0)

For a = 1 To 12
    Set data = Me.ListViewGrafik.ListItems.Add(a, , a)
'    Me.ListViewGrafik.ListItems(a).Text = a
    Me.ListViewGrafik.ListItems(a).SubItems(1) = Me.ListViewDataDouble.ListItems(a).SubItems(Val(Me.Text1.Text) + 1)
Next


Me.tabPenjualan.Tab = 3
Me.tabPenjualan.TabEnabled(0) = False
Me.tabPenjualan.TabEnabled(2) = False
Me.tabPenjualan.TabEnabled(1) = False
'Call Grafik

Dim X(1 To 13, 1 To 3) As Variant

'X(1, 1) = "Jagung"
X(1, 2) = "Penjualan"
X(1, 3) = "Peramalan"
'X(1, 4) = "kedelai"
'X(1, 5) = "Singkong"
'X(1, 6) = "Tebu"

X(2, 1) = "Januari"
X(2, 2) = Me.ListViewDataDouble.ListItems(1).SubItems(1)
X(2, 3) = Me.ListViewDataDouble.ListItems(1).SubItems(Val(Me.Text1.Text) + 1)
'X(2, 4) = 9
'X(2, 5) = 10
'X(2, 6) = 14

X(3, 1) = "Februari"
X(3, 2) = Me.ListViewDataDouble.ListItems(2).SubItems(1)
X(3, 3) = Me.ListViewDataDouble.ListItems(2).SubItems(Val(Me.Text1.Text) + 1)
'X(3, 4) = 10
'X(3, 5) = 8
'X(3, 6) = 19

X(4, 1) = "Maret"
X(4, 2) = Me.ListViewDataDouble.ListItems(3).SubItems(1)
X(4, 3) = Me.ListViewDataDouble.ListItems(3).SubItems(Val(Me.Text1.Text) + 1)

X(5, 1) = "April"
X(5, 2) = Me.ListViewDataDouble.ListItems(4).SubItems(1)
X(5, 3) = Me.ListViewDataDouble.ListItems(4).SubItems(Val(Me.Text1.Text) + 1)

X(6, 1) = "Mei"
X(6, 2) = Me.ListViewDataDouble.ListItems(5).SubItems(1)
X(6, 3) = Me.ListViewDataDouble.ListItems(5).SubItems(Val(Me.Text1.Text) + 1)

X(7, 1) = "Juni"
X(7, 2) = Me.ListViewDataDouble.ListItems(6).SubItems(1)
X(7, 3) = Me.ListViewDataDouble.ListItems(6).SubItems(Val(Me.Text1.Text) + 1)

X(8, 1) = "Juli"
X(8, 2) = Me.ListViewDataDouble.ListItems(7).SubItems(1)
X(8, 3) = Me.ListViewDataDouble.ListItems(7).SubItems(Val(Me.Text1.Text) + 1)

X(9, 1) = "Agustus"
X(9, 2) = Me.ListViewDataDouble.ListItems(8).SubItems(1)
X(9, 3) = Me.ListViewDataDouble.ListItems(8).SubItems(Val(Me.Text1.Text) + 1)

X(10, 1) = "September"
X(10, 2) = Me.ListViewDataDouble.ListItems(9).SubItems(1)
X(10, 3) = Me.ListViewDataDouble.ListItems(9).SubItems(Val(Me.Text1.Text) + 1)

X(11, 1) = "Oktober"
X(11, 2) = Me.ListViewDataDouble.ListItems(10).SubItems(1)
X(11, 3) = Me.ListViewDataDouble.ListItems(10).SubItems(Val(Me.Text1.Text) + 1)

X(12, 1) = "Nopember"
X(12, 2) = Me.ListViewDataDouble.ListItems(11).SubItems(1)
X(12, 3) = Me.ListViewDataDouble.ListItems(11).SubItems(Val(Me.Text1.Text) + 1)

X(13, 1) = "Desember"
X(13, 2) = Me.ListViewDataDouble.ListItems(12).SubItems(1)
X(13, 3) = Me.ListViewDataDouble.ListItems(12).SubItems(Val(Me.Text1.Text) + 1)

mscSingle.ChartData = X
mscSingle.ShowLegend = True
mscSingle.chartType = VtChChartType2dLine
mscSingle.Footnote = "Sumber dari data yang valid"
mscSingle.Title = "Metode Double"

With mscSingle.Plot.Axis(1, 1)
    .AxisTitle.Text = "Penjualan"
End With

With mscSingle.Plot.Axis(0, 1)
    .AxisTitle.Text = "Bulan"
End With
End Sub

Private Sub Form_Load()
Call TabMati
Call Buka_Database(1)

End Sub

