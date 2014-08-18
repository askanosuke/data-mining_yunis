VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM INPUT DATA"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBatal 
      Appearance      =   0  'Flat
      Caption         =   "Batal"
      Height          =   375
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdHapus 
      Appearance      =   0  'Flat
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdTambah 
      Appearance      =   0  'Flat
      Caption         =   "Tambah"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpWaktu 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM yyyy"
      Format          =   184287235
      CurrentDate     =   41833
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
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      Caption         =   "Edit"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListViewTampil 
      Height          =   3855
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6800
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
   Begin VB.TextBox txtID 
      Height          =   360
      Left            =   3600
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Jml. Penjualan"
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
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu"
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
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmInputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Tampil()
Me.ListViewTampil.ListItems.Clear
Me.ListViewTampil.ColumnHeaders.Clear

nama = Me.ListViewTampil.ColumnHeaders.Add(1, , "ID", 600, 0)
nama = Me.ListViewTampil.ColumnHeaders.Add(2, , "Jml Penjualan", 2000, 0)
nama = Me.ListViewTampil.ColumnHeaders.Add(3, , "Periode", 2000, 1)

If Not rsPenjualan.BOF Then
    rsPenjualan.MoveFirst
    While Not rsPenjualan.EOF
        Set data = Me.ListViewTampil.ListItems.Add(, , rsPenjualan.Fields(0))
        data.SubItems(1) = rsPenjualan.Fields(2)
        data.SubItems(2) = Format(rsPenjualan.Fields(1), "MMMM - yyyy")
        rsPenjualan.MoveNext
    Wend
End If
End Sub

Private Sub TampilText()
If Not rsPenjualan.BOF Then
    Me.txtID.Text = rsPenjualan.Fields(0)
    Me.txtPenjualan.Text = rsPenjualan.Fields(2)
    Me.dtpWaktu.Value = Format(rsPenjualan.Fields(1), "MMMM yyyy")
End If
End Sub
Private Sub cmdEdit_Click()
sqlUpdate = ""
sqlUpdate = "update penjualan set " _
            & "bulan='" & Format(Me.dtpWaktu.Value, "MM yyyy") & "'" _
            & "where jml_penjualan='" & Me.txtPenjualan.Text & "'"
koneksi.Execute sqlUpdate, , adCmdText
Call Form_Load
End Sub

Private Sub cmdHapus_Click()
If MsgBox("Anda Yakin Akan Menghapus Data Dengan ID = " & Me.txtID.Text & "?", vbInformation + vbYesNo, "Hapus data") = vbYes Then
    If Not rsPenjualan.BOF Or Not rsPenjualan.EOF Then
        sqlHapus = ""
        sqlHapus = "delete from penjualan where penjualan.jml_penjualan='" & Me.txtPenjualan.Text & "'"
        koneksi.Execute sqlHapus, , adCmdText
        Call Form_Load
    End If
End If
End Sub

Private Sub cmdTambah_Click()
sqlSimpan = ""
sqlSimpan = "insert into penjualan(jml_penjualan,bulan) values ('" & Me.txtPenjualan.Text & "','" & Me.dtpWaktu.Value & "')"
koneksi.Execute sqlSimpan, adCmdText
Call Form_Load
End Sub

Private Sub Form_Load()
Call Buka_Database(1)
Call Tampil
Me.txtPenjualan.Text = ""
Me.txtID.Text = ""
End Sub

Private Sub ListViewTampil_Click()
rsPenjualan.MoveFirst
rsPenjualan.Find ("id_penjualan='" & Me.ListViewTampil.SelectedItem.Text & "'")
If Not rsPenjualan.EOF Then
    Call TampilText
End If
End Sub
