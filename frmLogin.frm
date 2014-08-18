VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdMasuk 
      Caption         =   "Masuk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMasuk_Click()
Call Buka_Database(2)
'rsUser.Find ("username='" & Me.txtUserName.Text & "'")
On Error Resume Next
With rsUser
        .MoveFirst
        .Find "username = '" & Me.txtUserName.Text & "'"
         If Not .EOF Then
         If Me.txtPassword.Text = rsUser!Password Then
'         frmUtama.LblNama.Caption = rsKaryawan!nama
'           frmSplash.Show
           frmUtama.Show
           Unload frmLogin
           
         Else
    MsgBox "password atau karyawan salah", vbOKOnly, "Peringatan"
    
    Me.txtPassword.Text = ""
    Me.txtUserName.Text = ""
    Me.txtUserName.SetFocus
    End If
    Else
    MsgBox "password atau karyawan salah", vbOKOnly, "Peringatan"
    
    Me.txtPassword.Text = ""
    Me.txtUserName.Text = ""
    Me.txtUserName.SetFocus
    End If
    End With
'If Me.txtPassword.Text = "admin" Then
'    If Me.txtUserName.Text = "admin" Then
'        frmUtama.Show vbModal
'        Unload Me
'    Else
'        MsgBox "Username yang Anda masukkan salah", vbInformation, "Pemberitahuan"
'    End If
'Else
'    MsgBox "Password yang Anda masukkan salah", vbInformation, "Pemberitahuan"
'End If
End Sub
