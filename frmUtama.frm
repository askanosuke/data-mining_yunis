VERSION 5.00
Begin VB.Form frmUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FORM UTAMA"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   12750
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu file 
      Caption         =   "FILE"
      Begin VB.Menu fInputData 
         Caption         =   "INPUT DATA"
      End
   End
   Begin VB.Menu ramal 
      Caption         =   "RAMAL"
   End
   Begin VB.Menu exit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub inputdata_Click()

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub fInputData_Click()
frmInputData.Show vbModal
End Sub

Private Sub Form_Load()
Call Buka_Database(1)
End Sub

Private Sub ramal_Click()
frmRamal.Show vbModal
End Sub
