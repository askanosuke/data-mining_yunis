Attribute VB_Name = "Module1"
Public koneksi As New ADODB.Connection
Public rsPenjualan As New ADODB.Recordset
Public rsUser As New ADODB.Recordset

Public sqlSimpan As String
Public sqlUpdate As String
Public sqlHapus As String
Public sqlCari As String
Public Kriteria As String


Public Sub Buka_Database(data As Integer)

Call Konek
    Select Case data

        Case 1
            If rsPenjualan.State = adStateOpen Then
                rsPenjualan.Close
            ElseIf rsPenjualan.State = adStateClosed Then

            End If
            rsPenjualan.Open "select * from penjualan", koneksi, adOpenDynamic, adLockOptimistic
            
        Case 2
            If rsUser.State = adStateOpen Then
                rsUser.Close
            ElseIf rsUser.State = adStateClosed Then

            End If
            rsUser.Open "select * from data_user", koneksi, adOpenDynamic, adLockOptimistic
            
        
    End Select
End Sub


Public Sub Konek()
On Error GoTo AdaError
    If koneksi.State = adStateClosed Then
        koneksi.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\mining.mdb"
    End If
    Exit Sub
AdaError:
    MsgBox "Tidak Konek", vbOKOnly + vbCritical, "Koneksi Gagal"
    End
End Sub

'mengambil data grafik
