VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmProgress 
   Caption         =   "�i����"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "FrmProgress.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FrmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()
    
End Sub

'�i��0�ŕ\������
Public Sub initProgress(ByVal max As Integer)
    ProgressBar1.min = 0
    ProgressBar1.max = max
    ProgressBar1.Value = 0
    lblProgress.Caption = ProgressBar1.Value & "/" & ProgressBar1.max
    Show vbModeless
End Sub



'�i�����X�V����
Public Sub updateProgress(ByVal count As Integer)
    FrmProgress.ProgressBar1.Value = count
    lblProgress.Caption = count & "/" & ProgressBar1.max
    FrmProgress.Repaint

End Sub

