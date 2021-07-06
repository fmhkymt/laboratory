VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub cmdExec_Click()
    
    Dim i As Integer
    Dim fileName As String
    
    Dim tool As New SrcExportTool
    tool.outDirSuffix = Range("OutDirSuffix")
    
    logClear
    
    For i = Range("lblTargetFiles").Rows.Row To Range("lblLog").Rows.Row - 2
        fileName = Trim(Cells(i, 3).Text)
        If fileName <> "" Then
            tool.openAndExportSrc fileName
        End If
    Next i
 
    
    
    
End Sub


Private Sub cmdClear_Click()
    Dim i As Integer
    logClear
    For i = Range("lblTargetFiles").Rows.Row To Range("lblLog").Rows.Row - 2
        Cells(i, 3).Value = ""
    Next i

End Sub

Private Sub cmdFileSelect_Click()
    Dim myFile As Variant
    Dim i As Integer, j As Integer
    Dim fso As New FileSystemObject
    ChDrive fso.GetDriveName(ThisWorkbook.Path)
    ChDir ThisWorkbook.Path
    myFile = Application.GetOpenFilename("Excelマクロ有効ブック(*.xlsm),*.xlsm", , "処理対象ファイルの選択(複数選択可)", , True)
    For j = 1 To UBound(myFile)
        For i = Range("lblTargetFiles").Rows.Row To Range("lblLog").Rows.Row - 2
            If Trim(Cells(i, 3).Text) = "" Then
                Cells(i, 3).Value = myFile(j)
                Exit For
            End If
        Next i
        
    Next
    
        
End Sub
