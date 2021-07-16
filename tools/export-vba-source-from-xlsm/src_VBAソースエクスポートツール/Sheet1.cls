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
    
        
    If chkExcelOption = False Then Exit Sub '設定不備なら即終了
        
    
    Dim i As Integer
    Dim fileName As String
    
    Dim tool As New SrcExportTool
    Set tool.objExcel = CreateObject("Excel.Application")
    tool.outDirSuffix = Range("OutDirSuffix")
    
    logClear
    
    
    Dim trgFiles() As String, trgFileCount As Integer
    trgFileCount = 0
    
    
    For i = Range("lblTargetFiles").Rows.Row To Range("lblLog").Rows.Row - 2
        fileName = Trim(Cells(i, 3).Text)
        If fileName <> "" Then
            ReDim Preserve trgFiles(trgFileCount)
            trgFiles(trgFileCount) = fileName
            trgFileCount = trgFileCount + 1
        End If
    Next i
    
    If trgFileCount = 0 Then
        MsgBox "先に処理対象ファイルを選んでください。", vbExclamation
        Exit Sub
    End If
    
    
    
    'プログレスバー初期化
    FrmProgress.initProgress trgFileCount
    
    '処理実施
    For i = 0 To trgFileCount - 1
        FrmProgress.updateProgress i
        tool.openAndExportSrc trgFiles(i)
    Next i
    
    '処理終了
    FrmProgress.updateProgress trgFileCount
    FrmProgress.Hide
 
    tool.objExcel.Quit
    Set tool.objExcel = Nothing
    
    
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
    If VarType(myFile) And vbArray Then
        For j = 1 To UBound(myFile)
            For i = Range("lblTargetFiles").Rows.Row To Range("lblLog").Rows.Row - 2
                If Trim(Cells(i, 3).Text) = "" Then
                    Cells(i, 3).Value = myFile(j)
                    Exit For
                End If
            Next i
            
        Next
    End If
End Sub

'オブジェクトモデルのアクセスを信頼する設定になっているかどうかチェック
Private Function chkExcelOption() As Boolean
    
    Dim d As Integer
    On Error GoTo ErrorChkExcelOption
    
    d = ThisWorkbook.VBProject.VBComponents.count
    chkExcelOption = True
    Exit Function
    
ErrorChkExcelOption:
    'エラーが出たなら設定不備
    MsgBox "オブジェクトモデルのアクセスを信頼する設定が出来ておりません。" + vbCrLf + "「事前設定方法」シートに従いExcelの設定を変更してください。", vbCritical, "Excel設定エラー"
    Sheets("事前設定方法").Select
        
    chkExcelOption = False
    
    

End Function
