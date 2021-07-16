VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SrcExportTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public outDirSuffix As String
Public objExcel As Excel.Application

' Excelファイルを開きソースファイルをエクスポートする
' [引数]
' fileName Excelファイル名
Public Sub openAndExportSrc(ByVal fileName As String)

    Dim trgBook As Workbook
    Set trgBook = objExcel.Workbooks.Open(fileName, , True)
    
    logInfo fileName, "ファイルを開きました."
    If trgBook Is Nothing Then
        logError fileName, "Excelファイルのオープンに失敗しました。"
        Exit Sub
    End If
    logInfo fileName, "エクスポート開始"
    exportSrc trgBook
    logInfo fileName, "エクスポート終了"
    trgBook.Close False
    logInfo fileName, "ファイルを閉じました"
End Sub

'ソースファイルをエクスポートする
Private Sub exportSrc(ByRef trgBook As Workbook)
    Dim module As VBComponent
    Dim extension As String
    Dim outPath As String
    Dim outFileName As String
    Dim fso As New FileSystemObject
    
    '出力パス
    outPath = trgBook.Path + "\" + outDirSuffix + fso.GetBaseName(trgBook.Name)
    If Dir(outPath, vbDirectory) = "" Then
        MkDir outPath
    End If
    

    
    '全てのモジュール
    For Each module In trgBook.VBProject.VBComponents
    
        Select Case module.Type
        Case vbext_ct_ClassModule   'クラスモジュール
            extension = "cls"
        Case vbext_ct_MSForm        'フォーム
            extension = "frm"
        Case vbext_ct_StdModule     '標準モジュール
            extension = "bas"
        Case vbext_ct_Document      'Sheet中のソース
            extension = "cls"
        Case Else
            Err.Raise 99, "ExportSrc", "未知のモジュール"
        End Select
        
        outFileName = outPath & "\" & module.Name & "." & extension
        
        module.Export outFileName
        logInfo outFileName, "ソースをエクスポートしました。"
    Next
    
    
    
    
End Sub


