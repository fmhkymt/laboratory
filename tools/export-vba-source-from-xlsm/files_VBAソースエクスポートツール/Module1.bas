Attribute VB_Name = "Module1"
Option Explicit

Sub logClear()

    Sheet1.txtLog = ""
    
End Sub

Sub logBase(ByVal msg As String)
    Sheet1.txtLog = Sheet1.txtLog + msg + vbCrLf
End Sub


Sub logInfoBase(ByVal msg As String)
    logBase "[INFO] " + msg
End Sub

Sub logInfo(ByVal fileName As String, ByVal msg As String)
    logInfoBase fileName + " - " + msg
End Sub

Sub logErrorBase(ByVal msg As String)
    logBase "[ERROR] " + msg
End Sub

Sub logError(ByVal fileName As String, ByVal msg As String)
    logErrorBase fileName + " - " + msg
End Sub


