Attribute VB_Name = "modClipboard"
Option Compare Database
Option Explicit

' Module      : modClipboard

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function WinGetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrCopy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function WinSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipBoard Lib "user32" Alias "EmptyClipboard" () As Long
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function WinGetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrCopy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function WinSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function EmptyClipBoard Lib "user32" Alias "EmptyClipboard" () As Long
#End If

Private Const GHND As Long = &H42
Private Const CF_TEXT As Long = 1
Private Const mcintMaxSize As Integer = 4096

' Purpise: Clears the clipboard
Private Sub ClearClipboardData()
    On Error GoTo Err_Handler

    If OpenClipboard(0&) <> 0 Then
        ' Clear the Clipboard
        Call EmptyClipBoard
        Call CloseClipboard
    End If

Exit_Err_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.description & vbCrLf & "Procedure: ClearClipboardData", vbCritical + vbOKOnly, "ClearClipboardData - Error"
    Resume Exit_Err_Handler
End Sub

' Purpose: Get the text contents of the clipboard
' Returns : String
Public Function GetClipboardData() As String
    On Error GoTo Err_Handler

    #If VBA7 Then
        Dim lngClipMemory As LongPtr
    #Else
        Dim lngClipMemory As Long
    #End If
  
    Dim lngHandle As Long
    Dim strTmp As String
  
    If OpenClipboard(0&) <> 0 Then
    
        ' Get handle to global memory holding clipboard text
        lngHandle = WinGetClipboardData(CF_TEXT)
    
        ' Could we allocate the memory?
        If lngHandle <> 0 Then
      
            ' Lock memory so we can get the string
            lngClipMemory = GlobalLock(lngHandle)
      
            ' If we could lock it
            strTmp = Space$(mcintMaxSize)
            Call lstrCopy(strTmp, lngClipMemory)
            Call GlobalUnlock(lngHandle)
      
            ' Strip off any nulls and trim the result
            strTmp = Left$(strTmp, InStr(strTmp, Chr(0)) - 1)
      
        End If
    
        Call CloseClipboard
    End If
  
    GetClipboardData = strTmp

Exit_Err_Handler:
    Exit Function

Err_Handler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.description & vbCrLf & "Procedure: GetClipboardData", vbCritical + vbOKOnly, "GetClipboardData - Error"
    Resume Exit_Err_Handler
End Function

' Purpose: Writes the supplied string to the clipboard
' Params  : strText     Text to write
Public Sub SetClipboardData(ByVal strText As String)
    On Error GoTo Err_Handler

    #If VBA7 Then
        Dim lngHoldMem As LongPtr
        Dim lngGlobalMem As LongPtr
    #Else
        Dim lngHoldMem As Long
        Dim lngGlobalMem As Long
    #End If
  
    ' Allocate moveable global memory.
    lngHoldMem = GlobalAlloc(GHND, LenB(strText) + 1)
  
    ' Lock the block to get a far pointer to this memory.
    lngGlobalMem = GlobalLock(lngHoldMem)
  
    ' Copy the string to this global memory.
    lngGlobalMem = lstrCopy(lngGlobalMem, strText)
  
    ' Unlock the memory.
    If GlobalUnlock(lngHoldMem) = 0 Then
    
        ' Open the Clipboard to copy data to.
        If OpenClipboard(0&) <> 0 Then
      
            ' Clear the Clipboard.
            Call EmptyClipBoard
      
            ' Copy the data to the Clipboard.
            Call WinSetClipboardData(CF_TEXT, lngHoldMem)
      
            Call CloseClipboard
        End If
    End If

Exit_Err_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.description & vbCrLf & "Procedure: SetClipboardData", vbCritical + vbOKOnly, "SetClipboardData - Error"
    Resume Exit_Err_Handler
End Sub
