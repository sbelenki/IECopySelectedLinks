VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Copier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Clipboard functions:
Private Declare Function OpenClipboard Lib "USER32" _
    (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "USER32" _
    () As Long
Private Declare Function GetClipboardData Lib "USER32" _
    (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "USER32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
    As Long
Private Declare Function GlobalUnlock Lib "kernel32" ( _
    ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function lstrCopy Lib "kernel32" Alias "lstrcpyA" ( _
    ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function EmptyClipBoard Lib "USER32" Alias "EmptyClipboard" () _
    As Long

Private Const GHND = &H42
Private Const CF_TEXT = 1

Public Function AddDocument(doc As Object)
    Dim pxDoc As HTMLDocument
    Dim pxWindow As HTMLWindow2
    Dim selRange As Object
    Dim selText As String
    Dim sToClipboard As String
    Dim xSubDoc As Object
    Dim bIfCtrlKey As Boolean ' True if CTRL key pressed

    Set pxDoc = doc
    If pxDoc Is Nothing Then
        Exit Function ' nothing to do
    End If

    ' <Check if CTRL key pressed>
    Set pxWindow = pxDoc.parentWindow
    If Not pxWindow Is Nothing Then
        If Not pxWindow.event Is Nothing Then
            bIfCtrlKey = pxWindow.event.ctrlKey
        End If
    End If
    Set pxWindow = Nothing
    ' </Check if CTRL key pressed>

    ' <Find selection in the document>
    Set selRange = pxDoc.selection.createRange()
    selText = selRange.htmlText
    Set selRange = Nothing
    ' </Find selection in the document>
    
    sToClipboard = ActiveAnchorToString(doc.activeElement, bIfCtrlKey)

    If "" = sToClipboard Then
        If "" = selText Then
            ' Nothing is selected and no link is active - get all links from the document.
            sToClipboard = LinksToString(pxDoc.links, bIfCtrlKey)
        Else
            ' Selection is not empty, create a sub document from
            ' the selected HTML
            Set xSubDoc = doc.createDocumentFragment()
            If xSubDoc Is Nothing Then
                MsgBox "Error creating selected document."
            Else
                xSubDoc.appendChild (xSubDoc.createElement("head"))
                xSubDoc.appendChild (xSubDoc.createElement("body"))
                xSubDoc.body.insertAdjacentHTML "afterBegin", selText
                xSubDoc.recalc True
                ' xSubDoc.url is about:local, but links are being translated
                ' with the proper parent's base
                ' MsgBox Str(xSubDoc.links.length)
                
                sToClipboard = LinksToString(xSubDoc.links, bIfCtrlKey)
            End If
            Set xSubDoc = Nothing
        End If
    End If

    AllocateClipboardData sToClipboard

End Function

Private Function IsActiveAnchorInSelection( _
    pxFullDoc As HTMLDocument, _
    pxSubDoc As HTMLDocument _
    ) As Boolean

    If Not pxFullDoc.activeElement Is Nothing And _
        Not pxSubDoc.activeElement Is Nothing And _
        pxFullDoc.activeElement Is pxSubDoc.activeElement Then
        IsActiveAnchorInSelection = True
    Else
        IsActiveAnchorInSelection = False
    End If
End Function

Private Function ActiveAnchorToString( _
    pxActiveElem As Object, _
    bCtrlKey As Boolean) As String

    If Not pxActiveElem Is Nothing Then
        If "a" = LCase(pxActiveElem.tagName) Then ' check if Anchor is active
            ' copy active link URL
            With pxActiveElem
                If bCtrlKey Then
                    ActiveAnchorToString = .innerText & vbTab & .href
                Else
                    ActiveAnchorToString = .href
                End If
            End With
        End If
    End If
End Function

Private Function LinksToString( _
    pxLinks As IHTMLElementCollection, _
    bCtrlKey As Boolean) As String

    Dim i As Integer
    For i = 0 To pxLinks.length - 1
        If bCtrlKey Then
            LinksToString = LinksToString + pxLinks(i).innerText + _
                vbTab + pxLinks(i).href + vbCrLf
        Else
            LinksToString = LinksToString + pxLinks(i).href + vbCrLf
        End If
    Next i

End Function

' Comments : Allocates a string in the clipboard
' Parameters: strText - text to allocate in the Clipboard
' Returns : Nothing
Private Sub AllocateClipboardData(strText As String)

    Dim lngHoldMem As Long
    Dim lngGlobalMem As Long
    Dim lngClipMem As Long
    Dim lngTmp As Long

    On Error GoTo PROC_ERR

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
            lngTmp = EmptyClipBoard()
            ' Copy the data to the Clipboard.
            lngClipMem = SetClipboardData(CF_TEXT, lngHoldMem)
            lngTmp = CloseClipboard()
        End If
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "AllocateClipboardData"
    Resume PROC_EXIT

End Sub

Private Function SubstituteXmlMarkups(ByRef sHtmlText As String)
    sHtmlText = Replace(sHtmlText, "&amp;", "&")
End Function

