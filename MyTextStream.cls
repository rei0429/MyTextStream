VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyTextStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_buf As String

Public Sub WriteLine(ByVal line As String)
    'line delimited by LF
    m_buf = m_buf & line & vbLf
End Sub

Public Sub SaveAs(ByVal path As String)
    Dim bytes() As Byte
    
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2   'adTypeText
        .Charset = "UTF-8"
        .WriteText m_buf
        
        .Position = 0
        .Type = 1   'adTypeBynary
        .Position = 3 'remove(skip) BOM
        bytes = .Read
        
        .Close
    End With
    'write text file
    With CreateObject("ADODB.Stream")
        .Type = 1   'adTypeBynary
        .Open
        .Write bytes
        .SetEOS
        .SaveToFile path, 2 'OverWrite
        .Close
    End With
    
End Sub

