Attribute VB_Name = "Sample"
Option Explicit

Public Sub Main()
    Dim ts As MyTextStream: Set ts = New MyTextStream
    ts.WriteLine ("Hello, World!")
    ts.WriteLine ("����ɂ��͐��E")
    ts.SaveAs ("C:\test\utft8.txt")
    Set ts = Nothing
End Sub
