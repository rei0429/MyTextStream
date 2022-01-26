# MyTextStream
put UTF8-LF-without BOM text file via Offce VBA  
a simple replacement implementation of FileSystemObject.TextStream  
  
## Usage
import or paste MyTextStrema.cls  
see Sample.bas
```vba
Option Explicit

Public Sub Main()
    Dim ts As MyTextStream: Set ts = New MyTextStream
    ts.WriteLine ("Hello, World!")
    ts.WriteLine ("こんにちは世界")
    ts.SaveAs ("C:\test\utft8.txt")
    Set ts = Nothing
End Sub
```
