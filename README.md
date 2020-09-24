<div align="center">

## Zip and Unzip \- in VB6


</div>

### Description

we want to call "WinRar.exe" from our VB project .
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Zaid Markabi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/zaid-markabi.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/zaid-markabi-zip-and-unzip-in-vb6__1-72150/archive/master.zip)





### Source Code

```
' we need ZipCmd and UnZipCmd ( commands button ) .
Private Sub ZipCmd_Click()
' we like to zip [ C:\DataFiles ] folder to [ C:\Compressor.rar ] file.
Shell App.Path + "\WinRar.exe M C:\Compressor.rar C:\DataFiles"
MsgBox "Zipped"
End Sub
Private Sub UnZipCmd_Click()
Shell App.Path + "\WinRar.exe X C:\Compressor.rar C:\"
MsgBox "UnZipped"
End Sub
```

