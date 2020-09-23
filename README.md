<div align="center">

## GetTok \- How to tokenize a string ? \(Like $gettok\(\) in mIRC Scripting\)


</div>

### Description

It simply tokenizes the string by a specified separator.
 
### More Info
 
Like $gettok identifier in mIRC. Easy to understand and use. Please vote for my code because this is my first code on planet-soruce-code.com :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Roni Tovi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roni-tovi.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/roni-tovi-gettok-how-to-tokenize-a-string-like-gettok-in-mirc-scripting__1-48508/archive/master.zip)





### Source Code

```
'Paste it into a module and call from anywhere!
Public Function GetTok(strString As String, N As Integer, strSep As String)
On Error Resume Next
Dim GArray
GArray = Split(strString, strSep)
If N = 0 Then
'if you specify 0 as N, then the function returns how much tokens exists in your string
GetTok = UBound(GArray) + 1
Exit Function
End If
GetTok = GArray(n - 1)
End Function
```

