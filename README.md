<div align="center">

## brute force algorithm


</div>

### Description

it's supposed to return the next password in a sequence of passwords, not random passwords, or a dictionary attack. it's designed to return case-insensitive alphanumeric passwords, but that can easily be changed, i can't help you do that, it's up to you
 
### More Info
 
startPW - this is the pw you want the algorithm to start from

noChange - if this is True, it'll just poll for the next pw w/o changing it inside the function

startLength - ignored if startPW is used, must be > 1, it determines the length of the starting password.

the next base35 string (alphanumeric, lowercase)

none, as long as startLength is greater than 0


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Xeeky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/xeeky.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/xeeky-brute-force-algorithm__1-41252/archive/master.zip)





### Source Code

```
Public Function NextPW(Optional startPW As String, Optional noChange As Boolean, Optional startLength = 1) As String
  Dim pw() As String, i As Long, s As Long
  Static curPW As String
  If startPW <> "" And curPW = "" Then
    curPW = startPW
    NextPW = curPW
    Exit Function
  End If
  If curPW = "" Then
    curPW = String(startLength, 48)
    NextPW = curPW
    Exit Function
  End If
  If curPW = String(Len(curPW), 122) Then
    NextPW = String(Len(curPW) + 1, 48)
    If noChange = False Then curPW = NextPW
    Exit Function
  End If
  ReDim pw(Len(curPW))
  For i = 1 To Len(curPW)
    pw(i) = Mid(curPW, i, 1)
  Next
  i = UBound(pw)
donextchar:
  s = Asc(pw(i)) + 1
  Select Case s
    Case 58
      pw(i) = Chr(97)
    Case 123
      pw(i) = Chr(48)
      i = i - 1
      GoTo donextchar
    Case Else
      pw(i) = Chr(s)
  End Select
  For s = LBound(pw) To UBound(pw)
    NextPW = NextPW & pw(s)
  Next
  If noChange = False Then curPW = NextPW
End Function
```

