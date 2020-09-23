<div align="center">

## primechecker


</div>

### Description

prime number || not? Interesting: The fastest possibility 2 find out is not the recursive one!
 
### More Info
 
a number

true or false


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Max Christian Pohle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/max-christian-pohle.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/max-christian-pohle-primechecker__1-60538/archive/master.zip)





### Source Code

```
'1 : as recursive function
Function IsPrime(Num As Long, Optional Start As Long = 3) As Boolean
 If Num Mod 2 <> 0 Then
 If Num Mod Start <> 0 Then
  If Start > Sqr(Num) Then _
  IsPrime = True Else _
  IsPrime = IsPrime(Num, Start + 2)
 End If
 End If
End Function
'2 : as standard-function
Function IsPrime(Num As Long) As Boolean
 Dim L As Long
 If Num Mod 2 <> 0 Then
 For L = 3 To Sqr(Num) Step 2
 If Num Mod L = 0 Then Exit For
 Next L
 If L > Sqr(Num) Then IsPrime = True
 End If
End Function
'example how2 call
Sub StartAndWrite()
 Do
 Me.Tag = Val(Me.Tag) + 1
 If IsPrime(Me.Tag) Then
 Open App.Path & "\primes.dat" For Append As #1
 Print #1, val(Me.Tag) & ",";
 Close #1
 DoEvents
 End If
 Loop
End Sub
```

