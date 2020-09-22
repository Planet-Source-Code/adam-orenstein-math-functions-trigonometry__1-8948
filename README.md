<div align="center">

## Math Functions \(trigonometry\)


</div>

### Description

this here is a whole bunch of math functions! Too many for me to list right here so come in and check em' all out
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Orenstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-orenstein.md)
**Level**          |Advanced
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-orenstein-math-functions-trigonometry__1-8948/archive/master.zip)





### Source Code

```
Public Function InvSin(Number As Double) As Double
 InvSin = CutDecimal(Atn(Number / Sqr(-Number * Number + 1)), 87)
End Function
Public Function InvCos(Number As Double) As Double
 InvCos = Atn(-Number / Sqr(-Number * Number + 1)) + 2 * Atn(1)
End Function
Public Function InvSec(Number As Double) As Double
 InvSec = Atn(Number / Sqr(Number * Number - 1)) + Sgn((Number) - 1) * (2 * Atn(1))
End Function
Public Function InvCsc(Number As Double) As Double
 InvCsc = Atn(Number / Sqr(Number * Number - 1)) + (Sgn(Number) - 1) * (2 * Atn(1))
End Function
Public Function InvCot(Number As Double) As Double
 InvCot = Atn(Number) + 2 * Atn(1)
End Function
Public Function Sec(Number As Double) As Double
 Sec = 1 / Cos(Number * PI / 180)
End Function
Public Function Csc(Number As Double) As Double
 Csc = 1 / Sin(Number * PI / 180)
End Function
Public Function Cot(Number As Double) As Double
 Cot = 1 / Tan(Number * PI / 180)
End Function
Public Function HSin(Number As Double) As Double
 HSin = (Exp(Number) - Exp(-Number)) / 2
End Function
Public Function HCos(Number As Double) As Double
 HCos = (Exp(Number) + Exp(-Number)) / 2
End Function
Public Function HTan(Number As Double) As Double
 HTan = (Exp(Number) - Exp(-Number)) / (Exp(Number) + Exp(-Number))
End Function
Public Function HSec(Number As Double) As Double
 HSec = 2 / (Exp(Number) + Exp(-Number))
End Function
Public Function HCsc(Number As Double) As Double
 HCsc = 2 / (Exp(Number) + Exp(-Number))
End Function
Public Function HCot(Number As Double) As Double
 HCot = (Exp(Number) + Exp(-Number)) / (Exp(Number) - Exp(-Number))
End Function
Public Function InvHSin()
 InvHSin = Log(Number + Sqr(Number * Number + 1))
End Function
Public Function InvHCos(Number As Double) As Double
 InvHCos = Log(Number + Sqr(Number * Number - 1))
End Function
Public Function InvHTan(Number As Double) As Double
 InvHTan = Log((1 + Number) / (1 - Number)) / 2
End Function
Public Function InvHSec(Number As Double) As Double
 InvHSec = Log((Sqr(-Number * Number + 1) + 1) / Number)
End Function
Public Function InvHCsc(Number As Double) As Double
 InvHCsc = Log((Sgn(Number) * Sqr(Number * Number + 1) + 1) / Number)
End Function
Public Function InvHCot(Number As Double) As Double
 InvHCot = Log((Number + 1) / (Number - 1)) / 2
End Function
Public Function Percent(is_ As Double, of As Double) As Double
 Percent = is_ / of * 100
End Function
```

