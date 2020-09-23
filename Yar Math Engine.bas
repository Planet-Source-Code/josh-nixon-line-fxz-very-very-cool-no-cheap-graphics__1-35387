Attribute VB_Name = "Module1"
        '%****************************************************%
        '%             Yar Interactive Math Engine            %
        '%       By Joshua  Nixon                             %
        '%       Email: JNixon21@excite.com                   %
        '%       Aim Sn\Yahoo: Nit3shift                      %
        '%       Date: 2/17/02                                %
        '%    Purpose: The porpose of this module is to solve %
        '% 60 differant math functions that we dont like      %
        '% typing over and over.                              %
        '%****************************************************%
Public Sub CalcSin(Value As Double)
   ' CalcSin = Sin(Value * 0.0174533)
End Sub
Public Function InverseSin(Value As Double) As Double
    'InvSin =
End Function
Public Function InverseCos(Value As Double) As Double
    InvCos = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
End Function
Public Function InverseSec(Value As Double) As Double
    InvSec = Atn(Value / Sqr(Value * Value - 1)) + Sgn((Value) - 1) * (2 * Atn(1))
End Function
Public Function InverseCsc(Value As Double) As Double
    InvCsc = Atn(Value / Sqr(Value * Value - 1)) + (Sgn(Value) - 1) * (2 * Atn(1))
End Function
Public Function InverseCot(Value As Double) As Double
    InvCot = Atn(Value) + 2 * Atn(1)
End Function
Public Function Secant(Value As Double) As Double
    Sec = 1 / Cos(Value * pi / 180)
End Function
Public Function Csc(Value As Double) As Double
    Csc = 1 / Sin(Value * pi / 180)
End Function
Public Function Cot(Value As Double) As Double
    Cot = 1 / Tan(Value * pi / 180)
End Function
Public Function HSin(Value As Double) As Double
    HSin = (Exp(Value) - Exp(-Value)) / 2
End Function
Public Function HCos(Value As Long) As Double
    HCos = (Exp(Value) + Exp(-Value)) / 2
End Function
Public Function HTan(Value As Long) As Double
    HTan = (Exp(Value) - Exp(-Value)) / (Exp(Value) + Exp(-Value))
End Function
Public Function HSec(Value As Double) As Double
    HSec = 2 / (Exp(Value) + Exp(-Value))
End Function
Public Function HCsc(Value As Double) As Double
    HCsc = 2 / (Exp(Value) + Exp(-Value))
End Function
Public Function HCot(Value As Double) As Double
    HCot = (Exp(Value) + Exp(-Value)) / (Exp(Value) - Exp(-Value))
End Function
Public Function InverseHSin()
    InvHSin = Log(Value + Sqr(Value * Value + 1))
End Function
Public Function InverseHCos(Value As Double) As Double
On Error Resume Next
    InvHCos = Log(Value + Sqr(Value * Value - 1))
End Function
Public Function InverseHTan(Value As Double) As Double
    InvHTan = Log((1 + Value) / (1 - Value)) / 2
End Function
Public Function InverseHSec(Value As Double) As Double
    InvHSec = Log((Sqr(-Value * Value + 1) + 1) / Value)
End Function
Public Function InverseHCsc(Value As Double) As Double
    InvHCsc = Log((Sgn(Value) * Sqr(Value * Value + 1) + 1) / Value)
End Function
Public Function InverseHCot(Value As Double) As Double
    InvHCot = Log((Value + 1) / (Value - 1)) / 2
End Function
Public Function Logs(Value As Double) As Double
    Logo = Log(Value)
End Function
Public Function Exps(Value As Double) As Double
    Expo = Exp(Value)
End Function
Public Function Distance(Num1 As Double, Num2 As Double, Num3 As Double, Num4 As Double) As Double
    Distance = Sqr((Num2 - Num1) ^ 2 + (Num4 - Num3) ^ 2)
End Function

