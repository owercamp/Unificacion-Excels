VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "Counters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private accumulator As LongPtr, registers As LongPtr, data As LongPtr

Public Property Get counter() As Variant

counter = accumulator

End Property

Public Property Let counter(ByVal vNewValue As Variant)

accumulator = vNewValue

End Property

Public Property Get register() As Variant

register = registers

End Property

Public Property Let register(ByVal vNewValue As Variant)

registers = vNewValue

End Property

Public Property Get datainformation() As Variant

  datainformation = data

End Property

Public Property Let datainformation(ByVal vNewValue As Variant)

  data = vNewValue

End Property

Public Function sumError() As LongPtr

  Dim point As LongPtr
  point = counter + 1
  counter = point

End Function

Public Function sumRegister(ByVal totalRegister As Variant) As LongPtr

  Dim plus1 As LongPtr
  plus1 = register + totalRegister
  register = plus1

End Function

Public Function sumDataInformation(ByVal totalData As Variant) As LongPtr
  
  dim can As LongPtr
  can = data + totalData
  datainformation = can

End Function
