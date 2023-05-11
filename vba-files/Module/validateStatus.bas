Attribute VB_Name = "validateStatus"

Public Sub status_id(ByVal count As LongPtr,ByVal pos As String)
  Dim valorMenor As LongPtr, valorMayor As LongPtr
  valorMenor = count * 20 / 100
  valorMayor = count * 80 / 100

  Range(pos).FormatConditions.Delete
  Range(pos).FormatConditions.AddIconSetCondition
  Range(pos).FormatConditions(Range(pos).FormatConditions.count).SetFirstPriority
  With Range(pos).FormatConditions(1)
    .ReverseOrder = True
    .ShowIconOnly = True
    .IconSet = ActiveWorkbook.IconSets(xl3Symbols2)
  End With
  With Range(pos).FormatConditions(1).IconCriteria(2)
    .Type = xlConditionValueNumber
    .value = valorMenor
    .Operator = 7
  End With
  With Range(pos).FormatConditions(1).IconCriteria(3)
    .Type = xlConditionValueNumber
    .value = valorMayor
    .Operator = 7
  End With

End Sub

Public Function validateKey(dict As Object, key As String, item As Variant) As Variant
  If Not IsNull(dict.Item(key)) And dict.Item(key) <> "" And dict.Exists(key) = True Then
    validateKey = Trim$(item.Offset(, dict.Item(key)).Value)
  Else
    nCount.sumError
    validateKey = Trim$(item.Offset(, dict.Item("NOMBRE CONTRATO")).Value)
  End If
End Function
