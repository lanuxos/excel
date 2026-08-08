# excel notes by la

## count by background color
- Alt + F11
- Insert > Module
```vba
Function CountColor(CountRange As Range, ColorCell As Range) As Long
    Dim cell As Range
    Dim ColorCode As Long
    Dim Count As Long
    
    ColorCode = ColorCell.Interior.Color
    Count = 0
    
    For Each cell In CountRange
        If cell.Interior.Color = ColorCode Then
            Count = Count + 1
        End If
    Next cell
    
    CountColor = Count
End Function
```

## 
