<div align="center">

## EditFlexGrid


</div>

### Description

This code allows users to edit in a MSFlexGrid
 
### More Info
 
Just copy and paste this code in the KeyPress event of your MSFlexGrid

No side effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jared Odulio](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jared-odulio.md)
**Level**          |Advanced
**User Rating**    |4.4 (80 globes from 18 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jared-odulio-editflexgrid__1-13339/archive/master.zip)





### Source Code

```
Private Sub MyFlexGrid_KeyPress(KeyAscii As Integer)
'Provides manual data entry capability to flexgrid
  With MyFlexGrid
    Select Case KeyAscii
      Case vbKeyReturn
        If .Col + 1 <= .Cols - 1 Then
          .Col = .Cols - 1
          ElseIf .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1
            .Col = 0
          Else
            .Row = 1
            .Col = 0
        End If
      Case vbKeyBack
        If Trim(.Text) <> "" Then
          .Text = Mid(.Text, 1, Len(.Text) - 1)
        End If
      Case Is < 32
      Case Else
        If .Col = 0 Or .Row = 0 Then
          Exit Sub
          Else
            .Text = .Text & Chr(KeyAscii)
        End If
    End Select
  End With
End Sub
```

