<div align="center">

## Metallic Forms \*updated\*


</div>

### Description

Cut and Paste code... Turns your flat form into a metalic cylinder. Setting the AutoRedraw property is now done for you. Also thanks to Sebastjan for his tip about setting the ScaleMode to pixels.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[John Edward Colman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-edward-colman.md)
**Level**          |Beginner
**User Rating**    |4.3 (43 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/john-edward-colman-metallic-forms-updated__1-13788/archive/master.zip)





### Source Code

```
Private Sub GradientFill()
 Dim i As Long
 Dim c As Integer
 Dim r As Double
 r = ScaleHeight / 3.142
 'Hint: Multiplying r by differnt values give different effects (try 2.3)
 For i = 0 To ScaleHeight
  c = Abs(220 * Sin(i / r))
  'Hint: Changing sin to cos reverses range
  Me.Line (0, i)-(ScaleWidth, i), RGB(c, c, c + 30)
  'Hint: Notice the bias to blue. You can be more subtle by reducing this number (try 10). Try other colours too.
 Next
End Sub
Private Sub Form_Load()
 Me.ScaleMode = 3
 Me.AutoRedraw = True
End Sub
Private Sub Form_Resize()
 GradientFill
End Sub
```

