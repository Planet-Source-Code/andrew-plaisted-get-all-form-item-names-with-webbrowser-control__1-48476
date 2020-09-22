<div align="center">

## get all Form Item Names with webbrowser control


</div>

### Description

This will get all Form Item Names with webbrowser control.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew PLaisted](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-plaisted.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-plaisted-get-all-form-item-names-with-webbrowser-control__1-48476/archive/master.zip)





### Source Code

```
Public Function GetAllFormNames(doc As HTMLDocument, Form As Integer) As String
<p>
 Dim innames(20) As String
<p>
 Dim max As Integer
 <p>
 max = doc.Forms(Form).length
 <p>
 For i = 0 To max
<p>
If Not (doc.Forms(Form).Item(i) Is Nothing) Then
<p>
innames(i) = doc.Forms(Form).Item(i).name
 <p>
   Debug.Print innames(i)
<p>
  End If
<p>
 Next i
<p>
End Function
```

