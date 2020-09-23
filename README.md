<div align="center">

## Copy to Clipboard


</div>

### Description

how to copy from textbox to clipboard?
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[rizal](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rizal.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rizal-copy-to-clipboard__1-72085/archive/master.zip)





### Source Code

```
'harap faham2 sendiri la yek
Private Sub Command1_Click()
  Clipboard.Clear
  Clipboard.SetText Text1.Text
End Sub
Private Sub Command2_Click()
  Text1.Text = Clipboard.GetText
End Sub
```

