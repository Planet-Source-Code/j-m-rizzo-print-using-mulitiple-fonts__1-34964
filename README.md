<div align="center">

## Print using mulitiple fonts


</div>

### Description

print a single line using two fonts
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[J\. M\. Rizzo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/j-m-rizzo.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/j-m-rizzo-print-using-mulitiple-fonts__1-34964/archive/master.zip)





### Source Code

<p>'<br>
' Have you ever wanted an easy way to change fonts within a print line<br>
' when sending output to the printer? I reciently had a requirement to<br>
' print ascii data with embeded Barcode output on the same line.<br>
' Since the output was being generated from a database query, it made<br>
' sense to format the data using a proportional font (Courier New) then<br>
' change the font after printing the ascii data. Unfortunatly, the printer object<br>
' will only allow you to specify a font for the entire line being printed.<br>
'<br>
' Being an old fart programmer from way back, I remebered that Q-Basic had the<br>
' capability to stop the print head if there was a semicolon ";" after<br>
' the statCommand1_Click()<br>
'<br>
' Set up to display the Common ement. I could not find a reference to this anywhere in the VB<br>
' documentation, but tried it anyway. Wala, it works just fine.<br>
'<br>
' Create a form and place a command button and add the Common Dialog control to<br>
' the form. Then Cut and paste this code into the project.<br>
'<br>
' You probably will not have the Code 39 font but any other valid font name<br>
' will work fine.<br>
'<br>
Private Sub Command1_Click()<br>
'<br>
On Error Resume Next<br>
CommonDialog1.CancelError = True<br>
CommonDialog1.ShowPrinter<br>
'<br>
' check for errors or cancel selected<br>
'<br>
If Err &lt;> 0 Then<br>
  MsgBox Error(Err)<br>
  Exit Sub<br>
End If<br>
'<br>
' reset error checking<br>
'<br>
On Error GoTo HaveError<br>
'<br>
' Set the printer font to proportional font.<br>
'<br>
Printer.Font = "Courier New"<br>
Printer.FontSize = 10<br>
'<br>
' Print out the ASCII TEXT<br>
' NOTE THE SEMICOLON AT THE END!!!!<br>
' THIS TELLS THE PRINT METHOD NOT TO RETURN THE PRINTER<br>
' HEAD FOR THE NEXT LINE.<br>
'<br>
Printer.Print "This is the Ascii Text ";<br>
'<br>
' CHANGE THE PRINTER FONT FOR UPC 39 bar code font<br>
'<br>
Printer.Font = "Code 39"<br>
Printer.FontSize = 12<br>
'<br>
' print out the ascii text in barcode font<br>
' Notice the Leading and trailing "*" this<br>
' translates to the start/end barcode character<br>
' in the Code 39 font<br>
'<br>
Printer.Print "*This is Bar Code Text*"<br>
'<br>
' close out the printer object<br>
'<br>
Printer.EndDoc<br>
Exit Sub<br>
HaveError:<br>
MsgBox Error(Err)<br>
Resume Next<br>
<br>
End Sub<br>
<br>
</p>

