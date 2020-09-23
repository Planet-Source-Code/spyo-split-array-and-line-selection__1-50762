<div align="center">

## Split Array and Line Selection


</div>

### Description

The Split() Return a Array has a Few Hidding secrets, For VB6 only, Here is a quick simple code to explain and also the famous (what is the max array solution)... It Convert a TextBox into a Line by Line Array the match the TextBox Line Number
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Spyo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/spyo.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0, VB Script
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/spyo-split-array-and-line-selection__1-50762/archive/master.zip)





### Source Code

<Br>Option Explicit
<Br>' put this code in a form, add a <Br>RichTextBox1, Text1, and a Command1 Button for vb6 only
<Br>Private Sub Command1_Click()
<Br>Dim x As Integer, i As Integer
<Br>Dim Ray() As String
<Br>x = 0
<Br>RichTextBox1.Text = "Spyo Was Here, and got" & vbCrLf & "1: One Choice" & vbCrLf & "2: No Choice" & vbCrLf & "3: None of the Above" & vbCrLf & "4: All of the Above"
<Br>Ray() = Split(RichTextBox1.Text, "" & vbCrLf & "")
<Br>For i = 0 To UBound(Ray)
<Br>x = x + 1
<Br>Next i
<Br>RichTextBox1.Text = ""
<Br>x = x - 1
<Br>For i = 0 To x
<Br>RichTextBox1.Text = RichTextBox1.Text & Ray(i) & vbCrLf
<Br>Next i
<Br>Text1.Text = x & " Arrays As In Ray(0),Ray(1),Ray(2),Ray(3),Ray(4), but not Ray(5) or Above"
<Br>End Sub
<Br>'This is something very usefull, please improve it and share,,, vote ?

