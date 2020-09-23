<div align="center">

## Kewl 3D Star Field\! User can change camera angle\!


</div>

### Description

This code generates a moving 3D star field, almost identical to the windows 'flying through space' screen saver, except this runs in a window of any size, and when the user moves the mouse over it, they can change the camera angle, which I suppose could make a neet game back ground for an outter space filght sim.
 
### More Info
 
None!

If you know what the circle function does it helps I guess, all you need to know is that when you type

circle (100,120),30

it makes a circle at (100,120) on the form, 30 units (twips?) wide.

Nothin!

None, as far as I know.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Eagar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-eagar.md)
**Level**          |Unknown
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Games](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/games__1-38.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-eagar-kewl-3d-star-field-user-can-change-camera-angle__1-1958/archive/master.zip)

### API Declarations

Nope!


### Source Code

```
'first just start a new program, and insert a timer named timer1!
'Then set it's interval to 1! That's it!
Dim starX(0 To 100) As Double  'holds the X coords for the stars
Dim starY(0 To 100) As Double  'holds the Y coords for the stars
Dim starDist(0 To 100) As Double 'holds the size the stars should be
Dim starSpeed As Double   'holds the speed of the star field
Dim formMidX As Double 'holds the center X coord for the form
Dim formMidY As Double 'holds the center Y coord for the form
Private Sub Form_KeyPress(KeyAscii As Integer)
'end when the user presses a key
End
End Sub
Private Sub Form_Load()
'initialize the random number generator
Randomize
Form1.BackColor = &H0&
Form1.ForeColor = &HFFFFFF
Form1.FillColor = &HFFFFFF
Form1.FillStyle = 0
Form1.DrawWidth = 2
'the middle x and y coords of the form
formMidX = (Form1.Width / 2) 'set the center x axis of the form
formMidY = (Form1.Height / 2) 'set the center y axis of the form
'initialize the arrays
For X = 0 To 100
  'loops to check that the star is not in the exact center of the screen
  Do
    'set the stars (x,y) coords to random places
    starX(X) = Int(Rnd * Form1.Width)
    starY(X) = Int(Rnd * Form1.Height)
    starDist(X) = Int(Rnd * 5)
  Loop While (starX(X) = formMidY And starY(Y) = formMidY)
  'the size of each star
  starDist(X) = 0
Next X
'set the speed at which the stars are moving
starSpeed = 0.025
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'set the 0,0 lines for the x&y axis at the mouse co-ords.
formMidX = X
formMidY = Y
End Sub
Private Sub Timer1_Timer()
'loop for each star
For X = 0 To 100
  'set the fill color to black
  Form1.FillColor = Form1.BackColor
  'this circle draws a black star over the star's last location
  Circle (starX(X), starY(X)), starDist(X), BackColor
  'add 1 to the star distance (size of the star)
  starDist(X) = starDist(X) + 0.1
  'determine in which direction the star should be moving on the x axis
  If starX(X) > (formMidX) Then
    starX(X) = starX(X) + Int(Abs(formMidX - starX(X)) * starSpeed) * (starDist(X) * 0.2)
  Else
    starX(X) = starX(X) - Int(Abs(formMidX - starX(X)) * starSpeed) * (starDist(X) * 0.2)
  End If
  'determine in which direction the star should be moving on the y axis
  If starY(X) > (formMidY) Then
    starY(X) = starY(X) + Int(Abs(formMidY - starY(X)) * starSpeed) * (starDist(X) * 0.2)
  Else
    starY(X) = starY(X) - Int(Abs(formMidY - starY(X)) * starSpeed) * (starDist(X) * 0.2)
  End If
  'see if the star has left the edge of the screen
  If starX(X) > Form1.Width Or starX(X) < 0 Or starY(X) > Form1.Height Or starY(X) < 0 Then
    'loops to check that the star is not in the exact center of the screen
    Do
      'set the stars (x,y) coords to random places
      starX(X) = Int(Rnd * Form1.Width)
      starY(X) = Int(Rnd * Form1.Height)
    Loop While (starX(X) = formMidX Or starY(Y) = formMidY)
    starDist(X) = 1
  End If
  'make sure that the star isn't getting too close
  'like the user is holding the mouse over a star
  If starDist(X) > 30 Then
    starDist(X) = 1
    starX(X) = Int(Rnd * Form1.Width)
    starY(X) = Int(Rnd * Form1.Height)
  End If
  'draw the star, setting the fill color to white
  Form1.FillColor = &HFFFFFF
  Circle (starX(X), starY(X)), starDist(X)
Next X
End Sub
```

