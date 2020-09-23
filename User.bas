Attribute VB_Name = "User"
'---------------------------------------------------
' DIRECT X PLUG V 1.0
'---------------------
'INTRO

'Welcome to my latest and greatest invention,
'Direct X Plug.  Basically what this project
'represents is a means to jump in and start
'programming directX applications immediately.

'There are three modules included:

' 1- The DirectX module.
'      This module contains everything you need
'      to make a simple directX application.  It
'      has been broken down into a handful of
'      simple subs that can be called from anywhere
'      in the application.

' 2- The PRF module.
'      PRF is a file packaging system I created.
'      The idea is to pack all your program's image
'      files into one big file (like a wad file)
'      and extract them when needed.  This gives
'      things a more professional look than having
'      50 or so bmp files laying open on the user's
'      hard drive.

' 3- The User Module.
'      This is where YOUR program goes!  You don't
'      HAVE to use this module, you can use a form
'      if you like or create a new module.  What
'      you see in the module right now is a sample
'      Application that demonstrates mouse usage.

'The whole purpose of DirectX Plug is to seperate
'the complicated processes of DirectX from the user's
'code.  This has a few advantages:

' - Simpler, cleaner code.
' - The ability to easily upgrade to a new version
'   of DirectX in the future. (with a new release of)
'   DirectX Plug.
' - You don't have to know DirectX to use it.
' - Easily port your application from using DirectX
'   To using BitBLT, or have dual functionality

'As usual I've tried to comment as much of the project
'as possible and keep the code readable.
'---------------------------------------------------

'Now on to the Declarations!



'This statement creates 52 slides (-1 thru 50)
'You could just as easily create 100 or 500,
'but be wary of memory and speed constraints.

'A slide is a neat little user-defined type that
'includes the Surface, Surface Description, and
'Some extra information about the Surface.
'It makes alot of sense to group all this stuff
'together into one object.
Global Slide(-1 To 50) As KSurface
  
   
'These delarations are for mouse stuff
'GetCursorPos is a really fast way of getting the
'mouse's current position from windows.
'ShowCursor is to hide windows' cursor so we can
'use our own.
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Sub Main()

Dim pt As PointAPI 'Variable for tracking the mouse

'Initialize DirectX
'You can specify any resolution supported by the
'user's hardware.  ex. (640, 480, 32) is 640x480
'resolution at true color (32 bit) color depth.
vari = DirectX_Initialize(800, 600, 16)

'Hey, get that windows Cursor off the screen,
'I'm making my own!
ShowCursor False




'Load the mouse image onto slide number -1.
'And then load a second image into another slide.
'For purposes of this demo, I reserved -1 for the
'mouse, this is by no means mandatory.
LoadImageToSlide App.Path & "\mouse.bmp", -1, 32, 32
LoadImageToSlide App.Path & "\testwindow.bmp", 1, 507, 453

Clrscn 'Clear the screen.

'Program loop--------------------------------------------------
   Do
     'During each loop, clear the screen first
     Clrscn
     
     'DirectX_initialize gives us a keyboard object,
     'so lets use it.
     Keyboard.Acquire
     Keyboard.GetDeviceStateKeyboard KeyboardState
          
     If Err.Number = 0 Then
     
        'If the user hits the left arrow key, exit the program
        If (KeyboardState.key(DIK_LEFT) And &H80) <> 0 Then Exit Do
     End If

     'Display slide 1 at a certain place on the screen
     Image_Blt 1, 100, 120
          
     'Get the current cursor position
     'This is 1000 times more accurate than the form's
     'Mouse_move Sub
     Call GetCursorPos(pt)
     Image_Blt -1, pt.X, pt.Y, True
     
     '"Flip" the screen to apply the changes
     Flip
   Loop
'--------------------------------------------------------------



errOut:
     'Make sure I get my windows cursor back!
     ShowCursor True
     
     'Exit the program, turn off directX and clean up
     Cancel_and_Close
     
     
     
'This sample program (excluding comments) is 21 lines long.
'It doesn't do anything terribly useful except that it
'Shows a picture on the screen as well as a custom mouse
'cursor.

'Once you understand this sample, you should delete this
'subroutine and replace it with your own DirectX program!


End Sub


