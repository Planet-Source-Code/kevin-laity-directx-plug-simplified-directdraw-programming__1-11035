Attribute VB_Name = "DirectX"
Type KSurface
    Mode As Byte
    ScreenX As Integer
    screeny As Integer
    SURF As DirectDrawSurface7
    DESC As DDSURFACEDESC2
    Width As Integer
    Height As Integer
End Type

Public Type PointAPI
    X As Long
    Y As Long
End Type


Global binit As Boolean                  'Used by the program to see if everythings been started.
Global DirectX As New DirectX7           'Master Object, Everything is created from this.
Global ddraw As DirectDraw7              'The DirectDraw Object, this is created from the DirectX Object
  
Global Primary As KSurface               'Whats on the screen
Global Backbuffer As KSurface            'The Backbuffer
 
Global DirectInput As DirectInput
Global Keyboard As DirectInputDevice
Global KeyboardState As DIKEYBOARDSTATE
Global strKeyboardState As String

Global ScreenX As Long
Global screeny As Long
Global screendepth As Long

Sub Cancel_and_Close()

    Call ddraw.RestoreDisplayMode
    Call ddraw.SetCooperativeLevel(DirectXForm.hWnd, DDSCL_NORMAL)
    
    Keyboard.Unacquire
    Set Keyboard = Nothing
    Set DirectInput = Nothing

    Unload DirectXForm
    
    End

End Sub

Sub Clrscn()

Dim FillRect As RECT
With FillRect
   .Top = 0: .Left = 0
   .Right = 0: .Bottom = 0
End With
vari = Backbuffer.SURF.BltColorFill(FillRect, 0)

End Sub

Function DirectX_Initialize(Resolution_X As Long, Resolution_Y As Long, ColorDepth As Long)

    ScreenX = Resolution_X
    screeny = Resolution_Y
    screendepth = ColorDepth
    
    DirectXForm.Show

    Set DirectInput = DirectX.DirectInputCreate()
    Set Keyboard = DirectInput.CreateDevice("GUID_SysKeyboard")
    Keyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
    Keyboard.SetCooperativeLevel DirectXForm.hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND

    On Local Error GoTo errOut
    
    Set ddraw = DirectX.DirectDrawCreate("")
    Call ddraw.SetCooperativeLevel(DirectXForm.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    Call ddraw.SetDisplayMode(Resolution_X, Resolution_Y, ColorDepth, 0, DDSDM_DEFAULT)
    Primary.DESC.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    Primary.DESC.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    Primary.DESC.lBackBufferCount = 1
         
    DirectXForm.Picture = Nothing
     
    Set Primary.SURF = ddraw.CreateSurface(Primary.DESC)
    Dim caps As DDSCAPS2
    caps.lCaps = DDSCAPS_BACKBUFFER
    Set Backbuffer.SURF = Primary.SURF.GetAttachedSurface(caps)
    Backbuffer.SURF.GetSurfaceDesc Backbuffer.DESC

    Backbuffer.SURF.SetFontTransparency True
    Backbuffer.SURF.SetForeColor vbWhite
    'Backbuffer.SetFont
     
    binit = True
    brunning = True
    Exit Function
     
errOut:
    'If anything goes wrong, close the program
    Cancel_and_Close

End Function

Sub Flip()

 Primary.SURF.Flip Nothing, DDFLIP_WAIT


End Sub

Sub Image_Blt(idx, ByVal X As Long, ByVal Y As Long, Optional transparent As Boolean)

'This Sub is a fast way to display a slide.
'By passing the transparent variable as true,
'you can make all the black in the slide
'see-through

'It will automatically clip an image if it goes
'outside the screen,  which will simplify your
'program.


With Slide(idx)

    Dim temp As RECT

    temp.Top = 0: temp.Left = 0
    If X < 0 Then temp.Left = Abs(X): X = 0
    If Y < 0 Then temp.Top = Abs(Y): Y = 0
    
    If .DESC.lWidth + X > ScreenX Then
    temp.Right = .DESC.lWidth - (.DESC.lWidth + X - ScreenX)
    Else
    temp.Right = .DESC.lWidth
    End If
    
    If .DESC.lHeight + Y > screeny Then
    temp.Bottom = .DESC.lHeight - (.DESC.lHeight + Y - screeny)
    Else
    temp.Bottom = .DESC.lHeight
    End If
    
    
    Dim dest As RECT
    dest.Top = Y: dest.Left = X
    dest.Bottom = screeny: dest.Right = ScreenX
    
    If transparent Then
       rval = Backbuffer.SURF.BltFast(X, Y, .SURF, temp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
    Else
       rval = Backbuffer.SURF.BltFast(X, Y, .SURF, temp, DDBLTFAST_WAIT)
    End If
   
End With


End Sub

Sub LoadImageToSlide(filename As String, SlideIndex As Single, PixelWidth As Integer, PixelHeight As Integer)

Slide(SlideIndex).DESC.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
Slide(SlideIndex).DESC.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Slide(SlideIndex).DESC.lWidth = PixelWidth
Slide(SlideIndex).DESC.lHeight = PixelHeight

Set Slide(SlideIndex).SURF = ddraw.CreateSurfaceFromFile(filename, Slide(SlideIndex).DESC)



Dim key As DDCOLORKEY
key.low = 0
key.high = key.low


Slide(SlideIndex).SURF.SetColorKey DDCKEY_SRCBLT, key


End Sub


