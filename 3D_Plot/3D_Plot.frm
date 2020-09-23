VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "3D Plot Concepts Demo"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Draw_Axes_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw XYZ Axes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6525
      TabIndex        =   11
      ToolTipText     =   " XYZ Color Codes:     X = Red     Y = Green     Z = Blue "
      Top             =   720
      Width           =   1500
   End
   Begin VB.CommandButton Draw_Cube_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw Cube"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6525
      TabIndex        =   1
      Top             =   1125
      Width           =   1500
   End
   Begin VB.CommandButton Clear_Button 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6525
      TabIndex        =   10
      Top             =   1980
      Width           =   1500
   End
   Begin VB.TextBox Perspective_Factor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   5085
      TabIndex        =   5
      Text            =   "10000"
      Top             =   315
      Width           =   1320
   End
   Begin VB.TextBox Size_Factor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3825
      TabIndex        =   4
      Text            =   "10000"
      Top             =   315
      Width           =   1185
   End
   Begin VB.TextBox Alt_Ang 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1485
      TabIndex        =   3
      Text            =   "35"
      Top             =   315
      Width           =   1275
   End
   Begin VB.TextBox Theta_Ang 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      TabIndex        =   2
      Text            =   "30"
      Top             =   315
      Width           =   1365
   End
   Begin VB.PictureBox Plot 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   5505
      Left            =   45
      ScaleHeight     =   5445
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   720
      Width           =   6405
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Perspective"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   5130
      TabIndex        =   9
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Size Factor"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3825
      TabIndex        =   8
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Altitude of Eye"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1485
      TabIndex        =   7
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Direction of Eye"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   45
      TabIndex        =   6
      Top             =   90
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Option Explicit
'
' A simple program to demonstrate the basic essentials
' of plotting points in 3D space onto a 2D monitor screen
' by using the sine and cosine functions to perform the
' coordinate transformations in space and projections onto
' a plane (the monitor screen).
'
' The trig functions used by this program are modified to use
' degree instead of radian arguments for convenience.

' Written using Visual BASIC 6
'
' Author:  Jay Tanner
'          Jay@NeoProgrammics.com





' ===========================================
' Code to execute for the [Draw Cube] button

  Private Sub Draw_Cube_Button_Click()

' Define main plot parameters
  Dim Theta       As Single ' Azimuth of the eye
  Dim Alt         As Single ' Altitude of the eye
  Dim Size        As Single ' Image size scale
  Dim Perspective As Single ' Image perspective scale

' The angle (Theta) is the azimuth of the eye, or the
' direction to the eye as measured counterclockwise from
' the orgin of coordinates.

' The angle (Alt) is the angular altitude of the eye above
' the ground (XY) plane as measured upward from the origin.
' Zero is on the horizon and 90 degrees would be a view from
' directly above.

' Both the following (Size) and (Perspective) values alter the
' apparent size of the image in different ways.

' The (Size) factor controls the general apparent size of the
' image.  The larger the value, the larger the image.  It is
' the distance of the eye from the plane upon which the image
' is projected.
'
' The (Perspective) value controls the apparent perspective of
' the image and is a measure of how parallel the imaginary light
' rays projecting the image are.  It is the distance between the
' eye and the projected points.  The greater the value, the more
' parallel the light rays and the smaller the image and less the
' distortion caused by perspective.





' Display hourglass pointer while program is working
  Form1.MousePointer = vbHourglass

' Read scene parameters from interface text boxes.
' The angles Theta and Phi are in degrees.
  Perspective = Val(Perspective_Factor)
         Size = Val(Size_Factor)
          Alt = Val(Alt_Ang)
        Theta = Val(Theta_Ang)

  DRAW_A_CUBE Theta, Alt, Size, Perspective

' Restore mouse pointer to normal when plotting finished
  Form1.MousePointer = vbDefault

  End Sub

' ==============================================
' Code to execute for the [Draw XYZ Axes] button

  Private Sub Draw_Axes_Button_Click()

' Define main plot parameters
  Dim Theta       As Single ' Azimuth of eye
  Dim Alt         As Single ' Altitude of eye
  Dim Size        As Single ' Image size scale
  Dim Perspective As Single ' Image perspective scale

  Form1.MousePointer = vbHourglass

' Read scene parameters from interface text boxes.
' The angles Theta and Phi are in degrees.
  Perspective = Val(Perspective_Factor)
         Size = Val(Size_Factor)
          Alt = Val(Alt_Ang)
        Theta = Val(Theta_Ang)

  DRAW_XYZ_AXES Theta, Alt, Size, Perspective

  Form1.MousePointer = vbDefault

  End Sub



' =======================================================
' SUB to draw a 3D cube with a circle on front face side

  Private Sub DRAW_A_CUBE(Theta, Alt, Size, Perspective)

  Dim x     As Single
  Dim y     As Single
  Dim z     As Single
  Dim Angle As Single

  Plot.DrawWidth = 1
  Plot.ForeColor = RGB(92, 92, 92)

' Bottom right edge
  For x = -1000 To 1000
      y = -1000
      z = 0
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next x

' Bottom left edge
  For x = -1000 To 1000
      y = 1000
      z = 0
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next x

' Bottom front edge
  For y = -1000 To 1000
      x = 1000
      z = 0
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next y

' Bottom back edge
  For y = -1000 To 1000
      x = -1000
      z = 0
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next y

' Top left edge
  For x = -1000 To 1000
      y = -1000
      z = 2000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next x

' Top right edge
  For x = -1000 To 1000
      y = 1000
      z = 2000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next x

' Top front edge
  For y = -1000 To 1000
      x = 1000
      z = 2000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next y

' Top back edge
  For y = -1000 To 1000
      x = -1000
      z = 2000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next y

' Front left vertical edge
  For z = 0 To 2000
      x = 1000
      y = -1000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next z

' Front right vertical edge
  For z = 0 To 2000
      x = 1000
      y = 1000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next z

' Back left vertical edge
  For z = 0 To 2000
      x = -1000
      y = -1000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next z

' Back right vertical edge
  For z = 0 To 2000
      x = -1000
      y = 1000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next z

' Draw a dark yellow circle on top of cube
  Plot.DrawWidth = 2
  Plot.ForeColor = RGB(128, 128, 0)
  For Angle = 0 To 360
      y = 1000 * Cosine(Angle)
      z = 1000 * Sine(Angle) + 1000
      x = 1000
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next Angle

  End Sub


' ==========================================================
' SUB to plot a single dot for given XYZ and parameters.

  Private Sub Plot_Dot _
 (x, y, z, Theta, Alt, Size, Perspective)
  
' Define center coordinates of plotting area picture box
  Dim cX As Single
  Dim cY As Single

' Define 3D viewpoint (eye) coordinates
  Dim vX As Single
  Dim vY As Single
  Dim vZ As Single

' Define 2D screen (X,Y) plotting coordinates
  Dim pX As Single
  Dim pY As Single

' Define zenith distance angle.
  Dim Phi As Single
      Phi = 90 - Alt

' Define sines and cosines of Theta and Phi
  Dim Sin_Theta As Single
  Dim Cos_Theta As Single
  Dim Sin_Phi   As Single
  Dim Cos_Phi   As Single
  
' Set the center coordinate values of the plotting area
  cX = Plot.Width / 2
  cY = Plot.Height / 2

' Compute the sines and cosines of the Theta and Phi angles.
' This way they don't have to be computed more than once and
' it speeds things up a tiny bit.
  Sin_Theta = Sine(Theta)
  Cos_Theta = Cosine(Theta)
    Sin_Phi = Sine(Phi)
    Cos_Phi = Cosine(Phi)

' Compute viewpoint (eye) coordinates of point (X,Y,Z)
  vX = -x * Sin_Theta _
      + y * Cos_Theta

  vY = -x * Cos_Theta * Cos_Phi _
      - y * Sin_Theta * Cos_Phi _
      + z * Sin_Phi

  vZ = -x * Cos_Theta * Sin_Phi _
      - y * Sin_Theta * Sin_Phi _
      - z * Cos_Phi + Perspective

' Compute 2D screen plotting coordinates corresponding to
' point (X,Y,Z) in the 3D world.
  pX = cX + Size * vX / vZ
  pY = cY - Size * vY / vZ

' Plot the dot using the screen plotting area using the most
' recently set plotting (foreground) color.  The value 1000
' is an offset used to help center the image.
  Plot.PSet (pX, pY + 1000)

  End Sub

' ========================================
' Draw XYZ axes in colors RGB respectively

  Private Sub DRAW_XYZ_AXES(Theta, Alt, Size, Perspective)

' Define dimensions
  Dim x As Single
  Dim y As Single
  Dim z As Single

' Set to smallest line width
  Plot.DrawWidth = 1

' Draw X axis in dark red
  Plot.ForeColor = RGB(92, 0, 0)
  For x = 0 To 3000
      y = 0
      z = 0
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next x

' Draw Y axis in dark green
  Plot.ForeColor = RGB(0, 92, 0)
  For y = 0 To 3000
      x = 0
      z = 0
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next y

' Draw Z axis in dark blue
  Plot.ForeColor = RGB(0, 0, 128)
  For z = 0 To 3000
      x = 0
      y = 0
      Plot_Dot x, y, z, Theta, Alt, Size, Perspective
  Next z

  End Sub

' Clear the plotting area
  Private Sub Clear_Button_Click()
  Plot.Cls
  End Sub

' ============================================
' Modified trig functions for degree agruments

  Public Function Sine(Degrees_Arg)
' Level 00
  Sine = Sin(Degrees_Arg * Atn(1) / 45)
  End Function

  Public Function Cosine(Degrees_Arg)
' Level 00
  Cosine = Cos(Degrees_Arg * Atn(1) / 45)
  End Function

