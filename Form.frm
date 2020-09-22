VERSION 5.00
Begin VB.Form frmAnalog 
   Caption         =   "Analog watch"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Start 
      Caption         =   "Start clock"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Line lineSec 
      BorderColor     =   &H000000FF&
      X1              =   96
      X2              =   97
      Y1              =   96
      Y2              =   97
   End
   Begin VB.Line lineMin 
      BorderWidth     =   3
      X1              =   96
      X2              =   97
      Y1              =   96
      Y2              =   97
   End
   Begin VB.Line Dot 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   1
      X2              =   2
      Y1              =   1
      Y2              =   2
   End
   Begin VB.Line lineHour 
      BorderWidth     =   5
      X1              =   96
      X2              =   97
      Y1              =   96
      Y2              =   97
   End
End
Attribute VB_Name = "frmAnalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private angleDegSec As Integer           'Angle in degrees (Seconds)
Private angleRadSec As Double            'Angle in radians (Seconds)
Private angleDegMin As Integer           'Angle in degrees (Minutes)
Private angleRadMin As Double            'Angle in radians (Minutes)
Private angleDegHour As Double           'Angle in degrees (Hours)
Private angleRadHour As Double           'Angle in radians (Hours)
Private curSec As Byte                   'current Second (used for comparation)

Private Stoped As Boolean                'Is clock running or not
Const Pi As Double = 22 / 7              'Pi
Const lineLength As Integer = 90         'Length of the line indicating seconds
Option Explicit

Private Sub Start_Click()
  'Upon clicking the button you either activate or stop the clock
  'Here the caption for the button is chosed
  If Stoped Then
    Start.Caption = "Stop clock"
  Else
    Start.Caption = "Start clock"
  End If
  Stoped = Not (Stoped)                  'Stoped is changing value
  
  'Now the main loop
  Do
    'If the seconds have changed change the pointers as well
    If Second(Time) <> curSec Then
      'Setting angle of second pointer in degrees by current time.
      'Number of seconds is multiplied by 6 because there are 60 seconds in a minute,
      'and in a minute the second pointer passes a full cicrle(360 degs), thus
      'one second is equivalent to the angle of 360/60=6 degs
      angleDegSec = Second(Time) * 6
      'Turning degrees to radians (because of the sin and cos operations)
      angleRadSec = Pi * angleDegSec / 180
      'Sin(angle)=a/c
      'where a is X2-X1, and
      'c is lineLength, thus X2 is X1 + sin(angle) * lineLength
      lineSec.X2 = lineSec.X1 + Sin(angleRadSec) * lineLength
      'Cos(angle)=b/c
      'where b is Y1-Y2(because the system starts at the top-left corner of the form),
      'and c is lineLegth, thus Y2 is Y1 - cos(angle) * lineLength
      lineSec.Y2 = lineSec.Y1 - Cos(angleRadSec) * lineLength
      
      'Setting angle of minute pointer in degrees by current time.
      'Same as for seconds, but it's for minutes now
      angleDegMin = Minute(Time) * 6
      'Turning degrees to radians (because of the sin and cos operations)
      angleRadMin = Pi * angleDegMin / 180
      'Calculating X2 and Y2, with lineLength reduced by 5 pixels
      lineMin.X2 = lineMin.X1 + Sin(angleRadMin) * (lineLength - 5)
      lineMin.Y2 = lineMin.Y1 - Cos(angleRadMin) * (lineLength - 5)
    
      'Setting angle of hour pointer in degrees by current time.
      'There are twelve hours in a full circle, and for each of them the hour pointer
      'moves for 360/12=30 degs, and there are 60 minutes in a hour, and for each of
      'them the hour pointer moves for 30/60=0.5 degs
      angleDegHour = Hour(Time) * 30 + Minute(Time) * 0.5
      'Turning degrees to radians (because of the sin and cos operations)
      angleRadHour = Pi * angleDegHour / 180
      'Calculating X2 and Y2, with lineLength reduced by 20 pixels
      lineHour.X2 = lineHour.X1 + Sin(angleRadHour) * (lineLength - 20)
      lineHour.Y2 = lineHour.Y1 - Cos(angleRadHour) * (lineLength - 20)
  
      'Adjusting the current seconds number, this is used to execute this if statement
      'each time the number of seconds change
      curSec = Second(Time)
    End If
    'Some time to do the events
    DoEvents
  Loop Until Stoped
End Sub

Private Sub Form_Load()
  'Upon loading adjust current Second number, make sure that the clock is stoped, and
  'draw the dots
  curSec = Second(Time)
  Stoped = True
  LoadDots
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Stoped = True                           'Stop the clock (exit the loop)
End Sub

Private Sub LoadDots()
  Dim i As Byte                           'Number determing the dot number
  Dim angleRadDots As Double              'Angle of the dots' position in radians
  
  'Setting coordinates of the first dot(already created), using second pointer
  Dot(0).X1 = lineSec.X1
  Dot(0).X2 = Dot(0).X1
  Dot(0).Y1 = lineSec.Y1 - (lineLength + 5)
  Dot(0).Y2 = Dot(0).Y1
  
  For i = 1 To 59
    'Calculating angle in radians, as were the angle in degrees for minutes and seconds
    'multiplied by 6, so is the case with dots
    angleRadDots = Pi * (i * 6) / 180
    'Create new dot(dot is truly a line) at run-time
    Load Dot(i)
    'If it's not the dot that should be bigger, then reduce it and paint it black
    If (i Mod 5 <> 0) Then
      Dot(i).BorderWidth = 2
      Dot(i).BorderColor = vbBlack
    End If
    'Set the dot to be visible
    Dot(i).Visible = True
    'Calculate the coordinates(similar to calculating for seconds and minutes),
    'with lineLength extended by 5 pixels
    Dot(i).X1 = lineSec.X1 + Sin(angleRadDots) * (lineLength + 5)
    Dot(i).X2 = Dot(i).X1
    Dot(i).Y1 = lineSec.Y1 - Cos(angleRadDots) * (lineLength + 5)
    Dot(i).Y2 = Dot(i).Y1
  Next i
End Sub
