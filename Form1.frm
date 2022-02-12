VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Circular Buffer Test"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrRefresh 
      Interval        =   40
      Left            =   3120
      Top             =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bufferX_() As Byte
Private bufferY_() As Byte
Private Type tSize
    x As Long
    y As Long
End Type
Private Enum eDirection
    dUp = 0
    dRight
    dDown
    dLeft
End Enum
Private direction_ As eDirection
Private gridSize_ As Long
Private headSize_ As Long
Private moveCount_ As Currency
Private random_ As Long
Private snakeBodyX_ As CircularBufferByte
Private snakeBodyY_ As CircularBufferByte
Private snakeHeadX_ As Byte
Private snakeHeadY_ As Byte
Private turnChance_ As Double

Private Sub DrawSection(start As Long, length As Long)
    If length > 1 Then
        EnsureSizeX bufferX_, length
        EnsureSizeY bufferY_, length
        snakeBodyX_.CopyTo3 start, bufferX_, 0, length
        snakeBodyY_.CopyTo3 start, bufferY_, 0, length
        DrawLines bufferX_, bufferY_
    End If
End Sub
Private Sub DrawLines(ByRef bufferX() As Byte, ByRef bufferY() As Byte)
    Dim i&
    For i = 0 To (UBound(bufferX) - 2)
        Line (bufferX(i), bufferY(i))-(bufferX(i + 1), bufferY(i + 1)), vbBlue
    Next
End Sub
Private Sub EnsureSizeX(ByRef bufferX() As Byte, Size As Long)
    If (UBound(bufferX) + 1) <> Size Then
        ReDim bufferX(Size)
    End If
End Sub
Private Sub EnsureSizeY(ByRef bufferY() As Byte, Size As Long)
    If (UBound(bufferY) + 1) <> Size Then
        ReDim bufferY(Size)
    End If
End Sub
Private Function GetDistance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    Dim dx!, dy!
    dx = x1 - x2
    dy = y1 - y2
    GetDistance = CLng(Sqr((dx * dx) + (dy * dy)))
End Function
Private Function GetHeading() As tSize
    Dim x&, y&
    Dim result As tSize
    If direction_ = dUp Then
        result.x = 0
        result.y = -gridSize_
    ElseIf direction_ = dDown Then
        result.x = 0
        result.y = gridSize_
    ElseIf direction_ = dLeft Then
        result.x = -gridSize_
        result.y = 0
    ElseIf direction_ = dRight Then
        result.x = gridSize_
        result.y = 0
    Else
        result.x = 0
        result.y = 0
    End If
    GetHeading = result
End Function

Private Function GetRandomDirection() As eDirection
    Dim newDirection As eDirection
    Dim sample#
    sample = Rnd(random_)
    If sample < turnChance_ Then
        newDirection = direction_ - 1
        If newDirection < dUp Then
            newDirection = dLeft
        End If
    ElseIf sample > (1# - turnChance_) Then
        newDirection = newDirection + 1
        If newDirection > dLeft Then
            newDirection = dUp
        End If
    Else
        newDirection = direction_
    End If
    GetRandomDirection = newDirection
End Function

Private Sub NextMove()
    Dim x&, y&
    Dim heading As tSize
    Dim field As tSize
    moveCount_ = moveCount_ + 1
    snakeBodyX_.PutByte snakeHeadX_
    snakeBodyY_.PutByte snakeHeadY_
    
    heading = GetHeading()
    field.x = 255
    field.y = 255
    x = snakeHeadX_ + heading.x
    y = snakeHeadY_ + heading.y
    If x < 0 Then
        x = field.x - 1
    ElseIf x >= field.x Then
        x = 0
    End If
    If y < 0 Then
        y = field.y - 1
    ElseIf y >= field.y Then
        y = 0
    End If
    snakeHeadX_ = x
    snakeHeadY_ = y
    direction_ = GetRandomDirection()
    Me.Refresh
End Sub

Private Sub Form_Load()
    
    turnChance_ = 0.2
    Set snakeBodyX_ = New CircularBufferByte
    snakeBodyX_.Create 256, True
    Set snakeBodyY_ = New CircularBufferByte
    snakeBodyY_.Create 256, True
    snakeHeadX_ = 255 / 2
    snakeHeadY_ = 255 / 2
    direction_ = dRight
    ReDim bufferX_(0)
    ReDim bufferY_(0)
    gridSize_ = 12
    headSize_ = gridSize_ / 2
    random_ = 20200805
End Sub

Private Sub Form_Paint()

    Cls
    If snakeBodyX_.Size > 1 Or snakeBodyY_.Size > 1 Then
        Dim start&
        Dim previousX As Byte
        Dim previousY As Byte
        Dim currentX As Byte
        Dim currentY As Byte
        start = 0
        previousX = snakeBodyX_.PeekAtByte(0)
        previousY = snakeBodyY_.PeekAtByte(0)
        Dim i&
        For i = 1 To (snakeBodyX_.Size - 1)
            currentX = snakeBodyX_.PeekAtByte(i)
            currentY = snakeBodyY_.PeekAtByte(i)
            If GetDistance(previousX, previousY, currentX, currentY) > gridSize_ Then
                DrawSection start, i - start
                start = i
            End If
            previousX = currentX
            previousY = currentY
        Next
        If start < snakeBodyX_.Size Then
            DrawSection start, snakeBodyX_.Size - start
        End If
    End If
    Circle (snakeHeadX_, snakeHeadY_), headSize_

End Sub

Private Sub tmrRefresh_Timer()
    NextMove
End Sub
