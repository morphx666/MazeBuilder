Imports System.Drawing.Imaging
Imports System.Threading
Imports System.Runtime.InteropServices
Imports MazeBuilder.ColorSpace
Imports System.Runtime.CompilerServices

Public Class FormMain
    Private Class MazePoint
        Private mPoint As Point

        Public Property HLS As HLSRGB

        Public ReadOnly Property Point As Point
            Get
                Return mPoint
            End Get
        End Property

        Public Property X As Integer
            Get
                Return mPoint.X
            End Get
            Set(value As Integer)
                mPoint.X = value
            End Set
        End Property

        Public Property Y As Integer
            Get
                Return mPoint.Y
            End Get
            Set(value As Integer)
                mPoint.Y = value
            End Set
        End Property

        Public ReadOnly Property Color As Color
            Get
                Return HLS.Color
            End Get
        End Property

        Public Sub New(p As Point, c As HLSRGB)
            mPoint = p
            HLS = c
        End Sub

        Public Function Clone() As MazePoint
            Return New MazePoint(mPoint, New HLSRGB(HLS))
        End Function
    End Class

    Private mDotSize As Integer = 4
    Private mSpacing As Integer = 2

    Private mainSurface As Bitmap
    Private secondarySurface As Bitmap
    Private cancelThread As Boolean
    Private sourceData As BitmapData
    Private sourcePointer As System.IntPtr
    Private sourceStride As Integer
    Private bytesPerPixel As Integer
    Private moveStep As Integer
    Private delay As Integer = 10
    Private useColor As Boolean = True

    Private mainSurfaceRect As Rectangle
    Private helpRect As Rectangle = Rectangle.Empty

    Private lockForMainSurface As New Object()
    Private lockForSecondarySurface As New Object()
    Private threadsCount As Integer
    Private pixelFormat As PixelFormat

    Private sw As New Stopwatch()

    Private helpVisible As Boolean = True

    Private Enum DirectionConstants
        Top = 0
        Right = 1
        Bottom = 2
        Left = 3
        Any = 10
    End Enum
    Private direction As DirectionConstants

    Private Sub FormMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        StopThreads()
    End Sub

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
        Me.SetStyle(ControlStyles.ResizeRedraw, True)
        Me.SetStyle(ControlStyles.UserPaint, True)
        Me.KeyPreview = True

        SetMoveStep()
        Start()
    End Sub

    Private Sub FormMain_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        Select Case e.KeyCode
            Case Keys.PageUp : delay = Math.Min(delay + 1, 1000)
            Case Keys.PageDown : delay = Math.Max(delay - 1, 0)

            Case Keys.Up : DotSize += 1
            Case Keys.Down : DotSize -= 1
            Case Keys.Right : Spacing += 1
            Case Keys.Left : Spacing -= 1

            Case Keys.Escape : Me.Close()
            Case Keys.Enter : Start()

            Case Keys.H, Keys.F1 : helpVisible = Not helpVisible
            Case Keys.C : useColor = Not useColor
        End Select

        Me.Invalidate()
    End Sub

    Private Property DotSize As Integer
        Get
            Return mDotSize
        End Get
        Set(value As Integer)
            mDotSize = Math.Max(value, 1)
            SetMoveStep()
        End Set
    End Property

    Private Property Spacing As Integer
        Get
            Return mSpacing
        End Get
        Set(value As Integer)
            mSpacing = Math.Max(value, 1)
            SetMoveStep()
        End Set
    End Property

    Private Sub SetMoveStep()
        moveStep = mDotSize + mSpacing
    End Sub

    Private Sub Start()
        StopThreads()

        sw.Restart()
        cancelThread = False

        If mainSurface IsNot Nothing Then mainSurface.Dispose()
        If secondarySurface IsNot Nothing Then secondarySurface.Dispose()

        mainSurface = New Bitmap(Me.DisplayRectangle.Width, Me.DisplayRectangle.Height, PixelFormat.Format24bppRgb)
        secondarySurface = CType(mainSurface.Clone(), Bitmap)
        mainSurfaceRect = New Rectangle(0, 0, mainSurface.Width, mainSurface.Height)
        pixelFormat = mainSurface.PixelFormat

        sourceData = mainSurface.LockBits(mainSurfaceRect, ImageLockMode.ReadWrite, pixelFormat)
        sourcePointer = sourceData.Scan0
        sourceStride = sourceData.Stride
        bytesPerPixel = sourceStride \ mainSurfaceRect.Width
        mainSurface.UnlockBits(sourceData)

        Me.Invalidate()

        Dim rnd As Random = New Random(Now.Second)
        StartThread(New Point(rnd.Next(Me.DisplayRectangle.Width), rnd.Next(Me.DisplayRectangle.Height)))
    End Sub

    Private Sub StartThread(p As Point)
        Dim drawMazeThread As Thread = New Thread(New ParameterizedThreadStart(AddressOf DrawMaze))
        drawMazeThread.Start(p)
    End Sub

    Private Sub StopThreads()
        cancelThread = True
        Do
            Application.DoEvents()
        Loop While threadsCount > 0
    End Sub

    Private Sub DrawMaze(origin As Point)
        threadsCount += 1

        Dim rnd As Random = New Random(My.Computer.Clock.TickCount)
        Dim thisThreadIndex As Integer = threadsCount

        origin.X -= origin.X Mod moveStep
        origin.Y -= origin.Y Mod moveStep

        Dim p As MazePoint

        If useColor Then
            p = New MazePoint(origin, New HLSRGB(Color.Yellow))
            SyncLock lockForMainSurface
                SetHue(p)
            End SyncLock
        Else
            p = New MazePoint(origin, New HLSRGB(Color.White))
        End If

        Dim history((mainSurfaceRect.Width * mainSurfaceRect.Height) / (mDotSize * mSpacing) - 1)
        history(0) = p.Clone()

        Dim i As Integer = 1
        Do
            SyncLock lockForMainSurface
                sourceData = mainSurface.LockBits(mainSurfaceRect, ImageLockMode.ReadWrite, pixelFormat)
                If CanMove(p) Then
                    Do
                        direction = CType(rnd.Next(0, 4), DirectionConstants)
                    Loop Until CanMove(p, direction)

                    Select Case direction
                        Case DirectionConstants.Top : p.Y -= moveStep
                        Case DirectionConstants.Right : p.X += moveStep
                        Case DirectionConstants.Bottom : p.Y += moveStep
                        Case DirectionConstants.Left : p.X -= moveStep
                    End Select
                    history(i) = p.Clone()
                    ConnectPoints(history(i - 1), history(i))

                    If thisThreadIndex = threadsCount Then
                        UpdateSecondarySurface()
                        If helpVisible Then Me.Invalidate(helpRect)
                    End If
                    UpdateDisplay(GetRectFromPoints(history(i - 1).Point, history(i).Point))

                    If useColor Then SetHue(p)

                    i += 1
                    If delay > 0 Then Thread.Sleep(delay / threadsCount)
                Else
                    For i = i - 1 To 0 Step -1
                        If CanMove(history(i)) Then
                            p = history(i).Clone()
                            i += 1
                            Exit For
                        End If
                    Next
                End If
                mainSurface.UnlockBits(sourceData)
            End SyncLock
        Loop Until cancelThread OrElse i = -1

        threadsCount -= 1
        If threadsCount = 0 Then sw.Stop()

        UpdateDisplay(Me.DisplayRectangle)
    End Sub

    Private Sub SetHue(p As MazePoint)
        p.HLS.Hue = (p.X * p.Y) / (mainSurfaceRect.Width * mainSurfaceRect.Height) * 360
    End Sub

    Private Function GetRectFromPoints(p1 As Point, p2 As Point) As Rectangle
        Return New Rectangle(Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y), Math.Abs(p1.X - p2.X) + mDotSize, Math.Abs(p1.Y - p2.Y) + mDotSize)
    End Function

    Private Sub UpdateSecondarySurface()
        SyncLock lockForSecondarySurface
            If secondarySurface IsNot Nothing Then secondarySurface.Dispose()
            secondarySurface = New Bitmap(mainSurfaceRect.Width, mainSurfaceRect.Height, sourceStride, pixelFormat, sourcePointer)
        End SyncLock
    End Sub

    Private Sub UpdateDisplay(r As Rectangle)
        If Me.DisplayRectangle.Size <> mainSurfaceRect.Size Then
            r.Location = PointBitmapToDisplay(r.Location)
            r.Size = PointBitmapToDisplay(r.Size)
        End If
        Me.Invalidate(r, False)
    End Sub

    Private Sub ConnectPoints(p1 As MazePoint, p2 As MazePoint)
        Dim sign As Integer
        Dim c = p1.Color

        If p1.X = p2.X Then
            sign = Math.Sign(p2.Y - p1.Y)
            If sign = 0 Then sign = 1
            For y As Integer = p1.Y To p2.Y Step sign
                For dx As Integer = 0 To mDotSize - 1
                    For dy As Integer = 0 To mDotSize - 1
                        SetPixel(p1.X + dx, y + dy, c)
                    Next
                Next
            Next
        Else
            sign = Math.Sign(p2.X - p1.X)
            If sign = 0 Then sign = 1
            For x As Integer = p1.X To p2.X Step sign
                For dx As Integer = 0 To mDotSize - 1
                    For dy As Integer = 0 To mDotSize - 1
                        SetPixel(x + dx, p1.Y + dy, c)
                    Next
                Next
            Next
        End If
    End Sub

    Private Function IsPixelSet(x As Integer, y As Integer) As Boolean
        Dim c As Color = GetPixel(x, y)
        Return c.R <> 0 OrElse c.G <> 0 OrElse c.B <> 0
    End Function

    Private Function CanMove(p As MazePoint, Optional direction As DirectionConstants = DirectionConstants.Any) As Boolean
        Select Case direction
            Case DirectionConstants.Top : Return p.Y - moveStep > 0 AndAlso Not IsPixelSet(p.X, p.Y - moveStep)
            Case DirectionConstants.Right : Return p.X + moveStep < mainSurfaceRect.Width - moveStep AndAlso Not IsPixelSet(p.X + moveStep, p.Y)
            Case DirectionConstants.Bottom : Return p.Y + moveStep < mainSurfaceRect.Height - moveStep AndAlso Not IsPixelSet(p.X, p.Y + moveStep)
            Case DirectionConstants.Left : Return p.X - moveStep > 0 AndAlso Not IsPixelSet(p.X - moveStep, p.Y)
            Case DirectionConstants.Any
                Return CanMove(p, DirectionConstants.Top) OrElse
                        CanMove(p, DirectionConstants.Right) OrElse
                        CanMove(p, DirectionConstants.Bottom) OrElse
                        CanMove(p, DirectionConstants.Left)
        End Select
    End Function

    Private Function GetPixel(x As Integer, y As Integer) As Color
        Dim offset As Integer = x * bytesPerPixel + y * sourceStride
        Dim blue As Byte = Marshal.ReadByte(sourcePointer, offset + 0)
        Dim green As Byte = Marshal.ReadByte(sourcePointer, offset + 1)
        Dim red As Byte = Marshal.ReadByte(sourcePointer, offset + 2)

        Return Color.FromArgb(red, green, blue)
    End Function

    Private Sub SetPixel(point As Point, color As Color)
        SetPixel(point.X, point.Y, color)
    End Sub

    Private Sub SetPixel(x As Integer, y As Integer, color As Color)
        Dim offset As Integer = x * bytesPerPixel + y * sourceStride
        Marshal.WriteByte(sourcePointer, offset + 0, color.B)
        Marshal.WriteByte(sourcePointer, offset + 1, color.G)
        Marshal.WriteByte(sourcePointer, offset + 2, color.R)
    End Sub

    Private Function PointDisplayToBitmap(p As Point) As Point
        Return New Point(mainSurfaceRect.Width / Me.DisplayRectangle.Width * p.X,
                        mainSurfaceRect.Height / Me.DisplayRectangle.Height * p.Y)
    End Function

    Private Function PointBitmapToDisplay(p As Point) As Point
        Return New Point(Me.DisplayRectangle.Width / mainSurfaceRect.Width * p.X,
                        Me.DisplayRectangle.Height / mainSurfaceRect.Height * p.Y)
    End Function

    Private Sub FormMain_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Left Then StartThread(PointDisplayToBitmap(e.Location))
    End Sub

    Private Sub FormMain_MouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        delay += 1 * Math.Sign(e.Delta)
        If delay > 200 Then
            delay = 200
        ElseIf delay < 0 Then
            delay = 0
        End If
    End Sub

    Private Sub FormMain_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If secondarySurface Is Nothing Then Exit Sub

        Dim g As Graphics = e.Graphics

        SyncLock lockForSecondarySurface
            g.DrawImage(secondarySurface, 0, 0, Me.ClientRectangle.Width, Me.ClientRectangle.Height)
        End SyncLock

        If helpVisible Then
            Const helpLines As Integer = 13
            Dim s As SizeF = g.MeasureString("X", Me.Font)
            helpRect = New Rectangle(10, 10, s.Width * 34, 10 + s.Height * helpLines + 5)

            Using b As New SolidBrush(Color.FromArgb(220, Color.LightYellow))
                g.FillRectangle(b, helpRect)
            End Using

            Dim p As New Point(15, 15)
            g.DrawString("Click anywhere to start a new 'maze building' thread", Me.Font, Brushes.Black, p)
            p.Y += s.Height * 2
            g.DrawString($"PageUp/PageDown (or mouse scroll=wheel) changes delay ({delay}ms)", Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString($"Up/Down changes dot size ({mDotSize})", Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString($"Left/Right changes spacing ({mSpacing})", Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString($"'C' enable/disable rainbow effect ({If(useColor, "Enabled", "Disabled")})", Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString("'F1/H' display/hide this help", Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString("'ENTER' start over", Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString("'ESC' terminate the program", Me.Font, Brushes.Black, p)

            p.Y += s.Height
            g.DrawLine(Pens.Gray, p.X + 2, p.Y + s.Height / 2, helpRect.Right - 4, p.Y + s.Height / 2)

            p.Y += s.Height
            g.DrawString(String.Format("Bitmap Size: {0}x{1}", mainSurfaceRect.Width, mainSurfaceRect.Height), Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString(String.Format("Active Threads: {0}", threadsCount), Me.Font, Brushes.Black, p)
            p.Y += s.Height
            g.DrawString(String.Format("Elapsed: {0:00}:{1:00}:{2:00}:{3:000}", sw.Elapsed.Hours, sw.Elapsed.Minutes, sw.Elapsed.Seconds, sw.Elapsed.Milliseconds), Me.Font, Brushes.Black, p)
        End If
    End Sub
End Class
