Public Class ColorSpace
    <Serializable()>
    Public Class HLSRGB
        Private mRed As Byte = 0
        Private mGreen As Byte = 0
        Private mBlue As Byte = 0
        Private mAlpha As Byte = 255

        Private mHue As Single = 0
        Private mLuminance As Single = 0
        Private mSaturation As Single = 0

        Public Structure HueLumSat
            Private mH As Single
            Private mL As Single
            Private mS As Single

            Public Sub New(hue As Single, lum As Single, sat As Single)
                mH = hue
                mL = lum
                mS = sat
            End Sub

            Public Property Hue() As Single
                Get
                    Return mH
                End Get
                Set(value As Single)
                    mH = value
                End Set
            End Property

            Public Property Lum() As Single
                Get
                    Return mL
                End Get
                Set(value As Single)
                    mL = value
                End Set
            End Property

            Public Property Sat() As Single
                Get
                    Return mS
                End Get
                Set(value As Single)
                    mS = value
                End Set
            End Property
        End Structure

        Public Sub New(c As Color)
            mRed = c.R
            mGreen = c.G
            mBlue = c.B
            mAlpha = c.A
            ToHLS()
        End Sub

        Public Sub New(hue As Single, luminance As Single, saturation As Single)
            mHue = hue
            mLuminance = luminance
            mSaturation = saturation
            ToRGB()
        End Sub

        Public Sub New(red As Byte, green As Byte, blue As Byte)
            mRed = red
            mGreen = green
            mBlue = blue
            mAlpha = 255
        End Sub

        Public Sub New(alpha As Byte, red As Byte, green As Byte, blue As Byte)
            mRed = red
            mGreen = green
            mBlue = blue
            mAlpha = alpha
        End Sub

        Public Sub New(hlsrgb As HLSRGB)
            mRed = hlsrgb.Red
            mBlue = hlsrgb.Blue
            mGreen = hlsrgb.Green
            mLuminance = hlsrgb.Luminance
            mHue = hlsrgb.Hue
            mSaturation = hlsrgb.Saturation
        End Sub

        Public Sub New()
        End Sub

        Public Property Red() As Byte
            Get
                Return mRed
            End Get
            Set(value As Byte)
                mRed = value
                ToHLS()
            End Set
        End Property

        Public Property Green() As Byte
            Get
                Return mGreen
            End Get
            Set(value As Byte)
                mGreen = value
                ToHLS()
            End Set
        End Property

        Public Property Blue() As Byte
            Get
                Return mBlue
            End Get
            Set(value As Byte)
                mBlue = value
                ToHLS()
            End Set
        End Property

        Public Property Luminance() As Single
            Get
                Return mLuminance
            End Get
            Set(value As Single)
                mLuminance = ChkLum(value)
                ToRGB()
            End Set
        End Property

        Public Property Hue() As Single
            Get
                Return mHue
            End Get
            Set(value As Single)
                mHue = ChkHue(value)
                ToRGB()
            End Set
        End Property

        Public Property Saturation() As Single
            Get
                Return mSaturation
            End Get
            Set(value As Single)
                mSaturation = ChkSat(value)
                ToRGB()
            End Set
        End Property

        Public Property Alpha() As Byte
            Get
                Return mAlpha
            End Get
            Set(value As Byte)
                mAlpha = value
            End Set
        End Property

        Public Property HLS() As HueLumSat
            Get
                Return New HueLumSat(mHue, mLuminance, mSaturation)
            End Get
            Set(value As HueLumSat)
                mHue = ChkHue(value.Hue)
                mLuminance = ChkLum(value.Lum)
                mSaturation = ChkSat(value.Sat)
                ToRGB()
            End Set
        End Property

        Public Property Color() As Color
            Get
                Return Drawing.Color.FromArgb(mAlpha, mRed, mGreen, mBlue)
            End Get
            Set(value As Color)
                mRed = value.R
                mGreen = value.G
                mBlue = value.B
                mAlpha = value.A
                ToHLS()
            End Set
        End Property

        Public Sub LightenColor(lightenBy As Single)
            mLuminance *= (1.0 + lightenBy)
            If mLuminance > 1.0 Then mLuminance = 1.0
            ToRGB()
        End Sub

        Public Sub DarkenColor(darkenBy As Single)
            mLuminance *= darkenBy
            ToRGB()
        End Sub

        Private Sub ToHLS()
            Dim minVal As Integer = Math.Min(mRed, Math.Min(mGreen, mBlue))
            Dim maxVal As Integer = Math.Max(mRed, Math.Max(mGreen, mBlue))

            Dim mDiff As Single = maxVal - minVal
            Dim mSum As Single = maxVal + minVal

            mLuminance = mSum / 510.0

            If maxVal = minVal Then
                mSaturation = 0.0
                mHue = 0.0
            Else
                Dim rNorm As Single = (maxVal - mRed) / mDiff
                Dim gNorm As Single = (maxVal - mGreen) / mDiff
                Dim bNorm As Single = (maxVal - mBlue) / mDiff

                mSaturation = If(mLuminance <= 0.5, (mDiff / mSum), mDiff / (510.0 - mSum))

                If mRed = maxVal Then mHue = 60.0 * (6.0 + bNorm - gNorm)
                If mGreen = maxVal Then mHue = 60.0 * (2.0 + rNorm - bNorm)
                If mBlue = maxVal Then mHue = 60.0 * (4.0 + gNorm - rNorm)
                If mHue > 360.0 Then mHue = Hue - 360.0
            End If
        End Sub

        Private Sub ToRGB()
            If mSaturation = 0.0 Then
                Red = CByte(mLuminance * 255.0)
                mGreen = mRed
                mBlue = mRed
            Else
                Dim rm1 As Single
                Dim rm2 As Single

                If mLuminance <= 0.5 Then
                    rm2 = mLuminance + mLuminance * mSaturation
                Else
                    rm2 = mLuminance + mSaturation - mLuminance * mSaturation
                End If
                rm1 = 2.0 * mLuminance - rm2
                mRed = ToRGB1(rm1, rm2, mHue + 120.0)
                mGreen = ToRGB1(rm1, rm2, mHue)
                mBlue = ToRGB1(rm1, rm2, mHue - 120.0)
            End If
        End Sub

        Private Function ToRGB1(rm1 As Single, rm2 As Single, rh As Single) As Byte
            If rh > 360.0 Then
                rh -= 360.0
            ElseIf rh < 0.0 Then
                rh += 360.0
            End If

            If (rh < 60.0) Then
                rm1 += (rm2 - rm1) * rh / 60.0
            ElseIf (rh < 180.0) Then
                rm1 = rm2
            ElseIf (rh < 240.0) Then
                rm1 += (rm2 - rm1) * (240.0 - rh) / 60.0
            End If

            Return Math.Min(rm1 * 255, 255)
        End Function

        Private Function ChkHue(value As Single) As Single
            If value < 0.0 Then
                value = Math.Abs((360.0 + value) Mod 360.0)
            ElseIf value > 360.0 Then
                value = value Mod 360.0
            End If

            Return value
        End Function

        Private Function ChkLum(value As Single) As Single
            If (value < 0.0) Or (value > 1.0) Then
                If value < 0.0 Then value = Math.Abs(value)
                If value > 1.0 Then value = 1.0
            End If

            Return value
        End Function

        Private Function ChkSat(value As Single) As Single
            If value < 0.0 Then
                value = Math.Abs(value)
            ElseIf value > 1.0 Then
                value = 1.0
            End If

            Return value
        End Function
    End Class
End Class
