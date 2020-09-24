Attribute VB_Name = "calcWave"
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public O2F
Sub make330Hz() '11.05 samples = 11050/1000 Hz
O2F = 330       'Oscillator_Frequency
samPles = 11050 / O2F

gsx = 4 * Atn(1) * 2 / samPles


        For i = 0 To samPles
        n = i * gsx
        
        '---<
        Osc2Samp = Sin(n) * 127
        '---<
        
        myByte(i) = Osc2Samp + 128

        Next

End Sub

