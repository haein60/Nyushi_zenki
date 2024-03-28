Attribute VB_Name = "mdlMath"
Option Explicit

'Ø‚èã‚°ˆ—

Public Function RoundUp(dNum As Double, Optional iKeta As Integer = 0) As Double

Dim dWk As Double

    dWk = (dNum * 10 ^ iKeta)

    dWk = Int((Abs(dWk) + 0.5))

    RoundUp = dWk / 10 ^ iKeta

End Function
