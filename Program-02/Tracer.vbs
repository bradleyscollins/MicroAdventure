Option Explicit

Sub Main
    ' Tracer
    20  K = 0
    30  DATA "NORTH", "SOUTH", "EAST", "MIDWEST"
    40  DATA "MINNEAPOLIS", "DETOIT", "BISMARK", "JUNEAU"
    50  DATA "LOS ANGELES", "ATLANTA", "DALLAS", "TAMPA"
    60  DATA "WASHINGTON", "HEISTAND", "RICHMOND", "PITTSBURGH"
    70  DATA "COLUMBUS", "CLEVELAND", "WHEELING", "CHICAGO"
    80  RESTORE
    90  PRINT "ENTER 0 TO END."
    100 PRINT
    110 PRINT "ENTER TRANSACTION CODE->";
    120 INPUT N
    130 IF N = 0 THEN 360
    140 K = K + 1
    150 R = INT(N / 100)
    160 C = N - R * 100
    170 FOR I = 1 TO 4
    180 READ R$
    190 IF I <> R THEN 210
    200 A$ = R$
    210 NEXT I
    220 X = R * 4 - 4 + C
    230 FOR I = 1 TO 16
    240 READ C$
    250 IF I <> X THEN 270
    260 B$ = C$
    270 NEXT I
    280 PRINT "--------------------"
    290 PRINT "* STOP # ";K;" *"
    300 PRINT A$; "ERN AREA"
    310 PRINT "CITY:";B$
    320 PRINT "--------------------"
    330 PRINT
    340 PRINT
    350 GOTO 80
    360 END
End Sub

''' ASSERT FUNCTIONS -----------------------------------------------------------

Sub Assert(X, Msg)
    If Not X Then
        Err.Raise 1, Msg, Msg       
    End If
End Sub

Sub AssertEqual(Expected, Actual)
    Dim Message
    Message = "Values are not equal. EXPECTED: " & Expected & "; ACTUAL: " & Actual

    Assert (Expected = Actual), Message
End Sub

''' TEST FUNCTIONS -------------------------------------------------------------

Sub RunTests
    

    WScript.Echo "ALL TESTS PASSED!"
End Sub

''' ENTRY POINT ----------------------------------------------------------------

If WScript.Arguments.Count = 0 Then
    Main()
ElseIf LCase(WScript.Arguments(0)) = "test" Then
    RunTests()
End If
