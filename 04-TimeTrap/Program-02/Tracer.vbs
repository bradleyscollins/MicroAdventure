Option Explicit

Sub Main
    ' Tracer
    ' 20  K = 0
    Dim K
    K = 0
    ' 30  DATA "NORTH", "SOUTH", "EAST", "MIDWEST"
    ' 40  DATA "MINNEAPOLIS", "DETOIT", "BISMARK", "JUNEAU"
    ' 50  DATA "LOS ANGELES", "ATLANTA", "DALLAS", "TAMPA"
    ' 60  DATA "WASHINGTON", "HEISTAND", "RICHMOND", "PITTSBURGH"
    ' 70  DATA "COLUMBUS", "CLEVELAND", "WHEELING", "CHICAGO"
    Dim Data
    Data = Array("NORTH", "SOUTH", "EAST", "MIDWEST", _
                 "MINNEAPOLIS", "DETOIT", "BISMARK", "JUNEAU", _
                 "LOS ANGELES", "ATLANTA", "DALLAS", "TAMPA", _
                 "WASHINGTON", "HEISTAND", "RICHMOND", "PITTSBURGH", _
                 "COLUMBUS", "CLEVELAND", "WHEELING", "CHICAGO")
    ' 80  RESTORE
    Do While True ' End of loop at 350
      ' 90  PRINT "ENTER 0 TO END."
      ' 100 PRINT
      WScript.StdOut.WriteLine "ENTER 0 TO END."
      ' 110 PRINT "ENTER TRANSACTION CODE->";
      WScript.StdOut.Write "ENTER TRANSACTION CODE->"
      ' 120 INPUT N
      Dim N
      N = WScript.StdIn.ReadLine()
      N = CInt(N)
      ' 130 IF N = 0 THEN 360
      If N = 0 Then
        Exit Do ' To line 360
      End If
      ' 140 K = K + 1
      K = K + 1
      ' 150 R = INT(N / 100)
      Dim R
      R = INT(N / 100)
      ' 160 C = N - R * 100
      Dim C
      C = N - R * 100
      ' 170 FOR I = 1 TO 4
      Dim I
      For I = 1 To 4
        ' 180 READ R$
        Dim R_Dollar
        R_Dollar = Data(I - 1)
        ' 190 IF I <> R THEN 210
        If I = R Then
          ' 200 A$ = R$
          Dim A_Dollar
          A_Dollar = R_Dollar
        End If
      ' 210 NEXT I
      Next
      ' 220 X = R * 4 - 4 + C
      Dim X
      X = R * 4 - 4 + C
      ' 230 FOR I = 1 TO 16
      For I = 1 To 16
        ' 240 READ C$
        Dim C_Dollar
        C_Dollar = Data(I + 4 - 1)
        ' 250 IF I <> X THEN 270
        If I = X Then
          ' 260 B$ = C$
          Dim B_Dollar
          B_Dollar = C_Dollar
        End If
      ' 270 NEXT I
      Next
      ' 280 PRINT "--------------------"
      WScript.StdOut.WriteLine "--------------------"
      ' 290 PRINT "* STOP # ";K;" *"
      WScript.StdOut.WriteLine("* STOP # " & K & " *")
      ' 300 PRINT A$; "ERN AREA"
      WScript.StdOut.WriteLine(A_Dollar & "ERN AREA")
      ' 310 PRINT "CITY:";B$
      WScript.StdOut.WriteLine("CITY: " & B_Dollar)
      ' 320 PRINT "--------------------"
      WScript.StdOut.WriteLine "--------------------"
      ' 330 PRINT
      ' 340 PRINT
      WScript.StdOut.WriteLine()
      ' 350 GOTO 80
    ' 360 END
    Loop ' Back to 80
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
