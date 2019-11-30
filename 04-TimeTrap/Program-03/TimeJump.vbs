Option Explicit

Sub Main
  ' 10  REM TIMEJUMP
  ' 20  W = 40
  Dim W
  W = 40
  ' 30  PRINT "ENTER YEAR TO GO BACK TO:";
  WScript.StdOut.Write "ENTER YEAR TO GO BACK TO: "
  ' 40  INPUT D
  Dim D
  D = WScript.StdIn.ReadLine()
  D = CInt(D)
  ' 50  PRINT "ENTER PLACE";
  WScript.StdOut.Write "ENTER PLACE: "
  ' 60  INPUT P$
  Dim P_Dollar
  P_Dollar = WScript.StdIn.ReadLine()
  ' 70  PRINT "ENTER THIS YEAR";
  WScript.StdOut.Write "ENTER THIS YEAR: "
  ' 80  INPUT Y
  Dim Y
  Y = WScript.StdIn.ReadLine()
  Y = CInt(Y)
  ' 90  PRINT "ENTER SPEED FACTOR"
  WScript.StdOut.Write "ENTER SPEED FACTOR: "
  ' 100 INPUT S
  Dim S
  S = WScript.StdIn.ReadLine()
  S = CInt(S)
  ' 110 PRINT
  WScript.StdOut.WriteLine()
  ' 120 PRINT "ENTER JUMP-SPAN FACTOR"
  WScript.StdOut.WriteLine "ENTER JUMP-SPAN FACTOR"
  ' 130 PRINT "FROM MACHINE CONTROL:";
  WScript.StdOut.Write "FROM MACHINE CONTROL: "
  ' 140 INPUT M
  Dim M
  M = WScript.StdIn.ReadLine()
  M = CInt(M)
  ' 150 W = W - 6
  W = W - 6
  ' 160 Z = W - LEN(P$) - 2
  Dim Z
  Z = W - Len(P_Dollar) - 2
  ' 170 T = -1
  Dim T
  T = -1
  ' 180 IF Y < D THEN T = 1
  IF Y < D THEN T = 1
  ' 190 FOR I = Y TO D STEP M * T
  Dim I
  For I = Y To D Step M * T
    ' 200 J = ABS(I / M) - INT(ABS(I / M) / Z) * Z
    Dim J
    J = Abs(I / M) - Int(Abs(I / M) / Z) * Z
    ' 210 FOR K = 1 TO J
    Dim K
    For K = 1 To J
      ' 220 PRINT " ";
      WScript.StdOut.Write " "
    ' 230 NEXT K
    Next
    ' 240 PRINT I;" ->";P$
    WScript.StdOut.WriteLine(I & " -> " & P_Dollar)
    ' 250 FOR K = 1 TO S
    For K = 1 To S
    ' 260 NEXT K
    ''' S is speed factor. 1 is fastest; 50 is slowest.
    ''' Meant to simulate speed factor setting by chewing up clock longer for
    ''' higher values?
    Next
  ' 270 NEXT I
  Next
  ' 280 PRINT
  WScript.StdOut.WriteLine()
  ' 290 PRINT "...AT ";P$;" IN ";D
  WScript.StdOut.WriteLine("...AT " & P_Dollar & " IN " & D)
  ' 300 END
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
