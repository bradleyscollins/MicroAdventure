Option Explicit

Sub Main
  ' 10  REM BOMB RESET
  ' 20  T = 5
  Dim T
  T = 5
  ' 30  Y0 = INT(RND(1) * 5 + 1)
  Dim Y0
  Y0 = Int((Rnd(1) * 5) + 1)
  ' 40  X0 = INT(RND(1) * 5 + 1)
  Dim X0
  X0 = Int((Rnd(1) * 5) + 1)
  ' 50  PRINT "PROBE CORRECT SWITCH"
  WScript.StdOut.WriteLine "PROBE CORRECT SWITCH"
  ' 60  PRINT "DIAGONALS COUNT AS 2 UNITS"
  WScript.StdOut.WriteLine "DIAGONALS COUNT AS 2 UNITS"
  ' 70  PRINT
  WScript.StdOut.WriteLine

  Dim D
  Do
    ' 80  GOSUB 260
    Sub260()
    ' 90  PRINT "WHICH SWITCH";
    WScript.StdOut.Write "WHICH SWITCH "
    ' 100 INPUT G$
    Dim G_Dollar
    G_Dollar = WScript.StdIn.ReadLine ()
    ' 110 L = ASC(G$) - ASC("A")
    Dim L
    L = Asc(G_Dollar) - Asc("A")
    ' 120 Y = INT(L / 5) + 1
    Dim Y
    Y = Int(L / 5) + 1
    ' 130 X = L - INT(L / 5) * 5 + 1
    Dim X
    X = L - (Int(L / 5) * 5) + 1
    ' 140 D = ABS(Y0 - Y) + ABS(X0 - X)
    D = Abs(Y0 - Y) + Abs(X0 - X)
    ' 150 IF D = 0 THEN 240
    If D = 0 Then Exit Do
    ' 160 PRINT G$;" IS ";D;" UNITS AWAY"
    WScript.StdOut.WriteLine G_Dollar & " IS " & D & " UNITS AWAY"
    ' 170 PRINT
    WScript.StdOut.WriteLine ()
    ' 180 T = T - 1
    T = T - 1
    ' 190 IF T > 0 THEN 80
  Loop While T > 0

  If T = 0 Then
    ' 200 PRINT "TOO BAD..THE BOMB EXPLODED"
    WScript.StdOut.WriteLine "TOO BAD..THE BOMB EXPLODED"
    ' 210 PRINT "THE CORRECT SWITCH WAS ";
    WScript.StdOut.Write "THE CORRECT SWITCH WAS "
    ' 220 PRINT CHR$(Y0 * 5 - 5 + X0 + ASC("A") - 1)
    WScript.StdOut.Write Chr((Y0 * 5) - 5 + X0 + Asc("A") - 1)
    ' 230 END
  Else
    ' 240 PRINT "BOMB TIMER RESET"
    WScript.StdOut.WriteLine "BOMB TIMER RESET"
  End If
  ' 250 END
End Sub

Sub Sub260()
  ' 260 PRINT
  WScript.StdOut.WriteLine ()
  ' 270 PRINT "-------------"
  WScript.StdOut.WriteLine "-------------"
  ' 280 PRINT "! A B C D E !"
  WScript.StdOut.WriteLine "! A B C D E !"
  ' 290 PRINT "! F G H I J !"
  WScript.StdOut.WriteLine "! F G H I J !"
  ' 300 PRINT "! K L M N O !"
  WScript.StdOut.WriteLine "! K L M N O !"
  ' 310 PRINT "! P Q R S T !"
  WScript.StdOut.WriteLine "! P Q R S T !"
  ' 320 PRINT "! U V W X Y !"
  WScript.StdOut.WriteLine "! U V W X Y !"
  ' 330 PRINT "-------------"
  WScript.StdOut.WriteLine "-------------"
  ' 340 PRINT
  WScript.StdOut.WriteLine ()
  ' 350 RETURN
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
