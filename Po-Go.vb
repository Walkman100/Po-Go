Option Explicit On
Option Strict On
Option Compare Binary
Option Infer On

Public Partial Class Po_Go
    Public Sub New()
        Me.InitializeComponent()
        
        btnRestart_Click()
    End Sub
    
    Enum Player
        Red
        Black
        None
    End Enum
    
    Dim currentTurn As Player = Player.Red
    
    ' ==================== UI Handling subs ====================
    
    Sub gridButtons_Click(sender As Object, e As EventArgs) Handles _
      btnA1.Click, btnA2.Click, btnA3.Click, btnA4.Click, btnA5.Click, btnA6.Click, btnA7.Click, btnA8.Click, _
      btnB1.Click, btnB2.Click, btnB3.Click, btnB4.Click, btnB5.Click, btnB6.Click, btnB7.Click, btnB8.Click, _
      btnC1.Click, btnC2.Click, btnC3.Click, btnC4.Click, btnC5.Click, btnC6.Click, btnC7.Click, btnC8.Click, _
      btnD1.Click, btnD2.Click, btnD3.Click, btnD4.Click, btnD5.Click, btnD6.Click, btnD7.Click, btnD8.Click, _
      btnE1.Click, btnE2.Click, btnE3.Click, btnE4.Click, btnE5.Click, btnE6.Click, btnE7.Click, btnE8.Click, _
      btnF1.Click, btnF2.Click, btnF3.Click, btnF4.Click, btnF5.Click, btnF6.Click, btnF7.Click, btnF8.Click, _
      btnG1.Click, btnG2.Click, btnG3.Click, btnG4.Click, btnG5.Click, btnG6.Click, btnG7.Click, btnG8.Click, _
      btnH1.Click, btnH2.Click, btnH3.Click, btnH4.Click, btnH5.Click, btnH6.Click, btnH7.Click, btnH8.Click
        Dim senderName As String = CType(sender, Button).Name
        senderName = senderName.Substring(3)
        
        If DirectCast(DirectCast(sender, Button).Tag, Player) = Player.Red Or DirectCast(DirectCast(sender, Button).Tag, Player) = Player.Black Then
            txtStatus.Text = currentTurn.ToString & " cannot play in " & senderName & ": Existing Piece"
            
        Else
            Dim takenPieces As Boolean = False
            If CheckLeft(senderName) Then
                takenPieces = True
                ProcessLeft(senderName)
            End If
            If CheckRight(senderName) Then
                takenPieces = True
                ProcessRight(senderName)
            End If
            If CheckUp(senderName) Then
                takenPieces = True
                ProcessUp(senderName)
            End If
            If CheckDown(senderName) Then
                takenPieces = True
                ProcessDown(senderName)
            End If
            If CheckUpLeft(senderName) Then
                takenPieces = True
                ProcessUpLeft(senderName)
            End If
            If CheckUpRight(senderName) Then
                takenPieces = True
                ProcessUpRight(senderName)
            End If
            If CheckDownLeft(senderName) Then
                takenPieces = True
                ProcessDownLeft(senderName)
            End If
            If CheckDownRight(senderName) Then
                takenPieces = True
                ProcessDownRight(senderName)
            End If
            
            If takenPieces Then
                txtStatus.Text = currentTurn.ToString & " playing to " & senderName & ". "
                
                If currentTurn = Player.Red Then
                    SetRed(GetButton(senderName))
                    SetTurn(Player.Black)
                ElseIf currentTurn = Player.Black
                    SetBlack(GetButton(senderName))
                    SetTurn(Player.Red)
                End If
                
                txtStatus.Text &= "Old Scores: Red: " & lblRedScore.Text & " Black: " & lblBlackScore.Text & " "
                
                scoreCalculate()
                
                checkWin()
                
                checkCanPlace()
            Else
                txtStatus.Text = "No pieces to take for " & currentTurn.ToString & " at " & senderName & "!"
            End If
        End If
    End Sub
    
    Sub btnRestart_Click() Handles btnRestart.Click
        For i = 65 To 72
            For j = 1 To 8
                SetClear(GetButton(Chr(i) & j))
            Next
        Next
        
        SetBlack(btnD4)
        SetRed(btnE4)
        SetRed(btnD5)
        SetBlack(btnE5)
        
        scoreCalculate()
        
        txtStatus.Text = ""
        
        SetTurn(Player.Red)
    End Sub
    
    Sub gridButtons_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles _
      btnA1.PreviewKeyDown, btnA2.PreviewKeyDown, btnA3.PreviewKeyDown, btnA4.PreviewKeyDown, btnA5.PreviewKeyDown, btnA6.PreviewKeyDown, btnA7.PreviewKeyDown, btnA8.PreviewKeyDown, _
      btnB1.PreviewKeyDown, btnB2.PreviewKeyDown, btnB3.PreviewKeyDown, btnB4.PreviewKeyDown, btnB5.PreviewKeyDown, btnB6.PreviewKeyDown, btnB7.PreviewKeyDown, btnB8.PreviewKeyDown, _
      btnC1.PreviewKeyDown, btnC2.PreviewKeyDown, btnC3.PreviewKeyDown, btnC4.PreviewKeyDown, btnC5.PreviewKeyDown, btnC6.PreviewKeyDown, btnC7.PreviewKeyDown, btnC8.PreviewKeyDown, _
      btnD1.PreviewKeyDown, btnD2.PreviewKeyDown, btnD3.PreviewKeyDown, btnD4.PreviewKeyDown, btnD5.PreviewKeyDown, btnD6.PreviewKeyDown, btnD7.PreviewKeyDown, btnD8.PreviewKeyDown, _
      btnE1.PreviewKeyDown, btnE2.PreviewKeyDown, btnE3.PreviewKeyDown, btnE4.PreviewKeyDown, btnE5.PreviewKeyDown, btnE6.PreviewKeyDown, btnE7.PreviewKeyDown, btnE8.PreviewKeyDown, _
      btnF1.PreviewKeyDown, btnF2.PreviewKeyDown, btnF3.PreviewKeyDown, btnF4.PreviewKeyDown, btnF5.PreviewKeyDown, btnF6.PreviewKeyDown, btnF7.PreviewKeyDown, btnF8.PreviewKeyDown, _
      btnG1.PreviewKeyDown, btnG2.PreviewKeyDown, btnG3.PreviewKeyDown, btnG4.PreviewKeyDown, btnG5.PreviewKeyDown, btnG6.PreviewKeyDown, btnG7.PreviewKeyDown, btnG8.PreviewKeyDown, _
      btnH1.PreviewKeyDown, btnH2.PreviewKeyDown, btnH3.PreviewKeyDown, btnH4.PreviewKeyDown, btnH5.PreviewKeyDown, btnH6.PreviewKeyDown, btnH7.PreviewKeyDown, btnH8.PreviewKeyDown
        If e.KeyCode = Keys.Up Then
            Dim gridCode As String = CType(sender, Button).Name.Substring(3)
            gridCode = gridCode.Chars(0) & ( Integer.Parse(gridCode.Chars(1)) -1 )
            
            ' move row back to the right as the up key moves the row to the left
            If gridCode.Chars(0) <> "H" Then gridCode = Chr( Asc(gridCode.Chars(0)) +1 ) & gridCode.Chars(1)
            
            If gridCode.Chars(1) = "0" Then  ' loop vertically
                gridCode = gridCode.Chars(0) & "8"
            End If
            
            GetButton(gridCode).Focus
        ElseIf e.KeyCode = Keys.Down Then
            Dim gridCode As String = CType(sender, Button).Name.Substring(3)
            gridCode = gridCode.Chars(0) & ( Integer.Parse(gridCode.Chars(1)) +1 )
            
            ' move row back to the left as the down key moves the row to the right
            If gridCode.Chars(0) <> "A" Then gridCode = Chr( Asc(gridCode.Chars(0)) -1 ) & gridCode.Chars(1)
            
            If gridCode.Chars(1) = "9" Then  ' loop vertically
                gridCode = gridCode.Chars(0) & "1"
            End If
            
            GetButton(gridCode).Focus
        End If
    End Sub
    
    ' ==================== Checking if player can place ====================
    
    '       A   B   C   D   E   F   G   H
    '  1    A1  B1  C1  D1  E1  F1  G1  H1
    '  2    A2  B2  C2  D2  E2  F2  G2  H2
    '  3    A3  B3  C3  D3  E3  F3  G3  H3  e.t.c.
    
    Function opponentColour As Player
        If currentTurn = Player.Red Then
            Return Player.Black
        ElseIf currentTurn = Player.Black Then
            Return Player.Red
        Else
            Return Player.None
        End If
    End Function
    
    ' Orthogonal Directions
    
    Function CheckLeft(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim columnLeft As Char = Chr(Asc(column) - 1)
        
        If column = "A" Or column = "B" Then
            Return False
        ElseIf DirectCast(GetButton(columnLeft & row).Tag, Player) = opponentColour
            For i = Asc(columnLeft) To 65 Step -1
                If DirectCast(GetButton(Chr(i) & row).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(Chr(i) & row).Tag, Player) = Player.None Then
                    Return False
                End If
            Next
            
            Return False
        Else
            Return False
        End If
    End Function
    
    Function CheckRight(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim columnRight As Char = Chr(Asc(column) + 1)
        
        If column = "G" Or column = "H" Then
            Return False
        ElseIf DirectCast(GetButton(columnRight & row).Tag, Player) = opponentColour
            For i = Asc(columnRight) To 72 Step 1
                If DirectCast(GetButton(Chr(i) & row).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(Chr(i) & row).Tag, Player) = Player.None Then
                    Return False
                End If
            Next
            
            Return False
        Else
            Return False
        End If
    End Function
    
    Function CheckUp(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowUp As Integer = row - 1
        
        If row = 1 Or row = 2 Then
            Return False
        ElseIf DirectCast(GetButton(column & rowUp).Tag, Player) = opponentColour Then
            For i = rowUp To 1 Step -1
                If DirectCast(GetButton(column & i).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(column & i).Tag, Player) = Player.None Then
                    Return False
                End If
            Next
            
            Return False
        Else
            Return False
        End If
    End Function
    
    Function CheckDown(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowDown As Integer = row + 1
        
        If row = 7 Or row = 8 Then
            Return False
        ElseIf DirectCast(GetButton(column & rowDown).Tag, Player) = opponentColour Then
            For i = rowDown To 8 Step 1
                If DirectCast(GetButton(column & i).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(column & i).Tag, Player) = Player.None Then
                    Return False
                End If
            Next
            
            Return False
        Else
            Return False
        End If
    End Function
    
    ' Diagonally
    
    Function CheckUpLeft(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowUp As Integer = row - 1
        Dim columnLeft As Char = Chr(Asc(column) - 1)
        
        If column = "A" Or column = "B" Or row = 1 Or row = 2 Then
            Return False
        ElseIf DirectCast(GetButton(columnLeft & rowUp).Tag, Player) = opponentColour
            Do While Asc(columnLeft) > 64 AndAlso rowUp > 0
                If DirectCast(GetButton(columnLeft & rowUp).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(columnLeft & rowUp).Tag, Player) = Player.None Then
                    Return False
                End If
                
                rowUp -= 1
                columnLeft = Chr(Asc(columnLeft) - 1)
            Loop
            
            Return False
        Else
            Return False
        End If
    End Function
    
    Function CheckUpRight(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowUp As Integer = row - 1
        Dim columnRight As Char = Chr(Asc(column) + 1)
        
        If column = "G" Or column = "H" Or row = 1 Or row = 2 Then
            Return False
        ElseIf DirectCast(GetButton(columnRight & rowUp).Tag, Player) = opponentColour
            Do While Asc(columnRight) < 73 AndAlso rowUp > 0
                If DirectCast(GetButton(columnRight & rowUp).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(columnRight & rowUp).Tag, Player) = Player.None Then
                    Return False
                End If
                
                rowUp -= 1
                columnRight = Chr(Asc(columnRight) + 1)
            Loop
            
            Return False
        Else
            Return False
        End If
    End Function
    
    Function CheckDownLeft(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowDown As Integer = row + 1
        Dim columnLeft As Char = Chr(Asc(column) - 1)
        
        If column = "A" Or column = "B" Or row = 7 Or row = 8 Then
            Return False
        ElseIf DirectCast(GetButton(columnLeft & rowDown).Tag, Player) = opponentColour
            Do While Asc(columnLeft) > 64 AndAlso rowDown < 9
                If DirectCast(GetButton(columnLeft & rowDown).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(columnLeft & rowDown).Tag, Player) = Player.None Then
                    Return False
                End If
                
                rowDown += 1
                columnLeft = Chr(Asc(columnLeft) - 1)
            Loop
            
            Return False
        Else
            Return False
        End If
    End Function
    
    Function CheckDownRight(gridCode As String) As Boolean
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowDown As Integer = row + 1
        Dim columnRight As Char = Chr(Asc(column) + 1)
        
        If column = "G" Or column = "H" Or row = 7 Or row = 8 Then
            Return False
        ElseIf DirectCast(GetButton(columnRight & rowDown).Tag, Player) = opponentColour
            Do While Asc(columnRight) < 73 AndAlso rowDown < 9
                If DirectCast(GetButton(columnRight & rowDown).Tag, Player) = currentTurn Then
                    Return True
                ElseIf DirectCast(GetButton(columnRight & rowDown).Tag, Player) = Player.None Then
                    Return False
                End If
                
                rowDown += 1
                columnRight = Chr(Asc(columnRight) + 1)
            Loop
            
            Return False
        Else
            Return False
        End If
    End Function
    
    ' ==================== Doing the placement ====================
    ' Orthogonal Directions
    
    Sub ProcessLeft(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim columnLeft As Char = Chr(Asc(column) - 1)
        
        For i = Asc(columnLeft) To 65 Step -1
            If DirectCast(GetButton(Chr(i) & row).Tag, Player) = currentTurn Then
                Exit For
            End If
            SetColour(GetButton(Chr(i) & row), currentTurn)
        Next
    End Sub
    
    Sub ProcessRight(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim columnRight As Char = Chr(Asc(column) + 1)
        
        For i = Asc(columnRight) To 72 Step 1
            If DirectCast(GetButton(Chr(i) & row).Tag, Player) = currentTurn Then
                Exit For
            End If
            SetColour(GetButton(Chr(i) & row), currentTurn)
        Next
    End Sub
    
    Sub ProcessUp(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowUp As Integer = row - 1
        
        For i = rowUp To 1 Step -1
            If DirectCast(GetButton(column & i).Tag, Player) = currentTurn Then
                Exit For
            End If
            SetColour(GetButton(column & i), currentTurn)
        Next
    End Sub
    
    Sub ProcessDown(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowDown As Integer = row + 1
        
        For i = rowDown To 8 Step 1
            If DirectCast(GetButton(column & i).Tag, Player) = currentTurn Then
                Exit For
            End If
            SetColour(GetButton(column & i), currentTurn)
        Next
    End Sub
    
    ' Diagonally
    
    Sub ProcessUpLeft(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowUp As Integer = row - 1
        Dim columnLeft As Char = Chr(Asc(column) - 1)
        
        Do While Asc(columnLeft) > 64 AndAlso rowUp > 0
            If DirectCast(GetButton(columnLeft & rowUp).Tag, Player) = currentTurn Then
                Exit Do
            End If
            SetColour(GetButton(columnLeft & rowUp), currentTurn)
            
            rowUp -= 1
            columnLeft = Chr(Asc(columnLeft) - 1)
        Loop
    End Sub
    
    Sub ProcessUpRight(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowUp As Integer = row - 1
        Dim columnRight As Char = Chr(Asc(column) + 1)
        
        Do While Asc(columnRight) < 73 AndAlso rowUp > 0
            If DirectCast(GetButton(columnRight & rowUp).Tag, Player) = currentTurn Then
                Exit Do
            End If
            SetColour(GetButton(columnRight & rowUp), currentTurn)
            
            rowUp -= 1
            columnRight = Chr(Asc(columnRight) + 1)
        Loop
    End Sub
    
    Sub ProcessDownLeft(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowDown As Integer = row + 1
        Dim columnLeft As Char = Chr(Asc(column) - 1)
        
        Do While Asc(columnLeft) > 64 AndAlso rowDown < 9
            If DirectCast(GetButton(columnLeft & rowDown).Tag, Player) = currentTurn Then
                Exit Do
            End If
            SetColour(GetButton(columnLeft & rowDown), currentTurn)
            
            rowDown += 1
            columnLeft = Chr(Asc(columnLeft) - 1)
        Loop
    End Sub
    
    Sub ProcessDownRight(gridCode As String)
        Dim column As Char = gridCode.Chars(0)
        Dim row As Integer = Integer.Parse(gridCode.Chars(1))
        Dim rowDown As Integer = row + 1
        Dim columnRight As Char = Chr(Asc(column) + 1)
        
        Do While Asc(columnRight) < 73 AndAlso rowDown < 9
            If DirectCast(GetButton(columnRight & rowDown).Tag, Player) = currentTurn Then
                Exit Do
            End If
            SetColour(GetButton(columnRight & rowDown), currentTurn)
            
            rowDown += 1
            columnRight = Chr(Asc(columnRight) + 1)
        Loop
    End Sub
    
    ' ==================== Helper Subs ====================
    
    Sub scoreCalculate()
        Dim redScore As Integer = 0
        Dim blackScore As Integer = 0
        
        For i = 65 To 72
            For j = 1 To 8
                If DirectCast(GetButton(Chr(i) & j).Tag, Player) = Player.Red Then
                    redScore += 1
                ElseIf DirectCast(GetButton(Chr(i) & j).Tag, Player) = Player.Black Then
                    blackScore += 1
                End If
            Next
        Next
        
        lblRedScore.Text = redScore.ToString
        lblBlackScore.Text = blackScore.ToString
    End Sub
    
    Sub checkWin()
        For i = 65 To 72
            For j = 1 To 8
                If DirectCast(GetButton(Chr(i) & j).Tag, Player) = Player.None Then
                    Exit Sub
                End If
            Next
        Next
        
        SetTurn(Player.None)
        
        If Integer.Parse(lblRedScore.Text) > Integer.Parse(lblBlackScore.Text) Then
            txtStatus.Text &= "Red Wins!"
        ElseIf Integer.Parse(lblBlackScore.Text) > Integer.Parse(lblRedScore.Text) Then
            txtStatus.Text &= "Black Wins!"
        ElseIf Integer.Parse(lblRedScore.Text) = Integer.Parse(lblBlackScore.Text) Then
            txtStatus.Text &= "Tie!"
        End If
    End Sub
    
    Sub checkCanPlace()
        For i = 65 To 72
            For j = 1 To 8
                If DirectCast(GetButton(Chr(i) & j).Tag, Player) = Player.None Then
                    Dim gridCodeToCheck As String = GetButton(Chr(i) & j).Name.Substring(3)
                    If CheckLeft(gridCodeToCheck) Then Exit Sub
                    If CheckRight(gridCodeToCheck) Then Exit Sub
                    If CheckUp(gridCodeToCheck) Then Exit Sub
                    If CheckDown(gridCodeToCheck) Then Exit Sub
                    If CheckUpLeft(gridCodeToCheck) Then Exit Sub
                    If CheckUpRight(gridCodeToCheck) Then Exit Sub
                    If CheckDownLeft(gridCodeToCheck) Then Exit Sub
                    If CheckDownRight(gridCodeToCheck) Then Exit Sub
                End If
            Next
        Next
        
        If txtStatus.Text.StartsWith("Game Over.") = False Then
            MsgBox(currentTurn.ToString & " can't play! " & opponentColour.ToString & " goes again...", MsgBoxStyle.Information)
            txtStatus.Text &= currentTurn.ToString & " can't play! "
            SetTurn(opponentColour)
        End If
    End Sub
    
    Sub SetRed(btn As Button)
        SetColour(btn, Player.Red)
    End Sub
    Sub SetBlack(btn As Button)
        SetColour(btn, Player.Black)
    End Sub
    Sub SetClear(btn As Button)
        SetColour(btn, Player.None)
    End Sub
    
    Sub SetColour(btn As Button, playerColour As Player)
        Select Case playerColour
            Case Player.Red
                btn.Image = Resources.dragon_red
                btn.Tag = Player.Red
            Case Player.Black
                btn.Image = Resources.dragon_black
                btn.Tag = Player.Black
            Case Player.None
                btn.Image = Nothing
                btn.Tag = Player.None
        End Select
    End Sub
    
    Sub SetTurn(playerColour As Player)
        currentTurn = playerColour
        
        If playerColour = Player.Red Then
            pbxRedIcon.Image = Resources.dragon_red_selected
            pbxBlackIcon.Image = Resources.dragon_black
            txtStatus.Text &= "Red's Turn! "
            
        ElseIf playerColour = Player.Black Then
            pbxRedIcon.Image = Resources.dragon_red
            pbxBlackIcon.Image = Resources.dragon_black_selected
            txtStatus.Text &= "Black's Turn! "
            
        ElseIf playerColour = Player.None Then
            pbxRedIcon.Image = Resources.dragon_red
            pbxBlackIcon.Image = Resources.dragon_black
            txtStatus.Text = "Game Over. "
        End If
        
    End Sub
    
    ' ==================== GetButton function ====================
    
    Function GetButton(gridCode As String) As Button
        Select Case gridCode
            Case "A1": Return btnA1
            Case "A2": Return btnA2
            Case "A3": Return btnA3
            Case "A4": Return btnA4
            Case "A5": Return btnA5
            Case "A6": Return btnA6
            Case "A7": Return btnA7
            Case "A8": Return btnA8
            Case "B1": Return btnB1
            Case "B2": Return btnB2
            Case "B3": Return btnB3
            Case "B4": Return btnB4
            Case "B5": Return btnB5
            Case "B6": Return btnB6
            Case "B7": Return btnB7
            Case "B8": Return btnB8
            Case "C1": Return btnC1
            Case "C2": Return btnC2
            Case "C3": Return btnC3
            Case "C4": Return btnC4
            Case "C5": Return btnC5
            Case "C6": Return btnC6
            Case "C7": Return btnC7
            Case "C8": Return btnC8
            Case "D1": Return btnD1
            Case "D2": Return btnD2
            Case "D3": Return btnD3
            Case "D4": Return btnD4
            Case "D5": Return btnD5
            Case "D6": Return btnD6
            Case "D7": Return btnD7
            Case "D8": Return btnD8
            Case "E1": Return btnE1
            Case "E2": Return btnE2
            Case "E3": Return btnE3
            Case "E4": Return btnE4
            Case "E5": Return btnE5
            Case "E6": Return btnE6
            Case "E7": Return btnE7
            Case "E8": Return btnE8
            Case "F1": Return btnF1
            Case "F2": Return btnF2
            Case "F3": Return btnF3
            Case "F4": Return btnF4
            Case "F5": Return btnF5
            Case "F6": Return btnF6
            Case "F7": Return btnF7
            Case "F8": Return btnF8
            Case "G1": Return btnG1
            Case "G2": Return btnG2
            Case "G3": Return btnG3
            Case "G4": Return btnG4
            Case "G5": Return btnG5
            Case "G6": Return btnG6
            Case "G7": Return btnG7
            Case "G8": Return btnG8
            Case "H1": Return btnH1 
            Case "H2": Return btnH2
            Case "H3": Return btnH3
            Case "H4": Return btnH4
            Case "H5": Return btnH5
            Case "H6": Return btnH6
            Case "H7": Return btnH7
            Case "H8": Return btnH8
            Case Else: Throw New Exception("Invalid GridCode: " & gridCode)
        End Select
    End Function
End Class
