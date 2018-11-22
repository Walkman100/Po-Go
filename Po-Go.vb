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
        
        If currentTurn = Player.Red Then
            SetRed(GetButton(senderName))
            SetTurn(Player.Black)
        ElseIf currentTurn = Player.Black
            SetBlack(GetButton(senderName))
            SetTurn(Player.Red)
        End If
        
        scoreCalculate()
    End Sub
    
    Sub btnRestart_Click() Handles btnRestart.Click
        For i = 65 To 72
            For j = 1 To 8
                SetClear(GetButton(Chr(i) & j))
            Next
        Next
        
        SetRed(btnD4)
        SetBlack(btnE4)
        SetBlack(btnD5)
        SetRed(btnE5)
        
        scoreCalculate()
        
        SetTurn(Player.Red)
    End Sub
    
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
    
    Sub SetRed(btn As Button)
        btn.Image = Resources.dragon_red
        btn.Tag = Player.Red
    End Sub
    
    Sub SetBlack(btn As Button)
        btn.Image = Resources.dragon_black
        btn.Tag = Player.Black
    End Sub
    
    Sub SetClear(btn As Button)
        btn.Image = Nothing
        btn.Tag = Player.None
    End Sub
    
    Sub SetTurn(playerColor As Player)
        currentTurn = playerColor
        
        If playerColor = Player.Red Then
            pbxRedIcon.Image = Resources.dragon_red_selected
            pbxBlackIcon.Image = Resources.dragon_black
            txtStatus.Text = "Red's Turn!"
            
        ElseIf playerColor = Player.Black Then
            pbxRedIcon.Image = Resources.dragon_red
            pbxBlackIcon.Image = Resources.dragon_black_selected
            txtStatus.Text = "Black's Turn!"
            
        ElseIf playerColor = Player.None Then
            pbxRedIcon.Image = Resources.dragon_red
            pbxBlackIcon.Image = Resources.dragon_black
            txtStatus.Text = "Game Over"
        End If
        
    End Sub
    
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
