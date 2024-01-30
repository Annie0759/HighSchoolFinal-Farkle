Imports System.Data.SqlTypes
Imports System.Net.Mail
Imports System.Runtime.InteropServices.ComTypes

Public Class Form1
    Dim player As String = "1"
    Dim d1, d2, d3, d4, d5, d6 As Integer
    Dim d7, d8, d9, d10, d11, d12 As Integer
    Dim c1, c2, c3, c4, c5, c6 As Boolean
    Dim c7, c8, c9, c10, c11, c12 As Boolean
    Dim shows As Boolean = False
    Dim shows2 As Boolean = False
    Dim fark As Integer = 0
    Dim fark2 As Integer = 0
    'Dim di1 As Boolean = False
    'Dim di2 As Boolean = False
    'Dim di3 As Boolean = False
    'Dim di4 As Boolean = False
    'Dim di5 As Boolean = False
    'Dim di6 As Boolean = False
    Dim pick1 As Boolean = False
    Dim pick2 As Boolean = False
    Dim pick3 As Boolean = False
    Dim pick4 As Boolean = False
    Dim pick5 As Boolean = False
    Dim pick6 As Boolean = False
    Dim pick7 As Boolean = False
    Dim pick8 As Boolean = False
    Dim pick9 As Boolean = False
    Dim pick10 As Boolean = False
    Dim pick11 As Boolean = False
    Dim pick12 As Boolean = False
    Dim special As Boolean = False
    Dim special2 As Boolean = False
    Dim recalc As Boolean = True
    Dim recalc2 As Boolean = True
    Dim k1 As Boolean = False
    Dim k2 As Boolean = False
    Dim k3 As Boolean = False
    Dim k4 As Boolean = False
    Dim k5 As Boolean = False
    Dim k6 As Boolean = False
    Dim k7 As Boolean = False
    Dim k8 As Boolean = False
    Dim k9 As Boolean = False
    Dim k10 As Boolean = False
    Dim k11 As Boolean = False
    Dim k12 As Boolean = False
    Dim score As Integer = 0
    Dim score2 As Integer = 0
    Dim run As Integer = 0
    Dim five As Integer = 0
    Dim five2 As Integer = 0
    Dim one2 As Integer = 0
    Dim one As Integer = 0
    Dim keep As Boolean = False
    Dim keep2 As Boolean = False
    Dim bef As Integer = 0
    Dim bef2 As Integer = 0
    Dim aft2 As Integer = 0
    Dim aft As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MsgBox("Roll the dice alternatively with another player to earn points, whoever gets 10,000 points wins!")
        MsgBox("If your current roll didn't match any of the requirements on the right, you will be Farkled, after 3 Farkles, your score will be reset to 0!")
    End Sub
    Private Sub GoalToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GoalToolStripMenuItem.Click
        MsgBox("Roll the dice alternatively with another player to earn points, whoever gets 10,000 points wins!")
    End Sub
    Private Sub FarkleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FarkleToolStripMenuItem.Click
        MsgBox("If your roll didn't match any of the requirements on the right, you will be Farkled, after 3 Farkles, your score will be reset to 0!")
    End Sub
    Private Sub RollToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RollToolStripMenuItem.Click
        MsgBox("Click Roll to Roll your six die, there is a chance you get Farkled on your roll.")
    End Sub
    Private Sub RerollToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RerollToolStripMenuItem.Click
        MsgBox("If you're not satisfyed with your first roll, you can click reroll, but there is a chance you get Farkled on your reroll.")
    End Sub
    Private Sub ScoreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScoreToolStripMenuItem.Click
        MsgBox("Click score to add you rolled score to your total points and end your round")
    End Sub
    Private Sub KeepToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KeepToolStripMenuItem.Click
        MsgBox("You can keep your ones and fives by clicking this button, you will be able to roll again with the rest of the die, there is a chance that you will get farkled and lose your points for this round.")
    End Sub
    ' First Player
    Private Sub breroll_Click(sender As Object, e As EventArgs) Handles breroll.Click
        d1 = Int(Rnd() * 6) + 1
        d2 = Int(Rnd() * 6) + 1
        d3 = Int(Rnd() * 6) + 1
        d4 = Int(Rnd() * 6) + 1
        d5 = Int(Rnd() * 6) + 1
        d6 = Int(Rnd() * 6) + 1
        showdice1()
        showdice2()
        showdice3()
        showdice4()
        showdice5()
        showdice6()
        nofarkle()

        If d1 <> 1 And d1 <> 5 And d2 <> 1 And d2 <> 5 And d3 <> 1 And d3 <> 5 And d4 <> 1 And d4 <> 5 And d5 <> 1 And d5 <> 5 And d6 <> 1 And d6 <> 5 And special = False Then
            fark = fark + 1
            MsgBox("Player1 got Farkled " & fark & " time(s)!")
            bRoll.Visible = False
            bcroll.Visible = True
            bkeep.Visible = False
            breroll.Visible = False
            bscore.Visible = False
            If fark = 3 Then
                fark = 0
                score = 0
            End If
            ft.Text = fark
            pp.Text = score

        End If
    End Sub

    Private Sub bscore_Click(sender As Object, e As EventArgs) Handles bscore.Click
        determine()
        pp.Text = score

        bkeep.Visible = False
        breroll.Visible = False
        bscore.Visible = False
        bRoll.Visible = False
        If score > 10000 Then
            MsgBox("player1 Wins !")
            score = 0
            score2 = 0
            fark2 = 0
            fark = 0
        End If
        bcroll.Visible = True
        pp.Text = score
        ft.Text = fark
        cp.Text = score2
        ft2.Text = fark2
    End Sub

    Private Sub bkeep_Click(sender As Object, e As EventArgs) Handles bkeep.Click


        bkeep.Visible = False
        breroll.Visible = False
        bscore.Visible = False
        bRoll.Visible = True
        keep = True
        weatherkeep()
        If d1 = 1 Or d1 = 5 Then bef = bef + 1
        If d2 = 1 Or d2 = 5 Then bef = bef + 1
        If d3 = 1 Or d3 = 5 Then bef = bef + 1
        If d4 = 1 Or d4 = 5 Then bef = bef + 1
        If d5 = 1 Or d5 = 5 Then bef = bef + 1
        If d6 = 1 Or d6 = 5 Then bef = bef + 1
    End Sub


    Private Sub broll_Click(sender As Object, e As EventArgs) Handles bRoll.Click
        Randomize()
        If keep = False Then
            d1 = Int(Rnd() * 6) + 1
            d2 = Int(Rnd() * 6) + 1
            d3 = Int(Rnd() * 6) + 1
            d4 = Int(Rnd() * 6) + 1
            d5 = Int(Rnd() * 6) + 1
            d6 = Int(Rnd() * 6) + 1
            nofarkle()
        End If
        c1 = False
        c2 = False
        c3 = False
        c4 = False
        c5 = False
        c6 = False

        If keep = True Then
            If k1 = False And d1 <> 1 And d1 <> 5 Then
                d1 = Int(Rnd() * 6) + 1
                c1 = True
            End If
            If k2 = False And d2 <> 1 And d2 <> 5 Then
                d2 = Int(Rnd() * 6) + 1
                c2 = True
            End If
            If k3 = False And d3 <> 1 And d3 <> 5 Then
                d3 = Int(Rnd() * 6) + 1
                c3 = True
            End If
            If k4 = False And d4 <> 1 And d4 <> 5 Then
                d4 = Int(Rnd() * 6) + 1
                c4 = True
            End If
            If k5 = False And d5 <> 1 And d5 <> 5 Then
                d5 = Int(Rnd() * 6) + 1
                c5 = True
            End If
            If k6 = False And d6 <> 1 And d6 <> 5 Then
                d6 = Int(Rnd() * 6) + 1
                c6 = True
            End If
            c123456()
        End If
        keep = False
        showdice1()
        showdice2()
        showdice3()
        showdice4()
        showdice5()
        showdice6()
        calculate()

        If d1 = 1 Or d1 = 5 Then aft = aft + 1
        If d2 = 1 Or d2 = 5 Then aft = aft + 1
        If d3 = 1 Or d3 = 5 Then aft = aft + 1
        If d4 = 1 Or d4 = 5 Then aft = aft + 1
        If d5 = 1 Or d5 = 5 Then aft = aft + 1
        If d6 = 1 Or d6 = 5 Then aft = aft + 1
        If bef <> aft Or special = True Then
            bef = 0
            aft = 0
        Else
            fark = fark + 1
            MsgBox("Player1 got Farkled " & fark & " time(s)!")
            ft.Text = fark
            bkeep.Visible = False
            breroll.Visible = False
            bscore.Visible = False
            bRoll.Visible = False
            bcroll.Visible = True
            shows = True
            If fark = 3 Then
                fark = 0
                score = 0
                pp.Text = score
            End If
        End If



    End Sub
    Public Sub showdice1()
        If d1 = 1 Then dice1.Image = Image.FromFile("one.jpg")
        If d1 = 2 Then dice1.Image = Image.FromFile("two.jpg")
        If d1 = 3 Then dice1.Image = Image.FromFile("three.jpg")
        If d1 = 4 Then dice1.Image = Image.FromFile("four.jpg")
        If d1 = 5 Then dice1.Image = Image.FromFile("five.jpg")
        If d1 = 6 Then dice1.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice2()
        If d2 = 1 Then dice2.Image = Image.FromFile("one.jpg")
        If d2 = 2 Then dice2.Image = Image.FromFile("two.jpg")
        If d2 = 3 Then dice2.Image = Image.FromFile("three.jpg")
        If d2 = 4 Then dice2.Image = Image.FromFile("four.jpg")
        If d2 = 5 Then dice2.Image = Image.FromFile("five.jpg")
        If d2 = 6 Then dice2.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice3()
        If d3 = 1 Then dice3.Image = Image.FromFile("one.jpg")
        If d3 = 2 Then dice3.Image = Image.FromFile("two.jpg")
        If d3 = 3 Then dice3.Image = Image.FromFile("three.jpg")
        If d3 = 4 Then dice3.Image = Image.FromFile("four.jpg")
        If d3 = 5 Then dice3.Image = Image.FromFile("five.jpg")
        If d3 = 6 Then dice3.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice4()
        If d4 = 1 Then dice4.Image = Image.FromFile("one.jpg")
        If d4 = 2 Then dice4.Image = Image.FromFile("two.jpg")
        If d4 = 3 Then dice4.Image = Image.FromFile("three.jpg")
        If d4 = 4 Then dice4.Image = Image.FromFile("four.jpg")
        If d4 = 5 Then dice4.Image = Image.FromFile("five.jpg")
        If d4 = 6 Then dice4.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice5()
        If d5 = 1 Then dice5.Image = Image.FromFile("one.jpg")
        If d5 = 2 Then dice5.Image = Image.FromFile("two.jpg")
        If d5 = 3 Then dice5.Image = Image.FromFile("three.jpg")
        If d5 = 4 Then dice5.Image = Image.FromFile("four.jpg")
        If d5 = 5 Then dice5.Image = Image.FromFile("five.jpg")
        If d5 = 6 Then dice5.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice6()
        If d6 = 1 Then dice6.Image = Image.FromFile("one.jpg")
        If d6 = 2 Then dice6.Image = Image.FromFile("two.jpg")
        If d6 = 3 Then dice6.Image = Image.FromFile("three.jpg")
        If d6 = 4 Then dice6.Image = Image.FromFile("four.jpg")
        If d6 = 5 Then dice6.Image = Image.FromFile("five.jpg")
        If d6 = 6 Then dice6.Image = Image.FromFile("six.jpg")
    End Sub

    Public Sub calculate()
        If d1 = 5 Or d2 = 5 Or d3 = 5 Or d4 = 5 Or d5 = 5 Or d6 = 5 Or d1 = 1 Or d2 = 1 Or d3 = 1 Or d4 = 1 Or d5 = 1 Or d6 = 1 Or special = True Then
            bkeep.Visible = True
            If bef = 0 Or shows = True Then
                breroll.Visible = True
            End If
            bRoll.Visible = False
            bscore.Visible = True
            shows = False
        End If
    End Sub

    Public Sub determine()
        five = 0
        one = 0
        If d1 = d2 And d3 = d4 And d5 = d6 Or d1 = d2 And d3 = d5 And d4 = d6 Or d1 = d2 And d3 = d6 And d5 = d4 Or d1 = d3 And d2 = d4 And d5 = d6 Or d1 = d3 And d2 = d5 And d4 = d6 Or d1 = d3 And d2 = d6 And d5 = d4 Or d1 = d4 And d3 = d2 And d5 = d6 Or d1 = d4 And d2 = d5 And d3 = d6 Or d1 = d4 And d2 = d6 And d5 = d3 Or d1 = d5 And d2 = d3 And d4 = d6 Or d1 = d5 And d2 = d4 And d3 = d6 Or d1 = d5 And d2 = d6 And d3 = d4 Or d1 = d6 And d3 = d2 And d5 = d4 Or d1 = d6 And d2 = d4 And d5 = d3 Or d1 = d6 And d2 = d5 And d3 = d4 Then
            score = score + 1500
            recalc = False
        End If

        If d1 <> d2 And d1 <> d3 And d1 <> d4 And d1 <> d5 And d1 <> d6 And d2 <> d3 And d2 <> d4 And d2 <> d5 And d2 <> d6 And d3 <> d4 And d3 <> d5 And d3 <> d6 And d4 <> d5 And d4 <> d6 And d5 <> d6 Then
            score = score + 3000
            recalc = False
        End If

        If d1 + d2 + d3 = 18 Or d1 + d2 + d4 = 18 Or d1 + d2 + d5 = 18 Or d1 + d2 + d6 = 18 Or d1 + d3 + d4 = 18 Or d1 + d5 + d3 = 18 Or d1 + d6 + d3 = 18 Or d1 + d4 + d5 = 18 Or d1 + d4 + d6 = 18 Or d2 + d4 + d3 = 18 Or d5 + d2 + d3 = 18 Or d6 + d2 + d3 = 18 Or d5 + d2 + d4 = 18 Or d6 + d2 + d4 = 18 Or d1 + d5 + d6 = 18 Or d5 + d2 + d6 = 18 Or d5 + d4 + d3 = 18 Or d6 + d4 + d3 = 18 Or d5 + d6 + d3 = 18 Or d4 + d5 + d6 = 18 Then
            score = score + 600
            If d1 = 6 Then pick1 = True
            If d2 = 6 Then pick2 = True
            If d3 = 6 Then pick3 = True
            If d4 = 6 Then pick4 = True
            If d5 = 6 Then pick5 = True
            If d6 = 6 Then pick6 = True
        End If

        If d1 = 5 And d2 = 5 And d3 = 5 Or d1 = 5 And d2 = 5 And d4 = 5 Or d1 = 5 And d2 = 5 And d5 = 5 Or d1 = 5 And d2 = 5 And d6 = 5 Or d1 = 5 And d3 = 5 And d4 = 5 Or d1 = 5 And d5 = 5 And d3 = 5 Or d1 = 5 And d6 = 5 And d3 = 5 Or d1 = 5 And d4 = 5 And d5 = 5 Or d1 = 5 And d4 = 5 And d6 = 5 Or d2 = 5 And d4 = 5 And d3 = 5 Or d5 = 5 And d2 = 5 And d3 = 5 Or d6 = 5 And d2 = 5 And d3 = 5 Or d5 = 5 And d2 = 5 And d4 = 5 Or d6 = 5 And d2 = 5 And d4 = 5 Or d1 = 5 And d5 = 5 And d6 = 5 Or d5 = 5 And d2 = 5 And d6 = 5 Or d5 = 5 And d4 = 5 And d3 = 5 Or d6 = 5 And d4 = 5 And d3 = 5 Or d5 = 5 And d6 = 5 And d3 = 5 Or d4 = 5 And d5 = 5 And d6 = 5 Then
            score = score + 500
            five = 3
        End If

        If d1 = 4 And d2 = 4 And d3 = 4 Or d1 = 4 And d2 = 4 And d4 = 4 Or d1 = 4 And d2 = 4 And d5 = 4 Or d1 = 4 And d2 = 4 And d6 = 4 Or d1 = 4 And d3 = 4 And d4 = 4 Or d1 = 4 And d5 = 4 And d3 = 4 Or d1 = 4 And d6 = 4 And d3 = 4 Or d1 = 4 And d4 = 4 And d5 = 4 Or d1 = 4 And d4 = 4 And d6 = 4 Or d2 = 4 And d4 = 4 And d3 = 4 Or d5 = 4 And d2 = 4 And d3 = 4 Or d6 = 4 And d2 = 4 And d3 = 4 Or d5 = 4 And d2 = 4 And d4 = 4 Or d6 = 4 And d2 = 4 And d4 = 4 Or d1 = 4 And d5 = 4 And d6 = 4 Or d5 = 4 And d2 = 4 And d6 = 4 Or d5 = 4 And d4 = 4 And d3 = 4 Or d6 = 4 And d4 = 4 And d3 = 4 Or d5 = 4 And d6 = 4 And d3 = 4 Or d4 = 4 And d5 = 4 And d6 = 4 Then
            score = score + 400
        End If

        If d1 = 3 And d2 = 3 And d3 = 3 Or d1 = 3 And d2 = 3 And d4 = 3 Or d1 = 3 And d2 = 3 And d5 = 3 Or d1 = 3 And d2 = 3 And d6 = 3 Or d1 = 3 And d3 = 3 And d4 = 3 Or d1 = 3 And d5 = 3 And d3 = 3 Or d1 = 3 And d6 = 3 And d3 = 3 Or d1 = 3 And d4 = 3 And d5 = 3 Or d1 = 3 And d4 = 3 And d6 = 3 Or d2 = 3 And d4 = 3 And d3 = 3 Or d5 = 3 And d2 = 3 And d3 = 3 Or d6 = 3 And d2 = 3 And d3 = 3 Or d5 = 3 And d2 = 3 And d4 = 3 Or d6 = 3 And d2 = 3 And d4 = 3 Or d1 = 3 And d5 = 3 And d6 = 3 Or d5 = 3 And d2 = 3 And d6 = 3 Or d5 = 3 And d4 = 3 And d3 = 3 Or d6 = 3 And d4 = 3 And d3 = 3 Or d5 = 3 And d6 = 3 And d3 = 3 Or d4 = 3 And d5 = 3 And d6 = 3 Then
            score = score + 300
        End If

        If d1 = 2 And d2 = 2 And d3 = 2 Or d1 = 2 And d2 = 2 And d4 = 2 Or d1 = 2 And d2 = 2 And d5 = 2 Or d1 = 2 And d2 = 2 And d6 = 2 Or d1 = 2 And d3 = 2 And d4 = 2 Or d1 = 2 And d5 = 2 And d3 = 2 Or d1 = 2 And d6 = 2 And d3 = 2 Or d1 = 2 And d4 = 2 And d5 = 2 Or d1 = 2 And d4 = 2 And d6 = 2 Or d2 = 2 And d4 = 2 And d3 = 2 Or d5 = 2 And d2 = 2 And d3 = 2 Or d6 = 2 And d2 = 2 And d3 = 2 Or d5 = 2 And d2 = 2 And d4 = 2 Or d6 = 2 And d2 = 2 And d4 = 2 Or d1 = 2 And d5 = 2 And d6 = 2 Or d5 = 2 And d2 = 2 And d6 = 2 Or d5 = 2 And d4 = 2 And d3 = 2 Or d6 = 2 And d4 = 2 And d3 = 2 Or d5 = 2 And d6 = 2 And d3 = 2 Or d4 = 2 And d5 = 2 And d6 = 2 Then
            score = score + 200
        End If

        If d1 = 1 And d2 = 1 And d3 = 1 Or d1 = 1 And d2 = 1 And d4 = 1 Or d1 = 1 And d2 = 1 And d5 = 1 Or d1 = 1 And d2 = 1 And d6 = 1 Or d1 = 1 And d3 = 1 And d4 = 1 Or d1 = 1 And d5 = 1 And d3 = 1 Or d1 = 1 And d6 = 1 And d3 = 1 Or d1 = 1 And d4 = 1 And d5 = 1 Or d1 = 1 And d4 = 1 And d6 = 1 Or d2 = 1 And d4 = 1 And d3 = 1 Or d5 = 1 And d2 = 1 And d3 = 1 Or d6 = 1 And d2 = 1 And d3 = 1 Or d5 = 1 And d2 = 1 And d4 = 1 Or d6 = 1 And d2 = 1 And d4 = 1 Or d1 = 1 And d5 = 1 And d6 = 1 Or d5 = 1 And d2 = 1 And d6 = 1 Or d5 = 1 And d4 = 1 And d3 = 1 Or d6 = 1 And d4 = 1 And d3 = 1 Or d5 = 1 And d6 = 1 And d3 = 1 Or d4 = 1 And d5 = 1 And d6 = 1 Then
            score = score + 1000
            one = 3
        End If

        If d1 = 1 And one = 0 And recalc = True Then
            score = score + 100
        ElseIf d1 = 1 And one > 0 And recalc = True Then
            one = one - 1
        End If
        If d2 = 1 And one = 0 And recalc = True Then
            score = score + 100
        ElseIf d2 = 1 And one > 0 And recalc = True Then
            one = one - 1
        End If
        If d3 = 1 And one = 0 And recalc = True Then

            score = score + 100
        ElseIf d3 = 1 And one > 0 And recalc = True Then
            one = one - 1
        End If
        If d4 = 1 And one = 0 And recalc = True Then

            score = score + 100
        ElseIf d4 = 1 And one > 0 And recalc = True Then
            one = one - 1
        End If
        If d5 = 1 And one = 0 And recalc = True Then

            score = score + 100
        ElseIf d5 = 1 And one > 0 And recalc = True Then
            one = one - 1
        End If
        If d6 = 1 And one = 0 And recalc = True Then

            score = score + 100
        ElseIf d6 = 1 And one > 0 And recalc = True Then
            one = one - 1
        End If
        If d1 = 5 And five = 0 And recalc = True Then

            score = score + 50
        ElseIf d1 = 5 And five > 0 And recalc = True Then
            five = five - 1
        End If
        If d2 = 5 And five = 0 And recalc = True Then

            score = score + 50
        ElseIf d2 = 5 And five > 0 And recalc = True Then
            five = five - 1
        End If
        If d3 = 5 And five = 0 And recalc = True Then

            score = score + 50
        ElseIf d3 = 5 And five > 0 And recalc = True Then
            five = five - 1
        End If
        If d4 = 5 And five = 0 And recalc = True Then

            score = score + 50
        ElseIf d4 = 5 And five > 0 And recalc = True Then
            five = five - 1
        End If
        If d5 = 5 And five = 0 And recalc = True Then
            score = score + 50
        ElseIf d5 = 5 And five > 0 And recalc = True Then
            five = five - 1
        End If
        If d6 = 5 And five = 0 And recalc = True Then
            score = score + 50
        ElseIf d6 = 5 And five > 0 And recalc = True Then
            five = five - 1
        End If
        recalc = True
    End Sub

    Public Sub nofarkle()
        special = False
        If d1 = d2 And d3 = d4 And d5 = d6 Or d1 = d2 And d3 = d5 And d4 = d6 Or d1 = d2 And d3 = d6 And d5 = d4 Or d1 = d3 And d2 = d4 And d5 = d6 Or d1 = d3 And d2 = d5 And d4 = d6 Or d1 = d3 And d2 = d6 And d5 = d4 Or d1 = d4 And d3 = d2 And d5 = d6 Or d1 = d4 And d2 = d5 And d3 = d6 Or d1 = d4 And d2 = d6 And d5 = d3 Or d1 = d5 And d2 = d3 And d4 = d6 Or d1 = d5 And d2 = d4 And d3 = d6 Or d1 = d5 And d2 = d6 And d3 = d4 Or d1 = d6 And d3 = d2 And d5 = d4 Or d1 = d6 And d2 = d4 And d5 = d3 Or d1 = d6 And d2 = d5 And d3 = d4 Then
            special = True
        End If

        If d1 <> d2 And d1 <> d3 And d1 <> d4 And d1 <> d5 And d1 <> d6 And d2 <> d3 And d2 <> d4 And d2 <> d5 And d2 <> d6 And d3 <> d4 And d3 <> d5 And d3 <> d6 And d4 <> d5 And d4 <> d6 And d5 <> d6 Then
            special = True
        End If

        If d1 + d2 + d3 = 18 Or d1 + d2 + d4 = 18 Or d1 + d2 + d5 = 18 Or d1 + d2 + d6 = 18 Or d1 + d3 + d4 = 18 Or d1 + d5 + d3 = 18 Or d1 + d6 + d3 = 18 Or d1 + d4 + d5 = 18 Or d1 + d4 + d6 = 18 Or d2 + d4 + d3 = 18 Or d5 + d2 + d3 = 18 Or d6 + d2 + d3 = 18 Or d5 + d2 + d4 = 18 Or d6 + d2 + d4 = 18 Or d1 + d5 + d6 = 18 Or d5 + d2 + d6 = 18 Or d5 + d4 + d3 = 18 Or d6 + d4 + d3 = 18 Or d5 + d6 + d3 = 18 Or d4 + d5 + d6 = 18 Then
            special = True
        End If

        If d1 = 5 And d2 = 5 And d3 = 5 Or d1 = 5 And d2 = 5 And d4 = 5 Or d1 = 5 And d2 = 5 And d5 = 5 Or d1 = 5 And d2 = 5 And d6 = 5 Or d1 = 5 And d3 = 5 And d4 = 5 Or d1 = 5 And d5 = 5 And d3 = 5 Or d1 = 5 And d6 = 5 And d3 = 5 Or d1 = 5 And d4 = 5 And d5 = 5 Or d1 = 5 And d4 = 5 And d6 = 5 Or d2 = 5 And d4 = 5 And d3 = 5 Or d5 = 5 And d2 = 5 And d3 = 5 Or d6 = 5 And d2 = 5 And d3 = 5 Or d5 = 5 And d2 = 5 And d4 = 5 Or d6 = 5 And d2 = 5 And d4 = 5 Or d1 = 5 And d5 = 5 And d6 = 5 Or d5 = 5 And d2 = 5 And d6 = 5 Or d5 = 5 And d4 = 5 And d3 = 5 Or d6 = 5 And d4 = 5 And d3 = 5 Or d5 = 5 And d6 = 5 And d3 = 5 Or d4 = 5 And d5 = 5 And d6 = 5 Then
            special = True
        End If

        If d1 = 4 And d2 = 4 And d3 = 4 Or d1 = 4 And d2 = 4 And d4 = 4 Or d1 = 4 And d2 = 4 And d5 = 4 Or d1 = 4 And d2 = 4 And d6 = 4 Or d1 = 4 And d3 = 4 And d4 = 4 Or d1 = 4 And d5 = 4 And d3 = 4 Or d1 = 4 And d6 = 4 And d3 = 4 Or d1 = 4 And d4 = 4 And d5 = 4 Or d1 = 4 And d4 = 4 And d6 = 4 Or d2 = 4 And d4 = 4 And d3 = 4 Or d5 = 4 And d2 = 4 And d3 = 4 Or d6 = 4 And d2 = 4 And d3 = 4 Or d5 = 4 And d2 = 4 And d4 = 4 Or d6 = 4 And d2 = 4 And d4 = 4 Or d1 = 4 And d5 = 4 And d6 = 4 Or d5 = 4 And d2 = 4 And d6 = 4 Or d5 = 4 And d4 = 4 And d3 = 4 Or d6 = 4 And d4 = 4 And d3 = 4 Or d5 = 4 And d6 = 4 And d3 = 4 Or d4 = 4 And d5 = 4 And d6 = 4 Then
            special = True
        End If

        If d1 = 3 And d2 = 3 And d3 = 3 Or d1 = 3 And d2 = 3 And d4 = 3 Or d1 = 3 And d2 = 3 And d5 = 3 Or d1 = 3 And d2 = 3 And d6 = 3 Or d1 = 3 And d3 = 3 And d4 = 3 Or d1 = 3 And d5 = 3 And d3 = 3 Or d1 = 3 And d6 = 3 And d3 = 3 Or d1 = 3 And d4 = 3 And d5 = 3 Or d1 = 3 And d4 = 3 And d6 = 3 Or d2 = 3 And d4 = 3 And d3 = 3 Or d5 = 3 And d2 = 3 And d3 = 3 Or d6 = 3 And d2 = 3 And d3 = 3 Or d5 = 3 And d2 = 3 And d4 = 3 Or d6 = 3 And d2 = 3 And d4 = 3 Or d1 = 3 And d5 = 3 And d6 = 3 Or d5 = 3 And d2 = 3 And d6 = 3 Or d5 = 3 And d4 = 3 And d3 = 3 Or d6 = 3 And d4 = 3 And d3 = 3 Or d5 = 3 And d6 = 3 And d3 = 3 Or d4 = 3 And d5 = 3 And d6 = 3 Then
            special = True
        End If

        If d1 = 2 And d2 = 2 And d3 = 2 Or d1 = 2 And d2 = 2 And d4 = 2 Or d1 = 2 And d2 = 2 And d5 = 2 Or d1 = 2 And d2 = 2 And d6 = 2 Or d1 = 2 And d3 = 2 And d4 = 2 Or d1 = 2 And d5 = 2 And d3 = 2 Or d1 = 2 And d6 = 2 And d3 = 2 Or d1 = 2 And d4 = 2 And d5 = 2 Or d1 = 2 And d4 = 2 And d6 = 2 Or d2 = 2 And d4 = 2 And d3 = 2 Or d5 = 2 And d2 = 2 And d3 = 2 Or d6 = 2 And d2 = 2 And d3 = 2 Or d5 = 2 And d2 = 2 And d4 = 2 Or d6 = 2 And d2 = 2 And d4 = 2 Or d1 = 2 And d5 = 2 And d6 = 2 Or d5 = 2 And d2 = 2 And d6 = 2 Or d5 = 2 And d4 = 2 And d3 = 2 Or d6 = 2 And d4 = 2 And d3 = 2 Or d5 = 2 And d6 = 2 And d3 = 2 Or d4 = 2 And d5 = 2 And d6 = 2 Then
            special = True

        End If

        If d1 = 1 And d2 = 1 And d3 = 1 Or d1 = 1 And d2 = 1 And d4 = 1 Or d1 = 1 And d2 = 1 And d5 = 1 Or d1 = 1 And d2 = 1 And d6 = 1 Or d1 = 1 And d3 = 1 And d4 = 1 Or d1 = 1 And d5 = 1 And d3 = 1 Or d1 = 1 And d6 = 1 And d3 = 1 Or d1 = 1 And d4 = 1 And d5 = 1 Or d1 = 1 And d4 = 1 And d6 = 1 Or d2 = 1 And d4 = 1 And d3 = 1 Or d5 = 1 And d2 = 1 And d3 = 1 Or d6 = 1 And d2 = 1 And d3 = 1 Or d5 = 1 And d2 = 1 And d4 = 1 Or d6 = 1 And d2 = 1 And d4 = 1 Or d1 = 1 And d5 = 1 And d6 = 1 Or d5 = 1 And d2 = 1 And d6 = 1 Or d5 = 1 And d4 = 1 And d3 = 1 Or d6 = 1 And d4 = 1 And d3 = 1 Or d5 = 1 And d6 = 1 And d3 = 1 Or d4 = 1 And d5 = 1 And d6 = 1 Then
            special = True

        End If

    End Sub

    Public Sub weatherkeep()
        k1 = k2 = k3 = k4 = k5 = k6 = False
        If d1 = d2 And d3 = d4 And d5 = d6 Or d1 = d2 And d3 = d5 And d4 = d6 Or d1 = d2 And d3 = d6 And d5 = d4 Or d1 = d3 And d2 = d4 And d5 = d6 Or d1 = d3 And d2 = d5 And d4 = d6 Or d1 = d3 And d2 = d6 And d5 = d4 Or d1 = d4 And d3 = d2 And d5 = d6 Or d1 = d4 And d2 = d5 And d3 = d6 Or d1 = d4 And d2 = d6 And d5 = d3 Or d1 = d5 And d2 = d3 And d4 = d6 Or d1 = d5 And d2 = d4 And d3 = d6 Or d1 = d5 And d2 = d6 And d3 = d4 Or d1 = d6 And d3 = d2 And d5 = d4 Or d1 = d6 And d2 = d4 And d5 = d3 Or d1 = d6 And d2 = d5 And d3 = d4 Then
            k1 = True
            k2 = True
            k3 = True
            k4 = True
            k5 = True
            k6 = True
        End If
        If d1 <> d2 And d1 <> d3 And d1 <> d4 And d1 <> d5 And d1 <> d6 And d2 <> d3 And d2 <> d4 And d2 <> d5 And d2 <> d6 And d3 <> d4 And d3 <> d5 And d3 <> d6 And d4 <> d5 And d4 <> d6 And d5 <> d6 Then
            k1 = True
            k2 = True
            k3 = True
            k4 = True
            k5 = True
            k6 = True
        End If
        If d1 + d2 + d3 = 18 Then
            k1 = True
            k2 = True
            k3 = True
        End If
        If d1 + d2 + d4 = 18 Then
            k1 = True
            k2 = True
            k4 = True
        End If
        If d1 + d2 + d5 = 18 Then
            k1 = True
            k2 = True
            k5 = True
        End If
        If d1 + d2 + d6 = 18 Then
            k1 = True
            k2 = True
            k6 = True
        End If
        If d1 + d3 + d4 = 18 Then
            k1 = True
            k4 = True
            k3 = True
        End If
        If d1 + d5 + d3 = 18 Then
            k1 = True
            k5 = True
            k3 = True
        End If
        If d1 + d6 + d3 = 18 Then
            k1 = True
            k6 = True
            k3 = True
        End If
        If d1 + d4 + d5 = 18 Then
            k1 = True
            k5 = True
            k4 = True
        End If
        If d1 + d4 + d6 = 18 Then
            k1 = True
            k4 = True
            k6 = True
        End If
        If d2 + d4 + d3 = 18 Then
            k4 = True
            k2 = True
            k3 = True
        End If
        If d5 + d2 + d3 = 18 Then
            k5 = True
            k2 = True
            k3 = True
        End If
        If d6 + d2 + d3 = 18 Then
            k6 = True
            k2 = True
            k3 = True
        End If
        If d5 + d2 + d4 = 18 Then
            k4 = True
            k2 = True
            k5 = True
        End If
        If d6 + d2 + d4 = 18 Then
            k4 = True
            k2 = True
            k6 = True
        End If
        If d1 + d5 + d6 = 18 Then
            k6 = True
            k5 = True
            k1 = True
        End If
        If d5 + d2 + d6 = 18 Then
            k6 = True
            k2 = True
            k5 = True
        End If
        If d5 + d4 + d3 = 18 Then
            k4 = True
            k5 = True
            k3 = True
        End If
        If d6 + d4 + d3 = 18 Then
            k4 = True
            k6 = True
            k3 = True
        End If
        If d5 + d6 + d3 = 18 Then
            k6 = True
            k5 = True
            k3 = True
        End If
        If d4 + d5 + d6 = 18 Then
            k4 = True
            k5 = True
            k6 = True
        End If

        If d1 = 5 And d2 = 5 And d3 = 5 Then
            k1 = True
            k2 = True
            k3 = True
        End If
        If d1 = 5 And d2 = 5 And d4 = 5 Then
            k1 = True
            k2 = True
            k4 = True
        End If
        If d1 = 5 And d2 = 5 And d5 = 5 Then
            k1 = True
            k2 = True
            k5 = True
        End If
        If d1 = 5 And d2 = 5 And d6 = 5 Then
            k1 = True
            k2 = True
            k6 = True
        End If
        If d1 = 5 And d3 = 5 And d4 = 5 Then
            k1 = True
            k4 = True
            k3 = True
        End If
        If d1 = 5 And d5 = 5 And d3 = 5 Then
            k1 = True
            k5 = True
            k3 = True
        End If
        If d1 = 5 And d6 = 5 And d3 = 5 Then
            k1 = True
            k6 = True
            k3 = True
        End If
        If d1 = d4 = d5 = 5 Then
            k1 = True
            k5 = True
            k4 = True
        End If
        If d1 = d4 = d6 = 5 Then
            k1 = True
            k4 = True
            k6 = True
        End If
        If d2 = d4 = d3 = 5 Then
            k4 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d3 = 5 Then
            k5 = True
            k2 = True
            k3 = True
        End If
        If d6 = d2 = d3 = 5 Then
            k6 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d4 = 5 Then
            k4 = True
            k2 = True
            k5 = True
        End If
        If d6 = d2 = d4 = 5 Then
            k4 = True
            k2 = True
            k6 = True
        End If
        If d1 = d5 = d6 = 5 Then
            k6 = True
            k5 = True
            k1 = True
        End If
        If d5 = d2 = d6 = 5 Then
            k6 = True
            k2 = True
            k5 = True
        End If
        If d5 = d4 = d3 = 5 Then
            k4 = True
            k5 = True
            k3 = True
        End If
        If d6 = d4 = d3 = 5 Then
            k4 = True
            k6 = True
            k3 = True
        End If
        If d5 = d6 = d3 = 5 Then
            k6 = True
            k5 = True
            k3 = True
        End If
        If d4 = d5 = d6 = 5 Then
            k4 = True
            k5 = True
            k6 = True
        End If

        If d1 = 4 And d2 = 4 And d3 = 4 Then
            k1 = True
            k2 = True
            k3 = True
        End If
        If d1 = 4 And d2 = 4 And d4 = 4 Then
            k1 = True
            k2 = True
            k4 = True
        End If
        If d1 = 4 And d2 = 4 And d5 = 4 Then
            k1 = True
            k2 = True
            k5 = True
        End If
        If d1 = 4 And d2 = 4 And d6 = 4 Then
            k1 = True
            k2 = True
            k6 = True
        End If
        If d1 = 4 And d3 = 4 And d4 = 4 Then
            k1 = True
            k4 = True
            k3 = True
        End If
        If d1 = 4 And d5 = 4 And d3 = 4 Then
            k1 = True
            k5 = True
            k3 = True
        End If
        If d1 = 4 And d6 = 4 And d3 = 4 Then
            k1 = True
            k6 = True
            k3 = True
        End If
        If d1 = d4 = d5 = 4 Then
            k1 = True
            k5 = True
            k4 = True
        End If
        If d1 = d4 = d6 = 4 Then
            k1 = True
            k4 = True
            k6 = True
        End If
        If d2 = d4 = d3 = 4 Then
            k4 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d3 = 4 Then
            k5 = True
            k2 = True
            k3 = True
        End If
        If d6 = d2 = d3 = 4 Then
            k6 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d4 = 4 Then
            k4 = True
            k2 = True
            k5 = True
        End If
        If d6 = d2 = d4 = 4 Then
            k4 = True
            k2 = True
            k6 = True
        End If
        If d1 = d5 = d6 = 4 Then
            k6 = True
            k5 = True
            k1 = True
        End If
        If d5 = d2 = d6 = 4 Then
            k6 = True
            k2 = True
            k5 = True
        End If
        If d5 = d4 = d3 = 4 Then
            k4 = True
            k5 = True
            k3 = True
        End If
        If d6 = d4 = d3 = 4 Then
            k4 = True
            k6 = True
            k3 = True
        End If
        If d5 = d6 = d3 = 4 Then
            k6 = True
            k5 = True
            k3 = True
        End If
        If d4 = d5 = d6 = 4 Then
            k4 = True
            k5 = True
            k6 = True
        End If

        If d1 = 3 And d2 = 3 And d3 = 3 Then
            k1 = True
            k2 = True
            k3 = True
        End If
        If d1 = 3 And d2 = 3 And d4 = 3 Then
            k1 = True
            k2 = True
            k4 = True
        End If
        If d1 = 3 And d2 = 3 And d5 = 3 Then
            k1 = True
            k2 = True
            k5 = True
        End If
        If d1 = 3 And d2 = 3 And d6 = 3 Then
            k1 = True
            k2 = True
            k6 = True
        End If
        If d1 = 3 And d3 = 3 And d4 = 3 Then
            k1 = True
            k4 = True
            k3 = True
        End If
        If d1 = 3 And d5 = 3 And d3 = 3 Then
            k1 = True
            k5 = True
            k3 = True
        End If
        If d1 = 3 And d6 = 3 And d3 = 3 Then
            k1 = True
            k6 = True
            k3 = True
        End If
        If d1 = d4 = d5 = 3 Then
            k1 = True
            k5 = True
            k4 = True
        End If
        If d1 = d4 = d6 = 3 Then
            k1 = True
            k4 = True
            k6 = True
        End If
        If d2 = d4 = d3 = 3 Then
            k4 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d3 = 3 Then
            k5 = True
            k2 = True
            k3 = True
        End If
        If d6 = d2 = d3 = 3 Then
            k6 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d4 = 3 Then
            k4 = True
            k2 = True
            k5 = True
        End If
        If d6 = d2 = d4 = 3 Then
            k4 = True
            k2 = True
            k6 = True
        End If
        If d1 = d5 = d6 = 3 Then
            k6 = True
            k5 = True
            k1 = True
        End If
        If d5 = d2 = d6 = 3 Then
            k6 = True
            k2 = True
            k5 = True
        End If
        If d5 = d4 = d3 = 3 Then
            k4 = True
            k5 = True
            k3 = True
        End If
        If d6 = d4 = d3 = 3 Then
            k4 = True
            k6 = True
            k3 = True
        End If
        If d5 = d6 = d3 = 3 Then
            k6 = True
            k5 = True
            k3 = True
        End If
        If d4 = d5 = d6 = 3 Then
            k4 = True
            k5 = True
            k6 = True
        End If

        If d1 = 2 And d2 = 2 And d3 = 2 Then
            k1 = True
            k2 = True
            k3 = True
        End If
        If d1 = 2 And d2 = 2 And d4 = 2 Then
            k1 = True
            k2 = True
            k4 = True
        End If
        If d1 = 2 And d2 = 2 And d5 = 2 Then
            k1 = True
            k2 = True
            k5 = True
        End If
        If d1 = 2 And d2 = 2 And d6 = 2 Then
            k1 = True
            k2 = True
            k6 = True
        End If
        If d1 = 2 And d3 = 2 And d4 = 2 Then
            k1 = True
            k4 = True
            k3 = True
        End If
        If d1 = 2 And d5 = 2 And d3 = 2 Then
            k1 = True
            k5 = True
            k3 = True
        End If
        If d1 = 2 And d6 = 2 And d3 = 2 Then
            k1 = True
            k6 = True
            k3 = True
        End If
        If d1 = d4 = d5 = 2 Then
            k1 = True
            k5 = True
            k4 = True
        End If
        If d1 = d4 = d6 = 2 Then
            k1 = True
            k4 = True
            k6 = True
        End If
        If d2 = d4 = d3 = 2 Then
            k4 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d3 = 2 Then
            k5 = True
            k2 = True
            k3 = True
        End If
        If d6 = d2 = d3 = 2 Then
            k6 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d4 = 2 Then
            k4 = True
            k2 = True
            k5 = True
        End If
        If d6 = d2 = d4 = 2 Then
            k4 = True
            k2 = True
            k6 = True
        End If
        If d1 = d5 = d6 = 2 Then
            k6 = True
            k5 = True
            k1 = True
        End If
        If d5 = d2 = d6 = 2 Then
            k6 = True
            k2 = True
            k5 = True
        End If
        If d5 = d4 = d3 = 2 Then
            k4 = True
            k5 = True
            k3 = True
        End If
        If d6 = d4 = d3 = 2 Then
            k4 = True
            k6 = True
            k3 = True
        End If
        If d5 = d6 = d3 = 2 Then
            k6 = True
            k5 = True
            k3 = True
        End If
        If d4 = d5 = d6 = 2 Then
            k4 = True
            k5 = True
            k6 = True
        End If

        If d1 = 1 And d2 = 1 And d3 = 1 Then
            k1 = True
            k2 = True
            k3 = True
        End If
        If d1 = 1 And d2 = 1 And d4 = 1 Then
            k1 = True
            k2 = True
            k4 = True
        End If
        If d1 = 1 And d2 = 1 And d5 = 1 Then
            k1 = True
            k2 = True
            k5 = True
        End If
        If d1 = 1 And d2 = 1 And d6 = 1 Then
            k1 = True
            k2 = True
            k6 = True
        End If
        If d1 = 1 And d3 = 1 And d4 = 1 Then
            k1 = True
            k4 = True
            k3 = True
        End If
        If d1 = 1 And d5 = 1 And d3 = 1 Then
            k1 = True
            k5 = True
            k3 = True
        End If
        If d1 = 1 And d6 = 1 And d3 = 1 Then
            k1 = True
            k6 = True
            k3 = True
        End If
        If d1 = d4 = d5 = 1 Then
            k1 = True
            k5 = True
            k4 = True
        End If
        If d1 = d4 = d6 = 1 Then
            k1 = True
            k4 = True
            k6 = True
        End If
        If d2 = d4 = d3 = 1 Then
            k4 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d3 = 1 Then
            k5 = True
            k2 = True
            k3 = True
        End If
        If d6 = d2 = d3 = 1 Then
            k6 = True
            k2 = True
            k3 = True
        End If
        If d5 = d2 = d4 = 1 Then
            k4 = True
            k2 = True
            k5 = True
        End If
        If d6 = d2 = d4 = 1 Then
            k4 = True
            k2 = True
            k6 = True
        End If
        If d1 = d5 = d6 = 1 Then
            k6 = True
            k5 = True
            k1 = True
        End If
        If d5 = d2 = d6 = 1 Then
            k6 = True
            k2 = True
            k5 = True
        End If
        If d5 = d4 = d3 = 1 Then
            k4 = True
            k5 = True
            k3 = True
        End If
        If d6 = d4 = d3 = 1 Then
            k4 = True
            k6 = True
            k3 = True
        End If
        If d5 = d6 = d3 = 1 Then
            k6 = True
            k5 = True
            k3 = True
        End If
        If d4 = d5 = d6 = 1 Then
            k4 = True
            k5 = True
            k6 = True
        End If
    End Sub

    Public Sub c123456()
        special = False
        If c1 = True And c2 = True And c3 = True Then
            If d1 + d2 + d3 = 18 Then
                special = True
            End If
        End If
        If c1 = True Or c2 = True Or c4 = True Then
            If d1 + d2 + d4 = 18 Then
                special = True
            End If
        End If
        If c1 = True Or c2 = True Or c5 = True Then
            If d1 + d2 + d5 = 18 Then
                special = True
            End If
        End If
        If c1 = True Or c2 = True Or c6 = True Then
            If d1 + d2 + d6 = 18 Then
                special = True
            End If
        End If
        If c1 = True Or c3 = True Or c4 = True Then
            If d1 + d3 + d4 = 18 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c3 = True Then
            If d1 + d5 + d3 = 18 Then
                special = True
            End If
        End If
        If c1 = True And c6 = True And c3 = True Then
            If d1 + d6 + d3 = 18 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c5 = True Then
            If d1 + d4 + d5 = 18 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c6 = True Then
            If d1 + d4 + d6 = 18 Then
                special = True
            End If
        End If
        If c4 = True And c2 = True And c3 = True Then
            If d2 + d4 + d3 = 18 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c3 = True Then
            If d5 + d2 + d3 = 18 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c3 = True Then
            If d6 + d2 + d3 = 18 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c4 = True Then
            If d5 + d2 + d4 = 18 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c4 = True Then
            If d6 + d2 + d4 = 18 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c6 = True Then
            If d1 + d5 + d6 = 18 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c6 = True Then
            If d5 + d2 + d6 = 18 Then
                special = True
            End If
        End If
        If c4 = True And c5 = True And c3 = True Then
            If d5 + d4 + d3 = 18 Then
                special = True
            End If
        End If
        If c4 = True And c6 = True And c3 = True Then
            If d6 + d4 + d3 = 18 Then
                special = True
            End If
        End If
        If c6 = True And c5 = True And c3 = True Then
            If d5 + d6 + d3 = 18 Then
                special = True
            End If
        End If
        If c4 = True And c5 = True And c6 = True Then
            If d4 + d5 + d6 = 18 Then
                special = True
            End If
        End If

        If c1 = True And c2 = True And c3 = True Then
            If d1 = 5 And d2 = 5 And d3 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c4 = True Then
            If d1 = 5 And d2 = 5 And d4 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c5 = True Then
            If d1 = 5 And d2 = 5 And d5 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c6 = True Then
            If d1 = 5 And d2 = 5 And d6 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c3 = True Then
            If d1 = 5 And d3 = 5 And d4 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c3 = True Then
            If d1 = 5 And d5 = 5 And d3 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c6 = True And c3 = True Then
            If d1 = 5 And d6 = 5 And d3 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c5 = True Then
            If d1 = d4 = d5 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c6 = True Then
            If d1 = d4 = d6 = 5 Then
                special = True
            End If
        End If
        If c4 = True And c2 = True And c3 = True Then
            If d2 = d4 = d3 = 5 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c3 = True Then
            If d5 = d2 = d3 = 5 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c3 = True Then
            If d6 = d2 = d3 = 5 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c4 = True Then
            If d5 = d2 = d4 = 5 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c4 = True Then
            If d6 = d2 = d4 = 5 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c6 = True Then
            If d1 = d5 = d6 = 5 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c6 = True Then
            If d5 = d2 = d6 = 5 Then
                special = True
            End If
        End If
        If c5 = True And c4 = True And c3 = True Then
            If d5 = d4 = d3 = 5 Then
                special = True
            End If
        End If
        If c6 = True And c4 = True And c3 = True Then
            If d6 = d4 = d3 = 5 Then
                special = True
            End If
        End If
        If c5 = True And c6 = True And c3 = True Then
            If d5 = d6 = d3 = 5 Then
                special = True
            End If
        End If
        If c4 = True And c5 = True And c6 = True Then
            If d4 = d5 = d6 = 5 Then
                special = True
            End If
        End If

        If c1 = True And c2 = True And c3 = True Then
            If d1 = 4 And d2 = 4 And d3 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c4 = True Then
            If d1 = 4 And d2 = 4 And d4 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c5 = True Then
            If d1 = 4 And d2 = 4 And d5 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c6 = True Then
            If d1 = 4 And d2 = 4 And d6 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c3 = True Then
            If d1 = 4 And d3 = 4 And d4 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c3 = True Then
            If d1 = 4 And d5 = 4 And d3 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c6 = True And c3 = True Then
            If d1 = 4 And d6 = 4 And d3 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c5 = True Then
            If d1 = d4 = d5 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c6 = True Then
            If d1 = d4 = d6 = 4 Then
                special = True
            End If
        End If
        If c4 = True And c2 = True And c3 = True Then
            If d2 = d4 = d3 = 4 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c3 = True Then
            If d5 = d2 = d3 = 4 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c3 = True Then
            If d6 = d2 = d3 = 4 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c4 = True Then
            If d5 = d2 = d4 = 4 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c4 = True Then
            If d6 = d2 = d4 = 4 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c6 = True Then
            If d1 = d5 = d6 = 4 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c6 = True Then
            If d5 = d2 = d6 = 4 Then
                special = True
            End If
        End If
        If c5 = True And c4 = True And c3 = True Then
            If d5 = d4 = d3 = 4 Then
                special = True
            End If
        End If
        If c6 = True And c4 = True And c3 = True Then
            If d6 = d4 = d3 = 4 Then
                special = True
            End If
        End If
        If c5 = True And c6 = True And c3 = True Then
            If d5 = d6 = d3 = 4 Then
                special = True
            End If
        End If
        If c4 = True And c5 = True And c6 = True Then
            If d4 = d5 = d6 = 4 Then
                special = True
            End If
        End If

        If c1 = True And c2 = True And c3 = True Then
            If d1 = 3 And d2 = 3 And d3 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c4 = True Then
            If d1 = 3 And d2 = 3 And d4 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c5 = True Then
            If d1 = 3 And d2 = 3 And d5 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c6 = True Then
            If d1 = 3 And d2 = 3 And d6 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c3 = True Then
            If d1 = 3 And d3 = 3 And d4 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c3 = True Then
            If d1 = 3 And d5 = 3 And d3 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c6 = True And c3 = True Then
            If d1 = 3 And d6 = 3 And d3 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c5 = True Then
            If d1 = d4 = d5 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c6 = True Then
            If d1 = d4 = d6 = 3 Then
                special = True
            End If
        End If
        If c4 = True And c2 = True And c3 = True Then
            If d2 = d4 = d3 = 3 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c3 = True Then
            If d5 = d2 = d3 = 3 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c3 = True Then
            If d6 = d2 = d3 = 3 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c4 = True Then
            If d5 = d2 = d4 = 3 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c4 = True Then
            If d6 = d2 = d4 = 3 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c6 = True Then
            If d1 = d5 = d6 = 3 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c6 = True Then
            If d5 = d2 = d6 = 3 Then
                special = True
            End If
        End If
        If c5 = True And c4 = True And c3 = True Then
            If d5 = d4 = d3 = 3 Then
                special = True
            End If
        End If
        If c6 = True And c4 = True And c3 = True Then
            If d6 = d4 = d3 = 3 Then
                special = True
            End If
        End If
        If c5 = True And c6 = True And c3 = True Then
            If d5 = d6 = d3 = 3 Then
                special = True
            End If
        End If
        If c4 = True And c5 = True And c6 = True Then
            If d4 = d5 = d6 = 3 Then
                special = True
            End If
        End If

        If c1 = True And c2 = True And c3 = True Then
            If d1 = 2 And d2 = 2 And d3 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c4 = True Then
            If d1 = 2 And d2 = 2 And d4 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c5 = True Then
            If d1 = 2 And d2 = 2 And d5 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c6 = True Then
            If d1 = 2 And d2 = 2 And d6 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c3 = True Then
            If d1 = 2 And d3 = 2 And d4 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c3 = True Then
            If d1 = 2 And d5 = 2 And d3 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c6 = True And c3 = True Then
            If d1 = 2 And d6 = 2 And d3 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c5 = True Then
            If d1 = d4 = d5 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c6 = True Then
            If d1 = d4 = d6 = 2 Then
                special = True
            End If
        End If
        If c4 = True And c2 = True And c3 = True Then
            If d2 = d4 = d3 = 2 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c3 = True Then
            If d5 = d2 = d3 = 2 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c3 = True Then
            If d6 = d2 = d3 = 2 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c4 = True Then
            If d5 = d2 = d4 = 2 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c4 = True Then
            If d6 = d2 = d4 = 2 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c6 = True Then
            If d1 = d5 = d6 = 2 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c6 = True Then
            If d5 = d2 = d6 = 2 Then
                special = True
            End If
        End If
        If c5 = True And c4 = True And c3 = True Then
            If d5 = d4 = d3 = 2 Then
                special = True
            End If
        End If
        If c6 = True And c4 = True And c3 = True Then
            If d6 = d4 = d3 = 2 Then
                special = True
            End If
        End If
        If c5 = True And c6 = True And c3 = True Then
            If d5 = d6 = d3 = 2 Then
                special = True
            End If
        End If
        If c4 = True And c5 = True And c6 = True Then
            If d4 = d5 = d6 = 2 Then
                special = True
            End If
        End If

        If c1 = True And c2 = True And c3 = True Then
            If d1 = 1 And d2 = 1 And d3 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c4 = True Then
            If d1 = 1 And d2 = 1 And d4 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c5 = True Then
            If d1 = 1 And d2 = 1 And d5 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c2 = True And c6 = True Then
            If d1 = 1 And d2 = 1 And d6 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c3 = True Then
            If d1 = 1 And d3 = 1 And d4 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c3 = True Then
            If d1 = 1 And d5 = 1 And d3 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c6 = True And c3 = True Then
            If d1 = 1 And d6 = 1 And d3 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c5 = True Then
            If d1 = d4 = d5 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c4 = True And c6 = True Then
            If d1 = d4 = d6 = 1 Then
                special = True
            End If
        End If
        If c4 = True And c2 = True And c3 = True Then
            If d2 = d4 = d3 = 1 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c3 = True Then
            If d5 = d2 = d3 = 1 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c3 = True Then
            If d6 = d2 = d3 = 1 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c4 = True Then
            If d5 = d2 = d4 = 1 Then
                special = True
            End If
        End If
        If c6 = True And c2 = True And c4 = True Then
            If d6 = d2 = d4 = 1 Then
                special = True
            End If
        End If
        If c1 = True And c5 = True And c6 = True Then
            If d1 = d5 = d6 = 1 Then
                special = True
            End If
        End If
        If c5 = True And c2 = True And c6 = True Then
            If d5 = d2 = d6 = 1 Then
                special = True
            End If
        End If
        If c5 = True And c4 = True And c3 = True Then
            If d5 = d4 = d3 = 1 Then
                special = True
            End If
        End If
        If c6 = True And c4 = True And c3 = True Then
            If d6 = d4 = d3 = 1 Then
                special = True
            End If
        End If
        If c5 = True And c6 = True And c3 = True Then
            If d5 = d6 = d3 = 1 Then
                special = True
            End If
        End If
        If c4 = True And c5 = True And c6 = True Then
            If d4 = d5 = d6 = 1 Then
                special = True
            End If
        End If
    End Sub

    ' Second player

    Public Sub showdice7()
        If d7 = 1 Then dice7.Image = Image.FromFile("one.jpg")
        If d7 = 2 Then dice7.Image = Image.FromFile("two.jpg")
        If d7 = 3 Then dice7.Image = Image.FromFile("three.jpg")
        If d7 = 4 Then dice7.Image = Image.FromFile("four.jpg")
        If d7 = 5 Then dice7.Image = Image.FromFile("five.jpg")
        If d7 = 6 Then dice7.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice8()
        If d8 = 1 Then dice8.Image = Image.FromFile("one.jpg")
        If d8 = 2 Then dice8.Image = Image.FromFile("two.jpg")
        If d8 = 3 Then dice8.Image = Image.FromFile("three.jpg")
        If d8 = 4 Then dice8.Image = Image.FromFile("four.jpg")
        If d8 = 5 Then dice8.Image = Image.FromFile("five.jpg")
        If d8 = 6 Then dice8.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice9()
        If d9 = 1 Then dice9.Image = Image.FromFile("one.jpg")
        If d9 = 2 Then dice9.Image = Image.FromFile("two.jpg")
        If d9 = 3 Then dice9.Image = Image.FromFile("three.jpg")
        If d9 = 4 Then dice9.Image = Image.FromFile("four.jpg")
        If d9 = 5 Then dice9.Image = Image.FromFile("five.jpg")
        If d9 = 6 Then dice9.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice10()
        If d10 = 1 Then dice10.Image = Image.FromFile("one.jpg")
        If d10 = 2 Then dice10.Image = Image.FromFile("two.jpg")
        If d10 = 3 Then dice10.Image = Image.FromFile("three.jpg")
        If d10 = 4 Then dice10.Image = Image.FromFile("four.jpg")
        If d10 = 5 Then dice10.Image = Image.FromFile("five.jpg")
        If d10 = 6 Then dice10.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice11()
        If d11 = 1 Then dice11.Image = Image.FromFile("one.jpg")
        If d11 = 2 Then dice11.Image = Image.FromFile("two.jpg")
        If d11 = 3 Then dice11.Image = Image.FromFile("three.jpg")
        If d11 = 4 Then dice11.Image = Image.FromFile("four.jpg")
        If d11 = 5 Then dice11.Image = Image.FromFile("five.jpg")
        If d11 = 6 Then dice11.Image = Image.FromFile("six.jpg")
    End Sub
    Public Sub showdice12()
        If d12 = 1 Then dice12.Image = Image.FromFile("one.jpg")
        If d12 = 2 Then dice12.Image = Image.FromFile("two.jpg")
        If d12 = 3 Then dice12.Image = Image.FromFile("three.jpg")
        If d12 = 4 Then dice12.Image = Image.FromFile("four.jpg")
        If d12 = 5 Then dice12.Image = Image.FromFile("five.jpg")
        If d12 = 6 Then dice12.Image = Image.FromFile("six.jpg")
    End Sub

    Private Sub bcroll_Click(sender As Object, e As EventArgs) Handles bcroll.Click
        Randomize()
        If keep2 = False Then
            d12 = Int(Rnd() * 6) + 1
            d7 = Int(Rnd() * 6) + 1
            d8 = Int(Rnd() * 6) + 1
            d9 = Int(Rnd() * 6) + 1
            d10 = Int(Rnd() * 6) + 1
            d11 = Int(Rnd() * 6) + 1
            nofarkle2()
        End If
        c11 = False
        c12 = False
        c7 = False
        c8 = False
        c9 = False
        c10 = False

        If keep2 = True Then
            If k7 = False And d7 <> 1 And d7 <> 5 Then
                d7 = Int(Rnd() * 6) + 1
                c7 = True
            End If
            If k8 = False And d8 <> 1 And d8 <> 5 Then
                d8 = Int(Rnd() * 6) + 1
                c8 = True
            End If
            If k9 = False And d9 <> 1 And d9 <> 5 Then
                d9 = Int(Rnd() * 6) + 1
                c9 = True
            End If
            If k10 = False And d10 <> 1 And d10 <> 5 Then
                d10 = Int(Rnd() * 6) + 1
                c10 = True
            End If
            If k11 = False And d11 <> 1 And d11 <> 5 Then
                d11 = Int(Rnd() * 6) + 1
                c11 = True
            End If
            If k12 = False And d12 <> 1 And d12 <> 5 Then
                d12 = Int(Rnd() * 6) + 1
                c12 = True
            End If
            c789101112()
        End If
        keep2 = False
        showdice7()
        showdice12()
        showdice8()
        showdice9()
        showdice10()
        showdice11()
        calculate2()

        If d7 = 1 Or d7 = 5 Then aft2 = aft2 + 1
        If d8 = 1 Or d8 = 5 Then aft2 = aft2 + 1
        If d9 = 1 Or d9 = 5 Then aft2 = aft2 + 1
        If d10 = 1 Or d10 = 5 Then aft2 = aft2 + 1
        If d11 = 1 Or d11 = 5 Then aft2 = aft2 + 1
        If d12 = 1 Or d12 = 5 Then aft2 = aft2 + 1
        If bef2 <> aft2 Or special2 = True Then
            bef2 = 0
            aft2 = 0
        Else
            fark2 = fark2 + 1
            MsgBox("Player2 got Farkled " & fark2 & " time(s)!")
            ft2.Text = fark2
            bckeep.Visible = False
            bcreroll.Visible = False
            bcscore.Visible = False
            bcroll.Visible = False
            bRoll.Visible = True
            shows2 = True
            If fark2 = 3 Then
                fark2 = 0
                score2 = 0
                cp.Text = score2
            End If
        End If
    End Sub
    Public Sub calculate2()
        If d7 = 5 Or d8 = 5 Or d9 = 5 Or d10 = 5 Or d11 = 5 Or d12 = 5 Or d7 = 1 Or d8 = 1 Or d9 = 1 Or d12 = 1 Or d11 = 1 Or d10 = 1 Or special2 = True Then
            bckeep.Visible = True
            If bef2 = 0 Or shows2 = True Then
                bcreroll.Visible = True
            End If
            bcroll.Visible = False
            bcscore.Visible = True
            shows2 = False
        End If
    End Sub
    Public Sub c789101112()
        special2 = False
        If c7 = True And c8 = True And c9 = True Then
            If d7 + d8 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True Or c8 = True Or c10 = True Then
            If d7 + d8 + d10 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True Or c8 = True Or c11 = True Then
            If d7 + d8 + d11 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True Or c8 = True Or c12 = True Then
            If d7 + d8 + d12 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True Or c9 = True Or c10 = True Then
            If d7 + d9 + d10 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c9 = True Then
            If d7 + d11 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True And c12 = True And c9 = True Then
            If d7 + d12 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c11 = True Then
            If d7 + d10 + d11 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c12 = True Then
            If d7 + d10 + d12 = 18 Then
                special2 = True
            End If
        End If
        If c10 = True And c8 = True And c9 = True Then
            If d8 + d10 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c9 = True Then
            If d11 + d8 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c9 = True Then
            If d12 + d8 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c10 = True Then
            If d11 + d8 + d10 = 18 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c10 = True Then
            If d12 + d8 + d10 = 18 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c12 = True Then
            If d7 + d11 + d12 = 18 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c12 = True Then
            If d11 + d8 + d12 = 18 Then
                special2 = True
            End If
        End If
        If c10 = True And c11 = True And c9 = True Then
            If d11 + d10 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c10 = True And c12 = True And c9 = True Then
            If d12 + d10 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c12 = True And c11 = True And c9 = True Then
            If d11 + d12 + d9 = 18 Then
                special2 = True
            End If
        End If
        If c10 = True And c11 = True And c12 = True Then
            If d10 + d11 + d12 = 18 Then
                special2 = True
            End If
        End If

        If c7 = True And c8 = True And c9 = True Then
            If d7 = 5 And d8 = 5 And d9 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c10 = True Then
            If d7 = 5 And d8 = 5 And d10 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c11 = True Then
            If d7 = 5 And d8 = 5 And d11 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c12 = True Then
            If d7 = 5 And d8 = 5 And d12 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c9 = True Then
            If d7 = 5 And d9 = 5 And d10 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c9 = True Then
            If d7 = 5 And d11 = 5 And d9 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c12 = True And c9 = True Then
            If d7 = 5 And d12 = 5 And d9 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c11 = True Then
            If d7 = d10 = d11 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c12 = True Then
            If d7 = d10 = d12 = 5 Then
                special2 = True
            End If
        End If
        If c10 = True And c8 = True And c9 = True Then
            If d8 = d10 = d9 = 5 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c9 = True Then
            If d11 = d8 = d9 = 5 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c9 = True Then
            If d12 = d8 = d9 = 5 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c10 = True Then
            If d11 = d8 = d10 = 5 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c10 = True Then
            If d12 = d8 = d10 = 5 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c12 = True Then
            If d7 = d11 = d12 = 5 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c12 = True Then
            If d11 = d8 = d12 = 5 Then
                special2 = True
            End If
        End If
        If c11 = True And c10 = True And c9 = True Then
            If d11 = d10 = d9 = 5 Then
                special2 = True
            End If
        End If
        If c12 = True And c10 = True And c9 = True Then
            If d12 = d10 = d9 = 5 Then
                special2 = True
            End If
        End If
        If c11 = True And c12 = True And c9 = True Then
            If d11 = d12 = d9 = 5 Then
                special2 = True
            End If
        End If
        If c10 = True And c11 = True And c12 = True Then
            If d10 = d11 = d12 = 5 Then
                special2 = True
            End If
        End If

        If c7 = True And c8 = True And c9 = True Then
            If d7 = 4 And d8 = 4 And d9 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c10 = True Then
            If d7 = 4 And d8 = 4 And d10 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c11 = True Then
            If d7 = 4 And d8 = 4 And d11 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c12 = True Then
            If d7 = 4 And d8 = 4 And d12 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c9 = True Then
            If d7 = 4 And d9 = 4 And d10 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c9 = True Then
            If d7 = 4 And d11 = 4 And d9 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c12 = True And c9 = True Then
            If d7 = 4 And d12 = 4 And d9 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c11 = True Then
            If d7 = d10 = d11 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c12 = True Then
            If d7 = d10 = d12 = 4 Then
                special2 = True
            End If
        End If
        If c10 = True And c8 = True And c9 = True Then
            If d8 = d10 = d9 = 4 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c9 = True Then
            If d11 = d8 = d9 = 4 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c9 = True Then
            If d12 = d8 = d9 = 4 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c10 = True Then
            If d11 = d8 = d10 = 4 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c10 = True Then
            If d12 = d8 = d10 = 4 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c12 = True Then
            If d7 = d11 = d12 = 4 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c12 = True Then
            If d11 = d8 = d12 = 4 Then
                special2 = True
            End If
        End If
        If c11 = True And c10 = True And c9 = True Then
            If d11 = d10 = d9 = 4 Then
                special2 = True
            End If
        End If
        If c12 = True And c10 = True And c9 = True Then
            If d12 = d10 = d9 = 4 Then
                special2 = True
            End If
        End If
        If c11 = True And c12 = True And c9 = True Then
            If d11 = d12 = d9 = 4 Then
                special2 = True
            End If
        End If
        If c10 = True And c11 = True And c12 = True Then
            If d10 = d11 = d12 = 4 Then
                special2 = True
            End If
        End If

        If c7 = True And c8 = True And c9 = True Then
            If d7 = 3 And d8 = 3 And d9 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c10 = True Then
            If d7 = 3 And d8 = 3 And d10 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c11 = True Then
            If d7 = 3 And d8 = 3 And d11 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c12 = True Then
            If d7 = 3 And d8 = 3 And d12 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c9 = True Then
            If d7 = 3 And d9 = 3 And d10 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c9 = True Then
            If d7 = 3 And d11 = 3 And d9 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c12 = True And c9 = True Then
            If d7 = 3 And d12 = 3 And d9 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c11 = True Then
            If d7 = d10 = d11 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c12 = True Then
            If d7 = d10 = d12 = 3 Then
                special2 = True
            End If
        End If
        If c10 = True And c8 = True And c9 = True Then
            If d8 = d10 = d9 = 3 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c9 = True Then
            If d11 = d8 = d9 = 3 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c9 = True Then
            If d12 = d8 = d9 = 3 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c10 = True Then
            If d11 = d8 = d10 = 3 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c10 = True Then
            If d12 = d8 = d10 = 3 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c12 = True Then
            If d7 = d11 = d12 = 3 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c12 = True Then
            If d11 = d8 = d12 = 3 Then
                special2 = True
            End If
        End If
        If c11 = True And c10 = True And c9 = True Then
            If d11 = d10 = d9 = 3 Then
                special2 = True
            End If
        End If
        If c12 = True And c10 = True And c9 = True Then
            If d12 = d10 = d9 = 3 Then
                special2 = True
            End If
        End If
        If c11 = True And c12 = True And c9 = True Then
            If d11 = d12 = d9 = 3 Then
                special2 = True
            End If
        End If
        If c10 = True And c11 = True And c12 = True Then
            If d10 = d11 = d12 = 3 Then
                special2 = True
            End If
        End If

        If c7 = True And c8 = True And c9 = True Then
            If d7 = 2 And d8 = 2 And d9 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c10 = True Then
            If d7 = 2 And d8 = 2 And d10 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c11 = True Then
            If d7 = 2 And d8 = 2 And d11 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c12 = True Then
            If d7 = 2 And d8 = 2 And d12 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c9 = True Then
            If d7 = 2 And d9 = 2 And d10 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c9 = True Then
            If d7 = 2 And d11 = 2 And d9 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c12 = True And c9 = True Then
            If d7 = 2 And d12 = 2 And d9 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c11 = True Then
            If d7 = d10 = d11 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c12 = True Then
            If d7 = d10 = d12 = 2 Then
                special2 = True
            End If
        End If
        If c10 = True And c8 = True And c9 = True Then
            If d8 = d10 = d9 = 2 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c9 = True Then
            If d11 = d8 = d9 = 2 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c9 = True Then
            If d12 = d8 = d9 = 2 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c10 = True Then
            If d11 = d8 = d10 = 2 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c10 = True Then
            If d12 = d8 = d10 = 2 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c12 = True Then
            If d7 = d11 = d12 = 2 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c12 = True Then
            If d11 = d8 = d12 = 2 Then
                special2 = True
            End If
        End If
        If c11 = True And c10 = True And c9 = True Then
            If d11 = d10 = d9 = 2 Then
                special2 = True
            End If
        End If
        If c12 = True And c10 = True And c9 = True Then
            If d12 = d10 = d9 = 2 Then
                special2 = True
            End If
        End If
        If c11 = True And c12 = True And c9 = True Then
            If d11 = d12 = d9 = 2 Then
                special2 = True
            End If
        End If
        If c10 = True And c11 = True And c12 = True Then
            If d10 = d11 = d12 = 2 Then
                special2 = True
            End If
        End If

        If c7 = True And c8 = True And c9 = True Then
            If d7 = 1 And d8 = 1 And d9 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c10 = True Then
            If d7 = 1 And d8 = 1 And d10 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c11 = True Then
            If d7 = 1 And d8 = 1 And d11 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c8 = True And c12 = True Then
            If d7 = 1 And d8 = 1 And d12 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c9 = True Then
            If d7 = 1 And d9 = 1 And d10 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c9 = True Then
            If d7 = 1 And d11 = 1 And d9 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c12 = True And c9 = True Then
            If d7 = 1 And d12 = 1 And d9 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c11 = True Then
            If d7 = d10 = d11 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c10 = True And c12 = True Then
            If d7 = d10 = d12 = 1 Then
                special2 = True
            End If
        End If
        If c10 = True And c8 = True And c9 = True Then
            If d8 = d10 = d9 = 1 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c9 = True Then
            If d11 = d8 = d9 = 1 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c9 = True Then
            If d12 = d8 = d9 = 1 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c10 = True Then
            If d11 = d8 = d10 = 1 Then
                special2 = True
            End If
        End If
        If c12 = True And c8 = True And c10 = True Then
            If d12 = d8 = d10 = 1 Then
                special2 = True
            End If
        End If
        If c7 = True And c11 = True And c12 = True Then
            If d7 = d11 = d12 = 1 Then
                special2 = True
            End If
        End If
        If c11 = True And c8 = True And c12 = True Then
            If d11 = d8 = d12 = 1 Then
                special2 = True
            End If
        End If
        If c11 = True And c10 = True And c9 = True Then
            If d11 = d10 = d9 = 1 Then
                special2 = True
            End If
        End If
        If c12 = True And c10 = True And c9 = True Then
            If d12 = d10 = d9 = 1 Then
                special2 = True
            End If
        End If
        If c11 = True And c12 = True And c9 = True Then
            If d11 = d12 = d9 = 1 Then
                special2 = True
            End If
        End If
        If c10 = True And c11 = True And c12 = True Then
            If d10 = d11 = d12 = 1 Then
                special2 = True
            End If
        End If
    End Sub

    Private Sub bckeep_Click(sender As Object, e As EventArgs) Handles bckeep.Click

        bckeep.Visible = False
        bcreroll.Visible = False
        bcscore.Visible = False
        bcroll.Visible = True
        keep2 = True
        weatherkeep2()
        If d7 = 1 Or d7 = 5 Then bef2 = bef2 + 1
        If d8 = 1 Or d8 = 5 Then bef2 = bef2 + 1
        If d9 = 1 Or d9 = 5 Then bef2 = bef2 + 1
        If d10 = 1 Or d10 = 5 Then bef2 = bef2 + 1
        If d11 = 1 Or d11 = 5 Then bef2 = bef2 + 1
        If d12 = 1 Or d12 = 5 Then bef2 = bef2 + 1
    End Sub
    Public Sub weatherkeep2()
        k7 = k8 = k9 = k10 = k11 = k12 = False
        If d7 = d8 And d9 = d10 And d11 = d12 Or d7 = d8 And d9 = d11 And d10 = d12 Or d7 = d8 And d9 = d12 And d11 = d10 Or d7 = d9 And d8 = d10 And d11 = d12 Or d7 = d9 And d8 = d11 And d10 = d12 Or d7 = d9 And d8 = d12 And d11 = d10 Or d7 = d10 And d9 = d8 And d11 = d12 Or d7 = d10 And d8 = d11 And d9 = d12 Or d7 = d10 And d8 = d12 And d11 = d9 Or d7 = d11 And d8 = d9 And d10 = d12 Or d7 = d11 And d8 = d10 And d9 = d12 Or d7 = d11 And d8 = d12 And d9 = d10 Or d7 = d12 And d9 = d8 And d11 = d10 Or d7 = d12 And d8 = d10 And d11 = d9 Or d7 = d12 And d8 = d11 And d9 = d10 Then
            k7 = True
            k8 = True
            k9 = True
            k10 = True
            k11 = True
            k12 = True
        End If
        If d7 <> d8 And d7 <> d9 And d7 <> d10 And d7 <> d11 And d7 <> d12 And d8 <> d9 And d8 <> d10 And d8 <> d11 And d8 <> d12 And d9 <> d10 And d9 <> d11 And d9 <> d12 And d10 <> d11 And d10 <> d12 And d11 <> d12 Then
            k7 = True
            k8 = True
            k9 = True
            k10 = True
            k11 = True
            k12 = True
        End If
        If d7 + d8 + d9 = 18 Then
            k7 = True
            k8 = True
            k9 = True
        End If
        If d7 + d8 + d10 = 18 Then
            k7 = True
            k8 = True
            k10 = True
        End If
        If d7 + d8 + d11 = 18 Then
            k7 = True
            k8 = True
            k11 = True
        End If
        If d7 + d8 + d12 = 18 Then
            k7 = True
            k8 = True
            k12 = True
        End If
        If d7 + d9 + d10 = 18 Then
            k7 = True
            k10 = True
            k9 = True
        End If
        If d7 + d11 + d9 = 18 Then
            k7 = True
            k11 = True
            k9 = True
        End If
        If d7 + d12 + d9 = 18 Then
            k7 = True
            k12 = True
            k9 = True
        End If
        If d7 + d10 + d11 = 18 Then
            k7 = True
            k11 = True
            k10 = True
        End If
        If d7 + d10 + d12 = 18 Then
            k7 = True
            k10 = True
            k12 = True
        End If
        If d8 + d10 + d9 = 18 Then
            k10 = True
            k8 = True
            k9 = True
        End If
        If d11 + d8 + d9 = 18 Then
            k11 = True
            k8 = True
            k9 = True
        End If
        If d12 + d8 + d9 = 18 Then
            k12 = True
            k8 = True
            k9 = True
        End If
        If d11 + d8 + d10 = 18 Then
            k10 = True
            k8 = True
            k11 = True
        End If
        If d12 + d8 + d10 = 18 Then
            k10 = True
            k8 = True
            k12 = True
        End If
        If d7 + d11 + d12 = 18 Then
            k12 = True
            k11 = True
            k7 = True
        End If
        If d11 + d8 + d12 = 18 Then
            k12 = True
            k8 = True
            k11 = True
        End If
        If d11 + d10 + d9 = 18 Then
            k10 = True
            k11 = True
            k9 = True
        End If
        If d12 + d10 + d9 = 18 Then
            k10 = True
            k12 = True
            k9 = True
        End If
        If d11 + d12 + d9 = 18 Then
            k12 = True
            k11 = True
            k9 = True
        End If
        If d10 + d11 + d12 = 18 Then
            k10 = True
            k11 = True
            k12 = True
        End If

        If d7 = 5 And d8 = 5 And d9 = 5 Then
            k7 = True
            k8 = True
            k9 = True
        End If
        If d7 = 5 And d8 = 5 And d10 = 5 Then
            k7 = True
            k8 = True
            k10 = True
        End If
        If d7 = 5 And d8 = 5 And d11 = 5 Then
            k7 = True
            k8 = True
            k11 = True
        End If
        If d7 = 5 And d8 = 5 And d12 = 5 Then
            k7 = True
            k8 = True
            k12 = True
        End If
        If d7 = 5 And d9 = 5 And d10 = 5 Then
            k7 = True
            k10 = True
            k9 = True
        End If
        If d7 = 5 And d11 = 5 And d9 = 5 Then
            k7 = True
            k11 = True
            k9 = True
        End If
        If d7 = 5 And d12 = 5 And d9 = 5 Then
            k7 = True
            k12 = True
            k9 = True
        End If
        If d7 = d10 = d11 = 5 Then
            k7 = True
            k11 = True
            k10 = True
        End If
        If d7 = d10 = d12 = 5 Then
            k7 = True
            k10 = True
            k12 = True
        End If
        If d8 = d10 = d9 = 5 Then
            k10 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d9 = 5 Then
            k11 = True
            k8 = True
            k9 = True
        End If
        If d12 = d8 = d9 = 5 Then
            k12 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d10 = 5 Then
            k10 = True
            k8 = True
            k11 = True
        End If
        If d12 = d8 = d10 = 5 Then
            k10 = True
            k8 = True
            k12 = True
        End If
        If d7 = d11 = d12 = 5 Then
            k12 = True
            k11 = True
            k7 = True
        End If
        If d11 = d8 = d12 = 5 Then
            k12 = True
            k8 = True
            k11 = True
        End If
        If d11 = d10 = d9 = 5 Then
            k10 = True
            k11 = True
            k9 = True
        End If
        If d12 = d10 = d9 = 5 Then
            k10 = True
            k12 = True
            k9 = True
        End If
        If d11 = d12 = d9 = 5 Then
            k12 = True
            k11 = True
            k9 = True
        End If
        If d10 = d11 = d12 = 5 Then
            k10 = True
            k11 = True
            k12 = True
        End If

        If d7 = 4 And d8 = 4 And d9 = 4 Then
            k7 = True
            k8 = True
            k9 = True
        End If
        If d7 = 4 And d8 = 4 And d10 = 4 Then
            k7 = True
            k8 = True
            k10 = True
        End If
        If d7 = 4 And d8 = 4 And d11 = 4 Then
            k7 = True
            k8 = True
            k11 = True
        End If
        If d7 = 4 And d8 = 4 And d12 = 4 Then
            k7 = True
            k8 = True
            k12 = True
        End If
        If d7 = 4 And d9 = 4 And d10 = 4 Then
            k7 = True
            k10 = True
            k9 = True
        End If
        If d7 = 4 And d11 = 4 And d9 = 4 Then
            k7 = True
            k11 = True
            k9 = True
        End If
        If d7 = 4 And d12 = 4 And d9 = 4 Then
            k7 = True
            k12 = True
            k9 = True
        End If
        If d7 = d10 = d11 = 4 Then
            k7 = True
            k11 = True
            k10 = True
        End If
        If d7 = d10 = d12 = 4 Then
            k7 = True
            k10 = True
            k12 = True
        End If
        If d8 = d10 = d9 = 4 Then
            k10 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d9 = 4 Then
            k11 = True
            k8 = True
            k9 = True
        End If
        If d12 = d8 = d9 = 4 Then
            k12 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d10 = 4 Then
            k10 = True
            k8 = True
            k11 = True
        End If
        If d12 = d8 = d10 = 4 Then
            k10 = True
            k8 = True
            k12 = True
        End If
        If d7 = d11 = d12 = 4 Then
            k12 = True
            k11 = True
            k7 = True
        End If
        If d11 = d8 = d12 = 4 Then
            k12 = True
            k8 = True
            k11 = True
        End If
        If d11 = d10 = d9 = 4 Then
            k10 = True
            k11 = True
            k9 = True
        End If
        If d12 = d10 = d9 = 4 Then
            k10 = True
            k12 = True
            k9 = True
        End If
        If d11 = d12 = d9 = 4 Then
            k12 = True
            k11 = True
            k9 = True
        End If
        If d10 = d11 = d12 = 4 Then
            k10 = True
            k11 = True
            k12 = True
        End If

        If d7 = 3 And d8 = 3 And d9 = 3 Then
            k7 = True
            k8 = True
            k9 = True
        End If
        If d7 = 3 And d8 = 3 And d10 = 3 Then
            k7 = True
            k8 = True
            k10 = True
        End If
        If d7 = 3 And d8 = 3 And d11 = 3 Then
            k7 = True
            k8 = True
            k11 = True
        End If
        If d7 = 3 And d8 = 3 And d12 = 3 Then
            k7 = True
            k8 = True
            k12 = True
        End If
        If d7 = 3 And d9 = 3 And d10 = 3 Then
            k7 = True
            k10 = True
            k9 = True
        End If
        If d7 = 3 And d11 = 3 And d9 = 3 Then
            k7 = True
            k11 = True
            k9 = True
        End If
        If d7 = 3 And d12 = 3 And d9 = 3 Then
            k7 = True
            k12 = True
            k9 = True
        End If
        If d7 = d10 = d11 = 3 Then
            k7 = True
            k11 = True
            k10 = True
        End If
        If d7 = d10 = d12 = 3 Then
            k7 = True
            k10 = True
            k12 = True
        End If
        If d8 = d10 = d9 = 3 Then
            k10 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d9 = 3 Then
            k11 = True
            k8 = True
            k9 = True
        End If
        If d12 = d8 = d9 = 3 Then
            k12 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d10 = 3 Then
            k10 = True
            k8 = True
            k11 = True
        End If
        If d12 = d8 = d10 = 3 Then
            k10 = True
            k8 = True
            k12 = True
        End If
        If d7 = d11 = d12 = 3 Then
            k12 = True
            k11 = True
            k7 = True
        End If
        If d11 = d8 = d12 = 3 Then
            k12 = True
            k8 = True
            k11 = True
        End If
        If d11 = d10 = d9 = 3 Then
            k10 = True
            k11 = True
            k9 = True
        End If
        If d12 = d10 = d9 = 3 Then
            k10 = True
            k12 = True
            k9 = True
        End If
        If d11 = d12 = d9 = 3 Then
            k12 = True
            k11 = True
            k9 = True
        End If
        If d10 = d11 = d12 = 3 Then
            k10 = True
            k11 = True
            k12 = True
        End If

        If d7 = 2 And d8 = 2 And d9 = 2 Then
            k7 = True
            k8 = True
            k9 = True
        End If
        If d7 = 2 And d8 = 2 And d10 = 2 Then
            k7 = True
            k8 = True
            k10 = True
        End If
        If d7 = 2 And d8 = 2 And d11 = 2 Then
            k7 = True
            k8 = True
            k11 = True
        End If
        If d7 = 2 And d8 = 2 And d12 = 2 Then
            k7 = True
            k8 = True
            k12 = True
        End If
        If d7 = 2 And d9 = 2 And d10 = 2 Then
            k7 = True
            k10 = True
            k9 = True
        End If
        If d7 = 2 And d11 = 2 And d9 = 2 Then
            k7 = True
            k11 = True
            k9 = True
        End If
        If d7 = 2 And d12 = 2 And d9 = 2 Then
            k7 = True
            k12 = True
            k9 = True
        End If
        If d7 = d10 = d11 = 2 Then
            k7 = True
            k11 = True
            k10 = True
        End If
        If d7 = d10 = d12 = 2 Then
            k7 = True
            k10 = True
            k12 = True
        End If
        If d8 = d10 = d9 = 2 Then
            k10 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d9 = 2 Then
            k11 = True
            k8 = True
            k9 = True
        End If
        If d12 = d8 = d9 = 2 Then
            k12 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d10 = 2 Then
            k10 = True
            k8 = True
            k11 = True
        End If
        If d12 = d8 = d10 = 2 Then
            k10 = True
            k8 = True
            k12 = True
        End If
        If d7 = d11 = d12 = 2 Then
            k12 = True
            k11 = True
            k7 = True
        End If
        If d11 = d8 = d12 = 2 Then
            k12 = True
            k8 = True
            k11 = True
        End If
        If d11 = d10 = d9 = 2 Then
            k10 = True
            k11 = True
            k9 = True
        End If
        If d12 = d10 = d9 = 2 Then
            k10 = True
            k12 = True
            k9 = True
        End If
        If d11 = d12 = d9 = 2 Then
            k12 = True
            k11 = True
            k9 = True
        End If
        If d10 = d11 = d12 = 2 Then
            k10 = True
            k11 = True
            k12 = True
        End If

        If d7 = 1 And d8 = 1 And d9 = 1 Then
            k7 = True
            k8 = True
            k9 = True
        End If
        If d7 = 1 And d8 = 1 And d10 = 1 Then
            k7 = True
            k8 = True
            k10 = True
        End If
        If d7 = 1 And d8 = 1 And d11 = 1 Then
            k7 = True
            k8 = True
            k11 = True
        End If
        If d7 = 1 And d8 = 1 And d12 = 1 Then
            k7 = True
            k8 = True
            k12 = True
        End If
        If d7 = 1 And d9 = 1 And d10 = 1 Then
            k7 = True
            k10 = True
            k9 = True
        End If
        If d7 = 1 And d11 = 1 And d9 = 1 Then
            k7 = True
            k11 = True
            k9 = True
        End If
        If d7 = 1 And d12 = 1 And d9 = 1 Then
            k7 = True
            k12 = True
            k9 = True
        End If
        If d7 = d10 = d11 = 1 Then
            k7 = True
            k11 = True
            k10 = True
        End If
        If d7 = d10 = d12 = 1 Then
            k7 = True
            k10 = True
            k12 = True
        End If
        If d8 = d10 = d9 = 1 Then
            k10 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d9 = 1 Then
            k11 = True
            k8 = True
            k9 = True
        End If
        If d12 = d8 = d9 = 1 Then
            k12 = True
            k8 = True
            k9 = True
        End If
        If d11 = d8 = d10 = 1 Then
            k10 = True
            k8 = True
            k11 = True
        End If
        If d12 = d8 = d10 = 1 Then
            k10 = True
            k8 = True
            k12 = True
        End If
        If d7 = d11 = d12 = 1 Then
            k12 = True
            k11 = True
            k7 = True
        End If
        If d11 = d8 = d12 = 1 Then
            k12 = True
            k8 = True
            k11 = True
        End If
        If d11 = d10 = d9 = 1 Then
            k10 = True
            k11 = True
            k9 = True
        End If
        If d12 = d10 = d9 = 1 Then
            k10 = True
            k12 = True
            k9 = True
        End If
        If d11 = d12 = d9 = 1 Then
            k12 = True
            k11 = True
            k9 = True
        End If
        If d10 = d11 = d12 = 1 Then
            k10 = True
            k11 = True
            k12 = True
        End If
    End Sub

    Private Sub bcscore_Click(sender As Object, e As EventArgs) Handles bcscore.Click
        determine2()
        cp.Text = score2
        bckeep.Visible = False
        bcreroll.Visible = False
        bcscore.Visible = False
        bcroll.Visible = False
        If score2 > 10000 Then
            MsgBox("Player2 Wins !")
            score2 = 0
            fark2 = 0
            score = 0
            fark = 0
        End If
        pp.Text = score
        ft.Text = fark
        cp.Text = score2
        ft2.Text = fark2
        bRoll.Visible = True
    End Sub
    Public Sub determine2()
        five2 = 0
        one2 = 0
        If d7 = d8 And d9 = d10 And d11 = d12 Or d7 = d8 And d9 = d11 And d10 = d12 Or d7 = d8 And d9 = d12 And d11 = d10 Or d7 = d9 And d8 = d10 And d11 = d12 Or d7 = d9 And d8 = d11 And d10 = d12 Or d7 = d9 And d8 = d12 And d11 = d10 Or d7 = d10 And d9 = d8 And d11 = d12 Or d7 = d10 And d8 = d11 And d9 = d12 Or d7 = d10 And d8 = d12 And d11 = d9 Or d7 = d11 And d8 = d9 And d10 = d12 Or d7 = d11 And d8 = d10 And d9 = d12 Or d7 = d11 And d8 = d12 And d9 = d10 Or d7 = d12 And d9 = d8 And d11 = d10 Or d7 = d12 And d8 = d10 And d11 = d9 Or d7 = d12 And d8 = d11 And d9 = d10 Then
            score2 = score2 + 1500
            recalc2 = False
        End If

        If d7 <> d8 And d7 <> d9 And d7 <> d10 And d7 <> d11 And d7 <> d12 And d8 <> d9 And d8 <> d10 And d8 <> d11 And d8 <> d12 And d9 <> d10 And d9 <> d11 And d9 <> d12 And d10 <> d11 And d10 <> d12 And d11 <> d12 Then
            score2 = score2 + 3000
            recalc2 = False
        End If

        If d7 + d8 + d9 = 18 Or d7 + d8 + d10 = 18 Or d7 + d8 + d11 = 18 Or d7 + d8 + d12 = 18 Or d7 + d9 + d10 = 18 Or d7 + d11 + d9 = 18 Or d7 + d12 + d9 = 18 Or d7 + d10 + d11 = 18 Or d7 + d10 + d12 = 18 Or d8 + d10 + d9 = 18 Or d11 + d8 + d9 = 18 Or d12 + d8 + d9 = 18 Or d11 + d8 + d10 = 18 Or d12 + d8 + d10 = 18 Or d7 + d11 + d12 = 18 Or d11 + d8 + d12 = 18 Or d11 + d10 + d9 = 18 Or d12 + d10 + d9 = 18 Or d11 + d12 + d9 = 18 Or d10 + d11 + d12 = 18 Then
            score2 = score2 + 600
            If d7 = 6 Then pick7 = True
            If d8 = 6 Then pick8 = True
            If d9 = 6 Then pick9 = True
            If d10 = 6 Then pick10 = True
            If d11 = 6 Then pick11 = True
            If d12 = 6 Then pick12 = True
        End If

        If d7 = 5 And d8 = 5 And d9 = 5 Or d7 = 5 And d8 = 5 And d10 = 5 Or d7 = 5 And d8 = 5 And d11 = 5 Or d7 = 5 And d8 = 5 And d12 = 5 Or d7 = 5 And d9 = 5 And d10 = 5 Or d7 = 5 And d11 = 5 And d9 = 5 Or d7 = 5 And d12 = 5 And d9 = 5 Or d7 = 5 And d10 = 5 And d11 = 5 Or d7 = 5 And d10 = 5 And d12 = 5 Or d8 = 5 And d10 = 5 And d9 = 5 Or d11 = 5 And d8 = 5 And d9 = 5 Or d12 = 5 And d8 = 5 And d9 = 5 Or d11 = 5 And d8 = 5 And d10 = 5 Or d12 = 5 And d8 = 5 And d10 = 5 Or d7 = 5 And d11 = 5 And d12 = 5 Or d11 = 5 And d8 = 5 And d12 = 5 Or d11 = 5 And d10 = 5 And d9 = 5 Or d12 = 5 And d10 = 5 And d9 = 5 Or d11 = 5 And d12 = 5 And d9 = 5 Or d10 = 5 And d11 = 5 And d12 = 5 Then
            score2 = score2 + 500
            five2 = 3
        End If

        If d7 = 4 And d8 = 4 And d9 = 4 Or d7 = 4 And d8 = 4 And d10 = 4 Or d7 = 4 And d8 = 4 And d11 = 4 Or d7 = 4 And d8 = 4 And d12 = 4 Or d7 = 4 And d9 = 4 And d10 = 4 Or d7 = 4 And d11 = 4 And d9 = 4 Or d7 = 4 And d12 = 4 And d9 = 4 Or d7 = 4 And d10 = 4 And d11 = 4 Or d7 = 4 And d10 = 4 And d12 = 4 Or d8 = 4 And d10 = 4 And d9 = 4 Or d11 = 4 And d8 = 4 And d9 = 4 Or d12 = 4 And d8 = 4 And d9 = 4 Or d11 = 4 And d8 = 4 And d10 = 4 Or d12 = 4 And d8 = 4 And d10 = 4 Or d7 = 4 And d11 = 4 And d12 = 4 Or d11 = 4 And d8 = 4 And d12 = 4 Or d11 = 4 And d10 = 4 And d9 = 4 Or d12 = 4 And d10 = 4 And d9 = 4 Or d11 = 4 And d12 = 4 And d9 = 4 Or d10 = 4 And d11 = 4 And d12 = 4 Then
            score2 = score2 + 400
        End If

        If d7 = 3 And d8 = 3 And d9 = 3 Or d7 = 3 And d8 = 3 And d10 = 3 Or d7 = 3 And d8 = 3 And d11 = 3 Or d7 = 3 And d8 = 3 And d12 = 3 Or d7 = 3 And d9 = 3 And d10 = 3 Or d7 = 3 And d11 = 3 And d9 = 3 Or d7 = 3 And d12 = 3 And d9 = 3 Or d7 = 3 And d10 = 3 And d11 = 3 Or d7 = 3 And d10 = 3 And d12 = 3 Or d8 = 3 And d10 = 3 And d9 = 3 Or d11 = 3 And d8 = 3 And d9 = 3 Or d12 = 3 And d8 = 3 And d9 = 3 Or d11 = 3 And d8 = 3 And d10 = 3 Or d12 = 3 And d8 = 3 And d10 = 3 Or d7 = 3 And d11 = 3 And d12 = 3 Or d11 = 3 And d8 = 3 And d12 = 3 Or d11 = 3 And d10 = 3 And d9 = 3 Or d12 = 3 And d10 = 3 And d9 = 3 Or d11 = 3 And d12 = 3 And d9 = 3 Or d10 = 3 And d11 = 3 And d12 = 3 Then
            score2 = score2 + 300
        End If

        If d7 = 2 And d8 = 2 And d9 = 2 Or d7 = 2 And d8 = 2 And d10 = 2 Or d7 = 2 And d8 = 2 And d11 = 2 Or d7 = 2 And d8 = 2 And d12 = 2 Or d7 = 2 And d9 = 2 And d10 = 2 Or d7 = 2 And d11 = 2 And d9 = 2 Or d7 = 2 And d12 = 2 And d9 = 2 Or d7 = 2 And d10 = 2 And d11 = 2 Or d7 = 2 And d10 = 2 And d12 = 2 Or d8 = 2 And d10 = 2 And d9 = 2 Or d11 = 2 And d8 = 2 And d9 = 2 Or d12 = 2 And d8 = 2 And d9 = 2 Or d11 = 2 And d8 = 2 And d10 = 2 Or d12 = 2 And d8 = 2 And d10 = 2 Or d7 = 2 And d11 = 2 And d12 = 2 Or d11 = 2 And d8 = 2 And d12 = 2 Or d11 = 2 And d10 = 2 And d9 = 2 Or d12 = 2 And d10 = 2 And d9 = 2 Or d11 = 2 And d12 = 2 And d9 = 2 Or d10 = 2 And d11 = 2 And d12 = 2 Then
            score2 = score2 + 200
        End If

        If d7 = 1 And d8 = 1 And d9 = 1 Or d7 = 1 And d8 = 1 And d10 = 1 Or d7 = 1 And d8 = 1 And d11 = 1 Or d7 = 1 And d8 = 1 And d12 = 1 Or d7 = 1 And d9 = 1 And d10 = 1 Or d7 = 1 And d11 = 1 And d9 = 1 Or d7 = 1 And d12 = 1 And d9 = 1 Or d7 = 1 And d10 = 1 And d11 = 1 Or d7 = 1 And d10 = 1 And d12 = 1 Or d8 = 1 And d10 = 1 And d9 = 1 Or d11 = 1 And d8 = 1 And d9 = 1 Or d12 = 1 And d8 = 1 And d9 = 1 Or d11 = 1 And d8 = 1 And d10 = 1 Or d12 = 1 And d8 = 1 And d10 = 1 Or d7 = 1 And d11 = 1 And d12 = 1 Or d11 = 1 And d8 = 1 And d12 = 1 Or d11 = 1 And d10 = 1 And d9 = 1 Or d12 = 1 And d10 = 1 And d9 = 1 Or d11 = 1 And d12 = 1 And d9 = 1 Or d10 = 1 And d11 = 1 And d12 = 1 Then
            score2 = score2 + 1000
            one2 = 3
        End If

        If d7 = 1 And one2 = 0 And recalc2 = True Then
            score2 = score2 + 100
        ElseIf d7 = 1 And one2 > 0 And recalc2 = True Then
            one2 = one2 - 1
        End If
        If d8 = 1 And one2 = 0 And recalc2 = True Then
            score2 = score2 + 100
        ElseIf d8 = 1 And one2 > 0 And recalc2 = True Then
            one2 = one2 - 1
        End If
        If d9 = 1 And one2 = 0 And recalc2 = True Then

            score2 = score2 + 100
        ElseIf d9 = 1 And one2 > 0 And recalc2 = True Then
            one2 = one2 - 1
        End If
        If d10 = 1 And one2 = 0 And recalc2 = True Then

            score2 = score2 + 100
        ElseIf d10 = 1 And one2 > 0 And recalc2 = True Then
            one2 = one2 - 1
        End If
        If d11 = 1 And one2 = 0 And recalc2 = True Then

            score2 = score2 + 100
        ElseIf d11 = 1 And one2 > 0 And recalc2 = True Then
            one2 = one2 - 1
        End If
        If d12 = 1 And one2 = 0 And recalc2 = True Then

            score2 = score2 + 100
        ElseIf d12 = 1 And one2 > 0 And recalc2 = True Then
            one2 = one2 - 1
        End If
        If d7 = 5 And five2 = 0 And recalc2 = True Then

            score2 = score2 + 50
        ElseIf d7 = 5 And five2 > 0 And recalc2 = True Then
            five2 = five2 - 1
        End If
        If d8 = 5 And five2 = 0 And recalc2 = True Then

            score2 = score2 + 50
        ElseIf d8 = 5 And five2 > 0 And recalc2 = True Then
            five2 = five2 - 1
        End If
        If d9 = 5 And five2 = 0 And recalc2 = True Then

            score2 = score2 + 50
        ElseIf d9 = 5 And five2 > 0 And recalc2 = True Then
            five2 = five2 - 1
        End If
        If d10 = 5 And five2 = 0 And recalc2 = True Then

            score2 = score2 + 50
        ElseIf d10 = 5 And five2 > 0 And recalc2 = True Then
            five2 = five2 - 1
        End If
        If d11 = 5 And five2 = 0 And recalc2 = True Then
            score2 = score2 + 50
        ElseIf d11 = 5 And five2 > 0 And recalc2 = True Then
            five2 = five2 - 1
        End If
        If d12 = 5 And five2 = 0 And recalc2 = True Then
            score2 = score2 + 50
        ElseIf d12 = 5 And five2 > 0 And recalc2 = True Then
            five2 = five2 - 1
        End If
        recalc2 = True
    End Sub

    Private Sub bcreroll_Click(sender As Object, e As EventArgs) Handles bcreroll.Click
        d7 = Int(Rnd() * 6) + 1
        d8 = Int(Rnd() * 6) + 1
        d9 = Int(Rnd() * 6) + 1
        d10 = Int(Rnd() * 6) + 1
        d11 = Int(Rnd() * 6) + 1
        d12 = Int(Rnd() * 6) + 1
        showdice7()
        showdice8()
        showdice9()
        showdice12()
        showdice11()
        showdice10()
        nofarkle2()

        If d7 <> 1 And d7 <> 5 And d8 <> 1 And d8 <> 5 And d9 <> 1 And d9 <> 5 And d10 <> 1 And d10 <> 5 And d11 <> 1 And d11 <> 5 And d12 <> 1 And d12 <> 5 And special2 = False Then
            fark2 = fark2 + 1
            MsgBox("Player2 got Farkled " & fark & " time(s)!")
            bcroll.Visible = False
            bRoll.Visible = True
            bckeep.Visible = False
            bcreroll.Visible = False
            bcscore.Visible = False
            If fark2 = 3 Then
                fark2 = 0
                score2 = 0
            End If
            ft2.Text = fark2
            cp.Text = score2

        End If
    End Sub
    Public Sub nofarkle2()
        special2 = False
        If d7 = d8 And d9 = d10 And d11 = d12 Or d7 = d8 And d9 = d11 And d10 = d12 Or d7 = d8 And d9 = d12 And d11 = d10 Or d7 = d9 And d8 = d10 And d11 = d12 Or d7 = d9 And d8 = d11 And d10 = d12 Or d7 = d9 And d8 = d12 And d11 = d10 Or d7 = d10 And d9 = d8 And d11 = d12 Or d7 = d10 And d8 = d11 And d9 = d12 Or d7 = d10 And d8 = d12 And d11 = d9 Or d7 = d11 And d8 = d9 And d10 = d12 Or d7 = d11 And d8 = d10 And d9 = d12 Or d7 = d11 And d8 = d12 And d9 = d10 Or d7 = d12 And d9 = d8 And d11 = d10 Or d7 = d12 And d8 = d10 And d11 = d9 Or d7 = d12 And d8 = d11 And d9 = d10 Then
            special2 = True
        End If

        If d7 <> d8 And d7 <> d9 And d7 <> d10 And d7 <> d11 And d7 <> d12 And d8 <> d9 And d8 <> d10 And d8 <> d11 And d8 <> d12 And d9 <> d10 And d9 <> d11 And d9 <> d12 And d10 <> d11 And d10 <> d12 And d11 <> d12 Then
            special2 = True
        End If

        If d7 + d8 + d9 = 18 Or d7 + d8 + d10 = 18 Or d7 + d8 + d11 = 18 Or d7 + d8 + d12 = 18 Or d7 + d9 + d10 = 18 Or d7 + d11 + d9 = 18 Or d7 + d12 + d9 = 18 Or d7 + d10 + d11 = 18 Or d7 + d10 + d12 = 18 Or d8 + d10 + d9 = 18 Or d11 + d8 + d9 = 18 Or d12 + d8 + d9 = 18 Or d11 + d8 + d10 = 18 Or d12 + d8 + d10 = 18 Or d7 + d11 + d12 = 18 Or d11 + d8 + d12 = 18 Or d11 + d10 + d9 = 18 Or d12 + d10 + d9 = 18 Or d11 + d12 + d9 = 18 Or d10 + d11 + d12 = 18 Then
            special2 = True
        End If

        If d7 = 5 And d8 = 5 And d9 = 5 Or d7 = 5 And d8 = 5 And d10 = 5 Or d7 = 5 And d8 = 5 And d11 = 5 Or d7 = 5 And d8 = 5 And d12 = 5 Or d7 = 5 And d9 = 5 And d10 = 5 Or d7 = 5 And d11 = 5 And d9 = 5 Or d7 = 5 And d12 = 5 And d9 = 5 Or d7 = 5 And d10 = 5 And d11 = 5 Or d7 = 5 And d10 = 5 And d12 = 5 Or d8 = 5 And d10 = 5 And d9 = 5 Or d11 = 5 And d8 = 5 And d9 = 5 Or d12 = 5 And d8 = 5 And d9 = 5 Or d11 = 5 And d8 = 5 And d10 = 5 Or d12 = 5 And d8 = 5 And d10 = 5 Or d7 = 5 And d11 = 5 And d12 = 5 Or d11 = 5 And d8 = 5 And d12 = 5 Or d11 = 5 And d10 = 5 And d9 = 5 Or d12 = 5 And d10 = 5 And d9 = 5 Or d11 = 5 And d12 = 5 And d9 = 5 Or d10 = 5 And d11 = 5 And d12 = 5 Then
            special2 = True
        End If

        If d7 = 4 And d8 = 4 And d9 = 4 Or d7 = 4 And d8 = 4 And d10 = 4 Or d7 = 4 And d8 = 4 And d11 = 4 Or d7 = 4 And d8 = 4 And d12 = 4 Or d7 = 4 And d9 = 4 And d10 = 4 Or d7 = 4 And d11 = 4 And d9 = 4 Or d7 = 4 And d12 = 4 And d9 = 4 Or d7 = 4 And d10 = 4 And d11 = 4 Or d7 = 4 And d10 = 4 And d12 = 4 Or d8 = 4 And d10 = 4 And d9 = 4 Or d11 = 4 And d8 = 4 And d9 = 4 Or d12 = 4 And d8 = 4 And d9 = 4 Or d11 = 4 And d8 = 4 And d10 = 4 Or d12 = 4 And d8 = 4 And d10 = 4 Or d7 = 4 And d11 = 4 And d12 = 4 Or d11 = 4 And d8 = 4 And d12 = 4 Or d11 = 4 And d10 = 4 And d9 = 4 Or d12 = 4 And d10 = 4 And d9 = 4 Or d11 = 4 And d12 = 4 And d9 = 4 Or d10 = 4 And d11 = 4 And d12 = 4 Then
            special2 = True
        End If

        If d7 = 3 And d8 = 3 And d9 = 3 Or d7 = 3 And d8 = 3 And d10 = 3 Or d7 = 3 And d8 = 3 And d11 = 3 Or d7 = 3 And d8 = 3 And d12 = 3 Or d7 = 3 And d9 = 3 And d10 = 3 Or d7 = 3 And d11 = 3 And d9 = 3 Or d7 = 3 And d12 = 3 And d9 = 3 Or d7 = 3 And d10 = 3 And d11 = 3 Or d7 = 3 And d10 = 3 And d12 = 3 Or d8 = 3 And d10 = 3 And d9 = 3 Or d11 = 3 And d8 = 3 And d9 = 3 Or d12 = 3 And d8 = 3 And d9 = 3 Or d11 = 3 And d8 = 3 And d10 = 3 Or d12 = 3 And d8 = 3 And d10 = 3 Or d7 = 3 And d11 = 3 And d12 = 3 Or d11 = 3 And d8 = 3 And d12 = 3 Or d11 = 3 And d10 = 3 And d9 = 3 Or d12 = 3 And d10 = 3 And d9 = 3 Or d11 = 3 And d12 = 3 And d9 = 3 Or d10 = 3 And d11 = 3 And d12 = 3 Then
            special2 = True
        End If

        If d7 = 2 And d8 = 2 And d9 = 2 Or d7 = 2 And d8 = 2 And d10 = 2 Or d7 = 2 And d8 = 2 And d11 = 2 Or d7 = 2 And d8 = 2 And d12 = 2 Or d7 = 2 And d9 = 2 And d10 = 2 Or d7 = 2 And d11 = 2 And d9 = 2 Or d7 = 2 And d12 = 2 And d9 = 2 Or d7 = 2 And d10 = 2 And d11 = 2 Or d7 = 2 And d10 = 2 And d12 = 2 Or d8 = 2 And d10 = 2 And d9 = 2 Or d11 = 2 And d8 = 2 And d9 = 2 Or d12 = 2 And d8 = 2 And d9 = 2 Or d11 = 2 And d8 = 2 And d10 = 2 Or d12 = 2 And d8 = 2 And d10 = 2 Or d7 = 2 And d11 = 2 And d12 = 2 Or d11 = 2 And d8 = 2 And d12 = 2 Or d11 = 2 And d10 = 2 And d9 = 2 Or d12 = 2 And d10 = 2 And d9 = 2 Or d11 = 2 And d12 = 2 And d9 = 2 Or d10 = 2 And d11 = 2 And d12 = 2 Then
            special2 = True

        End If

        If d7 = 1 And d8 = 1 And d9 = 1 Or d7 = 1 And d8 = 1 And d10 = 1 Or d7 = 1 And d8 = 1 And d11 = 1 Or d7 = 1 And d8 = 1 And d12 = 1 Or d7 = 1 And d9 = 1 And d10 = 1 Or d7 = 1 And d11 = 1 And d9 = 1 Or d7 = 1 And d12 = 1 And d9 = 1 Or d7 = 1 And d10 = 1 And d11 = 1 Or d7 = 1 And d10 = 1 And d12 = 1 Or d8 = 1 And d10 = 1 And d9 = 1 Or d11 = 1 And d8 = 1 And d9 = 1 Or d12 = 1 And d8 = 1 And d9 = 1 Or d11 = 1 And d8 = 1 And d10 = 1 Or d12 = 1 And d8 = 1 And d10 = 1 Or d7 = 1 And d11 = 1 And d12 = 1 Or d11 = 1 And d8 = 1 And d12 = 1 Or d11 = 1 And d10 = 1 And d9 = 1 Or d12 = 1 And d10 = 1 And d9 = 1 Or d11 = 1 And d12 = 1 And d9 = 1 Or d10 = 1 And d11 = 1 And d12 = 1 Then
            special2 = True
        End If

    End Sub
End Class
