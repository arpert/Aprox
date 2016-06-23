Public Class Form1
    Inherits System.Windows.Forms.Form
   Dim lastY As Single
   Dim rootpa As String
   Dim xmin, xmax, ymin, ymax As Double

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
   Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
   Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
   Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
   Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents Button1 As System.Windows.Forms.Button
   Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
   Friend WithEvents ListView1 As System.Windows.Forms.ListView
   Friend WithEvents Plik As System.Windows.Forms.ColumnHeader
   Friend WithEvents Mapa As System.Windows.Forms.ColumnHeader
   Friend WithEvents Button2 As System.Windows.Forms.Button
   Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
   Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
   Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
   Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
   <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
      Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
      Me.PictureBox1 = New System.Windows.Forms.PictureBox()
      Me.TextBox1 = New System.Windows.Forms.TextBox()
      Me.TextBox2 = New System.Windows.Forms.TextBox()
      Me.TextBox3 = New System.Windows.Forms.TextBox()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.TextBox4 = New System.Windows.Forms.TextBox()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.Button1 = New System.Windows.Forms.Button()
      Me.PictureBox2 = New System.Windows.Forms.PictureBox()
      Me.ListView1 = New System.Windows.Forms.ListView()
      Me.Plik = New System.Windows.Forms.ColumnHeader()
      Me.Mapa = New System.Windows.Forms.ColumnHeader()
      Me.Button2 = New System.Windows.Forms.Button()
      Me.TextBox5 = New System.Windows.Forms.TextBox()
      Me.TextBox6 = New System.Windows.Forms.TextBox()
      Me.TextBox7 = New System.Windows.Forms.TextBox()
      Me.TextBox8 = New System.Windows.Forms.TextBox()
      Me.SuspendLayout()
      '
      'PictureBox1
      '
      Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
      Me.PictureBox1.Location = New System.Drawing.Point(8, 8)
      Me.PictureBox1.Name = "PictureBox1"
      Me.PictureBox1.Size = New System.Drawing.Size(72, 352)
      Me.PictureBox1.TabIndex = 0
      Me.PictureBox1.TabStop = False
      '
      'TextBox1
      '
      Me.TextBox1.Location = New System.Drawing.Point(176, 8)
      Me.TextBox1.Name = "TextBox1"
      Me.TextBox1.Size = New System.Drawing.Size(136, 20)
      Me.TextBox1.TabIndex = 1
      Me.TextBox1.Text = "TextBox1"
      '
      'TextBox2
      '
      Me.TextBox2.Location = New System.Drawing.Point(176, 40)
      Me.TextBox2.Name = "TextBox2"
      Me.TextBox2.Size = New System.Drawing.Size(136, 20)
      Me.TextBox2.TabIndex = 2
      Me.TextBox2.Text = "TextBox2"
      '
      'TextBox3
      '
      Me.TextBox3.Location = New System.Drawing.Point(176, 72)
      Me.TextBox3.Name = "TextBox3"
      Me.TextBox3.Size = New System.Drawing.Size(136, 20)
      Me.TextBox3.TabIndex = 3
      Me.TextBox3.Text = "TextBox3"
      '
      'Label1
      '
      Me.Label1.Location = New System.Drawing.Point(128, 8)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(40, 16)
      Me.Label1.TabIndex = 4
      Me.Label1.Text = "In pix"
      '
      'Label2
      '
      Me.Label2.Location = New System.Drawing.Point(128, 40)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(40, 16)
      Me.Label2.TabIndex = 5
      Me.Label2.Text = "Min"
      '
      'Label3
      '
      Me.Label3.Location = New System.Drawing.Point(128, 72)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(40, 16)
      Me.Label3.TabIndex = 6
      Me.Label3.Text = "Max"
      '
      'TextBox4
      '
      Me.TextBox4.Location = New System.Drawing.Point(176, 104)
      Me.TextBox4.Name = "TextBox4"
      Me.TextBox4.Size = New System.Drawing.Size(136, 20)
      Me.TextBox4.TabIndex = 7
      Me.TextBox4.Text = "TextBox4"
      '
      'Label4
      '
      Me.Label4.Location = New System.Drawing.Point(128, 104)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(40, 16)
      Me.Label4.TabIndex = 8
      Me.Label4.Text = "Calc"
      '
      'Button1
      '
      Me.Button1.Location = New System.Drawing.Point(328, 8)
      Me.Button1.Name = "Button1"
      Me.Button1.Size = New System.Drawing.Size(80, 16)
      Me.Button1.TabIndex = 9
      Me.Button1.Text = "Button1"
      '
      'PictureBox2
      '
      Me.PictureBox2.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                  Or System.Windows.Forms.AnchorStyles.Left) _
                  Or System.Windows.Forms.AnchorStyles.Right)
      Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
      Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Bitmap)
      Me.PictureBox2.Location = New System.Drawing.Point(96, 128)
      Me.PictureBox2.Name = "PictureBox2"
      Me.PictureBox2.Size = New System.Drawing.Size(496, 392)
      Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
      Me.PictureBox2.TabIndex = 10
      Me.PictureBox2.TabStop = False
      '
      'ListView1
      '
      Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Plik, Me.Mapa})
      Me.ListView1.FullRowSelect = True
      Me.ListView1.Location = New System.Drawing.Point(328, 32)
      Me.ListView1.Name = "ListView1"
      Me.ListView1.Size = New System.Drawing.Size(264, 88)
      Me.ListView1.TabIndex = 11
      Me.ListView1.View = System.Windows.Forms.View.Details
      '
      'Plik
      '
      Me.Plik.Text = "Plik"
      Me.Plik.Width = 80
      '
      'Mapa
      '
      Me.Mapa.Text = "Mapa"
      Me.Mapa.Width = 180
      '
      'Button2
      '
      Me.Button2.Location = New System.Drawing.Point(512, 8)
      Me.Button2.Name = "Button2"
      Me.Button2.Size = New System.Drawing.Size(80, 16)
      Me.Button2.TabIndex = 12
      Me.Button2.Text = "Button2"
      '
      'TextBox5
      '
      Me.TextBox5.Location = New System.Drawing.Point(8, 376)
      Me.TextBox5.Name = "TextBox5"
      Me.TextBox5.Size = New System.Drawing.Size(72, 20)
      Me.TextBox5.TabIndex = 13
      Me.TextBox5.Text = "TextBox5"
      '
      'TextBox6
      '
      Me.TextBox6.Location = New System.Drawing.Point(8, 408)
      Me.TextBox6.Name = "TextBox6"
      Me.TextBox6.Size = New System.Drawing.Size(72, 20)
      Me.TextBox6.TabIndex = 14
      Me.TextBox6.Text = "TextBox6"
      '
      'TextBox7
      '
      Me.TextBox7.Location = New System.Drawing.Point(8, 440)
      Me.TextBox7.Name = "TextBox7"
      Me.TextBox7.Size = New System.Drawing.Size(72, 20)
      Me.TextBox7.TabIndex = 15
      Me.TextBox7.Text = "TextBox7"
      '
      'TextBox8
      '
      Me.TextBox8.Location = New System.Drawing.Point(8, 472)
      Me.TextBox8.Name = "TextBox8"
      Me.TextBox8.Size = New System.Drawing.Size(72, 20)
      Me.TextBox8.TabIndex = 16
      Me.TextBox8.Text = "TextBox8"
      '
      'Form1
      '
      Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
      Me.ClientSize = New System.Drawing.Size(600, 522)
      Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBox8, Me.TextBox7, Me.TextBox6, Me.TextBox5, Me.Button2, Me.ListView1, Me.PictureBox2, Me.Button1, Me.Label4, Me.TextBox4, Me.Label3, Me.Label2, Me.Label1, Me.TextBox3, Me.TextBox2, Me.TextBox1, Me.PictureBox1})
      Me.Name = "Form1"
      Me.Text = "Form1"
      Me.ResumeLayout(False)

   End Sub

#End Region

   Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      lastY = -1
      TextBox1.Text = 0
      TextBox2.Text = 0
      TextBox3.Text = 100
      TextBox4.Text = -1
      xmin = 13.96666666666
      ymax = 54.68333333333
      xmax = 24.36666666666
      ymin = 49.22222222222

   End Sub

   Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
      TextBox4.Text = TextBox2.Text + (e.Y / PictureBox1.Height) * (TextBox3.Text - TextBox2.Text)
      DisplayBar(PictureBox1.CreateGraphics, e.Y)
      TextBox1.Text = lastY
   End Sub

   Private Sub PictureBox1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint
      Dim gr As Graphics
      gr = e.Graphics
      Dim wd, ht As Single
      wd = PictureBox1.Width
      ht = PictureBox1.Height
      Dim pnb, pnr As Pen
      pnb = New Pen(Drawing.Color.Blue)
      pnr = New Pen(Drawing.Color.Red)
      gr.DrawLine(pnb, wd / 2, 0, wd / 2, ht)
   End Sub

   Private Sub DisplayBar(ByVal gr As Graphics, ByVal y As Single)
      Dim pnb, pnr, pnh As Pen
      Dim wd As Single
      wd = PictureBox1.Width
      pnh = New Pen(PictureBox1.BackColor)
      pnb = New Pen(Drawing.Color.Blue)
      pnr = New Pen(Drawing.Color.Red)
      gr.DrawLine(pnh, wd / 2 - 4, lastY, wd / 2 + 4, lastY)
      gr.DrawLine(pnb, wd / 2, lastY - 1, wd / 2, lastY)
      gr.DrawLine(pnr, wd / 2 - 4, y, wd / 2 + 4, y)
      lastY = y

   End Sub

   Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
      Dim gr As Graphics
      Dim png, pnb As Pen
      png = New Pen(Drawing.Color.Green)
      pnb = New Pen(Drawing.Color.Blue)
      ' 13.58.00E, 54.35.00N
      ' 24.22.00E, 49.13.20N

      Dim fin As FileAttribute
      Dim s, sla, slo, sal As String
      gr = PictureBox2.CreateGraphics
      FileOpen(1, "C:\HAM\UI-View32\History Lists\ryp.log", OpenMode.Input, OpenAccess.Read, OpenShare.Default)
      Dim i, xn, xp, yp, x, y As Single
      xp = -1
      yp = -1
      Dim bDraw As Boolean
      While Not EOF(1)
         bDraw = False
         s = LineInput(1)
         If s.Chars(127) = "@" Then
            sla = Mid(s, 136, 7) / 100
            slo = Mid(s, 145, 8) / 100
            sla = Mid(s, 136, 2) + Mid(s, 138, 5) / 60
            slo = Mid(s, 145, 3) + Mid(s, 148, 5) / 60
            sal = Mid(s, 155, 3)
            'bDraw = True
         ElseIf s.Chars(127) = "'" Then
            Dim xi(7), yi(7) As Char
            sla = ""
            For i = 0 To 6
               xi(i) = s.Chars(132 + i)
               yi(i) = MicECharDec(s.Chars(48 + i))
               sla = sla & yi(i)
            Next
            sla = sla / 10000
            '            sla = LaLoDec(Mid(s, 129, 4)) / 380926
            '((yi(0) - 33) * (91 ^ 3) + (yi(1) - 33) * (91 ^ 2) + (yi(2) - 33) * 91 + yi(3) - 33) / 380926
            '           sla = 90 - sla
            slo = LaLoDec(Mid(s, 133, 4)) / 190463
            '((xi(0) - 33) * (91 ^ 3) + (xi(1) - 33) * (91 ^ 2) + (xi(2) - 33) * 91 + xi(3) - 33) / 190463
            slo = -180 + slo
            sal = 0
            bDraw = True
         End If
         If bDraw Then
            y = (sla - ymin) * PictureBox2.Height() / (ymax - ymin)
            x = (slo - xmin) * PictureBox2.Width() / (xmax - xmin)
            x = x
            y = y
            gr.DrawLine(png, xn, 0, xn, Int(sal) / 2)
            If (xp <> -1) Then
               gr.DrawLine(pnb, xp, PictureBox2.Height - yp, x, PictureBox2.Height - y)
            End If
            xp = x
            yp = y
            xn = xn + 2
            '            MsgBox(s & vbCrLf & sla & " - " & slo & " - " & sal)
            '            Exit While
         End If
      End While
      FileClose(1)
   End Sub
   Function LaLoEnc(ByVal deg As Double) As String

   End Function

   Function MicECharDec(ByVal v As Char) As Char
      v = UCase(v)
      If v = "0" Then
         MicECharDec = "0"
      ElseIf v = "1" Then
         MicECharDec = "1"
      ElseIf v = "2" Then
         MicECharDec = "2"
      ElseIf v = "3" Then
         MicECharDec = "3"
      ElseIf v = "4" Then
         MicECharDec = "4"
      ElseIf v = "5" Then
         MicECharDec = "5"
      ElseIf v = "6" Then
         MicECharDec = "6"
      ElseIf v = "7" Then
         MicECharDec = "7"
      ElseIf v = "8" Then
         MicECharDec = "8"
      ElseIf v = "9" Then
         MicECharDec = "9"
      ElseIf v = "A" Then
         MicECharDec = "0"
      ElseIf v = "B" Then
         MicECharDec = "1"
      ElseIf v = "C" Then
         MicECharDec = "2"
      ElseIf v = "D" Then
         MicECharDec = "3"
      ElseIf v = "E" Then
         MicECharDec = "4"
      ElseIf v = "F" Then
         MicECharDec = "5"
      ElseIf v = "G" Then
         MicECharDec = "6"
      ElseIf v = "H" Then
         MicECharDec = "7"
      ElseIf v = "I" Then
         MicECharDec = "8"
      ElseIf v = "J" Then
         MicECharDec = "9"
      ElseIf v = "K" Then
         MicECharDec = " "
      ElseIf v = "L" Then
         MicECharDec = " "
      ElseIf v = "P" Then
         MicECharDec = "0"
      ElseIf v = "Q" Then
         MicECharDec = "1"
      ElseIf v = "R" Then
         MicECharDec = "2"
      ElseIf v = "S" Then
         MicECharDec = "3"
      ElseIf v = "T" Then
         MicECharDec = "4"
      ElseIf v = "U" Then
         MicECharDec = "5"
      ElseIf v = "V" Then
         MicECharDec = "6"
      ElseIf v = "W" Then
         MicECharDec = "7"
      ElseIf v = "X" Then
         MicECharDec = "8"
      ElseIf v = "Y" Then
         MicECharDec = "9"
      ElseIf v = "Z" Then
         MicECharDec = " "
      End If
   End Function

   Function LaLoDec(ByVal cmp As String) As Double
      If Len(cmp) <> 4 Then
         LaLoDec = 0
      End If
      Dim b(4), i As Byte
      For i = 0 To 3
         b(i) = Asc(cmp.Chars(i))
      Next
      LaLoDec = ((b(0) - 33) * (91 ^ 3) + (b(1) - 33) * (91 ^ 2) + (b(2) - 33) * 91 + b(3) - 33)
   End Function

   Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
      Dim lvi As ListViewItem
      Dim fifi, ls As String
      Dim n As Integer
      '      lvi = ListView1.Items.Add("oko")
      '     lvi.SubItems(0).Text = "Olo"
      '    lvi.SubItems.Add("Ubu")
      rootpa = "C:\HAM\UI-View32\MAPS\"
      fifi = Dir(rootpa & "*.inf")
      While fifi <> ""
         lvi = ListView1.Items.Add(fifi)
         On Error Resume Next
         FileOpen(2, rootpa & fifi, OpenMode.Input, OpenAccess.Read, OpenShare.Default)
         n = 0
         ls = ""
         While Not EOF(2) And (n < 3)
            ls = LineInput(2)
            n = n + 1
         End While
         FileClose(2)
         lvi.SubItems.Add(ls)
         fifi = Dir()
      End While
   End Sub

   Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
      If ListView1.SelectedItems.Count <= 0 Then Exit Sub
      Dim path, imgpa As String
      PictureBox2.Image = Nothing

      path = rootpa & ListView1.SelectedItems(0).Text

      imgpa = Replace(path, ".inf", ".gif")
      If Dir(imgpa) = "" Then
         imgpa = Replace(path, ".inf", ".bmp")
         If Dir(imgpa) = "" Then
            imgpa = Replace(path, ".inf", ".png")
            If Dir(imgpa) = "" Then
               imgpa = Replace(path, ".inf", ".jpg")
            End If
         End If
      End If

      On Error GoTo errli
      PictureBox2.Image = Image.FromFile(imgpa)

      Dim n, m, k As Integer
      Dim ls As String

      FileOpen(2, path, OpenMode.Input, OpenAccess.Read, OpenShare.Default)
      n = 0
      k = 0
      ls = ""
      While Not EOF(2) And (n < 2)
         ls = LineInput(2)
         n = n + 1
         If (n = 1) Then
            m = InStr(ls, ",")
            If m <> 0 Then
               TextBox5.Text = Trim(Mid(ls, 1, m - 1))
               TextBox6.Text = Trim(Mid(ls, m + 1))
               k = k + 1
            Else
               TextBox5.Text = ls
               TextBox6.Text = ""
            End If
         End If
         If (n = 2) Then
            m = InStr(ls, ",")
            If m <> 0 Then
               TextBox7.Text = Trim(Mid(ls, 1, m - 1))
               TextBox8.Text = Trim(Mid(ls, m + 1))
               k = k + 2
            Else
               TextBox7.Text = ls
               TextBox8.Text = ""
            End If
         End If
      End While
      If Len(TextBox5.Text) < 10 Then TextBox5.Text = "0" & TextBox5.Text
      If Len(TextBox6.Text) < 10 Then TextBox6.Text = "0" & TextBox6.Text
      If Len(TextBox7.Text) < 10 Then TextBox7.Text = "0" & TextBox7.Text
      If Len(TextBox8.Text) < 10 Then TextBox8.Text = "0" & TextBox8.Text
      FileClose(2)
      If k = 3 Then
         Dim a5, a6, a7, a8, p As Double
         a5 = Mid(TextBox5.Text, 1, 3) + Mid(TextBox5.Text, 5, 2) / 60 + Mid(TextBox5.Text, 7, 2) / 3600
         a6 = Mid(TextBox6.Text, 1, 3) + Mid(TextBox6.Text, 5, 2) / 60 + Mid(TextBox6.Text, 7, 2) / 3600
         a7 = Mid(TextBox7.Text, 1, 3) + Mid(TextBox7.Text, 5, 2) / 60 + Mid(TextBox7.Text, 7, 2) / 3600
         a8 = Mid(TextBox8.Text, 1, 3) + Mid(TextBox8.Text, 5, 2) / 60 + Mid(TextBox8.Text, 7, 2) / 3600
         If InStr(TextBox5.Text, "N") Then ymin = a5
         If InStr(TextBox5.Text, "E") Then xmin = a5
         If InStr(TextBox6.Text, "N") Then ymin = a6
         If InStr(TextBox6.Text, "E") Then xmin = a6
         If InStr(TextBox7.Text, "N") Then ymax = a7
         If InStr(TextBox7.Text, "E") Then xmax = a7
         If InStr(TextBox8.Text, "N") Then ymax = a8
         If InStr(TextBox8.Text, "E") Then xmax = a8
         If (xmin > xmax) Then p = xmin : xmin = xmax : xmax = p
         If (ymin > ymax) Then p = ymin : ymin = ymax : ymax = p
         MsgBox("(" & xmin & ", " & xmax & ")-(" & ymin & ", " & ymax & ")")
      End If

errli:
   End Sub

   Private Sub PictureBox2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseMove
      TextBox2.Text = ymin + (PictureBox2.Height() - e.Y) / PictureBox2.Height() * (ymax - ymin)
      TextBox1.Text = xmin + e.X / PictureBox2.Width() * (xmax - xmin)

   End Sub
End Class
