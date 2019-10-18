Imports System.Math
Imports System.IO
Imports System.Numerics

Public Class Form1
    Const AlphaMask As Int32 = &HFF000000
    Const RedMask As Int32 = &HFF0000
    Const greenMask As Int32 = &HFF00
    Const blueMask As Int32 = &HFF
    '  Dim Cpu_Use As New PerformanceCounter("Processor", "% Processor Time", "_Total")
    '  Dim proc_count As Object = System.Environment.ProcessorCount
    ' Dim opti As New ParallelOptions
    Dim rec As New Rectangle(0, 0, 1, 1)
    Dim Swidth, Sheight As Int32
    Dim rng As New Random
    Dim time_stuff As New Stopwatch
    Dim picbox As New PictureBox
    Dim P_right As Boolean = True
    Dim screencenter As New Vector3 'Point
    Dim camdefault As Int32
    Dim camera As New Vector3
    Dim vertcount As Int32 = 0
    Dim polycount As Int32 = 0
    Dim MAX As Int32 = 10000000
    Dim Verts(MAX) As Vector3
    Dim polys(MAX) As Poly
    Dim normals(MAX) As Vector3
    Dim bigarray(10) As Int32
    Dim Zbuffer(10) As Int32
    Dim bkg_image(10) As Int32
    Dim trail As Int32 = 0
    Dim size_array As Int32
    Dim drawpolys As Boolean = True
    Dim modelcenter As New Vector3
    Dim tumble As Boolean = False
    Dim spinx As Boolean = False
    Dim spiny As Boolean = False
    Dim spinz As Boolean = False
    Dim bmp As Bitmap
    Dim bkg As Bitmap
    Dim spinspeed As Double
    Dim fillT As Boolean = False
    Dim light As Vector3
    Dim lightr, lightg, lightb As Int32
    Dim Obj_base_colour As Int32 = &HFF010101
    Dim light_brightness As Int32 = -8
    Dim ax, ay, az As Single 'rotatation angles
    Dim diffmult As Double = 0
    Dim mvelight As Boolean = False
    Dim metal As Boolean = False
    Dim LoadAModel As Boolean = True
    Dim filename As String = vbNullString
    Dim Lastpath As String = "x"
    Dim Lastpic As String = "x"
    Dim blendamount As Int32 = 172
    Dim pre_post As Boolean = False 'type of alpha blending to be used

    Public Structure Poly
        'the vertexes that belong to each polygon and a flag to set drawing of the polygon off or on
        Public vert1 As Int32
        Public vert2 As Int32
        Public vert3 As Int32
        Public Draw_poly As Boolean
    End Structure

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Load_Defaults() ' load custom defaults config file
        GroupBox2.Top = GroupBox1.Top
        GroupBox2.Left = GroupBox1.Left
        Make_logo()
        Nu_config()
        If Sheight > 1050 Then Label19.Left += 8
        If IO.File.Exists(CurDir() & "\logo") = True Then
            Choose(True)
        Else
            Choose(False)
        End If
    End Sub

    Private Sub Make_logo()
        Try
            Dim writer As New StreamWriter(CurDir() & "\logo")
            writer.WriteLine(My.Resources.logo)
            writer.Close()
            writer.Dispose()
        Catch oops As Exception
        End Try
    End Sub

    Public Sub Nu_config()
        Me.FormBorderStyle = FormBorderStyle.None
        Me.BackColor = Color.Black
        Swidth = Screen.PrimaryScreen.WorkingArea.Width '/ 2
        Sheight = Screen.PrimaryScreen.WorkingArea.Height '/ 2
        Me.Width = Swidth
        Me.Height = Sheight
        Panel1.Top = 0
        Panel1.Height = Sheight
        Swidth -= Panel1.Width
        If P_right = True Then Panel1.Left = Swidth Else Panel1.Left = 0
        bmp = New Bitmap(Swidth, Sheight, Imaging.PixelFormat.Format32bppRgb)
        Me.Left = 0
        Me.Top = 0
        rec.Width = Swidth
        rec.Height = Sheight
        size_array = Swidth * Sheight
        Array.Resize(bigarray, size_array)
        Array.Resize(Zbuffer, size_array)
        screencenter.X = Swidth >> 1
        screencenter.Y = Sheight >> 1
        screencenter.Z = 0
        camera.X = screencenter.X
        camera.Y = screencenter.Y
        camdefault = -2200
        camera.Z = camdefault
        picbox.Top = 0
        If P_right = True Then picbox.Left = 0 Else picbox.Left = Panel1.Width
        picbox.Width = Swidth
        picbox.Height = Sheight
        picbox.Cursor = Cursors.Hand
        picbox.WaitOnLoad = False
        picbox.Parent = Me
        AddHandler picbox.MouseMove, AddressOf Form1_MouseMove 'add mouse move handler for each picbox
        AddHandler picbox.MouseLeave, AddressOf Form1_MouseLeave
        AddHandler picbox.MouseEnter, AddressOf Form1_MouseEnter
        '  AddHandler picbox.KeyPress, AddressOf Form1_KeyPress 'keyboard
        Controls.Add(picbox)
        light.X = 0
        light.Y = 0
        light.Z = TrackBar7.Value * -1
        lightr = TrackBar3.Value
        lightg = TrackBar4.Value
        lightb = TrackBar5.Value
        '  Me.TopMost = True
        Panel1.BringToFront()
        Label23.BringToFront()
        picbox.Focus()
        Label12.Text = "Cam Z:" & ((TrackBar8.Maximum + TrackBar8.Minimum + 1 - TrackBar8.Value) * -1) * 20
        Me.Refresh()
    End Sub

    Private Sub Load_Defaults()
        Dim tempstr As String
        Dim cont As Object
        AddHandler picbox.MouseDown, AddressOf Form1_MouseDown 
        If IO.File.Exists(CurDir() & "\3dE_settings.cfg") = True Then
            Dim reader As New StreamReader(CurDir() & "\3dE_settings.cfg")
            Try
                For Each cont In Panel1.Controls
                    If TypeOf cont Is CheckBox Then
                        tempstr = reader.ReadLine
                        If tempstr = "False" Then
                            cont.checked = False
                        Else
                            cont.checked = True
                        End If
                    End If
                    If TypeOf cont Is TrackBar Then
                        tempstr = reader.ReadLine
                        cont.value = Val(tempstr)
                    End If
                Next
                light.X = CSng(Val(reader.ReadLine()))
                light.Y = CSng(Val(reader.ReadLine()))
                light.Z = CSng(Val(reader.ReadLine()))
                tempstr = reader.ReadLine
                Lastpath = reader.ReadLine '  writer.WriteLine(Lastpath)
                Lastpic = reader.ReadLine
                reader.Close()
                If tempstr = "False" Then
                    P_right = False
                Else
                    P_right = True
                End If
            Catch oops As Exception
                reader.Close()
                IO.File.Delete(CurDir() & "\3dE_settings.cfg")
                MessageBox.Show("The Configuration file was damaged or missing information and has been deleted")
            End Try
            lightr = TrackBar3.Value
            lightg = TrackBar4.Value
            lightb = TrackBar5.Value
            light_brightness = -TrackBar6.Value
        End If
        If CheckBox8.Checked = False Then CheckBox8.ForeColor = Color.Black
        Label12.Enabled = CheckBox6.Checked
        If CheckBox10.Checked = True Then CheckBox10.Text = "C" Else CheckBox10.Text = "c"
        If metal = True Then Button30.Text = "S" Else Button30.Text = "s"
        CheckBox11.Checked = False
        If IO.Directory.Exists(Lastpath) = False Then Lastpath = "x"
        If IO.Directory.Exists(Lastpic) = False Then Lastpic = "x"
    End Sub

    Public Sub Choose(ByVal logo As Boolean) ' a file to use
        If logo = False Then
            Dim filt As String
            If LoadAModel = True Then
                filt = "3D Models|*.obj;*.stl"
                If Lastpath.Length > 1 Then FileIO.FileSystem.CurrentDirectory = Lastpath
            Else
                filt = "Image Files|*.bmp;*.jpg;*.png"
                If Lastpic.Length > 1 Then FileIO.FileSystem.CurrentDirectory = Lastpic
            End If
            Dim import As New OpenFileDialog With {
            .Multiselect = False,
            .Filter = filt,
            .CheckFileExists = True,
            .AddExtension = True,
            .InitialDirectory = CurDir(),
            .CheckPathExists = True,
            .SupportMultiDottedExtensions = True,
            .AutoUpgradeEnabled = True
              }
            If LoadAModel = True Then
                import.ShowDialog()
                filename = import.FileName
                If filename.Length > 1 Then
                    Label1.Text = "importing"
                    If filename.Substring(filename.Length - 1) = "J" Or filename.Substring(filename.Length - 1) = "j" Then
                        Read_obj(filename)
                        Label2.Text = "Model:" & vbNewLine & import.SafeFileName
                    Else
                        Read_STL(filename)
                        Label2.Text = "Model:" & vbNewLine & import.SafeFileName
                    End If
                    Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
                Else
                    Exit Sub
                End If
                Lastpath = Mid(filename, 1, filename.Length - import.SafeFileName.Length)
            Else
                import.ShowDialog()
                If import.FileName.Length < 2 Then
                    filename = "xx"
                Else
                    filename = import.FileName
                    bkg = New Bitmap(filename)
                    Lastpic = Mid(filename, 1, filename.Length - import.SafeFileName.Length)
                End If
            End If
            LoadAModel = True
            import.Dispose()
        Else
            Read_obj(CurDir() & "\logo")
        End If
    End Sub

    Private Sub Clear_array()
        If CheckBox11.Checked = False Then
            Parallel.For(0, bigarray.Length, Sub(f As Int32)
                                                 If bigarray(f) <> &HFF000000 Then bigarray(f) = &HFF000000 'just clear the array 
                                                 If Zbuffer(f) < &H6F000000 Then Zbuffer(f) = &H6F000000 ' reset the zbuffer
                                             End Sub)
        Else
            Parallel.For(0, bigarray.Length, Sub(f As Int32)
                                                 bigarray(f) = bkg_image(f) 'copy data from background image 
                                                 Zbuffer(f) = &H6F000000
                                             End Sub)
        End If
    End Sub

    Private Sub Make_Normal(indx As Int32, vec0 As Vector3, vec1 As Vector3, vec2 As Vector3) ' not used
        polys(indx).Draw_poly = True
        Dim t As Vector3 = Vector3.Subtract(vec1, vec0)
        Dim t1 As Vector3 = Vector3.Subtract(vec2, vec0)
        normals(indx) = Vector3.Cross(t, t1)
        normals(indx) = Vector3.Normalize(normals(indx))
    End Sub

    Private Sub Nu_rasterPoly()
        diffmult = TrackBar9.Value * 0.01
        If diffmult > 0.99 Then diffmult = 0.99
        ' for each polygon (triangle) locate the 3 vetex positions,rotate them, generate a normal and map the vertexes to 2d plus perspective. Then draw the wireframe and or fill the triangles. 
        Parallel.For(0, polycount, Sub(f As Int32)
                                       Dim DrawAline As Boolean = CheckBox1.Checked
                                       Dim Lit_lines As Boolean = CheckBox4.Checked
                                       Dim colour As Int32 = 0
                                       'copy triangle vertexes to 3 temporary vectors for transforms etc.
                                       Dim vec(3) As Vector3
                                       vec(0) = Verts(polys(f).vert1)
                                       vec(1) = Verts(polys(f).vert2)
                                       vec(2) = Verts(polys(f).vert3)
                                       'rotate vertexes 
                                       Dim inputQ As Quaternion
                                       Dim resultQ As Quaternion
                                       For i As Int32 = 0 To 2
                                           vec(i) -= modelcenter
                                           inputQ = Quaternion.CreateFromYawPitchRoll(ay, ax, az) 'rotation angles converted to radians
                                           resultQ = (inputQ * New Quaternion(vec(i), 0)) * Quaternion.Conjugate(inputQ) ' rotation
                                           vec(i).X = resultQ.X
                                           vec(i).Y = resultQ.Y
                                           vec(i).Z = resultQ.Z
                                           vec(i) += modelcenter
                                       Next
                                       '   Make_Normal(f, vec(0), vec(1), vec(2)) 'generate the face normal
                                       polys(f).Draw_poly = True 'generate the face normal
                                       Dim t As Vector3 = Vector3.Subtract(vec(1), vec(0))
                                       Dim t1 As Vector3 = Vector3.Subtract(vec(2), vec(0))
                                       normals(f) = Vector3.Cross(t, t1)
                                       normals(f) = Vector3.Normalize(normals(f))
                                       Dim vtemp As Vector3 = normals(f)
                                       vtemp.Z = vtemp.Z - camera.Z
                                       Dim cp As Single = Vector3.Dot(vtemp, normals(f))
                                       If CheckBox5.Checked = True Then
                                           If cp >= 0 Then polys(f).Draw_poly = False  'back face cull
                                       Else
                                           If cp > 0 Then normals(f) *= -1 ' flips normal so the triangle is always faceing the camerea creating a "double sided" polygon
                                       End If
                                           If polys(f).Draw_poly = True Then
                                           If CheckBox6.Checked = True Then
                                               Dim cam As Int32 = CInt(camera.Z + 200) 'z clip distance (stops frame rate dropping to silly levels when geomatry starts get to close to the camera plane)
                                               For n As Int32 = 0 To 2 '   perspective transform(3D > 2D mapping) 
                                                   vec(n).X = camera.Z * (vec(n).X - camera.X)
                                                   vec(n).X /= (camera.Z - vec(n).Z)
                                                   vec(n).X += camera.X
                                                   vec(n).Y = camera.Z * (vec(n).Y - camera.Y)
                                                   vec(n).Y /= (camera.Z - vec(n).Z)
                                                   vec(n).Y += camera.Y
                                                   If cam > vec(n).Z = True Then
                                                       If Abs(vec(n).X) > (Swidth << 1) Or Abs(vec(n).Y) > (Sheight << 1) Then polys(f).Draw_poly = False '
                                                       If cam + 2 > (vec(n).Z) Then polys(f).Draw_poly = False '    clip z if it comes too close to the camera in the z plane.
                                                   End If
                                               Next
                                           End If
                                       End If
                                       'Vertex_sort by y axis lowest first
                                       If polys(f).Draw_poly = True Then
                                           If vec(2).Y < vec(0).Y Then
                                               vec(3) = vec(0)
                                               vec(0) = vec(2)
                                               vec(2) = vec(3)
                                           End If
                                           If vec(2).Y < vec(1).Y Then
                                               vec(3) = vec(1)
                                               vec(1) = vec(2)
                                               vec(2) = vec(3)
                                           End If
                                           If vec(1).Y < vec(0).Y Then
                                               vec(3) = vec(1)
                                               vec(1) = vec(0)
                                               vec(0) = vec(3)
                                           End If
                                       End If
                                       'if entire triangle is off screen then don't draw it. 
                                       If vec(2).Y < 0 OrElse vec(0).Y > Sheight Then polys(f).Draw_poly = False
                                       If vec(0).X > Swidth AndAlso vec(1).X > Swidth AndAlso vec(2).X > Swidth Then polys(f).Draw_poly = False
                                       If vec(0).X < 0 AndAlso vec(1).X < 0 AndAlso vec(2).X < 0 Then polys(f).Draw_poly = False
                                       If polys(f).Draw_poly = True Then
                                           'draw the Vertexes
                                           If CheckBox2.Checked = True Then
                                               For n As Int32 = 0 To 2
                                                   Dim Draw_vertex As Boolean = True
                                                   If vec(n).X < 0 OrElse vec(n).X + 1 >= Swidth Then Draw_vertex = False
                                                   If vec(n).Y < 0 OrElse vec(n).Y >= Sheight Then Draw_vertex = False
                                                   If Draw_vertex = True Then
                                                       'add a vertex marker into bigarray & the zbuffer
                                                       Dim loc As Int32 = CInt(vec(n).X) + (CInt(vec(n).Y) * Swidth) 'pixel location to array position
                                                       If loc < bigarray.Length - 1 Then
                                                           If vec(n).Z < Zbuffer(loc) Then
                                                               bigarray(loc) = &HFFFBE0C 'orangey yellow
                                                               Zbuffer(loc) = CInt(vec(n).Z)
                                                           End If
                                                       End If
                                                   End If
                                               Next
                                           End If
                                           'Pick a colour scheme & Draw the wire-frame
                                           If CheckBox10.Checked = True OrElse Lit_lines = False Then
                                               Dim Rcolour As Int32 = CInt(Abs(modelcenter.Z - vec(0).Z) * 0.18) Mod 255
                                               Dim Gcolour As Int32 = CInt(Abs(modelcenter.X - vec(0).X) * 0.179) Mod 255
                                               Dim Bcolour As Int32 = CInt(Abs(modelcenter.Y - vec(0).Y) * 0.181) Mod 255
                                               colour = AlphaMask + (Rcolour << 16) + (Gcolour << 8) + Bcolour
                                           End If
                                           If DrawAline = True Then
                                               If Lit_lines = True Then Lighting(normals(f), colour)
                                               Dim tcol As Int32 = colour
                                               If fillT = True Then colour = Not (&H66000000 Xor colour) 'negative colour for wireframe (toned down brightness)
                                               For i As Int32 = 0 To 2
                                                   DrawLine3d((CInt(vec((i + 1) Mod 3).X)), CInt((vec((i + 1) Mod 3).Y)), CInt((vec((i + 1) Mod 3).Z)), CInt(vec(i).X), CInt(vec(i).Y), CInt(vec(i).Z), colour)
                                               Next
                                               colour = tcol
                                           End If
                                           'Pick a colour scheme & Fill the triangle
                                           If fillT = True Then
                                               If DrawAline = True And Lit_lines = False Then
                                                   colour = Obj_base_colour
                                               Else
                                                   'Pointlight(normals(f), colour) 'lighting
                                                   Lighting(normals(f), colour) 'lighting
                                               End If
                                               If DrawAline = False AndAlso Lit_lines = False Then  '  non-lit polys
                                                   Dim test As Int32 = f >> 1 << 1
                                                   If test = f Then
                                                       colour = &H7F808080
                                                   Else
                                                       colour = &H7F401090
                                                   End If
                                               End If
                                               'force Y axis values of the vectors into int32 
                                               For n As Int32 = 0 To 2
                                                   vec(n).Y = CInt(vec(n).Y)
                                               Next
                                               'workout the triangle type and draw. 
                                               If vec(1).Y = vec(2).Y Then
                                                   Flatbottom(vec(0), vec(1), vec(2), colour)
                                               ElseIf vec(0).Y = vec(1).Y Then
                                                   Flattop(vec(0), vec(1), vec(2), colour)
                                               Else
                                                   vec(3).Y = (vec(1).Y - vec(0).Y) / (vec(2).Y - vec(0).Y) 'generate new temp vertex to split the triangle into 2
                                                   vec(3) = vec(0) + (vec(2) - vec(0)) * vec(3).Y
                                                   'If vec(1).X < vec(3).X Then
                                                   Flatbottom(vec(0), vec(1), vec(3), colour) 'generate a triangle from the top 3 vectors
                                                   Flattop(vec(1), vec(3), vec(2), colour) 'generate a  triangle from the bottom 3 vectors
                                                   ' Else
                                                   'Flatbottom(vec(0), vec(3), vec(1), colour) 'generate a lefthand triangle from the top 3 vectors 
                                                   'Flattop(vec(3), vec(1), vec(2), colour) 'generate a lefthand triangle from the bottom 3 vectors
                                                   'End If
                                               End If
                                           End If
                                       End If
                                   End Sub)
    End Sub

    Private Sub Flatbottom(ByVal vec0 As Vector3, ByVal vec1 As Vector3, ByVal vec2 As Vector3, ByVal colour As Int32)
        'this traces the two lines down the sides of a flat bottomed triangle in paralell generating x,y,z coords for the zbuffer and the two end points of the vertical line connecting them...
        'it then calls drawx which generates the points between the vertical line end points passed to it.
        Dim l1, l2 As Vector3
        For scanline As Int32 = vec0.Y + 1 To vec2.Y
            l1 = Vector3.Lerp(vec1, vec0, (scanline - vec2.Y) / (vec0.Y - vec1.Y))
            l2 = Vector3.Lerp(vec2, vec0, (scanline - vec2.Y) / (vec0.Y - vec2.Y))
            DrawX(l1, l2, colour)
        Next
    End Sub

    Private Sub Flattop(ByVal vec0 As Vector3, ByVal vec1 As Vector3, ByVal vec2 As Vector3, ByVal colour As Int32)
        'this traces the two lines down the sides of a flat topped triangle in paralell generating x,y,z coords for the zbuffer and the two end points of the vertical line connecting them...
        'it then calls drawx which generates the points between the vertical line end points passed to it.
        Dim l1, l2 As Vector3
        For scanline As Int32 = vec0.Y To vec2.Y - 1
            l1 = Vector3.Lerp(vec1, vec2, (scanline - vec2.Y) / (vec1.Y - vec2.Y))
            l2 = Vector3.Lerp(vec0, vec2, (scanline - vec2.Y) / (vec0.Y - vec2.Y))
            DrawX(l1, l2, colour)
        Next
    End Sub

    Private Sub DrawX(ByVal l1 As Vector3, ByVal l2 As Vector3, ByVal colour As Int32) 'draws a line along the x axis generating z axis coordinates for the zbuffer. (with or without alpha blending)
        If l2.X < l1.X Then
            Dim vline As Vector3 = l1
            l1 = l2
            l2 = vline
        End If
        l1.X = CInt(l1.X)
        l2.X = CInt(l2.X)
        If l2.X > l1.X Then
            Dim zslope As Single = (l1.Z - l2.Z - 0.5) / (l1.X + 1 - l2.X - 0.5)
            Dim zpos As Single = l1.Z
            Dim loc As Int32
            Dim scrnY As Int32 = (l2.Y * Swidth)
            If CheckBox12.Checked = True AndAlso pre_post = False Then  'alpha blending (uses the zbuffer to adjust polygon transparency - dirty hack)
                For n As Int32 = l1.X + 1 To l2.X
                    If n >= 0 AndAlso n < Swidth Then
                        Loc = n + scrnY
                        If loc < size_array AndAlso loc >= 0 Then
                            Dim bl As Int32 = blendamount
                            If CInt(zpos) < Zbuffer(loc) Then
                                bl -= 16
                                Zbuffer(loc) = CInt(zpos)
                            Else
                                bl = 255 - (blendamount >> 3)
                            End If
                            Dim rb As Int32 = colour And &HFF00FF
                            Dim g As Int32 = colour And &HFF00
                            rb += ((bigarray(loc) And &HFF00FF) - rb) * (bl) >> 8
                            g += ((bigarray(loc) And &HFF00) - g) * (bl) >> 8
                            bigarray(loc) = (rb And &HFF00FF) Or (g And &HFF00)
                        End If
                    End If
                    zpos += zslope
                Next n
            Else ' Not alpha blending
                For n As Int32 = l1.X + 1 To l2.X
                    If n >= 0 AndAlso n < Swidth Then
                        loc = n + scrnY
                        If loc < size_array AndAlso loc >= 0 Then
                            If CInt(zpos) < Zbuffer(loc) Then
                                bigarray(loc) = colour
                                Zbuffer(loc) = CInt(zpos)
                            End If
                        End If
                    End If
                    zpos += zslope
                Next n
            End If
        End If
    End Sub

    Private Sub AlphaBlend() ' post processing version of alpha blend ----will always obscure internal polygons
        If ToolStripMenuItem2.Checked = False Then
            Parallel.For(0, Zbuffer.Length - 1, Sub(f As Int32)
                                                    If Zbuffer(f) < &H6F000000 Then
                                                        Dim bkg As Int32
                                                        If CheckBox11.Checked = True Then bkg = bkg_image(f) Else bkg = &HFF000000
                                                        'fast full alpha blend.
                                                        Dim rb As Int32 = bigarray(f) And &HFF00FF
                                                        Dim g As Int32 = bigarray(f) And &HFF00
                                                        rb += ((bkg And &HFF00FF) - rb) * blendamount >> 8
                                                        g += ((bkg And &HFF00) - g) * blendamount >> 8
                                                        bigarray(f) = (rb And &HFF00FF) Or (g And &HFF00)
                                                    End If
                                                End Sub)
        End If
    End Sub

    Private Sub DrawLine3d(ByVal x As Int32, ByVal y As Int32, ByVal z As Int32, ByVal ex As Int32, ByVal ey As Int32, ByVal ez As Int32, ByVal col As Int32)
        'https://www.geeksforgeeks.org/bresenhams-algorithm-for-3-d-line-drawing/
        If x < 0 And ex < 0 Then Exit Sub
        If ey < 0 And y < 0 Then Exit Sub
        If x > Swidth And ex > Swidth Then Exit Sub
        If ey > Sheight And y > Sheight Then Exit Sub
        Dim dx As Int32 = Abs(ex - x)
        Dim dy As Int32 = Abs(ey - y)
        Dim dz As Int32 = Abs(ez - z)
        '  If dz < camera.Z + 5 Then Exit Sub
        Dim xs, ys, zs As Int32
        Dim p1, p2 As Int32
        Dim loc As Int32 = 0
        If ex > x Then xs = 1 Else xs = -1
        If ey > y Then ys = 1 Else ys = -1
        If ez > z Then zs = 1 Else zs = -1
        'first pixel in the line
        If x >= 0 AndAlso x < Swidth Then
            loc = x + (y * Swidth)
            If loc < size_array And loc >= 0 Then
                If z < Zbuffer(loc) Then
                    bigarray(loc) = col
                    Zbuffer(loc) = z - 1
                End If
            End If
        End If
        '.....
        If dx >= dy And dx >= dz Then
            p1 = ((dy + dy) - dx)
            p2 = ((dz + dz) - dx)
            Do While x <> ex
                x += xs
                If p1 >= 0 Then
                    y += ys
                    p1 -= (dx + dx)
                End If
                If p2 >= 0 Then
                    z += zs
                    p2 -= (dx + dx)
                End If
                p1 += dy + dy
                p2 += dz + dz
                If x >= 0 AndAlso x < Swidth Then
                    loc = x + (y * Swidth)
                    If loc < size_array And loc >= 0 Then
                        If z < Zbuffer(loc) Then
                            bigarray(loc) = col
                            Zbuffer(loc) = z - 1
                        End If
                    End If
                End If
            Loop
        ElseIf (dy >= dx And dy >= dz) Then
            p1 = ((dx + dx) - dy)
            p2 = ((dz + dz) - dy)
            Do While y <> ey
                y += ys
                If p1 >= 0 Then
                    x += xs
                    p1 -= (dy + dy)
                End If
                If p2 >= 0 Then
                    z += zs
                    p2 -= (dy + dy)
                End If
                p1 += (dx + dx)
                p2 += (dz + dz)
                If x >= 0 AndAlso x < Swidth Then
                    loc = x + (y * Swidth)
                    If loc < size_array And loc >= 0 Then
                        If z < Zbuffer(loc) Then
                            bigarray(loc) = col
                            Zbuffer(loc) = z - 1
                        End If
                    End If
                End If
            Loop
        Else
            p1 = (dy + dy) - dz
            p2 = (dx + dx) - dz
            Do While z <> ez
                z += zs
                If p1 > 0 Then
                    y += ys
                    p1 -= (dz + dz)
                End If
                If p2 >= 0 Then
                    x += xs
                    p2 -= (dz + dz)
                End If
                p1 += (dy + dy)
                p2 += (dx + dx)
                If x >= 0 AndAlso x < Swidth Then
                    loc = x + (y * Swidth)
                    If loc < size_array And loc >= 0 Then
                        If z < Zbuffer(loc) Then
                            bigarray(loc) = col
                            Zbuffer(loc) = z - 1
                        End If
                    End If
                End If
            Loop
        End If
    End Sub

    Private Sub Lighting(ByVal norm As Vector3, ByRef colour As Int32)
        Dim rdot As Int32
        Dim gdot As Int32
        Dim bdot As Int32
        If CheckBox10.Checked = True Then
            rdot = (colour And RedMask) >> 16
            gdot = (colour And greenMask) >> 8
            bdot = (colour And blueMask)
        Else
            rdot = (Obj_base_colour And RedMask) >> 16
            gdot = (Obj_base_colour And greenMask) >> 8
            bdot = (Obj_base_colour And blueMask)
        End If
        Dim LightDirection As Vector3 = Vector3.Subtract(light, norm)
        LightDirection = Vector3.Normalize(LightDirection)
        Dim diff As Double = Vector3.Dot(norm, LightDirection)
        If diff > diffmult Then
            diff *= (1 + diffmult)
            If metal = True Then diff += diff * 0.178
        Else
            If metal = True Then diff -= diff * 0.141
            ' If diff < 0 Then diff = 0
        End If
        rdot = CInt(rdot - light_brightness + ((diff * lightr)))
        gdot = CInt(gdot - light_brightness + ((diff * lightg)))
        bdot = CInt(bdot - light_brightness + ((diff * lightb)))
        If rdot > 255 Then rdot = 255
        If gdot > 255 Then gdot = 255
        If bdot > 255 Then bdot = 255
        If rdot < 0 Then rdot = 0
        If gdot < 0 Then gdot = 0
        If bdot < 0 Then bdot = 0
        colour = AlphaMask + (rdot << 16) + (gdot << 8) + bdot
    End Sub

    Private Sub Make_bkg()
        LoadAModel = False
        Choose(False)
        If filename = "xx" Then Exit Sub
        Try
            Array.Resize(Of Int32)(bkg_image, Swidth * Sheight)
            Dim Bkground As Bitmap 'bitmap to hold the background image
            Bkground = New Bitmap(Swidth, Sheight, Imaging.PixelFormat.Format32bppArgb)
            Dim g As Graphics = Graphics.FromImage(Bkground)
            g.DrawImage(bkg, 0, 0, Swidth, Sheight)
            Dim bkFinishedFramedata As System.Drawing.Imaging.BitmapData
            Dim bpixelcount As Int32
            bkFinishedFramedata = Bkground.LockBits(rec, Drawing.Imaging.ImageLockMode.ReadOnly, Bkground.PixelFormat)
            bpixelcount = (bkFinishedFramedata.Stride * Sheight) \ 4
            System.Runtime.InteropServices.Marshal.Copy(bkFinishedFramedata.Scan0, bkg_image, 0, bpixelcount)
            Bkground.UnlockBits(bkFinishedFramedata)
            ''free up resources/tidy
            Bkground.Dispose()
            g.Dispose()
        Catch oops As Exception
            MsgBox("Could not load the image try another")
            Exit Sub
        End Try
        CheckBox11.Enabled = True
        CheckBox11.Checked = True
        Makebmp()
    End Sub

    Private Sub Drawgrid()
        Dim zbool As Boolean = ToolStripMenuItem2.Checked
        Dim gsize As Int32 = 30
        For j As Int32 = 0 To Sheight - 1 Step gsize
            For i As Int32 = 0 To Swidth - 1 Step 3
                Dim loc As Int32 = i + (j * Swidth)
                If zbool = False Then bigarray(loc) = &H1444444 Else Zbuffer(loc) = &H6EFFFFFF
            Next
        Next
        For j As Int32 = 0 To Swidth - 1 Step gsize
            For i As Int32 = 0 To Sheight - 1 Step 3
                Dim loc As Int32 = j + (i * Swidth)
                If zbool = False Then bigarray(loc) = &H1444444 Else Zbuffer(loc) = &H6EFFFFFF
            Next
        Next
    End Sub

    Private Sub Makebmp() ' generate the image on screen
        If polycount > 1 Then
            time_stuff.Start()
            Clear_array()
            If mvelight = True Then
                light.X = -(screencenter.X - MousePosition.X)
                light.Y = -(screencenter.Y - MousePosition.Y)
            End If
            If CheckBox3.Checked = True Then Drawgrid()
            If fillT = True OrElse CheckBox1.Checked = True OrElse CheckBox2.Checked = True Then Nu_rasterPoly()
            If bigarray.Length = size_array Then
                If ToolStripMenuItem2.Checked = True Then
                    Parallel.For(0, Zbuffer.Length - 1, Sub(f As Int32) 'overwrite bigarray with a visualisation of the zbuffer
                                                            Dim zb As Int32 = 127 + CInt((Zbuffer(f) * 0.125)) Mod 256
                                                            If Zbuffer(f) = &H6F000000 Then zb = &HFF000000
                                                            bigarray(f) = &H6F000000 + ((zb << 16) + (zb << 8) + zb)
                                                        End Sub)
                End If
                If CheckBox12.Checked = True AndAlso pre_post = True Then AlphaBlend()
                Dim bmpdata As Imaging.BitmapData = bmp.LockBits(rec, Drawing.Imaging.ImageLockMode.WriteOnly, bmp.PixelFormat)
                Dim ptr As IntPtr = bmpdata.Scan0
                System.Runtime.InteropServices.Marshal.Copy(bigarray, 0, ptr, size_array)
                bmp.UnlockBits(bmpdata)
                picbox.Image = bmp
            End If
            time_stuff.Stop()
            Label3.Text = time_stuff.Elapsed.TotalMilliseconds.ToString("f2") & " ms " & "(" & (1000 / time_stuff.Elapsed.TotalMilliseconds).ToString("f2") & " fps)"
            If CheckBox4.Checked = True Then
                Label19.Text = "X:" & light.X.ToString("n0") & vbNewLine & "Y:" & light.Y.ToString("n0") & vbNewLine & "Z:" & -light.Z
            Else
                Label19.Text = vbNullString
            End If
            'If time_stuff.Elapsed.TotalMilliseconds > 2000 Then
            '    CheckBox1.Checked = False
            '    CheckBox2.Checked = True
            '    CheckBox7.Checked = False
            '    Application.DoEvents()
            'End If
            time_stuff.Reset()
        End If
    End Sub

    Private Sub Magnify(ByVal zoom As Single)
        Parallel.For(1, vertcount, Sub(f As Int32)
                                       Verts(f) = Vector3.Subtract(Verts(f), modelcenter)
                                       Verts(f) = Vector3.Multiply(Verts(f), zoom)
                                       Verts(f) = Vector3.Add(Verts(f), modelcenter)
                                   End Sub)
    End Sub
    Private Sub Center()
        For n As Int32 = 0 To 1
            Calc_centerpoint()
            Dim offset As Vector3
            For f As Int32 = 1 To vertcount
                offset = Vector3.Subtract(Verts(f), modelcenter)
                Verts(f) = Vector3.Subtract(screencenter, offset)
            Next
        Next
        Calc_centerpoint()
        Makebmp()
    End Sub
    Private Sub Calc_centerpoint()
        modelcenter.X = 0
        modelcenter.Y = 0
        modelcenter.Z = 0
        For counts As Int32 = 1 To vertcount
            modelcenter = Vector3.Add(modelcenter, Verts(counts))
        Next
        modelcenter = Vector3.Divide(modelcenter, vertcount)
    End Sub

    Private Sub Write_obj() 'export file
        Me.Cursor = Cursors.WaitCursor
        Dim desktp As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim writer As New StreamWriter(desktp & "\export.obj")
        Dim tempstring As String = vbNullString
        For f As Int32 = 1 To vertcount - 1
            tempstring = "v " & Verts(f).X & " " & Verts(f).Y & " " & Verts(f).Z
            writer.WriteLine(tempstring)
        Next
        For f As Int32 = 0 To polycount - 1
            tempstring = "f " & polys(f).vert1 & " " & polys(f).vert2 & " " & polys(f).vert3
            writer.WriteLine(tempstring)
        Next
        writer.Close()
        writer.Dispose()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Read_obj(ByVal fn As String) 'basic .obj parser
        Dim reader As New StreamReader(fn)
        Dim st As String = vbNullString
        Dim st_st As String = vbNullString
        Dim Axis_Track As Int32 = 1
        Dim counts As Int32 = 1
        Dim counts2 As Int32 = 0
        Dim Tcounts As Int32 = 1
        Dim Tcounts2 As Int32 = 0
        Dim num(3) As Single
        Do While reader.EndOfStream = False
            st = reader.ReadLine
            st = st.Replace("  ", " ")
            st = st.Replace("//", "/")
            st_st = Mid(st, 1, 2)
            'if the line does not start with vspace or fspace ignore it and move on
            Select Case st_st
                Case Is = "v " 'load vertex data
                    Dim startpos As Int32 = 3
                    Dim space As Int32
                    For f As Int32 = 0 To 1
                        space = InStr(startpos, st, " ")
                        If space > 0 Then num(f) = CSng(Val(Mid(st, startpos, space - startpos)))
                        startpos = space + 1
                    Next
                    num(2) = CSng(Val(Mid(st, startpos)))
                    Verts(counts).X = num(0) * -1 + screencenter.X
                    Verts(counts).Y = num(1) * -1 + screencenter.Y
                    Verts(counts).Z = num(2)
                    counts += 1
                    Tcounts += 1
                    If counts = MAX Then
                        MessageBox.Show("Maximum Vertex limit Reached" & vbNewLine & counts & vbNewLine & "Import canceled")
                        Exit Sub
                    End If
                    If Tcounts = 100000 Then
                        Tcounts = 0
                        Label1.Text = "Vertexes imported " & counts.ToString("n0") & vbNewLine & "Polygons imported " & counts2.ToString("n0")
                        Application.DoEvents()
                    End If
                Case Is = "f " 'load face data (expects traingles or quads only)
                    Dim startp As Int32 = 3
                    Dim space As Int32 = 0
                    Dim slash As Int32 = 0
                    slash = InStr(startp, st, "/")
                    If slash = 0 Then
                        'no slashes
                        For f As Int32 = 0 To 3
                            space = InStr(startp, st, " ")
                            If space > 0 Then
                                num(f) = CSng(Val(Mid(st, startp, space - startp)))
                            Else
                                num(f) = CSng(Val(Mid(st, startp)))
                            End If
                            startp = space + 1
                        Next
                        If num(3) = 0 Then
                            polys(counts2).vert1 = CInt(num(0))
                            polys(counts2).vert2 = CInt(num(1))
                            polys(counts2).vert3 = CInt(num(2))
                            polys(counts2).Draw_poly = True
                        Else ' convert 4 point poly into 2x 3 point poly's
                            polys(counts2).vert1 = CInt(num(0))
                            polys(counts2).vert2 = CInt(num(1))
                            polys(counts2).vert3 = CInt(num(2))
                            counts2 += 1
                            polys(counts2).vert1 = CInt(num(0))
                            polys(counts2).vert2 = CInt(num(2))
                            polys(counts2).vert3 = CInt(num(3))
                        End If
                    Else
                        'found slashes
                        For f As Int32 = 0 To 3
                            space = InStr(startp, st, " ")
                            slash = InStr(startp, st, "/")
                            If space > 0 Then
                                num(f) = CSng(Val(Mid(st, startp, slash - startp)))
                            Else
                                num(f) = CSng(Val(Mid(st, startp)))
                            End If
                            startp = space + 1
                        Next
                        If num(3) = 0 Then
                            polys(counts2).vert1 = CInt(num(0))
                            polys(counts2).vert2 = CInt(num(1))
                            polys(counts2).vert3 = CInt(num(2))
                        Else ' convert 4 point poly into 2x 3 point poly's
                            polys(counts2).vert1 = CInt(num(0))
                            polys(counts2).vert2 = CInt(num(1))
                            polys(counts2).vert3 = CInt(num(2))
                            counts2 += 1
                            Tcounts2 += 1
                            polys(counts2).vert1 = CInt(num(0))
                            polys(counts2).vert2 = CInt(num(2))
                            polys(counts2).vert3 = CInt(num(3))
                        End If
                    End If
                    counts2 += 1
                    Tcounts2 += 1
                    If counts2 = MAX Then
                        MessageBox.Show("Maximum Polygon limit Reached" & vbNewLine & vbNewLine & "Import canceled")
                        Exit Sub
                    End If
                    If Tcounts2 = 100000 Then
                        Tcounts2 = 0
                        Label1.Text = "Vertexes imported " & counts.ToString("n0") & vbNewLine & "Polygons imported " & counts2.ToString("n0")
                        Application.DoEvents()
                    End If
            End Select
        Loop
        reader.Close()
        reader.Dispose()
        Label1.Text = "Vertexes imported " & counts.ToString("n0") & vbNewLine & "Polygons imported " & counts2.ToString("n0")
        polycount = counts2
        vertcount = counts
        'Scale the model to a reasonable size based on screen height
        Firstscale()
    End Sub

    Private Sub Firstscale()
        Dim miny As Single
        Dim maxy As Single
        miny = 99999999999
        maxy = 0
        For ff As Int32 = 0 To 2
            For f As Int32 = 1 To vertcount - 1
                If Verts(f).Y > maxy Then maxy = Verts(f).Y
                If Verts(f).Y < miny Then miny = Verts(f).Y
            Next
            If maxy - modelcenter.Y < Sheight Then
                If maxy - modelcenter.Y < 10 Then
                    Magnify(15 - (maxy - modelcenter.Y))
                Else
                    Magnify(4.5)
                End If
            ElseIf maxy - modelcenter.Y > Sheight Then
                Magnify(0.18)
            Else
                Center()
                Exit For
            End If
            Center()
        Next ff
    End Sub

    Private Sub Write_Stl()
        Me.Cursor = Cursors.WaitCursor
        Dim tmp(100) As Byte
        Dim desktp As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim OutputFile As Stream = IO.File.Open(desktp & "\export.stl", IO.FileMode.Create)
        Dim Header As String = " PhilG's Model Viewer "
        Dim bytes(80) As Byte
        tmp = System.Text.Encoding.Unicode.GetBytes(Header)
        Array.Copy(tmp, bytes, tmp.Length - 1)
        OutputFile.Write(bytes, 0, 80) 'write header
        bytes = BitConverter.GetBytes(polycount)
        OutputFile.Write(bytes, 0, 4) ' tri count
        Dim countt As Int32 = 0
        Dim tempbytes(4) As Byte
        Array.Resize(Of Byte)(bytes, 50)
        For f As Int32 = 0 To polycount - 1
            tempbytes = BitConverter.GetBytes(normals(f).X) ' ....write data (normal xyz/vertex(1,2,3) x,y,z/2 zero byte values.
            bytes(0) = tempbytes(0)
            bytes(1) = tempbytes(1)
            bytes(2) = tempbytes(2)
            bytes(3) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(normals(f).Y)
            bytes(4) = tempbytes(0)
            bytes(5) = tempbytes(1)
            bytes(6) = tempbytes(2)
            bytes(7) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(normals(f).Z)
            bytes(8) = tempbytes(0)
            bytes(9) = tempbytes(1)
            bytes(10) = tempbytes(2)
            bytes(11) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert1).X)
            bytes(12) = tempbytes(0)
            bytes(13) = tempbytes(1)
            bytes(14) = tempbytes(2)
            bytes(15) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert1).Y)
            bytes(16) = tempbytes(0)
            bytes(17) = tempbytes(1)
            bytes(18) = tempbytes(2)
            bytes(19) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert1).Z)
            bytes(20) = tempbytes(0)
            bytes(21) = tempbytes(1)
            bytes(22) = tempbytes(2)
            bytes(23) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert2).X)
            bytes(24) = tempbytes(0)
            bytes(25) = tempbytes(1)
            bytes(26) = tempbytes(2)
            bytes(27) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert2).Y)
            bytes(28) = tempbytes(0)
            bytes(29) = tempbytes(1)
            bytes(30) = tempbytes(2)
            bytes(31) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert2).Z)
            bytes(32) = tempbytes(0)
            bytes(33) = tempbytes(1)
            bytes(34) = tempbytes(2)
            bytes(35) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert3).X)
            bytes(36) = tempbytes(0)
            bytes(37) = tempbytes(1)
            bytes(38) = tempbytes(2)
            bytes(39) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert3).Y)
            bytes(40) = tempbytes(0)
            bytes(41) = tempbytes(1)
            bytes(42) = tempbytes(2)
            bytes(43) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(polys(f).vert3).Z)
            bytes(44) = tempbytes(0)
            bytes(45) = tempbytes(1)
            bytes(46) = tempbytes(2)
            bytes(47) = tempbytes(3)
            bytes(48) = 0
            bytes(49) = 0
            OutputFile.Write(bytes, 0, 50)
        Next
        OutputFile.Close()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Read_STL(ByVal fn As String) '(binary stl files)
        Me.Cursor = Cursors.WaitCursor
        Dim buffersize As Int32 = 80
        Dim bytes(buffersize) As Byte
        Dim inputFile As Stream = IO.File.Open(fn, IO.FileMode.Open)
        inputFile.Read(bytes, 0, buffersize) 'read header
        Dim header As String = vbNullString
        For f As Int32 = 0 To 79
            header &= Chr(bytes(f))
        Next
        If InStr(header, "solid") > 0 Then
            MsgBox("This could be an ASCII format stl file" & vbNewLine & "It might not load correctly.")
        End If
        buffersize = 4
        inputFile.Read(bytes, 0, buffersize) 'read tri count
        Dim triangles As Int32 = BitConverter.ToInt32(bytes, 0)
        ' MsgBox(triangles)
        If triangles > MAX Then
            MsgBox("File contains too many polygons")
            Exit Sub
        End If
        buffersize = 50
        Dim vect As Vector3
        polycount = 0
        vertcount = 1
        For f As Int32 = 0 To triangles - 1
            inputFile.Read(bytes, 0, buffersize) 'read first triangle data
            vect.X = BitConverter.ToSingle(bytes, 12) + screencenter.X
            vect.Y = BitConverter.ToSingle(bytes, 16) + screencenter.Y
            vect.Z = BitConverter.ToSingle(bytes, 20) + 1000
            Verts(vertcount) = vect
            'check previous poly for matching vertexes and optimise.....
            polys(polycount).vert1 = vertcount
            vertcount += 1
            If vertcount + 2 > MAX Then
                MsgBox("File contains too many vertexes")
                Exit Sub
            End If
            vect.X = BitConverter.ToSingle(bytes, 24) + screencenter.X
            vect.Y = BitConverter.ToSingle(bytes, 28) + screencenter.Y
            vect.Z = BitConverter.ToSingle(bytes, 32) + 1000
            Verts(vertcount) = vect
            polys(polycount).vert2 = vertcount
            vertcount += 1
            vect.X = BitConverter.ToSingle(bytes, 36) + screencenter.X
            vect.Y = BitConverter.ToSingle(bytes, 40) + screencenter.Y
            vect.Z = BitConverter.ToSingle(bytes, 44) + 1000
            Verts(vertcount) = vect
            polys(polycount).vert3 = vertcount
            polys(polycount).Draw_poly = True
            vertcount += 1
            polycount += 1
        Next
        inputFile.Close()
        Label1.Text = "Vertexes imported " & vertcount.ToString("n0") & vbNewLine & "Polygons imported " & polycount.ToString("n0")
        Firstscale()
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub FlipToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FlipToolStripMenuItem1.Click
        If P_right = True Then
            Panel1.Left = 0
            picbox.Left = Panel1.Width
        Else
            picbox.Left = 0
            Panel1.Left = picbox.Width
        End If
        P_right = Not P_right
    End Sub 'flip the panel

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Try
            File.Delete(CurDir() & "\logo")
        Catch
        End Try
        bmp.Dispose()
        GC.Collect()
        Me.Close()
        End
    End Sub 'exit

    Private Sub OpenAModelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenAModelToolStripMenuItem.Click
        If spinx = True Then SpinnerX.PerformClick()
        If spiny = True Then SpinnerY.PerformClick()
        If spinz = True Then SpinnerZ.PerformClick()
        If tumble = True Then Tumbler.PerformClick()
        Nu_config() 'adjust for the display size
        CheckBox5.Checked = True
        Choose(False) ' select and load an object

    End Sub

    Private Sub ExportModelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportModelToolStripMenuItem.Click
        Write_obj()
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Make_bkg()
    End Sub

    Private Sub ToolStripMenuItem2_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.CheckedChanged
        Application.RaiseIdle(e) '.....prevents an occasional crash when trying to veiw the zbuffer
        Makebmp()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Write_Stl()
    End Sub

    Private Sub SaveSettingsAsDefaultToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveSettingsAsDefaultToolStripMenuItem.Click
        Dim writer As New StreamWriter(CurDir() & "\3dE_settings.cfg")
        Dim cont As Object
        For Each cont In Panel1.Controls
            If TypeOf cont Is CheckBox Then
                writer.WriteLine(cont.checked)
            End If
            If TypeOf cont Is TrackBar Then
                writer.WriteLine(cont.value)
            End If
        Next
        writer.WriteLine(light.X)
        writer.WriteLine(light.Y)
        writer.WriteLine(light.Z)
        writer.WriteLine(P_right)
        writer.WriteLine(Lastpath)
        writer.WriteLine(Lastpic)
        writer.Close()
        MessageBox.Show("The current program settings have been saved as the new default setup")
    End Sub

    Private Sub ClearDefaultSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearDefaultSettingsToolStripMenuItem.Click
        If IO.File.Exists(CurDir() & "\3dE_settings.cfg") = True Then
            IO.File.Delete(CurDir() & "\3dE_settings.cfg")
            MessageBox.Show("Custom defaults removed.")
        End If
    End Sub

    Private Sub ChangeBaseColourToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeBaseColourToolStripMenuItem.Click
        ColorDialog1.AnyColor = True
        ColorDialog1.FullOpen = True
        ColorDialog1.ShowDialog()
        Obj_base_colour = AlphaMask + (CInt(ColorDialog1.Color.R) << 16) + (CInt(ColorDialog1.Color.G) << 8) + CInt(ColorDialog1.Color.B)
        Makebmp()
    End Sub


    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        Makebmp()
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click, CheckBox4.Click
        Makebmp()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Makebmp()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        Makebmp()
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        TrackBar8_Scroll(Me, e)
        Label12.Enabled = CheckBox6.Checked
        Makebmp()
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        fillT = CheckBox7.Checked
        Makebmp()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        picbox.Focus()
        If CheckBox8.Checked = True Then
            CheckBox8.ForeColor = Color.Lavender
            '     Label23.ForeColor = Color.Yellow
            mvelight = True
        Else
            CheckBox8.ForeColor = Color.Black
            '     Label23.ForeColor = Color.White
            mvelight = False
        End If
        If CheckBox4.Checked = True Then Makebmp()
        Do Until CheckBox8.Checked = False
            Makebmp()
            Application.DoEvents()
        Loop
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then CheckBox10.Text = "C" Else CheckBox10.Text = "c"
        Makebmp()
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        Makebmp()
    End Sub


    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        blendamount = TrackBar1.Value
        Makebmp()
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        spinspeed = TrackBar2.Value >> 1
    End Sub

    Private Sub TrackBar3_Scroll(sender As Object, e As EventArgs) Handles TrackBar3.Scroll
        lightr = TrackBar3.Value
        Makebmp()
    End Sub

    Private Sub TrackBar4_Scroll(sender As Object, e As EventArgs) Handles TrackBar4.Scroll
        lightg = TrackBar4.Value
        Makebmp()
    End Sub

    Private Sub TrackBar5_Scroll(sender As Object, e As EventArgs) Handles TrackBar5.Scroll
        lightb = TrackBar5.Value
        Makebmp()
    End Sub

    Private Sub TrackBar6_Scroll(sender As Object, e As EventArgs) Handles TrackBar6.Scroll
        light_brightness = -TrackBar6.Value
        Makebmp()
    End Sub

    Private Sub TrackBar7_ValueChanged(sender As Object, e As EventArgs) Handles TrackBar7.ValueChanged
        light.Z = TrackBar7.Value * -1
        Makebmp()
    End Sub

    Private Sub TrackBar8_Scroll(sender As Object, e As EventArgs) Handles TrackBar8.Scroll
        camera.Z = ((TrackBar8.Maximum + TrackBar8.Minimum + 1 - TrackBar8.Value) * -1) * 20
            Clear_array()
        Label12.Text = "Cam Z:" & camera.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub TrackBar9_Scroll(sender As Object, e As EventArgs) Handles TrackBar9.Scroll
        Makebmp()
    End Sub

    Private Sub TrackBar10_Scroll(sender As Object, e As EventArgs) Handles TrackBar10.Scroll
        If TrackBar10.Value = 0 Then TrackBar10.Value = 360
        If TrackBar10.Value = 361 Then TrackBar10.Value = 1
        If CheckBox8.Checked = True Then CheckBox8.Checked = False
        ax = TrackBar10.Value - 1
        Label29.Text = ax.ToString
        ax = ax / 180 * PI
        Makebmp()
    End Sub
    Private Sub TrackBar11_Scroll(sender As Object, e As EventArgs) Handles TrackBar11.Scroll
        If TrackBar11.Value = 0 Then TrackBar11.Value = 360
        If TrackBar11.Value = 361 Then TrackBar11.Value = 1
        ay = TrackBar11.Value - 1
        Label30.Text = ay.ToString
        ay = ay / 180 * PI
        If CheckBox8.Checked = True Then CheckBox8.Checked = False
        Makebmp()
    End Sub
    Private Sub TrackBar12_Scroll(sender As Object, e As EventArgs) Handles TrackBar12.Scroll
        If TrackBar12.Value = 0 Then TrackBar12.Value = 360
        If TrackBar12.Value = 361 Then TrackBar12.Value = 1
        If CheckBox8.Checked = True Then CheckBox8.Checked = False
        az = TrackBar12.Value - 1
        Label31.Text = az.ToString
        az = az / 180 * PI
        Makebmp()
    End Sub


    Private Sub Button1_MouseDown(sender As Object, e As MouseEventArgs) Handles Button1.MouseDown 'move x-
        For f As Int32 = 0 To vertcount - 1
            Verts(f).X -= 50
        Next
        modelcenter.X -= 50
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim nx, ny, nz As Double
        nx = Val(TextBox1.Text)
        ny = Val(TextBox2.Text)
        nz = Val(TextBox3.Text)
        nx = nx - modelcenter.X
        ny = ny - modelcenter.Y
        nz = nz - modelcenter.Z
        For f As Int32 = 0 To vertcount - 1
            Verts(f).X += CSng(nx)
            Verts(f).Y += CSng(ny)
            Verts(f).Z += CSng(nz)
        Next
        Calc_centerpoint()
        Makebmp()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        GroupBox1.Visible = Not GroupBox1.Visible
        GroupBox2.Visible = Not GroupBox1.Visible
        If GroupBox2.Visible = True Then
            TextBox1.Text = modelcenter.X.ToString
            TextBox2.Text = modelcenter.Y.ToString
            TextBox3.Text = modelcenter.Z.ToString
        End If
        Label7.BringToFront()
        Button3.BringToFront()
    End Sub

    Private Sub Button5_MouseDown(sender As Object, e As MouseEventArgs) Handles Button5.MouseDown 'move x+
        For f As Int32 = 0 To vertcount - 1
            Verts(f).X += 50
        Next
        modelcenter.X += 50
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub Button6_MouseDown(sender As Object, e As MouseEventArgs) Handles Button6.MouseDown 'move y-
        For f As Int32 = 0 To vertcount - 1
            Verts(f).Y -= 50
        Next
        modelcenter.Y -= 50
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub Button7_MouseDown(sender As Object, e As MouseEventArgs) Handles Button7.MouseDown 'move y+
        For f As Int32 = 0 To vertcount - 1
            Verts(f).Y += 50
        Next
        modelcenter.Y += 50
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub Button8_MouseDown(sender As Object, e As MouseEventArgs) Handles Button8.MouseDown 'move Z-
        For f As Int32 = 0 To vertcount - 1
            Verts(f).Z -= 100
        Next
        modelcenter.Z -= 100
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub Button9_MouseDown(sender As Object, e As MouseEventArgs) Handles Button9.MouseDown 'move Z+
        For f As Int32 = 0 To vertcount - 1
            Verts(f).Z += 100
        Next
        modelcenter.Z += 100
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Magnify(1.25)
        Calc_centerpoint()
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
        Makebmp()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Magnify(0.875)
        Calc_centerpoint()
        Makebmp()
    End Sub

    Private Sub Tumbler_Click(sender As Object, e As EventArgs) Handles Tumbler.Click
        tumble = Not tumble
        If tumble = True Then Tumbler.Text = "Stop" Else Tumbler.Text = "Tumble"
        spinspeed = TrackBar2.Value * 0.5
        If spiny = True Then tumble = Not tumble : SpinnerY.PerformClick() : tumble = Not tumble
        If spinz = True Then tumble = Not tumble : SpinnerZ.PerformClick() : tumble = Not tumble
        If spinx = True Then tumble = Not tumble : SpinnerX.PerformClick() : tumble = Not tumble
        Dim angx As Single = ax
        Dim angy As Single = ay
        Dim angz As Single = az
        Dim randomx, randomy, randomz As Double
        Dim countt As Int32 = 0
        Do
            If countt = 0 Then
                randomx = rng.Next(50, 126) * 0.01
                randomy = rng.Next(50, 126) * 0.01
                randomz = rng.Next(50, 126) * 0.01
                Select Case rng.Next(2)
                    Case Is = 0
                        randomx *= -1
                    Case Is = 1
                        randomy *= -1
                    Case Is = 2
                        randomz *= -1
                End Select
            End If
            If CheckBox8.Checked = False Then
                ax += (spinspeed * randomx) / 180 * PI
                ay += (spinspeed * randomy) / 180 * PI
                az += (spinspeed * randomz) / 180 * PI
            End If
            Makebmp()
            Application.DoEvents()
            If tumble = False Then
                Exit Do
            End If
            If CheckBox9.Checked = True Then
                spinspeed = TrackBar2.Value >> 1
            Else
                spinspeed = -TrackBar2.Value >> 1
            End If
            If spinspeed = 0 Then spinspeed = 0.5
            If CheckBox8.Checked = False Then countt = (countt + 1) Mod 400
        Loop
        ax = angx
        ay = angy
        az = angz
        Makebmp()
    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles SpinnerX.Click 'spin demo
        spinx = Not spinx
        If spinx = True Then SpinnerX.Text = "Stop" Else SpinnerX.Text = "Spin (P)"
        If spiny = True Then spinx = Not spinx : SpinnerY.PerformClick() : spinx = Not spinx
        If spinz = True Then spinx = Not spinx : SpinnerZ.PerformClick() : spinx = Not spinx
        If tumble = True Then spinx = Not spinx : Tumbler.PerformClick() : spinx = Not spinx
        spinspeed = TrackBar2.Value * 0.5
        Dim angx As Single = ax
        Do Until spinx = False
            If CheckBox8.Checked = False Then ax += spinspeed / 180 * PI
            Makebmp()
            Application.DoEvents()
            If CheckBox9.Checked = True Then
                spinspeed = TrackBar2.Value >> 1
            Else
                spinspeed = -TrackBar2.Value >> 1
            End If
            If spinspeed = 0 Then spinspeed = 0.5
        Loop
        ax = angx
        Makebmp()
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles SpinnerY.Click 'spin demo
        spiny = Not spiny
        If spiny = True Then SpinnerY.Text = "Stop" Else SpinnerY.Text = "Spin (Y)"
        If spinx = True Then spiny = Not spiny : SpinnerX.PerformClick() : spiny = Not spiny
        If spinz = True Then spiny = Not spiny : SpinnerZ.PerformClick() : spiny = Not spiny
        If tumble = True Then spiny = Not spiny : Tumbler.PerformClick() : spiny = Not spiny
        Dim angy As Single = ay
        Do Until spiny = False
            If CheckBox8.Checked = False Then ay += spinspeed / 180 * PI
            Makebmp()
            If CheckBox9.Checked = True Then
                spinspeed = TrackBar2.Value >> 1
            Else
                spinspeed = -TrackBar2.Value >> 1
            End If
            If spinspeed = 0 Then spinspeed = 0.5
            Application.DoEvents()
        Loop
        ay = angy
        Makebmp()
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles SpinnerZ.Click 'spin demo
        spinz = Not spinz
        If spinz = True Then SpinnerZ.Text = "Stop" Else SpinnerZ.Text = "Spin (R)"
        If spinx = True Then spinz = Not spinz : SpinnerX.PerformClick() : spinz = Not spinz
        If spiny = True Then spinz = Not spinz : SpinnerY.PerformClick() : spinz = Not spinz
        If tumble = True Then spinz = Not spinz : Tumbler.PerformClick() : spinz = Not spinz
        spinspeed = TrackBar2.Value >> 1
        Dim angz As Single = az
        Do Until spinz = False
            If CheckBox8.Checked = False Then az += spinspeed / 180 * PI
            Makebmp()
            Application.DoEvents()
            If CheckBox9.Checked = True Then
                spinspeed = TrackBar2.Value >> 1
            Else
                spinspeed = -TrackBar2.Value >> 1
            End If
            If spinspeed = 0 Then spinspeed = 0.5
        Loop
        az = angz
        Makebmp()
        Tumbler.Enabled = True
        SpinnerX.Enabled = True
        SpinnerY.Enabled = True
    End Sub


    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Center()
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        metal = Not metal
        If metal = True Then
            Button30.Text = "S"
        Else
            Button30.Text = "s"
        End If
        Application.DoEvents()
        Makebmp()
    End Sub


    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Form2.Visible = Not Form2.Visible
        If Form2.Visible = False Then Me.Enabled = True Else Me.Enabled = False
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        Makebmp()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        pre_post = ToolStripMenuItem5.Checked
        Makebmp()
    End Sub



    Private Sub Form1_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        picbox.Focus()
        Label23.Visible = True
    End Sub

    Private Sub Form1_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
        Label23.Visible = False
        Panel1.PerformLayout()
    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            light.X = -(screencenter.X - MousePosition.X)
            light.Y = -(screencenter.Y - MousePosition.Y)
            If CheckBox8.Checked = True Then CheckBox8.Checked = False
            'If mvelight = True Then Label23.ForeColor = Color.Yellow Else Label23.ForeColor = Color.White
            If CheckBox4.Checked = True Then Makebmp()
            End If
            If e.Button = Windows.Forms.MouseButtons.Right Then
            light.X = 0
            light.Y = 0
            light.Z = -2
            TrackBar7.Value = 2
            If CheckBox4.Checked = True Then Makebmp()
        End If
    End Sub

    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        If MousePosition.X > Panel1.Right Or MousePosition.X < Panel1.Left Then
            If e.Delta > 0 Then
                If TrackBar7.Value + 50 <= TrackBar7.Maximum Then TrackBar7.Value += 50
            Else
                If TrackBar7.Value - 50 >= TrackBar7.Minimum Then TrackBar7.Value -= 50
            End If
            Makebmp()
        End If
    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        Dim x, y As Int32
        x = MousePosition.X - picbox.Left
        y = MousePosition.Y
        If CheckBox8.Checked = True Then
            Label23.ForeColor = Color.Yellow
        Else
            Label23.ForeColor = Color.White
        End If
        Label23.Text = "x " & x & " (" & x - CInt(screencenter.X) & ") , " & "y " & MousePosition.Y & " (" & MousePosition.Y - screencenter.Y & ")"
    End Sub
End Class
