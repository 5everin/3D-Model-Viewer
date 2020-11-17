Imports System.Math
Imports System.IO
Imports System.Numerics

Public Class Form1
    Const AlphaMask As Int32 = &HFF000000
    Const RedMask As Int32 = &HFF0000
    Const greenMask As Int32 = &HFF00
    Const blueMask As Int32 = &HFF
    Dim Cpu_Use As New PerformanceCounter("Processor", "% Processor Time", "_Total")
    Dim proc_count As String = System.Environment.ProcessorCount.ToString
    ' Dim opti As New ParallelOptions
    Dim rec As New Rectangle(0, 0, 1, 1)
    Dim Swidth, Sheight As Int32
    ReadOnly rng As New Random
    ReadOnly Time_stuff As New Stopwatch
    ReadOnly Picbox As New PictureBox
    Dim P_right As Boolean = True
    Dim screencenter As New Vector3 'Point
    Dim camdefault As Int32
    Dim camera As New Vector3
    Dim vertcount As Int32 = 0
    Dim polycount As Int32 = 0
    ReadOnly MAX As Int32 = 10000000
    ReadOnly Verts(MAX) As Vector3
    ReadOnly Polys(MAX) As Poly
    ReadOnly Normals(MAX) As Vector3
    Dim Bigarray(10) As Int32
    Dim Zbuffer(10) As Int32
    Dim bkg_image(10) As Int32
    Dim size_array As Int32
    Dim modelcenter As New Vector3
    Dim Mscale As Double = 1
    Dim tumble As Boolean = False
    Dim spinx As Boolean = False
    Dim spiny As Boolean = False
    Dim spinz As Boolean = False
    Dim bmp As Bitmap
    Dim bkg As Bitmap
    Dim spinspeed As Double
    Dim light As Vector3
    Dim lightr, lightg, lightb As Int32
    Dim Obj_base_colour As Int32 = &HFF010101
    Dim light_brightness As Int32 = -8
    Dim ax, ay, az As Single 'rotatation angles
    Dim diffmult As Double = 0
    Dim mvelight As Boolean = False
    Dim LoadAModel As Boolean = True
    Dim filename As String = vbNullString
    Dim Lastpath As String = "x"
    Dim Lastpic As String = "x"
    Dim blendamount As Int32 = 172
    Dim pre_post As Boolean = False 'type of alpha blending to be used
    ReadOnly cfgfolder As String = IO.Directory.GetCurrentDirectory.ToString

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
        'adjust the position of the panel controls as there will be no vertical scrollbar displayed  
        If Sheight >= 1050 Then
            Dim cont As Object
            For Each cont In Panel1.Controls
                cont.left += 8
            Next
        End If
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
        If CheckBox10.Checked = True Then CheckBox10.Text = "Shading + Light" Else CheckBox10.Text = "Light"
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
                        Label2.Text = "Model" & vbNewLine & import.SafeFileName
                    Else
                        Read_STL(filename)
                        Label2.Text = "Model" & vbNewLine & import.SafeFileName
                    End If
                    Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
                Else
                    Label2.Text = "Model" & vbNewLine & "Nothing Loaded"
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
            Label2.Text = "Model" & vbNewLine & "Default Model"
        End If
    End Sub

    Private Sub Clear_array()


        If CheckBox11.Checked = False Then
            Parallel.For(0, Bigarray.Length, Sub(f As Int32)
                                                 If Bigarray(f) <> &HFF000000 Then Bigarray(f) = &HFF000000 'just clear the array 
                                                 If Zbuffer(f) < &H6F000000 Then Zbuffer(f) = &H6F000000 ' reset the zbuffer
                                             End Sub)
        Else
            Parallel.For(0, bigarray.Length, Sub(f As Int32)
                                                 bigarray(f) = bkg_image(f) 'copy data from background image 
                                                 Zbuffer(f) = &H6F000000
                                             End Sub)
        End If
    End Sub

    Private Sub Nu_rasterPoly()
        diffmult = TrackBar9.Value * 0.001
        If diffmult >= 1 Then diffmult = 0.999
        ' for each triangle locate the 3 vetex positions,rotate them, generate a normal and map the vertexes to 2d with or without perspective.
        ' Sort the vertexes, split the triangle if required and Then call the draw wireframe and/or filled triangle routines. 
        Parallel.For(0, polycount, Sub(f As Int32)
                                       Dim camvec As Vector3 = camera
                                       Dim colour As Int32 = 0
                                       Dim norml As Vector3
                                       'copy triangle vertexes to 3 temp vectors for transforms etc. (original model data is not modified in any way. The image is generated from these copied triangles)
                                       Dim vec(3) As Vector3
                                       vec(0) = Verts(Polys(f).vert1)
                                       vec(1) = Verts(Polys(f).vert2)
                                       vec(2) = Verts(Polys(f).vert3)
                                       'rotate the vertexes 
                                       Dim inputQ As Quaternion
                                       Dim RotatedQ As Quaternion
                                       For i As Int32 = 0 To 2
                                           vec(i) -= modelcenter
                                           inputQ = Quaternion.CreateFromYawPitchRoll(ay, ax, az) 'rotation angles in radians
                                           RotatedQ = (inputQ * New Quaternion(vec(i), 0)) * Quaternion.Conjugate(inputQ)
                                           vec(i).X = RotatedQ.X
                                           vec(i).Y = RotatedQ.Y
                                           vec(i).Z = RotatedQ.Z
                                           vec(i) += modelcenter
                                       Next
                                       'generate the face normal
                                       Dim tmp As Vector3 = Vector3.Subtract(vec(1), vec(0))
                                       vec(3) = Vector3.Subtract(vec(2), vec(0))
                                       norml = Vector3.Cross(tmp, vec(3))
                                       norml = Vector3.Normalize(norml)
                                       tmp = norml
                                       tmp.Z -= camvec.Z
                                       Dim cp As Single = Vector3.Dot(tmp, norml)
                                       If CheckBox5.Checked = True Then
                                           If cp > 0 Then Exit Sub 'back face cull
                                       ElseIf cp >= 0 Then
                                           norml *= -1 ' flips normal so the triangle is always faceing the camerea creating a "double sided" polygon
                                       End If
                                       Normals(f) = norml
                                       If CheckBox6.Checked = True Then
                                           Dim cam As Int32 = Floor(camvec.Z + 200) 'z clip distance (stops frame rate dropping to silly levels when geomatry starts get to close to the camera )
                                           For n As Int32 = 0 To 2 '   perspective transform for each vertex(3D > 2D mapping) 
                                               vec(n).X = camvec.Z * (vec(n).X - camvec.X)
                                               vec(n).X /= (camvec.Z - vec(n).Z)
                                               vec(n).X += camvec.X
                                               vec(n).Y = camvec.Z * (vec(n).Y - camvec.Y)
                                               vec(n).Y /= (camvec.Z - vec(n).Z)
                                               vec(n).Y += camvec.Y
                                               If cam > vec(n).Z = True Then
                                                   If Abs(vec(n).X) > (Swidth << 1) Or Abs(vec(n).Y) > (Sheight << 1) Then Exit Sub
                                                   If cam + 2 > (vec(n).Z) Then Exit Sub ' clip z if it comes too close to the camera on the z axis.
                                               End If
                                           Next
                                       End If
                                       'force Y axis values of the vertexes into an integer values (for correct array locations.)
                                       For n As Int32 = 0 To 2
                                           vec(n).Y = Floor(vec(n).Y)
                                       Next
                                       'Vertex_sort by y axis lowest first
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
                                       'if entire triangle is off screen then don't attempt to draw it. 
                                       If vec(2).Y < 0 OrElse vec(0).Y > Sheight Then Exit Sub
                                       If vec(0).X > Swidth AndAlso vec(1).X > Swidth AndAlso vec(2).X > Swidth Then Exit Sub
                                       If vec(0).X < 0 AndAlso vec(1).X < 0 AndAlso vec(2).X < 0 Then Exit Sub
                                       'draw the Vertexes
                                       If CheckBox2.Checked = True Then
                                           For n As Int32 = 0 To 2
                                               Dim Draw_vertex As Boolean = True
                                               If vec(n).X < 0 OrElse vec(n).X + 1 >= Swidth Then Draw_vertex = False
                                               If vec(n).Y < 0 OrElse vec(n).Y >= Sheight Then Draw_vertex = False
                                               If Draw_vertex = True Then
                                                   'add a vertex marker into bigarray & the zbuffer
                                                   Dim loc As Int32 = Floor(vec(n).X) + (Floor(vec(n).Y) * Swidth) 'pixel location to array position
                                                   If loc + 1 < Bigarray.Length Then
                                                       If vec(n).Z < Zbuffer(loc) Then
                                                           Bigarray(loc) = &HFFFD3C00
                                                           Zbuffer(loc) = Floor(vec(n).Z) - 2
                                                       End If
                                                   End If
                                               End If
                                           Next
                                       End If
                                       'work out how to shade the image and set colours scheme.
                                       FillsnCols(f, vec(0), colour)
                                       'wireframe
                                       If CheckBox1.Checked = True Then
                                           Dim fw As Boolean = ToolStripMenuItem6.Checked
                                           Dim tmpc As Int32 = Not (colour)
                                           Dim dist As Int32 = Vector3.Distance(vec(0), vec(1))
                                           If fw = True Then dist >>= 1
                                           Triangle_wire(vec(0), vec(1), tmpc, dist)
                                           dist = Vector3.Distance(vec(1), vec(2))
                                           If fw = True Then dist >>= 1
                                           Triangle_wire(vec(1), vec(2), tmpc, dist)
                                           dist = Vector3.Distance(vec(2), vec(0))
                                           If fw = True Then dist >>= 1
                                           Triangle_wire(vec(2), vec(0), tmpc, dist)
                                       End If
                                       'workout the triangle type and draw. 
                                       If CheckBox7.Checked = True Then
                                           If vec(1).Y = vec(2).Y Then
                                               Flatbottom(vec(0), vec(1), vec(2), colour)
                                           ElseIf vec(0).Y = vec(1).Y Then
                                               Flattop(vec(0), vec(1), vec(2), colour)
                                           Else
                                               vec(3).Y = (vec(1).Y - vec(0).Y) / (vec(2).Y - vec(0).Y) 'generate new temp vertex to split the triangle into two
                                               vec(3) = vec(0) + (vec(2) - vec(0)) * vec(3).Y
                                               vec(3).Y = vec(1).Y
                                               Flatbottom(vec(0), vec(1), vec(3), colour)
                                               Flattop(vec(1), vec(3), vec(2), colour) 'generate a  triangle from the bottom 3 vectors'generate a triangle from the top 3 vectors
                                           End If
                                       End If
                                   End Sub)
    End Sub

    Public Sub FillsnCols(indx As Int32, vec As Vector3, ByRef colour As Int32)
        Dim DrawAline As Boolean = CheckBox1.Checked
        Dim Lit_lines As Boolean = CheckBox4.Checked
        'make a colour for non light based shading
        If CheckBox10.Checked = True OrElse Lit_lines = False Then
            Dim Rcolour As Int32 = Floor(Abs(modelcenter.Z * 0.168 - vec.X) * 0.138) Mod 255
            Dim Gcolour As Int32 = Floor(Abs(modelcenter.X * 0.159 - vec.Y) * 0.159) Mod 255
            Dim Bcolour As Int32 = Floor(Abs(modelcenter.Y * 0.141 - vec.Z) * 0.161) Mod 255
            colour = AlphaMask + (Rcolour << 16) + (Gcolour << 8) + Bcolour
        End If
        'Pick a colour scheme 
        If DrawAline = True AndAlso Lit_lines = False Then
            colour = Obj_base_colour
        Else
            Lighting(Normals(indx), colour) 'lighting
        End If
        If DrawAline = False AndAlso Lit_lines = False Then  '  non-lit polys
            Dim test As Int32 = indx >> 1 << 1
            If test = indx Then
                colour = &H7F808080
            Else
                colour = &H7F401090
            End If
        End If
    End Sub

    Private Sub Triangle_wire(vec0 As Vector3, vec1 As Vector3, colour As Int32, dist As Int32)
           Dim line As Vector3
        Dim loc As Int32
        For f As Int32 = 1 To Floor(dist)
            line = Vector3.Lerp(vec0, vec1, f / dist)
            loc = line.X + (Floor(line.Y) * Swidth)
            If line.X >= 0 AndAlso line.X + 1 < Swidth Then
                If loc < size_array AndAlso loc >= 0 Then
                    If line.Z <= Zbuffer(loc) Then
                        Bigarray(loc) = colour
            Zbuffer(loc) = line.Z - 2
          End If
                End If
            End If
        Next
    End Sub

  Private Sub Flatbottom(ByRef vec0 As Vector3, ByRef vec1 As Vector3, ByRef vec2 As Vector3, ByRef colour As Int32)
    'this traces the two lines down the sides of a flat bottomed triangle in paralell generating x,y,z coords for the zbuffer and the two end points of the vertical line connecting them...
    'it then calls drawx which generates the points between the vertical line end points passed to it.
    For scanline As Int32 = vec0.Y To vec2.Y
      Dim l1 As Vector3 = Vector3.Lerp(vec1, vec0, (scanline - vec2.Y) / (vec0.Y - vec1.Y))
      Dim l2 As Vector3 = Vector3.Lerp(vec2, vec0, (scanline - vec2.Y) / (vec0.Y - vec2.Y))
      If l2.X >= l1.X Then
        DrawX(l1, l2, colour)
      Else
        DrawX(l2, l1, colour)
      End If
    Next
  End Sub

  Private Sub Flattop(ByRef vec0 As Vector3, ByRef vec1 As Vector3, ByRef vec2 As Vector3, ByRef colour As Int32)
    'this traces the two lines down the sides of a flat topped triangle in paralell generating x,y,z coords for the zbuffer and the two end points of the vertical line connecting them...
    'it then calls drawx which generates the points between the vertical line end points passed to it.

    For scanline As Int32 = vec0.Y To vec2.Y - 1
      Dim l1 As Vector3 = Vector3.Lerp(vec1, vec2, (scanline - vec2.Y) / (vec1.Y - vec2.Y))
      Dim l2 As Vector3 = Vector3.Lerp(vec0, vec2, (scanline - vec2.Y) / (vec0.Y - vec2.Y))
      If l2.X >= l1.X Then
        DrawX(l1, l2, colour)
      Else
        DrawX(l2, l1, colour)
      End If
    Next
  End Sub


  Private Sub DrawX(ByVal l1 As Vector3, ByVal l2 As Vector3, ByVal colour As Int32) 'draws a line along the x axis generating z axis coordinates for the zbuffer. (with or without alpha blending)
    l1.X = Floor(l1.X)
    l2.X = Floor(l2.X)
    Dim zslope As Double = (l1.Z - l2.Z) / (l1.X - l2.X)
    Dim zpos As Double = l1.Z

    Dim loc As Int32
    For n As Int32 = l1.X + 1 To l2.X
      If n >= 0 AndAlso n < Swidth Then
        loc = n + (l2.Y * Swidth)
        If loc < size_array AndAlso loc >= 0 Then
          If zpos < Zbuffer(loc) Then
            If CheckBox12.Checked = True AndAlso pre_post = False Then  'alpha blending (uses the zbuffer to adjust polygon transparency - dirty hack)
              Dim bl As Int32 = blendamount
              If zpos < Zbuffer(loc) Then
                bl -= 16
                Zbuffer(loc) = zpos
              Else
                bl = 255 - (blendamount >> 3)
              End If
              Dim rb As Int32 = colour And &HFF00FF
              Dim g As Int32 = colour And &HFF00
              rb += ((Bigarray(loc) And &HFF00FF) - rb) * (bl) >> 8
              g += ((Bigarray(loc) And &HFF00) - g) * (bl) >> 8
              Bigarray(loc) = (rb And &HFF00FF) Or (g And &HFF00)
            Else
              Bigarray(loc) = colour
              Zbuffer(loc) = zpos
            End If
          End If
          End If
      End If
      zpos += zslope
    Next n
  End Sub

  Private Sub AlphaBlend() ' post processing version of alpha blend ----will always obscure internal polygons
        If ToolStripMenuItem2.Checked = False Then
            Parallel.For(0, Zbuffer.Length - 1, Sub(f As Int32)
                                                    If Zbuffer(f) < &H6F000000 Then
                                                        Dim bkg As Int32
                                                        If CheckBox11.Checked = True Then bkg = bkg_image(f) Else bkg = &HFF000000
                                                        'fast full alpha blend.
                                                        Dim rb As Int32 = Bigarray(f) And &HFF00FF
                                                        Dim g As Int32 = Bigarray(f) And &HFF00
                                                        rb += ((bkg And &HFF00FF) - rb) * blendamount >> 8
                                                        g += ((bkg And &HFF00) - g) * blendamount >> 8
                                                        Bigarray(f) = (rb And &HFF00FF) Or (g And &HFF00)
                                                    End If
                                                End Sub)
        End If
    End Sub

    Private Sub Lighting(ByRef norm As Vector3, ByRef colour As Int32)
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
        Dim LightDirection As Vector3 = Vector3.Add(light, norm)
        LightDirection = Vector3.Normalize(LightDirection)
        Dim diff As Double = Vector3.Dot(norm, LightDirection)
        If diff > diffmult Then
            diff *= (1 + diffmult)
            diff += diff * 0.158
        End If
        rdot = Floor(rdot - light_brightness + (diff * lightr))
        gdot = Floor(gdot - light_brightness + (diff * lightg))
        bdot = Floor(bdot - light_brightness + (diff * lightb))
        rdot = Math.Max(0, Math.Min(255, rdot))
        gdot = Math.Max(0, Math.Min(255, gdot))
        bdot = Math.Max(0, Math.Min(255, bdot))
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
  End Sub

  Private Sub Drawgrid()
        Dim zbool As Boolean = ToolStripMenuItem2.Checked
        Dim gsize As Int32 = 150
        For j As Int32 = 0 To Sheight - 1 Step gsize
            For i As Int32 = 0 To Swidth - 1 Step 3
                Dim loc As Int32 = i + (j * Swidth)
                If zbool = False Then Bigarray(loc) = &H1444444 Else Zbuffer(loc) = &H6EFFFFFF
            Next
        Next
        For j As Int32 = 0 To Swidth - 1 Step gsize
            For i As Int32 = 0 To Sheight - 1 Step 3
                Dim loc As Int32 = j + (i * Swidth)
                If zbool = False Then Bigarray(loc) = &H1444444 Else Zbuffer(loc) = &H6EFFFFFF
            Next
        Next
    End Sub

    Private Sub Makebmp() ' generate the image on screen
        If vertcount > 2 Then
            Time_stuff.Start()
            Clear_array()
            If mvelight = True Then
                light.X = (MousePosition.X - screencenter.X)
                light.Y = (MousePosition.Y - screencenter.Y)
            End If
            If CheckBox3.Checked = True Then Drawgrid()
            If CheckBox7.Checked = True OrElse CheckBox1.Checked = True OrElse CheckBox2.Checked = True Then Nu_rasterPoly()
            If Bigarray.Length = size_array Then
                If ToolStripMenuItem2.Checked = True Then
                    Parallel.For(0, Zbuffer.Length - 1, Sub(f As Int32) 'overwrite bigarray with a visualisation of the zbuffer
                                                            Dim zb As Int32 = 127 + Floor((Zbuffer(f) * 0.125)) Mod 256
                                                            If Zbuffer(f) = &H6F000000 Then zb = &HFF000000
                                                            Bigarray(f) = &H6F000000 + ((zb << 16) + (zb << 8) + zb)
                                                        End Sub)
                End If
                If CheckBox12.Checked = True AndAlso pre_post = True Then AlphaBlend()
                Dim bmpdata As Imaging.BitmapData = bmp.LockBits(rec, Drawing.Imaging.ImageLockMode.WriteOnly, bmp.PixelFormat)
                Dim ptr As IntPtr = bmpdata.Scan0
                System.Runtime.InteropServices.Marshal.Copy(Bigarray, 0, ptr, size_array)
                bmp.UnlockBits(bmpdata)
                Picbox.Image = bmp
            End If
            Time_stuff.Stop()
            Label3.Text = Time_stuff.Elapsed.TotalMilliseconds.ToString("f2") & " ms " & "(" & (1000 / Time_stuff.Elapsed.TotalMilliseconds).ToString("f2") & " fps)"
            If CheckBox4.Checked = True Then
                Label19.Text = "X:" & light.X.ToString("n0") & vbNewLine & "Y:" & light.Y.ToString("n0") & vbNewLine & "Z:" & -light.Z
            Else
                Label19.Text = vbNullString
            End If
            Time_stuff.Reset()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ToolStripMenuItem7.Checked = True Then
            Label33.Text = "Cpu:" & Cpu_Use.NextValue.ToString("n2") & "%" & " using" & proc_count & " logical cores"
        Else
            Label33.Text = vbNullString
        End If
    End Sub
    Private Sub Magnify(ByVal zoom As Single)
        Parallel.For(1, vertcount, Sub(f As Int32)
                                       Verts(f) = Vector3.Subtract(Verts(f), modelcenter)
                                       Verts(f) = Vector3.Multiply(Verts(f), zoom)
                                       Verts(f) = Vector3.Add(Verts(f), modelcenter)
                                   End Sub)
        Mscale *= zoom
        Label8.Text = "Scale x " & Mscale.ToString("n2")
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
        Label7.Text = "X:" & modelcenter.X.ToString("n0") & " Y:" & modelcenter.Y.ToString("n0") & " Z:" & modelcenter.Z.ToString("n0")
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
        Dim tempstring As String
        For f As Int32 = 1 To vertcount - 1
            tempstring = "v " & Verts(f).X & " " & Verts(f).Y & " " & Verts(f).Z
            writer.WriteLine(tempstring)
        Next
        For f As Int32 = 0 To polycount - 1
            tempstring = "f " & Polys(f).vert1 & " " & Polys(f).vert2 & " " & Polys(f).vert3
            writer.WriteLine(tempstring)
        Next
        writer.Close()
        writer.Dispose()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Read_obj(ByVal fn As String) 'basic .obj parser
        Dim reader As New StreamReader(fn)
        Dim st As String
        Dim st_st As String
        Dim counts As Int32 = 1
        Dim counts2 As Int32 = 0
        Dim Tcounts As Int32 = 1
        Dim Tcounts2 As Int32 = 0
        Dim num(3) As Double
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
                    Verts(counts).Z = num(2) + 1
                    '   Verts(counts) *= (Swidth * 2)
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
                    Dim space As Int32
                    Dim slash As Int32
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
                            Polys(counts2).vert1 = Floor(num(0))
                            Polys(counts2).vert2 = Floor(num(1))
                            Polys(counts2).vert3 = Floor(num(2))
                            Polys(counts2).Draw_poly = True
                        Else ' convert 4 point poly into 2x 3 point poly's
                            Polys(counts2).vert1 = Floor(num(0))
                            Polys(counts2).vert2 = Floor(num(1))
                            Polys(counts2).vert3 = Floor(num(2))
                            counts2 += 1
                            Polys(counts2).vert1 = Floor(num(0))
                            Polys(counts2).vert2 = Floor(num(2))
                            Polys(counts2).vert3 = Floor(num(3))
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
                            Polys(counts2).vert1 = Floor(num(0))
                            Polys(counts2).vert2 = Floor(num(1))
                            Polys(counts2).vert3 = Floor(num(2))
                        Else ' convert 4 point poly into 2x 3 point poly's
                            Polys(counts2).vert1 = Floor(num(0))
                            Polys(counts2).vert2 = Floor(num(1))
                            Polys(counts2).vert3 = Floor(num(2))
                            counts2 += 1
                            Tcounts2 += 1
                            Polys(counts2).vert1 = Floor(num(0))
                            Polys(counts2).vert2 = Floor(num(2))
                            Polys(counts2).vert3 = Floor(num(3))
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
        If polycount = 0 And vertcount > 1 Then polycount = counts
        'Scale the model to a reasonable size based on screen height
        Firstscale()
    End Sub


    Private Sub Firstscale_2()
        Dim minx, miny, minz As Double
        Dim maxx, maxy, maxz As Double

        minx = &HFFFF0000 '----- A
        maxx = 0 '------- B

        miny = &HFFFF0000 '-------A
        maxy = 0 '-------B

        minz = &HFFFF0000 '-------A
        maxz = 0 '-------B
        Dim Nrange_min As Vector3 = New Vector3(1, 1, 1) '-----a,a,a
        Dim Nrange_max As Vector3 = New Vector3(Sheight * 2, Sheight * 2, Sheight * 2) '-----b,b,b

        'Normalise ------- A=lowest value of input range, B=Highest value of input range
        '......... ------- a=lowest value of output range, a=Highest value of output range
        '......... ------- a+(x-A)*(b-a)/(B-A)............X = value to be normailsed


        For vert As Int32 = 1 To vertcount - 1
            If Verts(vert).X > maxy Then maxx = Verts(vert).X
            If Verts(vert).X < minx Then minx = Verts(vert).X
            If Verts(vert).Y > maxy Then maxy = Verts(vert).Y
            If Verts(vert).Y < miny Then miny = Verts(vert).Y
            If Verts(vert).Z > maxy Then maxy = Verts(vert).Z
            If Verts(vert).Z < miny Then miny = Verts(vert).Z
        Next

        For vert As Int32 = 1 To vertcount - 1
            Verts(vert).X = Nrange_min.X + (Verts(vert).X - minx) * (Nrange_max.X - Nrange_min.X) / (maxx - minx)
            Verts(vert).Y = Nrange_min.Y + (Verts(vert).Y - miny) * (Nrange_max.Y - Nrange_min.Y) / (maxy - miny)
            Verts(vert).Z = Nrange_min.Z + (Verts(vert).Z - minz) * (Nrange_max.Z - Nrange_min.Z) / (maxz - minz)
        Next

        Magnify(5)



        'If maxy - modelcenter.Y < Sheight Then
        '    If maxy - modelcenter.Y < 10 Then
        '        Magnify(15 - (maxy - modelcenter.Y))
        '    Else
        '        Magnify(4.5)
        '    End If
        'ElseIf maxy - modelcenter.Y > Sheight Then
        '    Magnify(0.18)
        'Else
        '    Center()
        '    Exit For
        'End If
        Center()

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
        Mscale = 1
        Magnify(1)
    End Sub

    Private Sub Write_Stl()
        Me.Cursor = Cursors.WaitCursor
        Dim tmp() As Byte
        Dim desktp As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim OutputFile As Stream = IO.File.Open(desktp & "\export.stl", IO.FileMode.Create)
        Dim Header As String = " PhilG's Model Viewer "
        Dim bytes(80) As Byte
        tmp = System.Text.Encoding.Unicode.GetBytes(Header)
        Array.Copy(tmp, bytes, tmp.Length - 1)
        OutputFile.Write(bytes, 0, 80) 'write header
        bytes = BitConverter.GetBytes(polycount)
        OutputFile.Write(bytes, 0, 4) ' tri count
        Dim tempbytes() As Byte
        Array.Resize(Of Byte)(bytes, 50)
        For f As Int32 = 0 To polycount - 1
            tempbytes = BitConverter.GetBytes(Normals(f).X) ' ....write data (normal xyz/vertex(1,2,3) x,y,z/2 zero byte values.
            bytes(0) = tempbytes(0)
            bytes(1) = tempbytes(1)
            bytes(2) = tempbytes(2)
            bytes(3) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Normals(f).Y)
            bytes(4) = tempbytes(0)
            bytes(5) = tempbytes(1)
            bytes(6) = tempbytes(2)
            bytes(7) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Normals(f).Z)
            bytes(8) = tempbytes(0)
            bytes(9) = tempbytes(1)
            bytes(10) = tempbytes(2)
            bytes(11) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert1).X)
            bytes(12) = tempbytes(0)
            bytes(13) = tempbytes(1)
            bytes(14) = tempbytes(2)
            bytes(15) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert1).Y)
            bytes(16) = tempbytes(0)
            bytes(17) = tempbytes(1)
            bytes(18) = tempbytes(2)
            bytes(19) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert1).Z)
            bytes(20) = tempbytes(0)
            bytes(21) = tempbytes(1)
            bytes(22) = tempbytes(2)
            bytes(23) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert2).X)
            bytes(24) = tempbytes(0)
            bytes(25) = tempbytes(1)
            bytes(26) = tempbytes(2)
            bytes(27) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert2).Y)
            bytes(28) = tempbytes(0)
            bytes(29) = tempbytes(1)
            bytes(30) = tempbytes(2)
            bytes(31) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert2).Z)
            bytes(32) = tempbytes(0)
            bytes(33) = tempbytes(1)
            bytes(34) = tempbytes(2)
            bytes(35) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert3).X)
            bytes(36) = tempbytes(0)
            bytes(37) = tempbytes(1)
            bytes(38) = tempbytes(2)
            bytes(39) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert3).Y)
            bytes(40) = tempbytes(0)
            bytes(41) = tempbytes(1)
            bytes(42) = tempbytes(2)
            bytes(43) = tempbytes(3)
            tempbytes = BitConverter.GetBytes(Verts(Polys(f).vert3).Z)
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
        If triangles > MAX Then
            MsgBox("File contains too many polygons")
            Exit Sub
        End If
        buffersize = 50
        Dim vect As Vector3
        polycount = 0
        vertcount = 1
        For f As Int32 = 0 To triangles - 1
            inputFile.Read(bytes, 0, buffersize) 'read first triangle data(skip the normal data)
            vect.X = BitConverter.ToSingle(bytes, 12) + screencenter.X
            vect.Y = BitConverter.ToSingle(bytes, 16) + screencenter.Y
            vect.Z = BitConverter.ToSingle(bytes, 20) + 1000
            Verts(vertcount) = vect
            'check previous poly for matching vertexes and optimise.....goes here
            Polys(polycount).vert1 = vertcount
            vertcount += 1
            If vertcount + 2 > MAX Then
                MsgBox("File contains too many vertexes")
                Exit For
            End If
            vect.X = BitConverter.ToSingle(bytes, 24) + screencenter.X
            vect.Y = BitConverter.ToSingle(bytes, 28) + screencenter.Y
            vect.Z = BitConverter.ToSingle(bytes, 32) + 1000
            Verts(vertcount) = vect
            Polys(polycount).vert2 = vertcount
            vertcount += 1
            vect.X = BitConverter.ToSingle(bytes, 36) + screencenter.X
            vect.Y = BitConverter.ToSingle(bytes, 40) + screencenter.Y
            vect.Z = BitConverter.ToSingle(bytes, 44) + 1000
            Verts(vertcount) = vect
            Polys(polycount).vert3 = vertcount
            Polys(polycount).Draw_poly = True
            vertcount += 1
            polycount += 1
        Next
        inputFile.Close()
        Label1.Text = "Vertexes imported " & vertcount.ToString("n0") & vbNewLine & "Polygons imported " & polycount.ToString("n0")
        If vertcount + 2 < MAX Then Firstscale()
        Me.Cursor = Cursors.Default
    End Sub


    Private Sub FlipToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FlipToolStripMenuItem1.Click
        If P_right = True Then
            Panel1.Left = 0
            Picbox.Left = Panel1.Width
        Else
            Picbox.Left = 0
            Panel1.Left = Picbox.Width
        End If
        P_right = Not P_right
    End Sub 'flip the panel

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Timer1.Stop()
        Try
            File.Delete(CurDir() & "\logo")
        Catch
        End Try
        Cpu_Use.Dispose()
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
        Array.Clear(Verts, 0, Verts.Length)
        Nu_config() 'adjust for the display size
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
        Dim writer As New StreamWriter(cfgfolder & "\3dE_settings.cfg")
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
        Obj_base_colour = AlphaMask + (Floor(ColorDialog1.Color.R) << 16) + (Floor(ColorDialog1.Color.G) << 8) + Floor(ColorDialog1.Color.B)
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
        Makebmp()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        Picbox.Focus()
        If CheckBox8.Checked = True Then
            CheckBox8.ForeColor = Color.Lavender
            mvelight = True
        Else
            CheckBox8.ForeColor = Color.Black
            mvelight = False
        End If
        If CheckBox4.Checked = True Then Makebmp()
        Do Until CheckBox8.Checked = False
            Makebmp()
            Application.DoEvents()
        Loop
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = True Then CheckBox10.Text = "Shading && Light" Else CheckBox10.Text = "Light"
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
        ax = -ax / 180 * PI
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
        nx -= modelcenter.X
        ny -= modelcenter.Y
        nz -= modelcenter.Z
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
      '     If CheckBox8.Checked = False Then
      ax += (spinspeed * randomx) / 180 * PI
      ay += (spinspeed * randomy) / 180 * PI
      az += (spinspeed * randomz) / 180 * PI
      '   End If
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
        TrackBar10.Value = 1
        TrackBar11.Value = 1
        TrackBar12.Value = 1
        TrackBar10_Scroll(Me, e)
        TrackBar11_Scroll(Me, e)
        TrackBar12_Scroll(Me, e)
    End Sub

    'Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
    '    Form2.Visible = Not Form2.Visible
    '    If Form2.Visible = False Then Me.Enabled = True Else Me.Enabled = False
    'End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        Makebmp()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        pre_post = ToolStripMenuItem5.Checked
        Makebmp()
    End Sub

    Private Sub ToolStripMenuItem6_CheckedChanged(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.CheckedChanged
        Makebmp()
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        ToolStripMenuItem6.Checked = Not ToolStripMenuItem6.Checked
    End Sub

    Private Sub Form1_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        Picbox.Focus()
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
            light.Z = -8
            TrackBar7.Value = 8
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
        x = MousePosition.X - Picbox.Left
        y = MousePosition.Y
        If CheckBox8.Checked = True Then
            Label23.ForeColor = Color.Yellow
        Else
            Label23.ForeColor = Color.White
        End If
        Label23.Text = "x " & x & " (" & x - Floor(screencenter.X) & ") , " & "y " & MousePosition.Y & " (" & MousePosition.Y - screencenter.Y & ")"
    End Sub
End Class
