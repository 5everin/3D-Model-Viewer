# 3D-Model-Viewer
A software renderer written in vb.net 

Requires .net 4.6 or newer. 
May require the adding of a reference to System.Numerics

This is not an example of great coding technique I'm afraid. (I'm just a hobbyist)
It's also my first attempt at anything using 3d and was a learning process for me.
So much of this I would have approached differently with the benefit of hindsight.

It is fast and an example of what you can achieve graphically if you are stubborn enough using only VB.Net.
It also shows that pixel shading and the use of texture maps would be feasible performance wise.


Quick feature list:

Perspective and Orthographic cameras.

Uses a Z-buffer.

Rotation, translation and scaling.

Lighting: A movable light that can be changed between a point source or a directional source.

Renders: Vertexs, wire-frames and flat shaded polygons.

Non light based shading.

multi-threaded

Limited Alias WaveFront .OBJ file support (vertexs, point elements and 3 or 4 sided polygon data only)

Binary .STL file suppport

Transparency:
The program does not z-sort the polygons each frame so partial transparency will probably cause rendering artifacts. 
Mileage will vary depending on the model and various settings in the program.
At worst the model will flicker badly (when moving) and look incorrect. At best it works reasonably well.

Fixing this would require something along the lines of: implementing an A-buffer or an out of order blending algorithm. (or a super fast z-sort)

*** update ***
Added a Point light.
The code runs much faster in .Net5 and above. 
