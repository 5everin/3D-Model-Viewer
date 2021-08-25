# 3D-Model-Viewer
A software renderer written in vb.net 

Requires .net 4.6 or newer. 
May require the adding of a reference to System.Numerics

This is not an example of great code or coding technique I'm afraid. (I'm just a hobbyist)
It is an example of what you can achieve graphically if you are stubborn enough.

Quick feature list:

Perspective and Orthographic cameras.

Uses a Z-buffer.

Rotation, translation and scaling.

Lighting: A movable light that can be changed between a point source or a directional source.

Vertex, wire-frame and flat shaded rendering.

Transparency*

multi-threaded

Limited Alias WaveFront .obj file support ( 3 & 4 sided polygon data only)

Binary .STL file suppport



*The program does not z-sort the polygons each frame so partial transparency will probably cause rendering artifacts. 
Mileage will vary depending on the model and various settings in the program.
At worst the model will flicker badly (when moving) and look incorrect. At best it works reasonably well.

Fixing this would require something along the lines of: implementing an A-buffer or an out of order blending algorithm. (or a super fast z-sort)

*** update ***
Added a Point light.
The code runs much faster in .Net5
The form layout may get messed up a bit. 
