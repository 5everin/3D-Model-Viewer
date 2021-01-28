# 3D-Model-Viewer
A software renderer written in vb.net 

Requires .net 4.6 or newer. 
May require the adding of a reference to System.Numerics

This is not an example of great code or coding technique I'm afraid. (I'm just a hobbyist)
It is an example of what you can achieve graphically if you are stubborn enough.

Quick feature list:

Perspective and Orthographic cameras.

Z buffer.

Rotation, translation and scaling.

A light

Vertex, wire-frame and flat shaded rendering.

Transparency*

multi-threaded

Limited Alias WaveFront .obj file support ( 3 & 4 sided polygon data only)

Binary .STL file suppport

*Transparency uses a dirty hack (for speed) to avoid having to z-sort the polygons each frame. 
Milage will vary depending on the model and various settings in the program.
At worst the model will flicker badly (when moving) and look incorrect. At best it is serviceable.

*** update ***
The code runs much faster in .Net5
The form layout may get messed up a bit. 
