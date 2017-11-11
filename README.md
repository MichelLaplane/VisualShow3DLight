VisualShow3DLight
===================
Create 3D View of Microsoft Visio Shape contained in a diagram

Copyright (c) MS-PL License
Michel Laplane (MVP Visio - ShareVisual)

What's new
===========

Release 0.6.2
	Display shapes using their boundingbox
	Apply color (if filled in the Visio Shape)
	Apply texture if ShapeData "Matériau" is filled with a valid texture file "Stone 01.bmp" (available in Textures directory)

Features
==========

Get any shape in a Diagram and display it in 3D with a default height
Get any shape that have being made VisualShow3DLight aware with correct height, elevation and Texture and
display it in 3D
Display your content in a Visio Document panel View
Display your content in a Window
Dynamic mode provide updating of the content in real time (be aware that this can consume CPU resources)

Configuration
==================

The config file could be customized to set different location of assets :
WebPath

<?xml version="1.0" encoding="utf-8"?>
<configuration> 
  <appSettings>
    <add key="WebPath" value="yourPath\Web" />
    <add key="StencilPath" value="yourPath\Stencils" />
    <add key="TemplatePath" value="yourPath\Template" />
  </appSettings>
</configuration>

Usage
==================

After installation, launch Microsoft Visio (2013 or 2016)
If at the first launch Visio ask you to rely on the content, click yes.

Comments
==================

-	Need a way to suppress the warning when Displaying BabylonJS Canvas.
-	Need a way to be able to call JavaScript function contained in js file (main.js or library.js or other) from Visio.
	It is working only in the javascript unction is in the index.html file.
-	Need Add GUI Mouse support to the 3D View.
-	Must add usage of Visio Shape geometry to enhance 3D View.
-	Must add light(direction and color) feature using special Visio Shape.
-	Must add export function to BabylonJS server.
-	Must add picking feature to select element in BabylonView and display extending Visio Shape content
-	...