# Bézier Curve Drawer
A program to generate Bézier curves.
## Table of Content
- [How to get](#how-to-get)
- [How to use](#how-to-use)
- [About](#about)
## How To Get
### To Use
```vb
If blnUse = True Then
  If isWindowsMachine = True Then
    Call downloadExecutable
  End If
End If
```
If you're on a machine running Windows, then you can download the executable [here](https://github.com/itsmarsss/Bezier-Curve-Drawer/blob/main/Quadratic%20Bezier%20Curve.exe) and run it
### To Edit
```vb
If blnEdit = True Then
  Call cloneRepository
End If
```
If you want to edit the code, then you can clone this GitHub;  
`git clone https://github.com/itsmarsss/Bezier-Curve-Drawer.git`
## How To Use
After running the program there should be a white canvas and a column of element/tools.
### Elements/Tools
**PictureBox/Canvas** - Click to create points (Only quadratic bézier curves)  
**Clear** - Clear canvas and/or curve list  
**List** - Stores all curves  
**Remove** - Removes selected item in List  
**Guides** - Toggle points and lines on curves  
**Save Image** - Save drawing as a JPG image  
**Draw** - Draws all saved curves  
**Draw Selected** - Draws selected item in List  
**Label** - Coordinated of mouse in pixels

## About
This program was created to better visualize Bézier curves for a Math research project.
