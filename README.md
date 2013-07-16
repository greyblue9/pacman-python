Pacman by David Reilly
======================
With contributions by Andy Sommerville (2007)

Github project page:
https://github.com/greyblue9/pacman-python

![Screenshot - 16x16 tile version](/screenshot-1.png)


Thank you for your interest in my tribute to everyone's favorite Pac-Man
games of yesteryear! Pacman features 12 colorful levels of varying
sizes, bouncing fruit, A* path-finding ghosts, and, more recently,
cross-platform support, joystick support, and a high score list. This
was the first large-scale project I ever developed in Python, so the
source is not the cleanest, but I would love to have your contributions!
Feel free to slice and dice it up however you like and redistribute it;
I just ask that you give me some credit and let me know about it, so I
can enjoy your creation too! Have fun!


Installation
------------

Pac-man requires Python 2.x (tested on versions 2.6 and 2.7), and the
corresponding version of the Pygame library, freely available online.
Make sure you install the matching (32- or 64-bit) version of Pygame
as your Python installation, and the one compatible with your Python
version number. The last I checked, only 32-bit binaries of Pygame for 
Windows were hosted on the official website, but there was a link to 
download unofficial 64-bit binaries of Pygame as well.


Running the Maze Editor
-----------------------

To run the maze editor in Windows 7 or 8 (64-bit), perform the following
steps depending on your version of Windows:

### Windows 7/8

Copy RICHTX32.OCX and COMDLG32.OCX in the libraries/ folder to
your Windows SysWOW64 directory for 64-bit Windows, or System32 if you 
are running 32-bit Windows. The paths are typically 
"C:\Windows\SysWOW64" and "C:\Windows\System32" respectively.
Second, open command prompt as Administrator (click Start, type "cmd"
and right-click > Run as Administrator) and run the following commands,
making sure to replace the path with the actual path to SysWOW64
(or System32 for 32-bit) on your machine:

   * 64-bit Windows
             regsvr32 C:\Windows\SysWOW64\RICHTX32.OCX
	         regsvr32 C:\Windows\SysWOW64\COMDLG32.OCX
	 
   * 32-bit Windows
			 regsvr32 C:\Windows\System32\RICHTX32.OCX
			 regsvr32 C:\Windows\System32\COMDLG32.OCX

Both commands should say the DLL was registered successfully. After
that, run pacman/maze_editor.exe and you should be good to go.

### Windows XP/Vista

I haven't tested it recently, but as far as I remember, the maze editor
will work if you either copy the .OCX files to System32, or just copy
them to the pacman/ folder (the same folder as maze_editor.exe). I never
ran the regsvr32 command on these versions of Windows, so I don't think
it is necessary. Please contact me if you know!

### Windows under WINE

Haven't tested this combination, but please let me know if you have,
along with any extra steps you had to take, and I will make the changes
to this document.

