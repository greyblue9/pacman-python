Pacman by David Reilly
======================
With contributions by Andy Sommerville (2007)

Github project page:
https://github.com/greyblue9/pacman-python

Personal project page:
http://pinproject.com/pacman/pacman.htm

![Screenshot - 16x16 tile version](/screenshot-1.png)


Thank you for your interest in my tribute to everyone's favorite Pac-Man games
of yesteryear! Pacman features 12 colorful levels of varying sizes, bouncing
fruit, A* path-finding ghosts, and, more recently, cross-platform support,
joystick support, and a high score list. This was the first large-scale project
I ever developed in Python, so the source is not the cleanest, but I would love
to have your contributions! Feel free to slice and dice it up however you like
and redistribute it; I just ask that you give me some credit and let me know
about it, so I can enjoy your creation too! Have fun!


Installation
------------

Pac-man requires **Python 3.x** (tested on 3.8), and the
corresponding version of the Pygame library, freely available online. Make sure
you install the matching (32- or 64-bit) version of Pygame as your Python
installation, and the one compatible with your Python version number. The last I
checked, only 32-bit binaries of Pygame for Windows were hosted on the official
website, but there was a link to download unofficial 64-bit binaries of Pygame
as well.


Running the Maze Editor
-----------------------

![Screenshot - Maze editor](/screenshot-maze-editor-1.png)
![Screenshot - Maze editor level properties](/screenshot-maze-editor-2.png)


To run the maze editor in Windows 7 or 8, perform the following steps depending
on your version of Windows:

### Windows 7/8

Copy RICHTX32.OCX and COMDLG32.OCX in the libraries/ folder to your Windows
SysWOW64 directory for 64-bit Windows, or System32 if you are running 32-bit
Windows. The paths are typically "C:\Windows\SysWOW64" and "C:\Windows\System32"
respectively. Second, open command prompt as **Administrator** (click Start,
type "cmd" and right-click > Run as Administrator) and run the following
commands, making sure to replace the path with the actual path to SysWOW64 (or
System32 for 32-bit) on your machine:

   * 64-bit Windows
   
        regsvr32 C:\Windows\SysWOW64\RICHTX32.OCX
	    regsvr32 C:\Windows\SysWOW64\COMDLG32.OCX
	 
   * 32-bit Windows
   
		regsvr32 C:\Windows\System32\RICHTX32.OCX
		regsvr32 C:\Windows\System32\COMDLG32.OCX

Both commands should say the DLL was registered successfully. After that, run
pacman/maze_editor.exe and you should be good to go. 

### Windows XP/Vista

I haven't tested it recently, but as far as I remember, the maze editor will
work if you either copy the .OCX files to System32, or just copy them to the
pacman/ folder (the same folder as maze_editor.exe). I never ran the regsvr32
command on these versions of Windows, so I don't think it is necessary. Please
contact me if you know!

### Windows under WINE

Haven't tested this combination, but please let me know if you have, along with
any extra steps you had to take, and I will make the changes to this document.

Since the maze editor was written and built in Visual Basic 6, I do have some
suggested steps that I've adapted from another app written in VB6,
[Heirowords](http://home.comcast.net/~thot/Linux.htm). Some of the example
commands are for a Ubuntu system, but a similar command should be available
under any modern Linux system.

   - Install [Wine](http://www.winehq.org/). 
    
     	sudo apt-get install wine

   - Install and run [winetricks](http://wiki.winehq.org/winetricks)
     to download the VB6 runtime libraries for Wine
     
		wget http://kegel.com/wine/winetricks
		chmod +x winetricks
		sudo mv winetricks /usr/bin/winetricks
		winetricks

   - When winetricks opens, select the package "vb6run" and install it.

   - Copy COMDLG32.OCX and RICHTX32.OCX from the libraries/ folder to
     your Wine system32 directory, probably 
     `~/.wine/drive_c/windows/system32/`.

Please let me know if you have any luck with these steps, or have any
suggestions, etc. Good luck and enjoy the maze editor!


Using the Maze Editor
---------------------

### Background and How it Works ###

Skip to the **Quick Start** section below if you're not interested in the
technical details! :)

The maze editor will expect to find a res/ directory in the path to the maze
editor executable; otherwise, it will not work. It should be set up this way
by default. In addition to the contents of the res/tiles/ directory, it also
looks for res/crossref.txt (for tile descriptions and code values), res/sprite
(for fruit graphics), and will default the level open dialog box to the
res/levels directory.

Each level is simply a text file with some level properties at the beginning,
followed by a `# startleveldata` tag which introduces the level data itself,
and then finally an `# endleveldata` tag and tag for `# sprites` (not used).

In the level data section, each line corresponds to one row of tiles on the
screen, and each tile consists of the tile's numeric value as found in
crossref.txt, separated from the next tile by a space character. If you're not
sure what a particular tile number corresponds to, you can look it up in
crossref.txt.

If you know what a tile looks like, but you're not sure of its number, you can
click on it in the maze editor tile pallette, and some information will appear
above, including the tile's name, description, and tile number. The latter is
what gets used in the level description files.

### Quick Start ###

Of course, the simplest way to make a level is just to use the maze editor and
make heavy use of the "x-paintwall" special tile, which will paint walls with
connections to adjacent walls, so you don't have to worry about it. The
"x-paintwall" tile is shown here:

![Screenshot - x-paintwall tile](/screenshot-maze-editor-x-paintwall.png)

Set your level size to have an odd number of columns (x-width), and turn on
"Symmetric editing mode" to make things really efficient. This technique was
used in nearly all levels that come with the game.

Besides the walls themselves, there are several other tile elements required for
a level to work properly. The easiest way to check this is to load an example
level, such as `1.txt`. The required elements are:

   - Maze must be fully closed in, with the exception of doors 
   - Doors must have a "door-h" or "door-v" tile blocking the exit. This tells
     the game to teleport pacman to the corresponding door tag on the other
     side of the board.
   - The ghost box itself must have room for 3 ghosts
	 inside, and must have a ghost-door in the top middle, directly
	 above/outside of which you should place Blinky (the red ghost).
	 Inside, the recommended order is Pinky, Inky, Sue (pink, blue, and
	 orange, respectively). See the ghost box in level `1.txt` for an example.
   - You must place Pacman! When the level starts, he will go to the right.
   - You must place at least one regular pellet.
   - Placing power pellets is recommmended, but not required.

If you want to edit the title screen, it can be found as level `0.txt`. It uses
a special tile to indicate where the game logo should be placed, which is found
in the res/text/ folder (`logo.gif`).

Happy editing! Please send me your creations; I may put level packs together for
the best ones I see and distribute them with the game.

### Contributors ###
* David Reilly https://github.com/greyblue9/pacman-python
* bumjoon.park https://github.com/Wollala/pacman-python
* Hobbledehoy https://github.com/chelyabinsk/pacman-multiplayer

Coming soon
===========

1. Ability to re-use same level sets, maze editor, and most of the res/ directory
between both versions of pacman.

2. Consolidating pacman original (16x16) with pacman-large (24x24) into one
source file, or set of source files.

3. Source for the Maze Editor (written in VB6)


