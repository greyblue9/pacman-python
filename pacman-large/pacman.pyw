#! /usr/bin/python

# pacman.pyw
# By David Reilly

# Code conversion and improvement by https://github.com/Wollala/pacman-python

# Modified by Andy Sommerville, 8 October 2007:
# - Changed hard-coded DOS paths to os.path calls
# - Added constant SCRIPT_PATH (so you don't need to have pacman.pyw and res in your cwd, as long
# -   as those two are in the same directory)
# - Changed text-file reading to accomodate any known EOLn method (\n, \r, or \r\n)
# - I (happily) don't have a Windows box to test this. Blocks marked "WIN???"
# -   should be examined if this doesn't run in Windows
# - Added joystick support (configure by changing JS_* constants)
# - Added a high-score list. Depends on wx for querying the user's name

# Modified by Andy Sommerville, 11 October 2007:
# - Mom's eyes aren't what they used to be, so I'm switching 16x16 tiles to 24x24
#   Added constants TILE_WIDTH,TILE_HEIGHT to make this easier to change later.

import os
import random
import sys

import pygame
from pygame.locals import *

# WIN???
if os.name == "nt":
    SCRIPT_PATH = os.getcwd()
else:
    SCRIPT_PATH = sys.path[0]    

SCREEN_TILE_SIZE_HEIGHT = 23
SCREEN_TILE_SIZE_WIDTH = 30

TILE_WIDTH = TILE_HEIGHT = 24

# constants for the high-score display
HS_FONT_SIZE = 16
HS_LINE_HEIGHT = 16
HS_WIDTH = 408
HS_HEIGHT = 120
HS_XOFFSET = 180
HS_YOFFSET = 400
HS_ALPHA = 200

# new constants for the score's position
SCORE_XOFFSET = 50  # pixels from left edge
SCORE_YOFFSET = 34  # pixels from bottom edge (to top of score)
SCORE_COLWIDTH = 13  # width of each character

# See GetCrossRef() -- where these colors occur in a GIF, they are replaced according to the level file
IMG_EDGE_LIGHT_COLOR = (255, 206, 255, 255)
IMG_FILL_COLOR = (132, 0, 132, 255)
IMG_EDGE_SHADOW_COLOR = (255, 0, 255, 255)
IMG_PELLET_COLOR = (128, 0, 128, 255)

# NO_GIF_TILES -- tile numbers which do not correspond to a GIF file
# currently only "23" for the high-score list
NO_GIF_TILES = [23]

NO_WX = 0  # if set, the high-score code will not attempt to ask the user his name
USER_NAME = "User"  # USER_NAME=os.getlogin() # the default user name if wx fails to load or NO_WX

# Joystick defaults - maybe add a Preferences dialog in the future?
JS_DEVNUM = 0  # device 0 (pygame joysticks always start at 0). if JS_DEVNUM is not a valid device, will use 0
JS_XAXIS = 0  # axis 0 for left/right (default for most joysticks)
JS_YAXIS = 1  # axis 1 for up/down (default for most joysticks)
JS_STARTBUTTON = 0  # button number to start the game. this is a matter of personal preference, and will vary from device to device


def get_image_surface(file_path):
    image = pygame.image.load(file_path).convert()
    # image_rect = image.get_rect()
    # image_surface = pygame.Surface((image_rect.width, image_rect.height))
    # image_surface.blit(image, image_rect)
    return image


# Must come before pygame.init()
pygame.mixer.pre_init(22050, -16, 1, 1024)
pygame.mixer.init()
pygame.mixer.set_num_channels(7)
channel_backgound = pygame.mixer.Channel(6)

clock = pygame.time.Clock()
pygame.init()

window = pygame.display.set_mode((1, 1))
pygame.display.set_caption("Pacman")

screen = pygame.display.get_surface()
pygame.mouse.set_visible(False)

img_Background = get_image_surface(os.path.join(SCRIPT_PATH, "res", "backgrounds", "1.gif"))

snd_pellet = {
    0: pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "pellet1.wav")),
    1: pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "pellet2.wav"))}
snd_levelintro = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "levelintro.wav"))
snd_default = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "default.wav"))
snd_extrapac = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "extrapac.wav"))
snd_gh2gohome = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "gh2gohome.wav"))
snd_death = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "death.wav"))
snd_powerpellet = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "powerpellet.wav"))
snd_eatgh = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "eatgh2.wav"))
snd_fruitbounce = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "fruitbounce.wav"))
snd_eatfruit = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "eatfruit.wav"))
snd_extralife = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "extralife.wav"))
snd_love = pygame.mixer.Sound(os.path.join(SCRIPT_PATH, "res", "sounds", "All You Need Is Love.wav"))

ghostcolor = {
    0: (255, 0, 0, 255),
    1: (255, 128, 255, 255),
    2: (128, 255, 255, 255),
    3: (255, 128, 0, 255),
    4: (50, 50, 255, 255),
    5: (255, 255, 255, 255)}

rect_list = []  # rect list for drawing


#      ___________________
# ___/  class definitions  \_______________________________________________

class game:

    def __init__(self):
        self.levelNum = 0
        self.score = 0
        self.lives = 3

        self.mode = 0
        self.modeTimer = 0
        self.ghostTimer = 0
        self.ghostValue = 0
        self.fruitTimer = 0
        self.fruitScoreTimer = 0
        self.fruitScorePos = (0, 0)

        self.SetMode(3)

        # camera variables
        self.screenTileSize = (SCREEN_TILE_SIZE_HEIGHT, SCREEN_TILE_SIZE_WIDTH)
        self.screenSize = (self.screenTileSize[1] * TILE_WIDTH, self.screenTileSize[0] * TILE_HEIGHT)

        self.screenPixelPos = (0, 0)  # absolute x,y position of the screen from the upper-left corner of the level
        self.screenNearestTilePos = (0, 0)  # nearest-tile position of the screen from the UL corner
        self.screenPixelOffset = (0, 0)  # offset in pixels of the screen from its nearest-tile position

        # numerical display digits
        self.digit = {}
        for i in range(0, 10, 1):
            self.digit[i] = get_image_surface(os.path.join(SCRIPT_PATH, "res", "text", str(i) + ".gif"))
        self.imLife = get_image_surface(os.path.join(SCRIPT_PATH, "res", "text", "life.gif"))
        self.imGameOver = get_image_surface(os.path.join(SCRIPT_PATH, "res", "text", "gameover.gif"))
        self.imReady = get_image_surface(os.path.join(SCRIPT_PATH, "res", "text", "ready.gif"))
        self.imLogo = get_image_surface(os.path.join(SCRIPT_PATH, "res", "text", "logo.gif"))
        self.imHiscores = self.makehiscorelist()

    @staticmethod
    def defaulthiscorelist():
        return [(100000, "David"), (80000, "Andy"), (60000, "Count Pacula"), (40000, "Cleopacra"),
                (20000, "Brett Favre"), (10000, "Sergei Pachmaninoff")]

    @staticmethod
    def writehiscores(hs):
        """Given a new list, write it to the default file."""
        fname = os.path.join(SCRIPT_PATH, "res", "hiscore.txt")
        f = open(fname, "w")
        for line in hs:
            f.write(str(line[0]) + " " + line[1] + "\n")
        f.close()

    @staticmethod
    def getplayername():
        """Ask the player his name, to go on the high-score list."""
        if NO_WX:
            return USER_NAME
        # noinspection PyBroadException
        try:
            import wx
        except:
            print("""Pacman Error: No module wx. Can not ask the user his name!
                     :(       Download wx from http://www.wxpython.org/"
                     :(       To avoid seeing this error again, set NO_WX in file pacman.pyw.""")
            return USER_NAME
        app = wx.App(None)
        dlog = wx.TextEntryDialog(None, "You made the high-score list! Name:")
        dlog.ShowModal()
        name = dlog.GetValue()
        dlog.Destroy()
        app.Destroy()
        return name

    @staticmethod
    def PlayBackgoundSound(snd):
        channel_backgound.stop()
        channel_backgound.play(snd, loops=-1)

    def gethiscores(self):
        """If res/hiscore.txt exists, read it. If not, return the default high scores.
           Output is [ (score,name) , (score,name) , .. ]. Always 6 entries."""
        try:
            f = open(os.path.join(SCRIPT_PATH, "res", "hiscore.txt"))
            hs = []
            for line in f:
                while len(line) > 0 and (line[0] == "\n" or line[0] == "\r"):
                    line = line[1:]
                while len(line) > 0 and (line[-1] == "\n" or line[-1] == "\r"):
                    line = line[:-1]
                score = int(line.split(" ")[0])
                name = line.partition(" ")[2]
                if score > 99999999: score = 99999999
                if len(name) > 22: name = name[:22]
                hs.append((score, name))
            f.close()
            if len(hs) > 6: hs = hs[:6]
            while len(hs) < 6: hs.append((0, ""))
            return hs
        except IOError:
            return self.defaulthiscorelist()

    def updatehiscores(self, newscore):
        """Add newscore to the high score list, if appropriate."""
        hs = self.gethiscores()
        for line in hs:
            if newscore >= line[0]:
                hs.insert(hs.index(line), (newscore, self.getplayername()))
                hs.pop(-1)
                break
        self.writehiscores(hs)

    def makehiscorelist(self):
        global rect_list
        "Read the High-Score file and convert it to a useable Surface."
        # My apologies for all the hard-coded constants.... -Andy
        f = pygame.font.Font(os.path.join(SCRIPT_PATH, "res", "zig_____.ttf"), HS_FONT_SIZE)
        scoresurf = pygame.Surface((HS_WIDTH, HS_HEIGHT), pygame.SRCALPHA)
        scoresurf.set_alpha(HS_ALPHA)
        linesurf = f.render("HIGH SCORES".center(28), 1, (255, 255, 0))
        scoresurf.blit(linesurf, (0, 0))
        hs = self.gethiscores()
        vpos = 0
        for line in hs:
            vpos += HS_LINE_HEIGHT
            linesurf = f.render(line[1].ljust(18) + str(line[0]).rjust(10), 1, (255, 255, 255))
            scoresurf.blit(linesurf, (0, vpos))
        return scoresurf

    def drawmidgamehiscores(self):
        """Redraw the high-score list image after pacman dies."""
        self.imHiscores = self.makehiscorelist()

    def StartNewGame(self):
        self.levelNum = 1
        self.score = 0
        self.lives = 3

        self.SetMode(0)
        thisLevel.LoadLevel(thisGame.GetLevelNum())

    def AddToScore(self, amount):
        extraLifeSet = [25000, 50000, 100000, 150000]

        for specialScore in extraLifeSet:
            if self.score < specialScore <= self.score + amount:
                snd_extralife.play()
                thisGame.lives += 1

        self.score += amount

    def DrawScore(self):
        global rect_list
        self.DrawNumber(self.score, (SCORE_XOFFSET, self.screenSize[1] - SCORE_YOFFSET))

        for i in range(0, self.lives, 1):
            screen.blit(self.imLife, (34 + i * 10 + 16, self.screenSize[1] - 18))

        screen.blit(thisFruit.imFruit[thisFruit.fruitType], (4 + 16, self.screenSize[1] - 28))

        if self.mode == 3:
            screen.blit(self.imGameOver, (self.screenSize[0] / 2 - (self.imGameOver.get_width() / 2),
                                          self.screenSize[1] / 2 - (self.imGameOver.get_height() / 2)))
        elif self.mode == 0 or self.mode == 4:
            screen.blit(self.imReady, (self.screenSize[0] / 2 - 20, self.screenSize[1] / 2 + 12))

        self.DrawNumber(self.levelNum, (0, self.screenSize[1] - 20))

    def DrawNumber(self, number, x_y):
        global rect_list
        (x, y) = x_y

        strNumber = str(number)

        for i in range(0, len(str(number)), 1):
            if strNumber[i] == '.':
                break
            iDigit = int(strNumber[i])
            screen.blit(self.digit[iDigit], (x + i * SCORE_COLWIDTH, y))

    def SmartMoveScreen(self):
        possibleScreenX = player.x - self.screenTileSize[1] / 2 * TILE_WIDTH
        possibleScreenY = player.y - self.screenTileSize[0] / 2 * TILE_HEIGHT

        if self.screenSize[0] >= thisLevel.lvlWidth * TILE_WIDTH:
            possibleScreenX = -(self.screenSize[0] - thisLevel.lvlWidth * TILE_WIDTH) / 2
        elif possibleScreenX < 0:
            possibleScreenX = 0
        elif possibleScreenX > thisLevel.lvlWidth * TILE_WIDTH - self.screenSize[0]:
            possibleScreenX = thisLevel.lvlWidth * TILE_WIDTH - self.screenSize[0]

        if self.screenSize[1] >= thisLevel.lvlHeight * TILE_HEIGHT:
            possibleScreenY = -(self.screenSize[1] - thisLevel.lvlHeight * TILE_HEIGHT) / 2
        elif possibleScreenY < 0:
            possibleScreenY = 0
        elif possibleScreenY > thisLevel.lvlHeight * TILE_HEIGHT - self.screenSize[1]:
            possibleScreenY = thisLevel.lvlHeight * TILE_HEIGHT - self.screenSize[1]

        thisGame.MoveScreen((possibleScreenX, possibleScreenY))

    def MoveScreen(self, newX_newY):
        (newX, newY) = newX_newY
        self.screenPixelPos = (newX, newY)
        self.screenNearestTilePos = (
            int(newY / TILE_HEIGHT), int(newX / TILE_WIDTH))  # nearest-tile position of the screen from the UL corner
        self.screenPixelOffset = (
            newX - self.screenNearestTilePos[1] * TILE_WIDTH, newY - self.screenNearestTilePos[0] * TILE_HEIGHT)

    def GetScreenPos(self):
        return self.screenPixelPos

    def GetLevelNum(self):
        return self.levelNum

    def SetNextLevel(self):
        self.levelNum += 1

        self.SetMode(0)
        thisLevel.LoadLevel(thisGame.GetLevelNum())

        player.velX = 0
        player.velY = 0
        player.anim_pacmanCurrent = player.anim_pacmanS

    def SetMode(self, newMode):
        self.mode = newMode
        self.modeTimer = 0

        if newMode == 0:
            self.PlayBackgoundSound(snd_levelintro)
        elif newMode == 1:
            self.PlayBackgoundSound(snd_default)
        elif newMode == 2:
            self.PlayBackgoundSound(snd_death)
        elif newMode == 8:
            self.PlayBackgoundSound(snd_gh2gohome)
        elif newMode == 9:
            self.PlayBackgoundSound(snd_extrapac)
        elif newMode == 11:
            self.PlayBackgoundSound(snd_love)
        else:
            channel_backgound.stop()


class node:
    def __init__(self):
        self.g = -1  # movement cost to move from previous node to this one (usually +10)
        self.h = -1  # estimated movement cost to move from this node to the ending node (remaining horizontal and vertical steps * 10)
        self.f = -1  # total movement cost of this node (= g + h)
        # parent node - used to trace path back to the starting node at the end
        self.parent = (-1, -1)
        # node type - 0 for empty space, 1 for wall (optionally, 2 for starting node and 3 for end)
        self.type = -1


class path_finder:
    def __init__(self):
        # map is a 1-DIMENSIONAL array.
        # use the Unfold( (row, col) ) function to convert a 2D coordinate pair
        # into a 1D index to use with this array.
        self.map = {}
        self.size = (-1, -1)  # rows by columns

        self.pathChainRev = ""
        self.pathChain = ""

        # starting and ending nodes
        self.start = (-1, -1)
        self.end = (-1, -1)

        # current node (used by algorithm)
        self.current = (-1, -1)

        # open and closed lists of nodes to consider (used by algorithm)
        self.openList = []
        self.closedList = []

        # used in algorithm (adjacent neighbors path finder is allowed to consider)
        self.neighborSet = [(0, -1), (0, 1), (-1, 0), (1, 0)]

    def ResizeMap(self, numRows_numCols):
        (numRows, numCols) = numRows_numCols
        self.map = {}
        self.size = (numRows, numCols)

        # initialize path_finder map to a 2D array of empty nodes
        for row in range(0, self.size[0], 1):
            for col in range(0, self.size[1], 1):
                self.Set((row, col), node())
                self.SetType((row, col), 0)

    def CleanUpTemp(self):
        # this resets variables needed for a search (but preserves the same map / maze)
        self.pathChainRev = ""
        self.pathChain = ""
        self.current = (-1, -1)
        self.openList = []
        self.closedList = []

    def FindPath(self, startPos, endPos):
        self.CleanUpTemp()
        # (row, col) tuples
        self.start = startPos
        self.end = endPos

        # add start node to open list
        self.AddToOpenList(self.start)
        self.SetG(self.start, 0)
        self.SetH(self.start, 0)
        self.SetF(self.start, 0)

        thisLowestFNode = None
        doContinue = True
        while doContinue:

            thisLowestFNode = self.GetLowestFNode()

            if thisLowestFNode != self.end and thisLowestFNode != False:
                self.current = thisLowestFNode
                self.RemoveFromOpenList(self.current)
                self.AddToClosedList(self.current)

                for offset in self.neighborSet:
                    thisNeighbor = (self.current[0] + offset[0], self.current[1] + offset[1])

                    if not thisNeighbor[0] < 0 and not thisNeighbor[1] < 0 and not thisNeighbor[0] > self.size[
                        0] - 1 and not thisNeighbor[1] > self.size[1] - 1 and not self.GetType(thisNeighbor) == 1:
                        cost = self.GetG(self.current) + 10

                        if self.IsInOpenList(thisNeighbor) and cost < self.GetG(thisNeighbor):
                            self.RemoveFromOpenList(thisNeighbor)

                        # if self.IsInClosedList( thisNeighbor ) and cost < self.GetG( thisNeighbor ):
                        #	self.RemoveFromClosedList( thisNeighbor )

                        if not self.IsInOpenList(thisNeighbor) and not self.IsInClosedList(thisNeighbor):
                            self.AddToOpenList(thisNeighbor)
                            self.SetG(thisNeighbor, cost)
                            self.CalcH(thisNeighbor)
                            self.CalcF(thisNeighbor)
                            self.SetParent(thisNeighbor, self.current)
            else:
                doContinue = False

        if not thisLowestFNode:
            return False

        # reconstruct path
        self.current = self.end
        while not self.current == self.start:
            # build a string representation of the path using R, L, D, U
            if self.current[1] > self.GetParent(self.current)[1]:
                self.pathChainRev += 'R'
            elif self.current[1] < self.GetParent(self.current)[1]:
                self.pathChainRev += 'L'
            elif self.current[0] > self.GetParent(self.current)[0]:
                self.pathChainRev += 'D'
            elif self.current[0] < self.GetParent(self.current)[0]:
                self.pathChainRev += 'U'
            self.current = self.GetParent(self.current)
            self.SetType(self.current, 4)

        # because pathChainRev was constructed in reverse order, it needs to be reversed!
        for i in range(len(self.pathChainRev) - 1, -1, -1):
            self.pathChain += self.pathChainRev[i]

        # set start and ending positions for future reference
        self.SetType(self.start, 2)
        self.SetType(self.end, 3)

        return self.pathChain

    def Unfold(self, row_col):
        # this function converts a 2D array coordinate pair (row, col)
        # to a 1D-array index, for the object's 1D map array.
        (row, col) = row_col
        return (row * self.size[1]) + col

    def Set(self, row_col, newNode):
        # sets the value of a particular map cell (usually refers to a node object)
        (row, col) = row_col
        self.map[self.Unfold((row, col))] = newNode

    def GetType(self, row_col):
        (row, col) = row_col
        return self.map[self.Unfold((row, col))].type

    def SetType(self, row_col, newValue):
        (row, col) = row_col
        self.map[self.Unfold((row, col))].type = newValue

    def GetF(self, row_col):
        (row, col) = row_col
        return self.map[self.Unfold((row, col))].f

    def GetG(self, row_col):
        (row, col) = row_col
        return self.map[self.Unfold((row, col))].g

    def GetH(self, row_col):
        (row, col) = row_col
        return self.map[self.Unfold((row, col))].h

    def SetG(self, row_col, newValue):
        (row, col) = row_col
        self.map[self.Unfold((row, col))].g = newValue

    def SetH(self, row_col, newValue):
        (row, col) = row_col
        self.map[self.Unfold((row, col))].h = newValue

    def SetF(self, row_col, newValue):
        (row, col) = row_col
        self.map[self.Unfold((row, col))].f = newValue

    def CalcH(self, row_col):
        (row, col) = row_col
        self.map[self.Unfold((row, col))].h = abs(row - self.end[0]) + abs(col - self.end[0])

    def CalcF(self, row_col):
        (row, col) = row_col
        unfoldIndex = self.Unfold((row, col))
        self.map[unfoldIndex].f = self.map[unfoldIndex].g + self.map[unfoldIndex].h

    def AddToOpenList(self, row_col):
        (row, col) = row_col
        self.openList.append((row, col))

    def RemoveFromOpenList(self, row_col):
        (row, col) = row_col
        self.openList.remove((row, col))

    def IsInOpenList(self, row_col):
        (row, col) = row_col
        if self.openList.count((row, col)) > 0:
            return True
        else:
            return False

    def GetLowestFNode(self):
        lowestValue = 1000  # start arbitrarily high
        lowestPair = (-1, -1)

        for iOrderedPair in self.openList:
            if self.GetF(iOrderedPair) < lowestValue:
                lowestValue = self.GetF(iOrderedPair)
                lowestPair = iOrderedPair

        if not lowestPair == (-1, -1):
            return lowestPair
        else:
            return False

    def AddToClosedList(self, row_col):
        (row, col) = row_col
        self.closedList.append((row, col))

    def IsInClosedList(self, row_col):
        (row, col) = row_col
        if self.closedList.count((row, col)) > 0:
            return True
        else:
            return False

    def SetParent(self, row_col, parentRow_parentCol):
        (row, col) = row_col
        (parentRow, parentCol) = parentRow_parentCol
        self.map[self.Unfold((row, col))].parent = (parentRow, parentCol)

    def GetParent(self, row_col):
        (row, col) = row_col
        return self.map[self.Unfold((row, col))].parent

    def draw(self):
        global rect_list
        for row in range(0, self.size[0], 1):
            for col in range(0, self.size[1], 1):
                thisTile = self.GetType((row, col))
                screen.blit(tileIDImage[thisTile], (col * (TILE_WIDTH * 2), row * (TILE_WIDTH * 2)))


class ghost:
    def __init__(self, ghostID):
        self.x = 0
        self.y = 0
        self.velX = 0
        self.velY = 0
        self.speed = 2

        self.nearestRow = 0
        self.nearestCol = 0

        self.id = ghostID

        # ghost "state" variable
        # 1 = normal
        # 2 = vulnerable
        # 3 = spectacles
        self.state = 1

        self.homeX = 0
        self.homeY = 0

        self.currentPath = ""

        self.anim = {}
        for i in range(1, 7, 1):
            self.anim[i] = get_image_surface(
                os.path.join(SCRIPT_PATH, "res", "sprite", "ghost " + str(i) + ".gif"))

            # change the ghost color in this frame
            for y in range(0, TILE_HEIGHT, 1):
                for x in range(0, TILE_WIDTH, 1):

                    if self.anim[i].get_at((x, y)) == (255, 0, 0, 255):
                        # default, red ghost body color
                        self.anim[i].set_at((x, y), ghostcolor[self.id])

        self.animFrame = 1
        self.animDelay = 0

    def Draw(self):
        global rect_list
        pupilSet = None

        if thisGame.mode == 3:
            return False

        # ghost eyes --
        for y in range(6, 12, 1):
            for x in [5, 6, 8, 9]:
                self.anim[self.animFrame].set_at((x, y), (248, 248, 248, 255))
                self.anim[self.animFrame].set_at((x + 9, y), (248, 248, 248, 255))

                if player.x > self.x and player.y > self.y:
                    # player is to lower-right
                    pupilSet = (8, 9)
                elif player.x < self.x and player.y > self.y:
                    # player is to lower-left
                    pupilSet = (5, 9)
                elif player.x > self.x and player.y < self.y:
                    # player is to upper-right
                    pupilSet = (8, 6)
                elif player.x < self.x and player.y < self.y:
                    # player is to upper-left
                    pupilSet = (5, 6)
                else:
                    pupilSet = (5, 9)

        for y in range(pupilSet[1], pupilSet[1] + 3, 1):
            for x in range(pupilSet[0], pupilSet[0] + 2, 1):
                self.anim[self.animFrame].set_at((x, y), (0, 0, 255, 255))
                self.anim[self.animFrame].set_at((x + 9, y), (0, 0, 255, 255))
        # -- end ghost eyes

        if self.state == 1:
            # draw regular ghost (this one)
            screen.blit(self.anim[self.animFrame],
                        (self.x - thisGame.screenPixelPos[0], self.y - thisGame.screenPixelPos[1]))
        elif self.state == 2:
            # draw vulnerable ghost

            if thisGame.ghostTimer > 100:
                # blue
                screen.blit(ghosts[4].anim[self.animFrame],
                            (self.x - thisGame.screenPixelPos[0], self.y - thisGame.screenPixelPos[1]))
            else:
                # blue/white flashing
                tempTimerI = int(thisGame.ghostTimer / 10)
                if tempTimerI == 1 or tempTimerI == 3 or tempTimerI == 5 or tempTimerI == 7 or tempTimerI == 9:
                    screen.blit(ghosts[5].anim[self.animFrame],
                                (self.x - thisGame.screenPixelPos[0], self.y - thisGame.screenPixelPos[1]))
                else:
                    screen.blit(ghosts[4].anim[self.animFrame],
                                (self.x - thisGame.screenPixelPos[0], self.y - thisGame.screenPixelPos[1]))

        elif self.state == 3:
            # draw glasses
            screen.blit(tileIDImage[tileID['glasses']],
                        (self.x - thisGame.screenPixelPos[0], self.y - thisGame.screenPixelPos[1]))

        if thisGame.mode == 6 or thisGame.mode == 7:
            # don't animate ghost if the level is complete
            return False

        self.animDelay += 1

        if self.animDelay == 2:
            self.animFrame += 1

            if self.animFrame == 7:
                # wrap to beginning
                self.animFrame = 1

            self.animDelay = 0

    def Move(self):
        if (thisGame.levelNum == 2):
            return
        self.x += self.velX
        self.y += self.velY

        self.nearestRow = int(((self.y + (TILE_HEIGHT / 2)) / TILE_HEIGHT))
        self.nearestCol = int(((self.x + (TILE_HEIGHT / 2)) / TILE_WIDTH))

        if (self.x % TILE_WIDTH) == 0 and (self.y % TILE_HEIGHT) == 0:
            # if the ghost is lined up with the grid again
            # meaning, it's time to go to the next path item

            if self.currentPath is not False and (len(self.currentPath) > 0):
                self.currentPath = self.currentPath[1:]
                self.FollowNextPathWay()

            else:
                self.x = self.nearestCol * TILE_WIDTH
                self.y = self.nearestRow * TILE_HEIGHT

                # chase pac-man
                self.currentPath = path.FindPath((self.nearestRow, self.nearestCol),
                                                 (player.nearestRow, player.nearestCol))
                self.FollowNextPathWay()

    def FollowNextPathWay(self):
        # print "Ghost " + str(self.id) + " rem: " + self.currentPath
        # only follow this pathway if there is a possible path found!
        if not self.currentPath == False:

            if len(self.currentPath) > 0:
                if self.currentPath[0] == "L":
                    (self.velX, self.velY) = (-self.speed, 0)
                elif self.currentPath[0] == "R":
                    (self.velX, self.velY) = (self.speed, 0)
                elif self.currentPath[0] == "U":
                    (self.velX, self.velY) = (0, -self.speed)
                elif self.currentPath[0] == "D":
                    (self.velX, self.velY) = (0, self.speed)

            else:
                # this ghost has reached his destination!!
                if not self.state == 3:
                    # chase pac-man
                    self.currentPath = path.FindPath((self.nearestRow, self.nearestCol),
                                                     (player.nearestRow, player.nearestCol))
                    self.FollowNextPathWay()

                else:
                    # glasses found way back to ghost box
                    self.state = 1
                    self.speed = self.speed / 4

                    # give ghost a path to a random spot (containing a pellet)
                    (randRow, randCol) = (0, 0)

                    while not thisLevel.GetMapTile((randRow, randCol)) == tileID['pellet'] or (randRow, randCol) == (
                            0, 0):
                        randRow = random.randint(1, thisLevel.lvlHeight - 2)
                        randCol = random.randint(1, thisLevel.lvlWidth - 2)

                    self.currentPath = path.FindPath((self.nearestRow, self.nearestCol), (randRow, randCol))
                    self.FollowNextPathWay()


class fruit:
    def __init__(self):
        # when fruit is not in use, it's in the (-1, -1) position off-screen.
        self.slowTimer = 0
        self.x = -TILE_WIDTH
        self.y = -TILE_HEIGHT
        self.velX = 0
        self.velY = 0
        self.speed = 2
        self.active = False

        self.bouncei = 0
        self.bounceY = 0

        self.nearestRow = (-1, -1)
        self.nearestCol = (-1, -1)

        self.imFruit = {}
        for i in range(0, 5, 1):
            self.imFruit[i] = get_image_surface(
                os.path.join(SCRIPT_PATH, "res", "sprite", "fruit " + str(i) + ".gif"))

        self.currentPath = ""
        self.fruitType = 1

    def Draw(self):
        global rect_list
        if thisGame.mode == 3 or self.active == False:
            return False

        screen.blit(self.imFruit[self.fruitType],
                    (self.x - thisGame.screenPixelPos[0], self.y - thisGame.screenPixelPos[1] - self.bounceY))

    def Move(self):
        if not self.active:
            return False

        self.bouncei += 1
        if self.bouncei == 1:
            self.bounceY = 2
        elif self.bouncei == 2:
            self.bounceY = 4
        elif self.bouncei == 3:
            self.bounceY = 5
        elif self.bouncei == 4:
            self.bounceY = 5
        elif self.bouncei == 5:
            self.bounceY = 6
        elif self.bouncei == 6:
            self.bounceY = 6
        elif self.bouncei == 9:
            self.bounceY = 6
        elif self.bouncei == 10:
            self.bounceY = 5
        elif self.bouncei == 11:
            self.bounceY = 5
        elif self.bouncei == 12:
            self.bounceY = 4
        elif self.bouncei == 13:
            self.bounceY = 3
        elif self.bouncei == 14:
            self.bounceY = 2
        elif self.bouncei == 15:
            self.bounceY = 1
        elif self.bouncei == 16:
            self.bounceY = 0
            self.bouncei = 0
            snd_fruitbounce.play()

        self.slowTimer += 1
        if self.slowTimer == 2:
            self.slowTimer = 0

            self.x += self.velX
            self.y += self.velY

            self.nearestRow = int(((self.y + (TILE_WIDTH / 2)) / TILE_WIDTH))
            self.nearestCol = int(((self.x + (TILE_HEIGHT / 2)) / TILE_HEIGHT))

            if (self.x % TILE_WIDTH) == 0 and (self.y % TILE_HEIGHT) == 0:
                # if the fruit is lined up with the grid again
                # meaning, it's time to go to the next path item

                if len(self.currentPath) > 0:
                    self.currentPath = self.currentPath[1:]
                    self.FollowNextPathWay()

                else:
                    self.x = self.nearestCol * TILE_WIDTH
                    self.y = self.nearestRow * TILE_HEIGHT

                    self.active = False
                    thisGame.fruitTimer = 0

    def FollowNextPathWay(self):
        # only follow this pathway if there is a possible path found!
        if not self.currentPath == False:

            if len(self.currentPath) > 0:
                if self.currentPath[0] == "L":
                    (self.velX, self.velY) = (-self.speed, 0)
                elif self.currentPath[0] == "R":
                    (self.velX, self.velY) = (self.speed, 0)
                elif self.currentPath[0] == "U":
                    (self.velX, self.velY) = (0, -self.speed)
                elif self.currentPath[0] == "D":
                    (self.velX, self.velY) = (0, self.speed)


class pacman:
    def __init__(self):
        self.x = 0
        self.y = 0
        self.velX = 0
        self.velY = 0
        self.speed = 3

        self.nearestRow = 0
        self.nearestCol = 0

        self.homeX = 0
        self.homeY = 0

        self.anim_pacmanL = {}
        self.anim_pacmanR = {}
        self.anim_pacmanU = {}
        self.anim_pacmanD = {}
        self.anim_pacmanS = {}
        self.anim_pacmanCurrent = {}

        self.animFrame = 1

        for i in range(1, 9, 1):
            self.anim_pacmanL[i] = get_image_surface(
                os.path.join(SCRIPT_PATH, "res", "sprite", "pacman-l " + str(i) + ".gif"))
            self.anim_pacmanR[i] = get_image_surface(
                os.path.join(SCRIPT_PATH, "res", "sprite", "pacman-r " + str(i) + ".gif"))
            self.anim_pacmanU[i] = get_image_surface(
                os.path.join(SCRIPT_PATH, "res", "sprite", "pacman-u " + str(i) + ".gif"))
            self.anim_pacmanD[i] = get_image_surface(
                os.path.join(SCRIPT_PATH, "res", "sprite", "pacman-d " + str(i) + ".gif"))
            self.anim_pacmanS[i] = get_image_surface(os.path.join(SCRIPT_PATH, "res", "sprite", "pacman.gif"))

        self.pelletSndNum = 0

    def Move(self):
        self.nearestRow = int(((self.y + (TILE_WIDTH / 2)) / TILE_WIDTH))
        self.nearestCol = int(((self.x + (TILE_HEIGHT / 2)) / TILE_HEIGHT))

        # make sure the current velocity will not cause a collision before moving
        if not thisLevel.CheckIfHitWall((self.x + self.velX, self.y + self.velY), (self.nearestRow, self.nearestCol)):
            # it's ok to Move
            self.x += self.velX
            self.y += self.velY

            # check for collisions with other tiles (pellets, etc)
            thisLevel.CheckIfHitSomething((self.x, self.y), (self.nearestRow, self.nearestCol))

            # check for collisions with the ghosts
            for i in range(0, 4, 1):
                if thisLevel.CheckIfHit((self.x, self.y), (ghosts[i].x, ghosts[i].y), TILE_WIDTH / 2):
                    # hit a ghost

                    if ghosts[i].state == 1:
                        # ghost is normal
                        thisGame.SetMode(2)

                    elif ghosts[i].state == 2:
                        # ghost is vulnerable
                        # give them glasses
                        # make them run
                        thisGame.AddToScore(thisGame.ghostValue)
                        thisGame.ghostValue = thisGame.ghostValue * 2
                        snd_eatgh.play()

                        ghosts[i].state = 3
                        ghosts[i].speed = ghosts[i].speed * 4
                        # and send them to the ghost box
                        ghosts[i].x = ghosts[i].nearestCol * TILE_WIDTH
                        ghosts[i].y = ghosts[i].nearestRow * TILE_HEIGHT
                        ghosts[i].currentPath = path.FindPath((ghosts[i].nearestRow, ghosts[i].nearestCol), (
                            thisLevel.GetGhostBoxPos()[0] + 1, thisLevel.GetGhostBoxPos()[1]))
                        ghosts[i].FollowNextPathWay()

                        # set game mode to brief pause after eating
                        thisGame.SetMode(5)

            # check for collisions with the fruit
            if thisFruit.active:
                if thisLevel.CheckIfHit((self.x, self.y), (thisFruit.x, thisFruit.y), TILE_WIDTH / 2):
                    thisGame.AddToScore(2500)
                    thisFruit.active = False
                    thisGame.fruitTimer = 0
                    thisGame.fruitScoreTimer = 80
                    snd_eatfruit.play()

        else:
            # we're going to hit a wall -- stop moving
            self.velX = 0
            self.velY = 0

        # deal with power-pellet ghost timer
        if thisGame.ghostTimer > 0:
            thisGame.ghostTimer -= 1

            if thisGame.ghostTimer == 0:
                thisGame.PlayBackgoundSound(snd_default)
                for i in range(0, 4, 1):
                    if ghosts[i].state == 2:
                        ghosts[i].state = 1
                thisGame.ghostValue = 0

        # deal with fruit timer
        thisGame.fruitTimer += 1
        if thisGame.fruitTimer == 380:
            pathwayPair = thisLevel.GetPathwayPairPos()

            if not pathwayPair == False:
                pathwayEntrance = pathwayPair[0]
                pathwayExit = pathwayPair[1]

                thisFruit.active = True

                thisFruit.nearestRow = pathwayEntrance[0]
                thisFruit.nearestCol = pathwayEntrance[1]

                thisFruit.x = thisFruit.nearestCol * TILE_WIDTH
                thisFruit.y = thisFruit.nearestRow * TILE_HEIGHT

                thisFruit.currentPath = path.FindPath((thisFruit.nearestRow, thisFruit.nearestCol), pathwayExit)
                thisFruit.FollowNextPathWay()

        if thisGame.fruitScoreTimer > 0:
            thisGame.fruitScoreTimer -= 1

    def Draw(self):
        global rect_list
        if thisGame.mode == 3:
            return False

        # set the current frame array to match the direction pacman is facing
        if self.velX > 0:
            self.anim_pacmanCurrent = self.anim_pacmanR
        elif self.velX < 0:
            self.anim_pacmanCurrent = self.anim_pacmanL
        elif self.velY > 0:
            self.anim_pacmanCurrent = self.anim_pacmanD
        elif self.velY < 0:
            self.anim_pacmanCurrent = self.anim_pacmanU

        screen.blit(self.anim_pacmanCurrent[self.animFrame],
                    (self.x - thisGame.screenPixelPos[0], self.y - thisGame.screenPixelPos[1]))

        if thisGame.mode == 1 or thisGame.mode == 8 or thisGame.mode == 9:
            if self.velX != 0 or self.velY != 0:
                # only Move mouth when pacman is moving
                self.animFrame += 1

            if self.animFrame == 9:
                # wrap to beginning
                self.animFrame = 1


class level:
    def __init__(self):
        self.lvlWidth = 0
        self.lvlHeight = 0
        self.edgeLightColor = (255, 255, 0, 255)
        self.edgeShadowColor = (255, 150, 0, 255)
        self.fillColor = (0, 255, 255, 255)
        self.pelletColor = (255, 255, 255, 255)

        self.map = {}

        self.pellets = 0
        self.powerPelletBlinkTimer = 0

    def SetMapTile(self, row_col, newValue):
        (row, col) = row_col
        self.map[(row * self.lvlWidth) + col] = newValue

    def GetMapTile(self, row_col):
        (row, col) = row_col
        if 0 <= row < self.lvlHeight and 0 <= col < self.lvlWidth:
            return self.map[(row * self.lvlWidth) + col]
        else:
            return 0

    @staticmethod
    def IsWall(row_col):
        (row, col) = row_col
        if row > thisLevel.lvlHeight - 1 or row < 0:
            return True

        if col > thisLevel.lvlWidth - 1 or col < 0:
            return True

        # check the offending tile ID
        result = thisLevel.GetMapTile((row, col))

        # if the tile was a wall
        if 100 <= result <= 199:
            return True
        else:
            return False

    def CheckIfHitWall(self, possiblePlayerX_possiblePlayerY, row_col):
        (possiblePlayerX, possiblePlayerY) = possiblePlayerX_possiblePlayerY
        (row, col) = row_col
        numCollisions = 0

        # check each of the 9 surrounding tiles for a collision
        for iRow in range(row - 1, row + 2, 1):
            for iCol in range(col - 1, col + 2, 1):

                if (possiblePlayerX - (iCol * TILE_WIDTH) < TILE_WIDTH) and (
                        possiblePlayerX - (iCol * TILE_WIDTH) > -TILE_WIDTH) and (
                        possiblePlayerY - (iRow * TILE_HEIGHT) < TILE_HEIGHT) and (
                        possiblePlayerY - (iRow * TILE_HEIGHT) > -TILE_HEIGHT):

                    if self.IsWall((iRow, iCol)):
                        numCollisions += 1

        if numCollisions > 0:
            return True
        else:
            return False

    @staticmethod
    def CheckIfHit(playerX_playerY, x_y, cushion):
        (playerX, playerY) = playerX_playerY
        (x, y) = x_y
        if (playerX - x < cushion) and (playerX - x > -cushion) and (playerY - y < cushion) and (
                playerY - y > -cushion):
            return True
        else:
            return False

    @staticmethod
    def CheckIfHitSomething(playerX_playerY, row_col):
        (playerX, playerY) = playerX_playerY
        (row, col) = row_col
        for iRow in range(row - 1, row + 2, 1):
            for iCol in range(col - 1, col + 2, 1):

                if (playerX - (iCol * TILE_WIDTH) < TILE_WIDTH) and (
                        playerX - (iCol * TILE_WIDTH) > -TILE_WIDTH) and (
                        playerY - (iRow * TILE_HEIGHT) < TILE_HEIGHT) and (
                        playerY - (iRow * TILE_HEIGHT) > -TILE_HEIGHT):
                    # check the offending tile ID
                    result = thisLevel.GetMapTile((iRow, iCol))

                    if result == tileID['pellet']:
                        # got a pellet
                        thisLevel.SetMapTile((iRow, iCol), 0)
                        snd_pellet[player.pelletSndNum].play()
                        player.pelletSndNum = 1 - player.pelletSndNum

                        thisLevel.pellets -= 1

                        thisGame.AddToScore(10)

                        if thisLevel.pellets == 0:
                            # no more pellets left!
                            # WON THE LEVEL
                            thisGame.SetMode(6)


                    elif result == tileID['pellet-power']:
                        # got a power pellet
                        thisGame.SetMode(9)
                        thisLevel.SetMapTile((iRow, iCol), 0)
                        snd_powerpellet.play()

                        thisGame.AddToScore(100)
                        thisGame.ghostValue = 200

                        thisGame.ghostTimer = 360
                        for i in range(0, 4, 1):
                            if ghosts[i].state == 1:
                                ghosts[i].state = 2

                                """
                                # Must line up with grid before invoking a new path (for now)
                                ghosts[i].x = ghosts[i].nearestCol * TILE_HEIGHT
                                ghosts[i].y = ghosts[i].nearestRow * TILE_WIDTH								
                                
                                # give each ghost a path to a random spot (containing a pellet)
                                (randRow, randCol) = (0, 0)

                                while not self.GetMapTile((randRow, randCol)) == tileID[ 'pellet' ] or (randRow, randCol) == (0, 0):
                                    randRow = random.randint(1, self.lvlHeight - 2)
                                    randCol = random.randint(1, self.lvlWidth - 2)
                                ghosts[i].currentPath = path.FindPath( (ghosts[i].nearestRow, ghosts[i].nearestCol), (randRow, randCol) )
                                
                                ghosts[i].FollowNextPathWay()
                                """

                    elif result == tileID['door-h']:
                        # ran into a horizontal door
                        for i in range(0, thisLevel.lvlWidth, 1):
                            if not i == iCol:
                                if thisLevel.GetMapTile((iRow, i)) == tileID['door-h']:
                                    player.x = i * TILE_WIDTH

                                    if player.velX > 0:
                                        player.x += TILE_WIDTH
                                    else:
                                        player.x -= TILE_WIDTH

                    elif result == tileID['door-v']:
                        # ran into a vertical door
                        for i in range(0, thisLevel.lvlHeight, 1):
                            if not i == iRow:
                                if thisLevel.GetMapTile((i, iCol)) == tileID['door-v']:
                                    player.y = i * TILE_HEIGHT

                                    if player.velY > 0:
                                        player.y += TILE_HEIGHT
                                    else:
                                        player.y -= TILE_HEIGHT

                    elif result == tileID['heart']:
                        thisGame.SetMode(11)

    def GetGhostBoxPos(self):
        for row in range(0, self.lvlHeight, 1):
            for col in range(0, self.lvlWidth, 1):
                if self.GetMapTile((row, col)) == tileID['ghost-door']:
                    return row, col

        return False

    def GetPathwayPairPos(self):
        doorArray = []

        for row in range(0, self.lvlHeight, 1):
            for col in range(0, self.lvlWidth, 1):
                if self.GetMapTile((row, col)) == tileID['door-h']:
                    # found a horizontal door
                    doorArray.append((row, col))
                elif self.GetMapTile((row, col)) == tileID['door-v']:
                    # found a vertical door
                    doorArray.append((row, col))

        if len(doorArray) == 0:
            return False

        chosenDoor = random.randint(0, len(doorArray) - 1)

        if self.GetMapTile(doorArray[chosenDoor]) == tileID['door-h']:
            # horizontal door was chosen
            # look for the opposite one
            for i in range(0, thisLevel.lvlWidth, 1):
                if not i == doorArray[chosenDoor][1]:
                    if thisLevel.GetMapTile((doorArray[chosenDoor][0], i)) == tileID['door-h']:
                        return doorArray[chosenDoor], (doorArray[chosenDoor][0], i)
        else:
            # vertical door was chosen
            # look for the opposite one
            for i in range(0, thisLevel.lvlHeight, 1):
                if not i == doorArray[chosenDoor][0]:
                    if thisLevel.GetMapTile((i, doorArray[chosenDoor][1])) == tileID['door-v']:
                        return doorArray[chosenDoor], (i, doorArray[chosenDoor][1])

        return False

    def PrintMap(self):
        for row in range(0, self.lvlHeight, 1):
            outputLine = ""
            for col in range(0, self.lvlWidth, 1):
                outputLine += str(self.GetMapTile((row, col))) + ", "

    def DrawMap(self):
        global rect_list
        self.powerPelletBlinkTimer += 1
        if self.powerPelletBlinkTimer == 40:
            self.powerPelletBlinkTimer = 0

        for row in range(-1, thisGame.screenTileSize[0] + 1, 1):
            for col in range(-1, thisGame.screenTileSize[1] + 1, 1):

                # row containing tile that actually goes here
                actualRow = thisGame.screenNearestTilePos[0] + row
                actualCol = thisGame.screenNearestTilePos[1] + col

                useTile = self.GetMapTile((actualRow, actualCol))
                if useTile != 0 and useTile != tileID['door-h'] and useTile != tileID['door-v']:
                    # if this isn't a blank tile
                    if useTile == tileID['pellet-power']:
                        if self.powerPelletBlinkTimer < 20:
                            screen.blit(tileIDImage[useTile], (col * TILE_WIDTH - thisGame.screenPixelOffset[0],
                                                               row * TILE_HEIGHT - thisGame.screenPixelOffset[1]))

                    elif useTile == tileID['showlogo']:
                        screen.blit(thisGame.imLogo, (col * TILE_WIDTH - thisGame.screenPixelOffset[0],
                                                      row * TILE_HEIGHT - thisGame.screenPixelOffset[1]))

                    elif useTile == tileID['hiscores']:
                        screen.blit(thisGame.imHiscores, (col * TILE_WIDTH - thisGame.screenPixelOffset[0],
                                                          row * TILE_HEIGHT - thisGame.screenPixelOffset[1]))

                    else:
                        screen.blit(tileIDImage[useTile], (col * TILE_WIDTH - thisGame.screenPixelOffset[0],
                                                           row * TILE_HEIGHT - thisGame.screenPixelOffset[1]))

    def LoadLevel(self, levelNum):
        self.map = {}
        self.pellets = 0

        f = open(os.path.join(SCRIPT_PATH, "res", "levels", str(levelNum) + ".txt"), 'r')
        lineNum = -1
        rowNum = 0
        isReadingLevelData = False

        for line in f:

            lineNum += 1

            # print " ------- Level Line " + str(lineNum) + " -------- "
            while len(line) > 0 and (line[-1] == "\n" or line[-1] == "\r"): line = line[:-1]
            while len(line) > 0 and (line[0] == "\n" or line[0] == "\r"): line = line[1:]
            str_splitBySpace = line.split(' ')

            j = str_splitBySpace[0]

            if j == "'" or j == "":
                # comment / whitespace line
                # print " ignoring comment line.. "
                useLine = False
            elif j == "#":
                # special divider / attribute line
                useLine = False

                firstWord = str_splitBySpace[1]

                if firstWord == "lvlwidth":
                    self.lvlWidth = int(str_splitBySpace[2])

                elif firstWord == "lvlheight":
                    self.lvlHeight = int(str_splitBySpace[2])

                elif firstWord == "edgecolor":
                    # edge color keyword for backwards compatibility (single edge color) mazes
                    red = int(str_splitBySpace[2])
                    green = int(str_splitBySpace[3])
                    blue = int(str_splitBySpace[4])
                    self.edgeLightColor = (red, green, blue, 255)
                    self.edgeShadowColor = (red, green, blue, 255)

                elif firstWord == "edgelightcolor":
                    red = int(str_splitBySpace[2])
                    green = int(str_splitBySpace[3])
                    blue = int(str_splitBySpace[4])
                    self.edgeLightColor = (red, green, blue, 255)

                elif firstWord == "edgeshadowcolor":
                    red = int(str_splitBySpace[2])
                    green = int(str_splitBySpace[3])
                    blue = int(str_splitBySpace[4])
                    self.edgeShadowColor = (red, green, blue, 255)

                elif firstWord == "fillcolor":
                    red = int(str_splitBySpace[2])
                    green = int(str_splitBySpace[3])
                    blue = int(str_splitBySpace[4])
                    self.fillColor = (red, green, blue, 255)

                elif firstWord == "pelletcolor":
                    red = int(str_splitBySpace[2])
                    green = int(str_splitBySpace[3])
                    blue = int(str_splitBySpace[4])
                    self.pelletColor = (red, green, blue, 255)

                elif firstWord == "fruittype":
                    thisFruit.fruitType = int(str_splitBySpace[2])

                elif firstWord == "startleveldata":
                    isReadingLevelData = True
                    # print "Level data has begun"
                    rowNum = 0

                elif firstWord == "endleveldata":
                    isReadingLevelData = False
            # print "Level data has ended"

            else:
                useLine = True

            # this is a map data line
            if useLine:

                if isReadingLevelData:

                    # print str( len(str_splitBySpace) ) + " tiles in this column"

                    for k in range(0, self.lvlWidth, 1):
                        self.SetMapTile((rowNum, k), int(str_splitBySpace[k]))

                        thisID = int(str_splitBySpace[k])
                        if thisID == 4:
                            # starting position for pac-man

                            player.homeX = k * TILE_WIDTH
                            player.homeY = rowNum * TILE_HEIGHT
                            self.SetMapTile((rowNum, k), 0)

                        elif 10 <= thisID <= 13:
                            # one of the ghosts

                            ghosts[thisID - 10].homeX = k * TILE_WIDTH
                            ghosts[thisID - 10].homeY = rowNum * TILE_HEIGHT
                            self.SetMapTile((rowNum, k), 0)

                        elif thisID == 2:
                            # pellet

                            self.pellets += 1

                    rowNum += 1
        f.close()
        # reload all tiles and set appropriate colors
        GetCrossRef()

        # load map into the pathfinder object
        path.ResizeMap((self.lvlHeight, self.lvlWidth))

        for row in range(0, path.size[0], 1):
            for col in range(0, path.size[1], 1):
                if self.IsWall((row, col)):
                    path.SetType((row, col), 1)
                else:
                    path.SetType((row, col), 0)

        # do all the level-starting stuff
        self.Restart()

    def Restart(self):
        if thisGame.levelNum == 2:
            player.speed = 4

        for i in range(0, 4, 1):
            # move ghosts back to home
            ghosts[i].x = ghosts[i].homeX
            ghosts[i].y = ghosts[i].homeY
            ghosts[i].velX = 0
            ghosts[i].velY = 0
            ghosts[i].state = 1
            ghosts[i].speed = 2
            ghosts[i].Move()

            # give each ghost a path to a random spot (containing a pellet)
            (randRow, randCol) = (0, 0)

            while not self.GetMapTile((randRow, randCol)) == tileID['pellet'] or (randRow, randCol) == (0, 0):
                randRow = random.randint(1, self.lvlHeight - 2)
                randCol = random.randint(1, self.lvlWidth - 2)

            # print "Ghost " + str(i) + " headed towards " + str((randRow, randCol))
            ghosts[i].currentPath = path.FindPath((ghosts[i].nearestRow, ghosts[i].nearestCol), (randRow, randCol))
            ghosts[i].FollowNextPathWay()

        thisFruit.active = False

        thisGame.fruitTimer = 0

        player.x = player.homeX
        player.y = player.homeY
        player.velX = 0
        player.velY = 0

        player.anim_pacmanCurrent = player.anim_pacmanS
        player.animFrame = 3


def CheckIfCloseButton(events):
    for event in events:
        if event.type == QUIT:
            sys.exit(0)


def CheckInputs():
    if thisGame.mode == 1 or thisGame.mode == 8 or thisGame.mode == 9:
        if pygame.key.get_pressed()[pygame.K_RIGHT] or (js is not None and js.get_axis(JS_XAXIS) > 0.5):
            if not (player.velX == player.speed and player.velY == 0) and not thisLevel.CheckIfHitWall(
                    (player.x + player.speed, player.y), (player.nearestRow, player.nearestCol)):
                player.velX = player.speed
                player.velY = 0

        elif pygame.key.get_pressed()[pygame.K_LEFT] or (js is not None and js.get_axis(JS_XAXIS) < -0.5):
            if not (player.velX == -player.speed and player.velY == 0) and not thisLevel.CheckIfHitWall(
                    (player.x - player.speed, player.y), (player.nearestRow, player.nearestCol)):
                player.velX = -player.speed
                player.velY = 0

        elif pygame.key.get_pressed()[pygame.K_DOWN] or (js is not None and js.get_axis(JS_YAXIS) > 0.5):
            if not (player.velX == 0 and player.velY == player.speed) and not thisLevel.CheckIfHitWall(
                    (player.x, player.y + player.speed), (player.nearestRow, player.nearestCol)):
                player.velX = 0
                player.velY = player.speed

        elif pygame.key.get_pressed()[pygame.K_UP] or (js is not None and js.get_axis(JS_YAXIS) < -0.5):
            if not (player.velX == 0 and player.velY == -player.speed) and not thisLevel.CheckIfHitWall(
                    (player.x, player.y - player.speed), (player.nearestRow, player.nearestCol)):
                player.velX = 0
                player.velY = -player.speed

    if pygame.key.get_pressed()[pygame.K_ESCAPE] or (js is not None and js.get_button(7)) :
        sys.exit(0)

    elif thisGame.mode == 3:
        if pygame.key.get_pressed()[pygame.K_RETURN] or (js is not None and js.get_button(JS_STARTBUTTON)):
            thisGame.StartNewGame()


#      _____________________________________________
# ___/  function: Get ID-Tilename Cross References  \______________________________________

def GetCrossRef():
    f = open(os.path.join(SCRIPT_PATH, "res", "crossref.txt"), 'r')

    lineNum = 0

    for i in f.readlines():
        # print " ========= Line " + str(lineNum) + " ============ "
        while len(i) > 0 and (i[-1] == '\n' or i[-1] == '\r'): i = i[:-1]
        while len(i) > 0 and (i[0] == '\n' or i[0] == '\r'): i = i[1:]
        str_splitBySpace = i.split(' ')

        j = str_splitBySpace[0]

        if j == "'" or j == "" or j == "#":
            # comment / whitespace line
            # print " ignoring comment line.. "
            useLine = False
        else:
            # print str(wordNum) + ". " + j
            useLine = True

        if useLine:
            tileIDName[int(str_splitBySpace[0])] = str_splitBySpace[1]
            tileID[str_splitBySpace[1]] = int(str_splitBySpace[0])

            thisID = int(str_splitBySpace[0])
            if not thisID in NO_GIF_TILES:
                tileIDImage[thisID] = get_image_surface(
                    os.path.join(SCRIPT_PATH, "res", "tiles", str_splitBySpace[1] + ".gif"))
            else:
                tileIDImage[thisID] = pygame.Surface((TILE_WIDTH, TILE_HEIGHT))

            # change colors in tileIDImage to match maze colors
            for y in range(0, TILE_WIDTH, 1):
                for x in range(0, TILE_HEIGHT, 1):

                    if tileIDImage[thisID].get_at((x, y)) == IMG_EDGE_LIGHT_COLOR:
                        # wall edge
                        tileIDImage[thisID].set_at((x, y), thisLevel.edgeLightColor)

                    elif tileIDImage[thisID].get_at((x, y)) == IMG_FILL_COLOR:
                        # wall fill
                        tileIDImage[thisID].set_at((x, y), thisLevel.fillColor)

                    elif tileIDImage[thisID].get_at((x, y)) == IMG_EDGE_SHADOW_COLOR:
                        # pellet color
                        tileIDImage[thisID].set_at((x, y), thisLevel.edgeShadowColor)

                    elif tileIDImage[thisID].get_at((x, y)) == IMG_PELLET_COLOR:
                        # pellet color
                        tileIDImage[thisID].set_at((x, y), thisLevel.pelletColor)

        lineNum += 1
    f.close()


#      __________________
# ___/  main code block  \_____________________________________________________

# create the pacman
player = pacman()

# create a path_finder object
path = path_finder()

# create ghost objects
ghosts = {}
for i in range(0, 6, 1):
    # remember, ghost[4] is the blue, vulnerable ghost
    ghosts[i] = ghost(i)

# create piece of fruit
thisFruit = fruit()

tileIDName = {}  # gives tile name (when the ID# is known)
tileID = {}  # gives tile ID (when the name is known)
tileIDImage = {}  # gives tile image (when the ID# is known)
oldEdgeLightColor = None
oldEdgeShadowColor = None
oldFillColor = None

# create game and level objects and load first level
thisGame = game()
thisLevel = level()
thisLevel.LoadLevel(thisGame.GetLevelNum())

#window = pygame.display.set_mode(thisGame.screenSize, pygame.FULLSCREEN)
window = pygame.display.set_mode(thisGame.screenSize)

# initialise the joystick
if pygame.joystick.get_count() > 0:
    if JS_DEVNUM < pygame.joystick.get_count():
        js = pygame.joystick.Joystick(JS_DEVNUM)
    else:
        js = pygame.joystick.Joystick(0)
    js.init()
else:
    js = None

# game "mode" variable
# 0 = ready to level start
# 1 = normal
# 2 = hit ghost
# 3 = game over
# 4 = wait to start
# 5 = wait after eating ghost
# 6 = wait after finishing level
# 7 = flashing maze after finishing level
# 8 = extra pacman, small ghost mode
# 9 = changed ghost to glasses
# 10 = blank screen before changing levels

while True:
    CheckIfCloseButton(pygame.event.get())
    if thisGame.mode == 0:
        # ready to level start
        thisGame.modeTimer += 1

        if thisGame.modeTimer == 150:
            thisGame.SetMode(1)

    if thisGame.mode == 1:
        # normal gameplay mode
        CheckInputs()
        thisGame.modeTimer += 1

        player.Move()
        for i in range(0, 4, 1):
            ghosts[i].Move()
        thisFruit.Move()

    elif thisGame.mode == 2:
        # waiting after getting hit by a ghost
        thisGame.modeTimer += 1

        if thisGame.modeTimer == 60:
            thisLevel.Restart()

            thisGame.lives -= 1
            if thisGame.lives == -1:
                thisGame.updatehiscores(thisGame.score)
                thisGame.SetMode(3)
                thisGame.drawmidgamehiscores()
            else:
                thisGame.SetMode(4)

    elif thisGame.mode == 3:
        # game over
        CheckInputs()

    elif thisGame.mode == 4:
        # waiting to start
        thisGame.modeTimer += 1

        if thisGame.modeTimer == 60:
            thisGame.SetMode(1)
            player.velX = player.speed

    elif thisGame.mode == 5:
        # brief pause after munching a vulnerable ghost
        thisGame.modeTimer += 1

        if thisGame.modeTimer == 20:
            thisGame.SetMode(8)

    elif thisGame.mode == 6:
        # pause after eating all the pellets
        thisGame.modeTimer += 1

        if thisGame.modeTimer == 40:
            thisGame.SetMode(7)
            oldEdgeLightColor = thisLevel.edgeLightColor
            oldEdgeShadowColor = thisLevel.edgeShadowColor
            oldFillColor = thisLevel.fillColor

    elif thisGame.mode == 7:
        # flashing maze after finishing level
        thisGame.modeTimer += 1

        whiteSet = [10, 30, 50, 70]
        normalSet = [20, 40, 60, 80]

        if not whiteSet.count(thisGame.modeTimer) == 0:
            # member of white set
            thisLevel.edgeLightColor = (255, 255, 254, 255)
            thisLevel.edgeShadowColor = (255, 255, 254, 255)
            thisLevel.fillColor = (0, 0, 0, 255)
            GetCrossRef()
        elif not normalSet.count(thisGame.modeTimer) == 0:
            # member of normal set
            thisLevel.edgeLightColor = oldEdgeLightColor
            thisLevel.edgeShadowColor = oldEdgeShadowColor
            thisLevel.fillColor = oldFillColor
            GetCrossRef()
        elif thisGame.modeTimer == 100:
            thisGame.SetMode(10)

    elif thisGame.mode == 8:
        CheckInputs()
        ghostState = 1
        thisGame.modeTimer += 1

        player.Move()

        for i in range(0, 4, 1):
            ghosts[i].Move()

        for i in range(0, 4, 1):
            if ghosts[i].state == 3:
                ghostState = 3
                break
            elif ghosts[i].state == 2:
                ghostState = 2

        if thisLevel.pellets == 0:
            # WON THE LEVEL
            thisGame.SetMode(6)
        elif ghostState == 1:
            thisGame.SetMode(1)
        elif ghostState == 2:
            thisGame.SetMode(9)

        thisFruit.Move()

    elif thisGame.mode == 9:
        CheckInputs()
        thisGame.modeTimer += 1

        player.Move()
        for i in range(0, 4, 1):
            ghosts[i].Move()
        thisFruit.Move()

    elif thisGame.mode == 10:
        # blank screen before changing levels
        thisGame.modeTimer += 1
        if thisGame.modeTimer == 10:
            thisGame.SetNextLevel()

    elif thisGame.mode == 11:
        # flashing maze after finishing level
        thisGame.modeTimer += 1

        whiteSet = [10, 30, 50, 70]
        normalSet = [20, 40, 60, 80]

        if not whiteSet.count(thisGame.modeTimer) == 0:
            # member of white set
            thisLevel.edgeLightColor = (255, 255, 254, 255)
            thisLevel.edgeShadowColor = (255, 255, 254, 255)
            thisLevel.fillColor = (0, 0, 0, 255)
            GetCrossRef()
        elif not normalSet.count(thisGame.modeTimer) == 0:
            # member of normal set
            thisLevel.edgeLightColor = oldEdgeLightColor
            thisLevel.edgeShadowColor = oldEdgeShadowColor
            thisLevel.fillColor = oldFillColor
            GetCrossRef()
        elif thisGame.modeTimer == 100:
            thisGame.modeTimer = 1

    thisGame.SmartMoveScreen()

    screen.blit(img_Background, (0, 0))

    if not thisGame.mode == 10:
        thisLevel.DrawMap()

        if thisGame.fruitScoreTimer > 0:
            if thisGame.modeTimer % 2 == 0:
                thisGame.DrawNumber(2500, (
                    thisFruit.x - thisGame.screenPixelPos[0] - 16, thisFruit.y - thisGame.screenPixelPos[1] + 4))

        for i in range(0, 4, 1):
            ghosts[i].Draw()
        thisFruit.Draw()
        player.Draw()

        if thisGame.mode == 3:
            screen.blit(thisGame.imHiscores, (HS_XOFFSET, HS_YOFFSET))

    if thisGame.mode == 5:
        thisGame.DrawNumber(thisGame.ghostValue / 2,
                            (player.x - thisGame.screenPixelPos[0] - 4, player.y - thisGame.screenPixelPos[1] + 6))

    thisGame.DrawScore()

    pygame.display.update()
    del rect_list[:]

    clock.tick(40)
