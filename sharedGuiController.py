import sharedViewHelpers
import os
from sharedDefs import *
import tkinter as tk

class GUIController:
    def __init__(self, window):
        self.onDarkMode = False
        self.window = window
        self.logoImg = None

    def standardInitialize(self):
        # default color mode == dark
        self.changeColorMode()

        # bindings
        self.window.bind('<Control-Key-w>', lambda _: self.window.destroy())
        self.window.bind('<Control-Key-W>', lambda _: self.window.destroy())
        self.window.bind("<Button-2>", lambda _: self.changeColorMode()) # middle click
        
        #
        self.setLogoIcon()

    ###
    def changeColorMode(self):
        if self.onDarkMode:
            self.onDarkMode = False
            sharedViewHelpers.changeToLightMode(self.window)
        else:
            self.onDarkMode = True
            sharedViewHelpers.changeToDarkMode(self.window)

    def isOnDarkMode(self):
        return self.onDarkMode

    ###
    def setLogoIcon(self):
        # set icon image (if available)
        self.logoImg = None
        if os.path.exists(LOGO_PATH):
            self.logoImg = tk.PhotoImage(file=LOGO_PATH)
            self.window.iconphoto(False, self.logoImg)

    def getLogoIcon(self):
        return self.logoImg
