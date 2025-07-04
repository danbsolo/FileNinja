from FileNinjaSuite.Shared import sharedViewHelpers
from FileNinjaSuite.Shared.sharedDefs import *
from FileNinjaSuite.Shared import sharedCommon
import os
import tkinter as tk
from idlelib.tooltip import Hovertip
import sys
import threading


class GUIController:
    def __init__(self, window, title):
        self.window = window
        self.title = title

        self.onDarkMode = False
        self.logoImg = None
        
        self.currentStatusPair = None
        self.setStatusIdle()

        self.executeButton = None
        self.executionThread = None

    def standardInitialize(self):
        # default color mode == dark
        self.changeColorMode()

        # bindings
        self.window.bind('<Control-Key-w>', lambda _: self.window.destroy())
        self.window.bind('<Control-Key-W>', lambda _: self.window.destroy())
        self.window.bind("<Button-2>", lambda _: self.changeColorMode()) # middle click
        self.window.protocol("WM_DELETE_WINDOW", self.closeWindow)
        #
        self.setLogoIcon()

    def changeColorMode(self):
        if self.onDarkMode:
            self.onDarkMode = False
            sharedViewHelpers.changeToLightMode(self.window)
        else:
            self.onDarkMode = True
            sharedViewHelpers.changeToDarkMode(self.window)

    def isOnDarkMode(self):
        return self.onDarkMode

    def setLogoIcon(self):
        # set icon image (if available)
        self.logoImg = None
        if os.path.exists(LOGO_PATH):
            self.logoImg = tk.PhotoImage(file=LOGO_PATH)
            self.window.iconphoto(False, self.logoImg)

    def getLogoIcon(self):
        return self.logoImg
    
    def closeWindow(self):
        if self.currentStatusPair[0] == STATUS_RUNNING:
            sys.exit()  # Force close
        else:
            self.window.destroy()  # Graceful close

    def setCurrentStatus(self, exitCode, description=None):
        if isinstance(exitCode, int):
            self.currentStatusPair = (exitCode, description)
        else:
            self.currentStatusPair = exitCode

    def setStatusIdle(self):
        self.setCurrentStatus(STATUS_IDLE, None)
    
    def setStatusRunning(self):
        self.setCurrentStatus(STATUS_RUNNING, None)
    
    def setStatusSuccessful(self):
        self.setCurrentStatus(STATUS_SUCCESSFUL, None)
    
    def getCurrentStatus(self):
        return self.currentStatusPair


    def createHoverTips(self, hoverTipDictionary):
        for widget in hoverTipDictionary.keys():
            self.createHoverTip(widget, hoverTipDictionary[widget])

    def createHoverTip(self, widget, description):
        Hovertip(widget, description, hover_delay=0)


    def setThreadVars(self, executeButton, threadAliveFunction=lambda: None, threadDeadFunction=lambda: None):
        self.executeButton = executeButton
        self.threadAliveFunction = threadAliveFunction
        self.threadDeadFunction = threadDeadFunction

    def checkIfDone(self):
        # If the thread has finished
        if not self.executionThread.is_alive():
            self.window.title(self.title)
            self.executeButton.config(text="Execute", state="normal")
            
            self.threadDeadFunction()

            exitCode, exitMessage = self.currentStatusPair
            if exitCode == STATUS_SUCCESSFUL:
                return
            tk.messagebox.showerror(f"Error: {exitCode}", sharedCommon.interpretError(exitCode, exitMessage))
            self.setStatusIdle()
        else:
            # Otherwise check again after the specified number of milliseconds.
            self.threadAliveFunction()
            self.scheduleCheckIfDone()

    def scheduleCheckIfDone(self):
        self.window.after(500, self.checkIfDone)

    def initiateControllerThread(self, launchControllerThreadFunction):
        self.setStatusRunning()
        
        self.window.title(f"{self.title}: RUNNING...")
        self.executeButton.config(text="RUNNING...", state="disabled")

        self.executionThread = threading.Thread(target=lambda: self.setCurrentStatus(launchControllerThreadFunction()))
        self.executionThread.daemon = True  # When the main thread closes, this daemon thread will also close alongside it
        self.executionThread.start()

        self.scheduleCheckIfDone()
