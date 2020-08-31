import sqlite3
import os
import sys
import psutil
import platform
import wmi
import time
from random import randint
from uuid import uuid4
import warnings
defaultDir = os.path.join(os.getenv('UserProfile'), 'AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\\LocalState\\')

# A new note could be created without a reference by inserting a note, getting position

class StickyNotes:
  def __init__(self, dir=defaultDir):
    if os.name != 'nt':
      raise OSError("Operating system is not supported")

    os.system('explorer.exe shell:appsFolder\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')

    if not os.path.isdir(defaultDir):
      raise IOError("Microsoft Sticky Notes not found")

    self.directory = dir
    self.connectToDB()

    self.__managedPosition = self.getManagedPosition()
    self.theme = self.__Theme()

  # This is not how I wanted to do this but due to a recent update it is required
  def getManagedPosition(self):
    self.__cursor.execute('SELECT WindowPosition FROM Note')
    for windowPosition in self.__cursor.fetchall():
      if not windowPosition[0] is None:
        notePosition = windowPosition[0].split(";")
        if notePosition[0][0:25] == "ManagedPosition=DeviceId:":
          return notePosition[0]
        
  def getWindowPositionString(self, note):
    return ";".join([self.__managedPosition, note.getPositionString(), note.getSizeString()])

  def getDir(self):
    return self.directory

  def connectToDB(self):
    if os.path.isfile(os.path.join(self.directory, 'plum.sqlite')):
      self.__db = sqlite3.connect(os.path.join(self.directory, 'plum.sqlite'))
      self.__cursor = self.__db.cursor()
    else:
      raise FileNotFoundError('Database file not found')

  def closeDB(self):
    self.__db.close()

  def getNotes(self, id):
    self.__cursor.execute('SELECT * FROM Note WHERE Id = ?', [id])
    return list(map(lambda x: Note(x[0], x[5], x[2]), self.__cursor.fetchall()))
   
  def deleteNote(self, id):
    self.__cursor.execute('DELETE FROM Note WHERE Id = ?', [id])

  def writeNote(self, note):
    if type(note) == Note:
      if (note.theme in self.theme.themes and
        type(note.text) == str):
        
        if self.__managedPosition is None and note.getIsOpen() == True:
          self.__cursor.execute(
            'INSERT INTO Note(Text, Theme, IsOpen, Id) values (?, ?, ?, ?)',
            ["tmp", note.theme, 1, "tmp"])
          self.commit()
          self.reloadNotes()

          # Welp this is ghetto but it works
          while not self.__managedPosition:
            time.sleep(.1)
            self.__managedPosition = self.getManagedPosition()
            pass

        self.deleteNote("tmp")

        self.__cursor.execute(
          'INSERT INTO Note(Text , WindowPosition,Theme, IsOpen, Id) values (?, ?, ?, ?, ?)',
          [note.text, self.getWindowPositionString(note), note.theme, note.getIsOpen(), note.id])
      else:
        raise TypeError('Incorrect type within Note')
    else:
      raise TypeError('Note expected')

  def writeNotes(self, *notes):
    for note in notes:
      self.writeNote(note)

  def reloadNotes(self):
    for p in psutil.process_iter():
      if "Microsoft.Notes.exe" in p.name():
        p.kill()
        # ??? interesting command, took me a while to find
        os.system('explorer.exe shell:appsFolder\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')

  def commit(self): self.__db.commit()

  class __Theme:
    def __init__(self):
      self.themes = [
        "Yellow",
        "White",
        "Green",
        "Pink",
        "Purple",
        "Blue",
        "Gray",
        "Charcoal",
      ] 
      self.yellow = 'Yellow'
      self.white = 'White'
      self.green = 'Green'
      self.pink = 'Pink'
      self.purple = 'Purple'
      self.blue = 'Blue'
      self.gray = 'Gray'
      self.charcoal = 'Charcoal'


class Note():
  def __init__(self, text, theme=None, isOpen=False):
    self.text = text
    self.setIsOpen(isOpen)
    self.setTheme(theme)
    self.id = str(uuid4())

    self.size = self.calculateSize()
    
    self.position = {
      "x": randint(0, 500),
      "y": randint(0, 500),
    }

  def calculateSize(self):
    textArray = self.text.replace("\r", "").split('\n')
    # Fixed window height + line height, is rough but works in most scenarios
    y = 90 + len(textArray) * 18.65
    if y < 320: y = 320
    return {
      "width": 320,
      "height": y
    }

  def getPositionString(self):
    return "Position={},{}".format(self.position["x"], self.position["y"])
  
  def getSizeString(self):
    return "Size={},{}".format(self.size["width"], self.size["height"])

  def setTheme(self, theme):
    if theme == None:
      self.theme = 'White'
    else:
      self.theme = theme

  # Probably a better way of doing this
  def setIsOpen(self, isOpen):
    isOpenBool = bool(isOpen)
    if (isOpenBool == False):
      self.__isOpen = 0
    if (isOpenBool == True):
      self.__isOpen = 1

  def getIsOpen(self): return self.__isOpen

  # def findMonitors(self):
  #   self.__monitors = []
  #   obj = wmi.WMI().Win32_PnPEntity(ConfigManagerErrorCode=0)
  #   displays = [x for x in obj if 'DISPLAY' in str(x)]
  #   for monitor in displays:
  #     if 'UID' in monitor.DeviceID:
  #       print(monitor)
  #       self.__monitors.append(monitor.DeviceID)

  # def writeNote(self): 
    # self.__cursor.execute('''INSERT INTO Note())

  # def setPos(self):

  
# db = sqlite3.connect(stickyNoteDBPath)
# cursor = db.cursor()
# cursor.execute('SELECT * FROM Note')
# print(cursor.fetchone())
