import sqlite3
import os
import sys
import psutil
import platform
import wmi
import time
import warnings
from datetime import datetime, timedelta
from random import randint
from uuid import uuid4

DEFAULT_DIR = os.path.join(os.getenv('UserProfile'), 'AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\\LocalState\\')

class StickyNotes:
  def __init__(self, dir=DEFAULT_DIR):
    if os.name != 'nt':
      raise OSError("Operating system is not supported")

    os.system('explorer.exe shell:appsFolder\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')

    if not os.path.isdir(DEFAULT_DIR):
      raise IOError("Microsoft Sticky Notes not found")

    self.directory = dir
    self.connect_to_db()

    self._managed_position = self.get_managed_position()
    self.theme = self.__Theme()

  # This is not how I wanted to do this but due to a recent update it is required
  def get_managed_position(self):
    self._cursor.execute('SELECT WindowPosition FROM Note')
    for window_position in self._cursor.fetchall():
      if not window_position[0] is None:
        note_position = window_position[0].split(";")
        if note_position[0][0:25] == "ManagedPosition=DeviceId:":
          return note_position[0]

  def get_window_position_string(self, note):
    return ";".join([self._managed_position, note.get_position_string(), note.get_size_string()])

  def get_dir(self):
    return self.directory

  def connect_to_db(self):
    if os.path.isfile(os.path.join(self.directory, 'plum.sqlite')):
      self.__db = sqlite3.connect(os.path.join(self.directory, 'plum.sqlite'))
      self._cursor = self.__db.cursor()
    else:
      raise FileNotFoundError('Database file not found')

  def close_db(self):
    self.__db.close()

  def get_notes(self, id):
    self._cursor.execute('SELECT * FROM Note WHERE Id = ?', [id])
    return list(map(lambda x: Note(x[0], x[5], x[2]), self._cursor.fetchall()))
   
  def delete_note(self, id):
    self._cursor.execute('DELETE FROM Note WHERE Id = ?', [id])

  def write_note(self, note):
    if type(note) == Note:
      temp_uuid = str(uuid4())
      if (note.theme in self.theme.themes and type(note.text) == str):
        if self._managed_position is None and note.get_is_open() == True:
          self._cursor.execute(
            'INSERT INTO Note(Text, Theme, IsOpen, Id) values (?, ?, ?, ?)',
            ["", note.theme, 1, temp_uuid])
          self.commit()
          self.reload_notes()

          # Welp this is ghetto but it works
          while not self._managed_position:
            time.sleep(.1)
            self._managed_position = self.get_managed_position()
            pass

        self.delete_note(temp_uuid)
        self._cursor.execute(
          'INSERT INTO Note(Text , WindowPosition,Theme, IsOpen, Id, CreatedAt, UpdatedAt) values (?, ?, ?, ?, ?, ?, ?)',
          [note.text, self.get_window_position_string(note), note.theme, note.get_is_open(), note.id, note.get_ticks(), note.get_ticks()])
      else:
        raise TypeError('Incorrect type within Note')
    else:
      raise TypeError('Note expected')

  def write_notes(self, *notes):
    for note in notes:
      self.write_note(note)

  @staticmethod
  def reload_notes():
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
    self.set_is_open(isOpen)
    self.set_theme(theme)
    self.id = str(uuid4())
    self.size = {
      "width": 320,
      "height": 320
    }
    self._creation_ticks = self._calc_creation_ticks(datetime.utcnow())
    self.position = {
      "x": randint(0, 500),
      "y": randint(0, 500),
    }
    
    self.auto_height = True

  @staticmethod
  def _calc_creation_ticks(dt):
    return (dt - datetime(1, 1, 1)).total_seconds() * 10000000

  def get_creation_date(self):
    return datetime.datetime(1, 1, 1) + datetime.timedelta(microseconds = self._creation_ticks/10)

  def set_creation_date(self, dt):
    self._creation_ticks = self._calc_creation_ticks(dt)

  def get_ticks(self):
    return self._creation_ticks

  def calculate_size(self):
    text_array = self.text.replace("\r", "").split('\n')
    # Fixed window height + line height, is rough but works in most scenarios
    y = 90 + len(text_array) * 18.65
    if y < 320: y = 320
    return {
      "width": 320,
      "height": y
    }

  def get_position_string(self):
    return "Position={},{}".format(self.position["x"], self.position["y"])
  
  def get_size_string(self):
    if self.auto_height: self.size = self.calculate_size()
    return "Size={},{}".format(self.size["width"], self.size["height"])

  def set_theme(self, theme):
    if theme == None:
      self.theme = 'White'
    else:
      self.theme = theme

  # Probably a better way of doing this
  def set_is_open(self, isOpen):
    is_open_bool = bool(isOpen)
    if (is_open_bool == False):
      self._is_open = 0
    if (is_open_bool == True):
      self._is_open = 1

  def get_is_open(self): return self._is_open