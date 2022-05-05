import sqlite3
import os
import psutil
import time
import warnings
from datetime import datetime, timedelta
from random import randint
from uuid import uuid4
from enum import Enum

DEFAULT_DIR = os.path.join(os.getenv(
    'UserProfile'), 'AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\\LocalState\\')


class StickyNotes:
    """class representing the Microsoft Sticky Notes application
    """

    def __init__(self, dir=DEFAULT_DIR):
        """Initilizes interaction with the Microsoft Sticky Notes application

        :param dir: Sticky Notes Database directory, defaults to DEFAULT_DIR
        :type dir: string, optional
        :raises OSError: incompatible operating system
        :raises IOError: sticky Notes database not found
        """
        if os.name != 'nt':
            raise OSError("Operating system is not supported")

        os.system(
            'explorer.exe shell:appsFolder\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')

        if not os.path.isdir(DEFAULT_DIR):
            raise IOError("Microsoft Sticky Notes not found")

        self.directory = dir
        self.connect_to_db()

        self._managed_position = self.get_managed_position()
        self.theme = self.__Theme()

    # This is not how I wanted to do this but due to a recent update it is required
    def get_managed_position(self):
        """attempts to find a managed position from a sticky note

        :return: managed position string
        :rtype: str
        """
        self._cursor.execute('SELECT WindowPosition FROM Note')
        for window_position in self._cursor.fetchall():
            if not window_position[0] is None:
                note_position = window_position[0].split(";")
                if note_position[0][0:25] == "ManagedPosition=DeviceId:":
                    return note_position[0]

    def get_window_position_string(self, note):
        """generates the window position string

        :param note: a note to generate the position string for
        :type note: Note
        :return: the position string used by Microsoft Sticky Notes
        :rtype: str
        """
        return ";".join([self._managed_position, note.get_position_string(), note.get_size_string()])

    def connect_to_db(self):
        """connects to the Microsoft Sticky Notes databases

        :raises FileNotFoundError: database file not found
        """
        if os.path.isfile(os.path.join(self.directory, 'plum.sqlite')):
            self.__db = sqlite3.connect(
                os.path.join(self.directory, 'plum.sqlite'))
            self._cursor = self.__db.cursor()
        else:
            raise FileNotFoundError('Database file not found')

    def close_db(self):
        """closes the database
        """
        self.__db.close()

    def get_notes(self, id):
        """grabs a note from the id

        :param id: the unique identifier within the database to search for
        :type id: str
        :return: a list of notes found
        :rtype: list of Note
        """
        self._cursor.execute('SELECT * FROM Note WHERE Id = ?', [id])
        return list(map(lambda x: Note(x[0], x[5], x[2]), self._cursor.fetchall()))

    def delete_note(self, id):
        """deletes a note using the id

        :param id: the id of the note to delete
        :type id: str
        """
        self._cursor.execute('DELETE FROM Note WHERE Id = ?', [id])

    def generate_managed_position_from_temp(self):
        temp_uuid = str(uuid4())
        self._cursor.execute(
            'INSERT INTO Note(Text, Theme, IsOpen, Id) values (?, ?, ?, ?)',
            ["", self.theme.charcoal, 1, temp_uuid])
        self.commit()
        self.reload_notes()

        # Welp this is ghetto but it works
        while not self._managed_position:
            time.sleep(.1)
            self._managed_position = self.get_managed_position()
            pass

        self.delete_note(temp_uuid)

    def write_notes(self, *notes):
        if len(notes) >= 4:
            warnings.warn(
                "For an unknown reason adding more than 3 notes at a single time will cause the notes list to stop working")

        has_open_without_reference = False
        for note in notes:
            if not type(note) == Note:
                raise TypeError('Note expected')

            if not note.theme in self.theme.themes or type(note.text) != str:
                raise TypeError('Incorrect type within Note')

            if self._managed_position is None and note.get_is_open():
                has_open_without_reference = True

        if has_open_without_reference:
            self.generate_managed_position_from_temp()

        mapped_notes = list(map(lambda note: (note.text, self.get_window_position_string(
            note), note.theme, note.get_is_open(), note.id, note.get_ticks(), note.get_ticks()), notes))

        self._cursor.executemany(
            'INSERT INTO Note(Text, WindowPosition, Theme, IsOpen, Id, CreatedAt, UpdatedAt) values (?, ?, ?, ?, ?, ?, ?)', mapped_notes)

    @staticmethod
    def reload_notes():
        """reloads the Microsoft Stick Notes application
        """
        for p in psutil.process_iter():
            if "Microsoft.Notes.exe" in p.name():
                p.kill()
                # ??? interesting command, took me a while to find
                os.system(
                    'explorer.exe shell:appsFolder\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe!App')

    def commit(self):
        """commits all changes to the database
        """
        self.__db.commit()


class Note():
    """class representing a sticky note
    """

    def __init__(self, text, theme=None, is_open=False):
        """Makes the basic sticky note

        :param text: text within the sticky note
        :type text: str
        :param theme: colour of the sticky note, defaults to None
        :type theme: str, optional
        :param is_open: determines if the sticky note should be open or not, defaults to False
        :type is_open: bool, optional
        """
        self.text = text
        self.set_is_open(is_open)
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
        """calculates the .Net ticks for the creation date

        :param dt: datetime for when the creation date should be
        :type dt: datetime
        :return: value for the creation date
        :rtype: float
        """
        return (dt - datetime(1, 1, 1)).total_seconds() * 10000000

    def get_creation_date(self):
        """converts .Net ticks to a human readable time

        :return: [description]
        :rtype: [type]
        """
        return datetime(1, 1, 1) + timedelta(microseconds=self._creation_ticks/10)

    def set_creation_date(self, dt):
        """sets the creation date of the note

        :param dt: the datetime to set the creation date to
        :type dt: datetime
        """
        self._creation_ticks = self._calc_creation_ticks(dt)

    def get_ticks(self):
        """gets the .Net ticks for the creation date

        :return: the ticks
        :rtype: int
        """
        return self._creation_ticks

    def calculate_size(self):
        """calculates the size of the note based off the number of lines

        :return: the width and height of the sticky note
        :rtype: dict of (int, int)
        """
        text_array = self.text.replace("\r", "").split('\n')
        # Fixed window height + line height, is rough but works in most scenarios
        y = int(90 + len(text_array) * 18.65)
        if y < 320:
            y = 320
        return {
            "width": 320,
            "height": y
        }

    def get_position_string(self):
        """generates the position string for WindowPosition

        :return: position string
        :rtype: str
        """
        return "Position={},{}".format(self.position["x"], self.position["y"])

    def get_size_string(self):
        """generates the size string for WindowPosition

        :return: size string
        :rtype: str
        """
        if self.auto_height:
            self.size = self.calculate_size()
        return "Size={},{}".format(self.size["width"], self.size["height"])

    def set_theme(self, theme):
        """sets the theme of the sticky note

        :param theme: the theme to use
        :type theme: str
        """
        if theme == None:
            self.theme = 'White'
        else:
            self.theme = theme

    # Probably a better way of doing this
    def set_is_open(self, is_open):
        """sets the open status of the sticky note

        :param is_open: whether the note is open or not
        :type is_open: bool
        """
        is_open_bool = bool(is_open)
        if (is_open_bool == False):
            self._is_open = 0
        if (is_open_bool == True):
            self._is_open = 1

    def get_is_open(self):
        """get if the note is open or not

        :return: the open variable of the note
        :rtype: bool
        """
        return self._is_open

    class Theme(Enum):
        yellow = 'Yellow'
        white = 'White'
        green = 'Green'
        pink = 'Pink'
        purple = 'Purple'
        blue = 'Blue'
        gray = 'Gray'
        charcoal = 'Charcoal'
