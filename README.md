# microsoft-sticky-notes-py
## A module that allows for interaction with the Microsoft Sticky Notes application
![Example Sticky Notes](https://i.imgur.com/I9YW2sf.png)


## TODO:
- Implement searching for sticky notes other than exact text
- Finish implementing match exact text
- Add editing of sticky notes
- Add usage of markdown within the sticky notes
- Add custom creation date

## A simple example to create new sticky notes
```py
from StickyNotes import StickyNotes, Note

sticky_notes = StickyNotes()
newNote = Note('Hello', Note.Theme.yellow, True)
pinkNote = Note('World', Note.Theme.pink, True)
sticky_notes.write_notes(newNote, pinkNote)
sticky_notes.commit()
sticky_notes.reload_notes()
sticky_notes.close_db()
```

## The same example using a custom db directory
```py
import os
import sys
from StickyNotes import StickyNotes, Note

stickyNoteDBPath = '/custompath/'

if os.path.isdir(stickyNoteDBPath):
  stickyNotes = StickyNotes(stickyNoteDBPath)
  newNote = Note('Hello', Note.Theme.yellow, True)
  pinkNote = Note('World', Note.Theme.pink, True)
  stickyNotes.write_notes(newNote, pinkNote)
  stickyNotes.commit()
  stickyNotes.reload_notes()
  stickyNotes.close_db()
else:
  print('Windows 10 Sticky Notes app not found')
  sys.exit()
```
---
Note: The default Microsoft Sticky Notes db directory is:
```
%UserProfile%\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\
```
