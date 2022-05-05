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

stickyNotes = StickyNotes()
newNote = Note('Hello', stickyNotes.Theme.blue, True)
pinkNote = Note('World', stickyNotes.Theme.pink, True)
stickyNotes.writeNotes(newNote, pinkNote)
stickyNotes.commit()
stickyNotes.reloadNotes()
stickyNotes.closeDB()
```

## The same example using a custom db directory
```py
import os
import sys
from StickyNotes import StickyNotes, Note

stickyNoteDBPath = '/custompath/

if os.path.isdir(stickyNoteDBPath):
  stickyNotes = StickyNotes(stickyNoteDBPath)
  newNote = Note('Hello', stickyNotes.Theme.blue, True)
  pinkNote = Note('World', stickyNotes.Theme.pink, True)
  stickyNotes.writeNotes(newNote, pinkNote)
  stickyNotes.commit()
  stickyNotes.reloadNotes()
  stickyNotes.closeDB()
else:
  print('Windows 10 Sticky Notes app not found')
  sys.exit()

```
---
Note: The default Microsoft Sticky Notes db directory is:
```
%UserProfile%\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\
```
