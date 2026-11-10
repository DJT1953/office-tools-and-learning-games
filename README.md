# Office Tools and Learning Games

Free fresh Windows tools for working through batches of documents, photos, and files, six remedial learning spoken games for practicing reading and counting/money, and a Force Account/Change Order cost Excel workbook for South Dakota DOT.

**[Open the page with all games and downloads](https://djt1953.github.io/office-tools-and-learning-games/)**

You may download and use everything here for free, but do not modify or redistribute it without credit to this author. See [LICENSE](LICENSE).

## Learning games

The games speak out loud and run in a web browser, with nothing to install. Turn the sound on. They were built for adults practicing remedial reading and everyday money/counting skills. Each has a coach panel for adults adjusting the lessons, and progress stays in that browser on that computer.

| Game | Practice | Play |
|---|---|---|
| Word Race | Hear a word, then find it, while racing a rival | [Play](https://djt1953.github.io/office-tools-and-learning-games/games/word-race.html) |
| Oval Showdown | Match words in upper and lower case on an oval track | [Play](https://djt1953.github.io/office-tools-and-learning-games/games/oval-showdown.html) |
| Haunted House Payday | Count out dollar bills and get change back | [Play](https://djt1953.github.io/office-tools-and-learning-games/games/haunted-house-payday.html) |
| Christmas Gift Run | Pay with bills and coins and watch your change | [Play](https://djt1953.github.io/office-tools-and-learning-games/games/christmas-gift-run.html) |
| Menu Run | Read menu words and learn where food comes from | [Play](https://djt1953.github.io/office-tools-and-learning-games/games/menu-run.html) |
| Carol Match | Match real photos to words; each match spells the word aloud and plays a note of a Christmas carol | [Play](https://djt1953.github.io/office-tools-and-learning-games/games/carol-match.html) |

[Download all six games](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/learning-games.zip) to play without an internet connection. The fonts need internet, and plain fonts are used without it.

Carol Match's photos come from Wikimedia Commons and keep their own licenses (public domain, CC0, CC BY and CC BY-SA), listed in the game's Photo credits and in [the photo credits](docs/games/carol-match-photos/CREDITS.md).

## Windows tools

Unzip any tool below, then start it as shown. Your originals are left alone unless a tool's instructions say otherwise, but keep a backup when working on important files.

| Tool | What it does | Needs | Start with | Download |
|---|---|---|---|---|
| Document Reformatter | Reformats folders of .rtf, .doc and .docx files | Microsoft Word; LibreOffice optional | `Create desktop shortcut.cmd` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/document-reformatter.zip) |
| Document Renamer | Renames files from the first characters of their text, or with typed text, numbers, dates and more | Microsoft Word for names from text; LibreOffice optional | `Create desktop shortcut.cmd` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/document-renamer.zip) |
| Find and Replace | Changes words, marks and formatting across many documents, and deletes sections between two phrases | Microsoft Word | `Create desktop shortcut.cmd` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/find-and-replace.zip) |
| Duplicate File Finder | Finds exact copies, matching documents, and look-alike photos and videos | Python 3 | `Duplicate_File_Finder.py` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/duplicate-file-finder.zip) |
| Cross-Folder Compare | Finds photos and documents in one folder that match another | Python 3 | `Cross_Folder_Compare.py` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/cross-folder-compare.zip) |
| Clean Up Scans | Straightens, crops and tick box for whitening photographed pages | Python 3 | `Run_CleanScans.bat` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/clean-up-scans.zip) |
| Combine Word Files | Joins a folder of .docx files into one, each keeping its look | Python 3 | `Run_CombineDocx.bat` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/combine-word-files.zip) |
| Photos text to Word | Turns photos of text into Word files, keeping the layout | Python 3, Tesseract OCR | `Run_PhotoText.bat` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/photos-to-word.zip) |
| Video Player | Plays MP4, MKV, AVI, MOV, WMV and other common videos with sound, and saves any frame as a PNG picture | Python 3; downloads a 150 MB add-on the first time | `Run_VideoPlayer.bat` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/video-player.zip) |
| Video Editor | Trims, splits, inserts and splices videos and sound files, with volume, fades and added music or narration, then saves in the format and folder you choose | Python 3; downloads about 190 MB the first time (40 MB with the Video Player); see the memory warning below | `Run_VideoEditor.bat` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/video-editor.zip) |
| Music Player | Plays the music on this PC, internet radio with pause and rewind, and Spotify, with big buttons and a sleep timer | Python 3.10 or newer; downloads about 190 MB of add-ons the first time; the Spotify app for Spotify; see Music Player needs below | `Run_MusicPlayer.bat` | [Zip](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/music-player.zip) |

**Video Editor memory warning:** saving long 4K edits needs a lot of memory, about 5 GB for 20 pieces of 4K video (HD needs far less). On a PC with 8 GB or less, keep 4K edits short or save them in parts.

**Music Player needs:** Windows 10 or 11 and [Python 3.10 or newer](https://www.python.org/downloads/windows/) (tested with 3.14). The first time it starts, it offers to download and install its free add-ons, about 190 MB (less if the Video Player or Video Editor is already set up): [PySide6](https://pypi.org/project/PySide6/), [Mutagen](https://pypi.org/project/mutagen/), [imageio-ffmpeg](https://pypi.org/project/imageio-ffmpeg/), [winrt-Windows.Media.Control](https://pypi.org/project/winrt-Windows.Media.Control/), [winrt-Windows.Media](https://pypi.org/project/winrt-Windows.Media/) and [pycaw](https://pypi.org/project/pycaw/). Radio needs an internet connection; the stations come from [Radio Browser](https://www.radio-browser.info/). Spotify is optional and needs the Spotify app ([Microsoft Store](https://apps.microsoft.com/detail/9ncbcszsjrsb) or [spotify.com](https://www.spotify.com/download/windows/)) and a free or Premium Spotify account. The zip's `How to use.txt` has the same list.

Get Python 3 from [python.org](https://www.python.org/downloads/). Then the tools install any free add-ons they need the first time they run, which takes an internet connection.

## SDDOT force account workbook

An Excel workbook that totals labor, materials, equipment, overhead, profit, the prime contractor's administrative allowance, and excise tax under Section 9.4 D of the South Dakota DOT specifications.
[Download the workbook](https://github.com/DJT1953/office-tools-and-learning-games/releases/latest/download/sddot-force-account-workbook.zip)

It is an independent aid, not affiliated with or endorsed by the South Dakota Department of Transportation. Confirm it against your own contract and with your Engineer before relying on its figures. The `Inputs` tab lists the places where the specification leaves room for interpretation.
