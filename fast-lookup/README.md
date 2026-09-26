# Fast Lookup (Windows sidebar)

A floating sidebar for looking up profit centers, GL accounts, bank details and user IDs. It works fully offline.

- **Open:** move the mouse to the **right edge of the screen** and hold it there for a moment, or press **Ctrl+Shift+F**, or click the tray icon.
- **Close:** move the mouse away and it slides back. You can also press **Esc** or click **»**.
- **Pin (📌):** keeps the sidebar open until you unpin it.
- **Search:** type one or more words. Every word must match somewhere in the row.
- **Copy:** click any value to copy it. **Enter** copies the first value of the top result.
- **Menu (⋯, or right-click the tray icon):** Import JSON files, Reload, Open data folder, Clear database, Start with Windows, Exit.

## Get the exe

Pick one:

1. **Build it on your PC:** double-click `build.bat`. It uses the C# compiler that comes with Windows, so there's nothing to install. It creates `FastLookup.exe`.
2. **Download it:** on GitHub, open **Actions → Fast Lookup (Windows build)**, pick the latest run and download **FastLookup-exe**.

Copy `FastLookup.exe` anywhere you like and run it.

## Data

On first run, choose **⋯ → Import JSON files** and select your files (for example `CAO.json`, `finance_database.json` and `usernames.json`). The app copies them to `%LOCALAPPDATA%\FastFinanceLookup\data`, so you only do this once.

- Importing a file with the same name replaces the old copy.
- A `data` folder placed next to the exe is also read.
- The files must be JSON arrays of rows, like the old extension used. Excel exports that aren't UTF-8 are handled.

**Keep the JSON files out of this repo, because the repo is public.** `.gitignore` already excludes them.
