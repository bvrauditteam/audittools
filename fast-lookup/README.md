# Fast Lookup (Windows sidebar)

A floating sidebar for looking up profit centers, GL accounts, bank details and user IDs. It works fully offline.

- **Open:** move the mouse to the **right edge of the screen** and hold it there for a moment, or press **Ctrl+Shift+F**, or click the tray icon.
- **Close:** move the mouse away and it slides back. You can also press **Esc** or click **»**.
- **Pin (📌):** keeps the sidebar open until you unpin it.
- **Search:** type one or more words. Every word must match somewhere in the row.
- **Copy:** click any value to copy it. **Enter** copies the first value of the top result.
- **Menu (⋯, or right-click the tray icon):** Import Excel / CSV / JSON, Reload, Open data folder, Clear database, Start with Windows, Exit.

## Get the exe

Pick one:

1. **Build it on your PC:** double-click `build.bat`. It uses the C# compiler that comes with Windows, so there's nothing to install. It creates `FastLookup.exe`.
2. **Download it:** on GitHub, open **Actions → Fast Lookup (Windows build)**, pick the latest run and download **FastLookup-exe**.

Copy `FastLookup.exe` anywhere you like and run it.

## Data

Import your Excel file directly. There's no need to make JSON first.

1. Open the sidebar and choose **⋯ → Import Excel / CSV / JSON**, or drag the file onto the sidebar.
2. Pick your `.xlsx` or `.xlsm` workbook. Every sheet that has data becomes its own searchable list, named like `Masters - GL`.

What your Excel file needs:

- **Row 1 of each sheet holds the column headings**, for example `G/L account | Long Text` or `Region | RO | name | plant | profit center`. Every row below that is one lookup record.
- Blank rows and empty sheets are skipped. Numbers such as GL accounts come through as plain text (`1000000010`).
- You can import while the workbook is still open in Excel.
- For an old `.xls` file, in Excel use **File → Save As → Excel Workbook (.xlsx)** first.
- `.csv` files work too, whether they use commas or semicolons. So do the `.json` files from the old extension.

When your data changes, import the file again. Each sheet replaces its previous copy.

The app keeps its copy in `%LOCALAPPDATA%\FastFinanceLookup\data`, so everything works offline. A `data` folder placed next to the exe is also read.

**Keep your data files out of this repo, because the repo is public.**
