![PSArrange_Logo](https://github.com/user-attachments/assets/fcdf715b-e427-47df-8e13-20161a944e60)

# PSArrange

**Tidy up your directories effortlessly with rule-based organization!**

---

## The Problem

Are your Desktop and Downloads folders constantly overflowing? Do you find yourself spending precious time manually sorting photos, documents, code files, and archives into the right places? It's a common frustration!

## The Solution: PSArrange ✨

PSArrange is a flexible PowerShell script designed to bring order to your digital chaos. It automatically organizes files and folders based on a set of clear rules you define in a simple JSON configuration file. Let the script handle the sorting, so you can focus on more important things!

## Features

* **Rule-Based Power:** Define exactly how you want things organized using a central `config.json` file.
* **Flexible File Rules:**
    * Match files by one or multiple **extensions** (e.g., `.jpg`, `.png`, `.gif`).
    * Match files using **name patterns** with wildcards (e.g., `Invoice_*.pdf`, `*Report*.*`).
    * Filter files based on **age** (older or newer than a specified number of days).
    * Move files to designated **target folders**.
    * Create nested **subfolder** structures.
    * Automatically generate **date-based subfolders** (`YYYY/YYYY-MM`) for chronological sorting (great for photos!).
    * Choose how to handle **duplicates** if a file already exists at the destination (Skip, Rename with Timestamp, Overwrite) – globally or per-rule.
* **Smart Folder Rules:**
    * **Rename** folders matching specific patterns (e.g., add `_ARCHIVED` to old project folders).
    * Use **placeholders** like `{OriginalName}`, `{JJJJ}` (Year), `{MM}` (Month), `{TT}` (Day) in rename templates.
    * **Move** entire folders matching patterns to designated archive or category locations.
    * Apply folder rules based on **age**.
* **Global Exclusions:** Protect important system folders, version control directories (`.git`), specific temporary files, or cloud storage sync folders from being processed.
* **Catch-All Default:** Any file not matching a specific rule can be routed to a designated default folder (e.g., "ToSort") so nothing gets lost.
* **Safety First (Preview Mode):** Run the script with the `-Preview` switch to see a detailed log of exactly what *would* happen (which files moved where, which folders created) without actually changing anything on your disk. **Always use this first!**
* **Optional Cleanup:** Use the `-DeleteEmptyFolders` switch to automatically remove any folders left empty after the organization process.
* **Logging:** Every action (or proposed action in preview mode) and any errors are recorded in the `organisation.log` file for review.

## Prerequisites

* **Operating System:** Windows
* **PowerShell:** Version 5.1 or higher (usually included with modern Windows versions). You can check your version by opening PowerShell and running `$PSVersionTable`.

## How to Use

1.  **Get the Files:** Clone this repository or download the `PSArrange.ps1` script and an example `config.json`.
2.  **Configure `config.json`:** Edit the `config.json` file to define *your* directories and organization rules (see Configuration section below).
3.  **Open PowerShell:** Navigate to the directory where you saved `PSArrange.ps1` and `config.json`.
4.  **Run in Preview Mode (CRITICAL FIRST STEP!):**
    ```powershell
    .\PSArrange.ps1 -ConfigFile .\config.json -Preview
    ```
5.  **Review:** Carefully check the console output and the `organisation.log` file. Does the proposed organization match what you expect? Are the rules matching the correct files? Are the target paths correct?
6.  **Adjust `config.json`:** Modify your rules in `config.json` based on the preview until it looks perfect. Run the preview again after changes.
7.  **Perform Organization (After Backup!):** Once you are completely satisfied with the preview:
    * **Strongly Recommended:** Back up the directories you are about to organize!
    * Run the script without the `-Preview` switch:
        ```powershell
        .\PSArrange.ps1 -ConfigFile .\config.json
        ```
8.  **Optional - Delete Empty Folders:** If you want to clean up empty folders *after* organizing, add the `-DeleteEmptyFolders` switch (can be combined with or without `-Preview`):
    ```powershell
    # Preview empty folder deletion
    .\PSArrange.ps1 -ConfigFile .\config.json -DeleteEmptyFolders -Preview

    # Actually organize AND delete empty folders (Use with caution!)
    .\PSArrange.ps1 -ConfigFile .\config.json -DeleteEmptyFolders
    ```

## Configuration (`config.json`)

This file tells PSArrange what to do. Here's a breakdown of the main sections:

* **`Directories`**: An array of strings specifying the full paths of the main folders you want to organize. If left empty (`[]`), the script might attempt to automatically find common user folders (like Desktop, Downloads, Documents - *current implementation might vary*).
    ```json
    "Directories": [
        "C:\\Users\\YourName\\Desktop",
        "C:\\Users\\YourName\\Downloads",
        "D:\\ProjectsToSort"
    ],
    ```
* **`GlobalExclusions`**: Defines patterns for files and folders that should *never* be touched.
    * `FilePatterns`: Array of file patterns (e.g., `"*.tmp"`, `"~*.*"`).
    * `FolderPatterns`: Array of folder path patterns (e.g., `"*\\node_modules"`, `"C:\\Windows*"`). Wildcards (`*`) are supported.
    ```json
     "GlobalExclusions": {
        "FilePatterns": [ "*.tmp", "~*" ],
        "FolderPatterns": [ "*\\$RECYCLE.BIN*", "*\\.git*", "*\\AppData*" ]
      },
    ```
* **`GlobalDuplicateHandling`**: Default strategy if a file is moved to a location where a file with the same name already exists. Options:
    * `"Skip"` (Default): Leaves the source file untouched.
    * `"RenameWithTimestamp"`: Moves the file, renaming it like `original_name_YYYYMMDDHHmmssfff.ext`.
    * `"Overwrite"`: Replaces the existing destination file (Use with caution!).
    * `"Ask"`: (Not implemented in non-interactive mode, currently acts like "Skip").
    ```json
    "GlobalDuplicateHandling": "RenameWithTimestamp",
    ```
* **`FileRules`**: An array of rule objects. The *first* rule that matches a file determines its fate.
    ```json
    "FileRules": [
      {
        "Description": "Holiday Pictures", // Just for your reference
        "Extensions": [ ".jpg", ".jpeg", ".png", ".heic", ".raw" ], // Which file types
        "NamePatterns": [ "IMG_*", "DSC*", " Urlaub*" ], // Optional: Match filename patterns
        "OlderThanDays": 0, // Optional: Process only if older than X days (0=ignore)
        "NewerThanDays": 0, // Optional: Process only if newer than X days (0=ignore)
        "TargetFolder": "Bilder", // Destination relative to the base directory
        "SubFolder": "Urlaub", // Optional: Subfolder within "Bilder"
        "OrganizeByDate": true, // Creates YYYY/YYYY-MM folders inside "Urlaub"
        "Action": "Move", // Must be "Move" currently
        "DuplicateHandling": "RenameWithTimestamp" // Optional: Override global setting
      },
      {
         "Description": "Important Documents",
         "Extensions": [".pdf", ".docx"],
         "NamePatterns": ["Contract*", "Invoice*", "Report*"],
         "TargetFolder": "Dokumente",
         "SubFolder": "Wichtig"
      }
      // ... more file rules
    ],
    ```
* **`DefaultTargetFolder`**: Folder name (relative to the base directory) where files go if *no* `FileRules` match them.
    ```json
    "DefaultTargetFolder": "Sonstiges - Zu Prüfen",
    ```
* **`FolderRules`**: An array of rule objects applied *recursively* to folders within the target directories. The first matching rule triggers an action.
    ```json
    "FolderRules": [
        {
            "Description": "Rename old Backup Folders",
            "RenamePattern": "*Backup*", // Folder name pattern to trigger rename
            "RenameOlderThanDays": 90, // Only rename if folder hasn't been modified in 90 days
            "NewNameTemplate": "ALT_{OriginalName}_{JJJJ-MM}" // How to rename it
            // MovePattern, MoveOlderThanDays, TargetFolder are null/0 if only renaming
        },
        {
            "Description": "Archive Folders marked 'Temp'",
            "MovePattern": "*Temp*", // Folder name pattern to trigger move
            "MoveOlderThanDays": 7, // Only move if folder hasn't been modified in 7 days
            "TargetFolder": "Archiv\\Temporary" // Destination relative to the base directory
             // RenamePattern, RenameOlderThanDays, NewNameTemplate are null/0 if only moving
        }
      // ... more folder rules
    ]
    ```

## Known Issues / Quirks

* **`$false` Recognition Workaround:** During development in the original testing environment, a strange issue occurred where PowerShell sometimes failed to recognize the built-in `$false` variable during assignments (e.g., `$myVar = $false`), causing a "term not recognized" error. A workaround using the equivalent `!($true)` (logical NOT true) was implemented on the affected lines within the script (`$isMatch = !($true)`, `$appliesRenameRule = !($true)`, etc.). If you modify the script or run it in your environment and encounter this specific error with `$false`, this workaround might be necessary. However, standard PowerShell environments should recognize `$false` correctly, so ideally, investigate your specific environment if this happens. The current script includes the workaround for stability based on prior testing.
* **Console Encoding:** Depending on your terminal/console settings, special characters like German Umlauts (ä, ö, ü) might not display correctly in the *live output*. Running `[System.Console]::OutputEncoding = [System.Text.Encoding]::UTF8` in PowerShell before executing the script can often fix this. The `organisation.log` file *should* correctly use UTF-8 encoding.

## Contributing

Feel free to open an issue on GitHub if you find bugs or have suggestions for improvements!

## License

MIT License

Copyright (c) [year] [fullname]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

---

Happy Organizing!
