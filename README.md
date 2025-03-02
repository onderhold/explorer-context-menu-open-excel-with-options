[//]: # (see https://www.markdownguide.org/basic-syntax/ or https://commonmark.org/help/ for Markdown syntax )

# Context menu extension for Windows file explorer: Opening Excel files in various modes

## Using Excel in different modes

When developing reports, templates and applications in Excel, it is often useful to be able to open Excel files in different modes:

- **write-protected** in order to ensure from the outset that no unintentional modifications are saved to the file
- in a **separate instance** in order to be able to open the same file multiple times
- in a **separate instance** in order to be able to open files with the same name from different folders at the same time
- using a file as the **template** for a new file
- in **safe** mode ("safe")
- in a combination of these modes

## Solution

see ["Command-line switches for Microsoft Office products"](https://support.microsoft.com/en-gb/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6#Category=Excel)

- A cascaded context menu is defined for the current user (HKCU) using registry entries
- This menu is displayed at the top of the context menu when an Excel file is selected
- This menu is also displayed when it is called up in an empty area of the file explorer or when a folder is selected.<br />
However, in that case, the context menu is no longer placed at the top of all menu items, and contains only menu items for starting Excel without opening any file
- the start option "/p", which is supposed to set the working directory for Excel, unfortunately does not work and is therefore not used.
- The menu entries are given shortcuts. As by default in Windows 11 the shortcuts in the context menu are not underlined (at the current time 2023-09-28), the character for the shortcut key is set in capital letters.

[//]: # (> Note: Shortcuts are only displayed underlined if the corresponding system setting is activated:)
[//]: # (>)
[//]: # (> ![shortcut_underline_settings] ./img/keyboard_settings_shortcut_underline.de.png)
[//]: # (<br />)

### Context Submenus

***

These are the shortcut keys that you can use:

In the file explorer, select an Excel file, then start with \[SHIFT-F10\] \[x\] or a right mouse click,<br />
afterwards, for opening the file, choose:
- \[r\] &rarr; open <u>R</u>eadonly
- \[s\] &rarr; open in <u>S</u>eparate instance
- \[x\] &rarr; open separate instance, readonly \[<u>X</u>\]
- \[t\] &rarr; use this file as <u>T</u>emplate, readonly
- \[m\] &rarr; use this file as te<u>M</u>plate, in separate instance
----
or for starting Excel without opening a file, choose:
- \[n\] &rarr; start Excel <u>N</u>ormally, without splash
- \[f\] &rarr; start Excel in sa<u>F</u>e mode, without splash
- \[i\] &rarr; start Excel in separate <u>I</u>nstance, without splash
- \[w\] &rarr; start Excel in separate instance and safe mode, <u>W</u>ithout splash

![kontextmenü_datei](./img/file.png)


### Submenu in an empty area of the file explorer

***

The menu item is no longer displayed at the top, and of course only the options for starting Excel without opening a file are available

![kontextmenü_verzeichnis](./img/dir.png)

## Installation

- Double-click on the file "***\[+\]myExcel***_dev_menu.hkcu.reg" to import the settings into the registry
- Double-click on the file "***\[-\]myExcel***_dev_menu.hkcu.reg" to remove the settings from the registry

## More Information

- The .reg file contains comments which, among other things, refer to the sources used on the Internet.
- It's valid to add comments to a .reg file. They are ignored when the file is imported into the registry.