MERB: Multi-EditROM Builder  (c)2017-2021 Steve J. Gray
===========================  Version 1.5, Feb 5, 2021

This is a tool to build binary images for my Multi-ROM projects:

* Multi-EditorROM
* Multi-EditorROM+
* Multi-EditorROM with 40/80 column switcher
* PET/CBM Multi-ROM

These boards allow you to have 16 Editor, Firmware, or Character ROMs in one EPROM and select
which one you want via DIP switch or 40/80 switch. Each ROM can be either 2K or 4K in size
and will be padded to fill a 4K slot.

About
-----

Click "About" to see the copyright, date and version# of the program.

This program was written in VB6, and requires the proper VB6 runtimes for Windows.


Selecting Files
---------------

The program will open with a single screen. There are 16 numbered "slots" where you can
specify a BINARY/ROM file to be included. One slot will be highlighted in RED. This is the
currently selected slot. To set a filename in the slot you can:

1) Click in any filename box and type a valid filename.
2) Double-click the slot number to open the file selector dialog.
3) Select a slot, then click "Add Binary".
4) Drag a file, or files, from a Folder to any slot.

Make sure your ROM/BIN files are pure images.

Do not use CRT files, P00 files or any file which contains header or any other non-ROM
content. Proper ROM files should be either 2048 or 4096 bytes long. If a file is 2050
or 4098 bytes long the program will assume the first two bytes are a load address and will
ignore them. If you know the file is correct but is missing 1 or more bytes, select the
"Allow short files" option. The file size of the selected file will be shown in the file
info area.

When you load a file that is 4096 bytes it MAY contain text or copyright embedded inside.
This will be displayed to the right of the slots along with the actual file size. Other
types of files will most likely show random text or nothing at all.

The program will hide the full path of the file except while editing the filename.


Ordering Files
--------------

If you find the files out of order you can use the bottom left buttons to arrange them.

Delete Entry..... Deletes the selected slot and moves lower slots up
Insert Entry..... Moves selected slot and below down leaving an empty slot
UP............... Moves the selected slot up
DOWN............. Moves the selected slot down.


Comparing Files
---------------

You can compare two or more files. First select the file you want to compare to, then click on the
Compare button. It will load the selected file into memory then read each of the other files,
comparing each byte. It will report if the files is longer or shorter and if the file is identical
or how many bytes differ. Empty slots, or invalid files will be ignored.

Working With Sets
-----------------

All 16 slots make a set. When all your slots are filled you can click "Save Set" to save
to a TXT file. Click "Load Set" to load a saved set.


Building a Set Image
--------------------

The Mode option lets you chose how 2K ROMs are handled. Pad will pad the file with zeros.
Duplicate will duplicate the file. For Firmware (Editor, Kernal, BASIC etc) ROMs you'll
probably want to Pad the file, unless you are creating a set for a ROM socket of 2K with
the high address set HI. For character sets you'll probably want to duplicate.

The "Allow short files" option lets you select binaries less than exactly 2K.

When your slots are set click "Build It!". It will check all 16 slots to make sure:

1) The file exists
2) The file is 2048 to 4096 bytes.

Files are automatically padded to 4096 bytes.
 
If all files are ok then it will ask you for a filename to save to. Enter one and click SAVE.
Your EPROM binary file is created!


Burning an Image
----------------

The resulting ROM/BIN file will be 65535 bytes (64K) long and should be burned to a 27512
(512KBit/64KByte) EPROM.


Conclusion
----------

This utility was written specifically to support my Multi-EditorROM projects and not as a general
usage tool. It was written to be a quick tool and may contain errors, bugs etc, and may crash with
incorrect input.

If you have comments or suggestions please contact me at:

sjgray@rogers.com
www.stevegray.ca

Thank-you!
