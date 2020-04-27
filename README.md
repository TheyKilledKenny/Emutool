# Emutool
EmuTool
Move, Copy, Convert, Backup and Restore your SXOS and Atmosphere Emunand/emuMMC.

For support go to GBATemp, <a href="https://gbatemp.net/threads/emutool-move-partition-emu-on-another-sd-switch-emu-type-on-sxos-and-more.550756/">click here</a>: 

<center><img src="https://gbatemp.net/attachments/upload_2019-11-18_20-25-34-png.187157/" /></center>
(Art by @CrazyKing93)​

(Emunand or emuMMC are the same concept, below will be referred to simply as Emu)

Italian language guide <a href="https://graph.org/EmuTools-migrazione-backup-e-dintorni-per-la-tua-emuMMC-eo-Emunand-10-21">here</a>

Warning!
Antivirus program can block and delete EmuTools.<br />
EmuTool uses these functions of Windows Kernel32: CreateFile, CloseHandle, DeviceIoControl, SetFilePointer, ReadFile, GetFileSize, WriteFile, FlushFileBuffer, LockFile, UnlockFile<br />
Check source code if you need.
<br />

1. What is it for?

    - Move the Hidden partition Emu (Atmosphere\SXOS) on a new (bigger?) SD
	- On SXOS allows to switch between hidden partition and file Emunand mode, so you can have two Emunand on the same SD card.
	- Change your Emu format from Atmosphere to SX OS and viceversa and from Hidden partition Emu to Emu on file
	- Backup and restore of every type of Emu.
	- Create a new Emu (for experiments?) starting from you current Hekate/SXOS backup or from your current Emu
	- Split an Hekate backup in multiple files ready for fat32 partitions
	- Create the relevant configuration files to boot the Emu (emummc.ini, raw_based, file_based and folder structure)


2. Using the Tool
To access the SD partitions run EmuTool.exe <strong>with administrator rights</strong>, it should already ask for it, if not please do a right click and choose Run as Administrator...
To start, double click on the EmuTool.exe file and confirm administrative rights if necessary.
EmuTool requires that you select a source (Source), ie where to read the Emu, and a destination (Target), ie where to write the copy of the Emu.
Both Source and Target support Partition and Files.
When Source and Destination are set press the Start button to start copying.
(I apologize for the UI quality. This was a tool created just for me and I wanted to keep it lightweight, free of dependencies, no installation and easily usable, as I think a tool so limited and specific should be)​
<center><img src="https://gbatemp.net/proxy.php?image=https%3A%2F%2Ftelegra.ph%2Ffile%2Ffefd4f169c294b50ab045.png&hash=90ea5814b82e156322c14505799948d4"></center>


3. How to Select SD card and File
After selecting the type of Emu you want as a Source or Target, click on the "white" box with the words "Click to select SD Card" in the frame relating to the Emu type selection.
A navigation window will appear depending on the Emu type selected
In case of Partition type, the following window will appear​
<center><img src="​https://gbatemp.net/proxy.php?image=https%3A%2F%2Ftelegra.ph%2Ffile%2F2f2820ee5df6c9f967084.png&hash=dec5094e5f5588bc5ef55e7d6ee5bf2b"></center>
​
Selecting the drive containing the SD shows a list of the partitions present on the SD card.
The Sector field located at the bottom right is important.
The first value is read by emummc.ini if it is correct, check the emummc.ini file in the emummc folder of the SD card if this value is incorrect. If you are using Kosmos simply select the Emu from the emuMMC menu.
If you select a partition from the list, then the initial sector of the partition, added to the 16Mbyte offset, will be shown in the Sector field.
No partition selection for SXOS as it is fixed to 0x2 on the first patition.
If something is wrong then you can correct the value, in hex (with notation 0x as in 0x02AC2300) or in decimal (for example the value read from Minitool Partition Wizard)

After confirmed with Ok the main screen displays the data related to the Emu to read
<center><img src="​https://gbatemp.net/proxy.php?image=https%3A%2F%2Ftelegra.ph%2Ffile%2F2ddc5984d870dcf2e91c2.png&hash=498cc8582bff0f9143fc5faf595c32a9"></center>

If you choose The File Type Emu the following window will appear
[<center><img src="https://gbatemp.net/proxy.php?image=https%3A%2F%2Ftelegra.ph%2Ffile%2F8332cd2bcd1ae3168afb9.png&hash=92d1113a6b2a882892a1fc2307fffc0b"></center>

Select the destination folder and the path will be shown on the main screen in the white box
<center><img src="​https://gbatemp.net/proxy.php?image=https%3A%2F%2Ftelegra.ph%2Ffile%2Fd50f510cc9fecee6194f1.png&hash=f9796b2411c5e16cfdf287828f4e87a4"></center>

When everything is set as desired press Start.​

4. Enabling/disabling Emunand SXOS on partition (allows to start emunand on file)
Select Source SXOS hidden partition and click the white box to select the SD card drive. When SD is selected two new buttons will appear in the main window
[​IMG]​

WARNING!!!
No check is made on the actual existence of Emunand, so you can enable emu on partition even if this partition does not exist.

The state of the buttons indicates the current status of Emunand on partition.
Click on "Enable Partition Emu" to enable reading the Emunand from the hidden partition.
Click on "Disable Partition Emu" to disable it. In this case the File Emunand will be loaded if there is a valid SXOS File Emunand on SD in the sxos\Emunand folder.
You can now prepare a SXOS Emunand file.
Select the SXOS File type as Target, select the SD root as the path and press Start to create a copy of your Emunand partition in Emunand file, without the need to use hard disk and without having to reload the cfw files on the SD.
EmuTool Create the sxos/Emunand folder starting from the point you choose as the destination folder. Inside the Emunand foledr you will find the Emunand files.

​
5. Change SD for those who have Emu on partition
To bring the Emu to the new SD there are two possibilities:
1 - Copy the Emu to file and use it in File mode (see section 6)
2 - Create a special partition of at least 30GByte on the new SD and transfer the Emu on it, however, requires the passage at point 1 (first go to section 6 and then to section 7)
​
6. Convert partition Emu in Emu on File
Select as Source the type of partiton Emu, Atmosphere or SXOS, you want to read/copy and then select the SD card by clicking on the white box.
Select as Target the type of Emu on file you want to get and click on the white box to select a destination path.
The Emu-related folders will be created, ready to be copied to the SD root. In the case of the Atmosphere type, the emummc.ini and file_based files are also created.
For SXOS the sxos folder is created and within it the Emunand folder is created. For Atmosphere the emuMMC folder is created, the emummc.ini file compiled and the HPE0 folder which will contain the eMMC folder with the Emu files and the file_based file needed by Kosmos\Hekate\Nyx.

Press Start to begin.
​
7. Transfer of the Emu to partition
Select as Source the Emu File type and select the folder that contains the Emu files, not the root folder (ie choose sxos\Emunand or emummc\HE0\eMMC folders).
Select as Target the type of partition Emu that you want to restore, in case of Atmosphere type you will also have to indicate the initial sector of the partition you created to host the Emu. Please check the other tutorial on how to create a suitable partition for emuMMC (you can use the free minitool partition software).

Press the Start button to start copying.​

6. Conclusion
For any other operation you can think of, the way to select the SD reader and browse through the folders does not change, so I guess I shouldn't bother you with unnecessary chatter ;-)

I do not offer any guarantee for the use of this software. I am a very bad programmer, so if you use EmuTool it will be at your own risk.
If using this software the processor should start like a rocket and cut your head off (Hey They killed kenny! ... You Bastard!), I'm not responsible.

With this software you can do whatever you want (copy, distribute, decompile, write the user TheyKilledKenny is idiot, etc.), but NOT SELL in any ways. If you recover even a single penny with this software you must immediately donate to charity, otherwise you are a thief.


Thanks to the user @GraFfiX420 for pointing me in the right direction with his tutorial
https://gbatemp.net/threads/moving-from-sx-os-sd-emunand-to-sd-hidden_emunand.526587/


Ciao!



Changelog:
Version 0.2.9
Fixed restore from Hekate backup file, now rawnand.bin or rawnand.bin.xx are supported as single file backup or up to 51 files splitted backup (rawnand.bin.50)
Added partition selection for Atmosphere
Minor fixes in partition selection for Atmosphere

Version 0.2.8
Solved bad starting sector report when Atmosphere hidden partition was chosen as Source
Fixed some minor bugs found during more tests

Version 0.2.7
Added a minimum of error trapping, useful to debug errors
Added partition selection, 16MByte offset will be added to the real partition start sector because Kosmos\Hekate do it when I create a new emu on partition from starting menu. Text filed is always editable to correct if needed.
Some other minor error traping and correction around.
Not solved overflow error that someone reported, maybe the error trap can be useful.

Version 0.2.6
Added different file size for Amosphere eMMC and SXOS size, it should solve some slowness problem during eMMC boot.
Trapped an overflow error during SD read at the start if not executed as Admin
Sector field for Atmosphere partition is now always editable
Changed the icon of the main form that caused crashes on some Windows7 systems.

Version 0.2.5
Test Version

Version 0.2.4
First public release