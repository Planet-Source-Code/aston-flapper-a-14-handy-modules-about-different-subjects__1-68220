In this map are some modules that can come in very handy :), because there are so much functions and subs 
I made this function list:
			self 
			made
what	sub/funct/val	or not	information
---------------------------------------------------------------------------------------------------------

modBatch:			### This module is here to use the batch language in visualbasic ###
prop	batString	yes	Batch data
prop	batPath		yes	Path to create the batchfile in
sub	batClear	yes	Clear batString
sub	batPrint	yes	Prints something in batString
sub	batKill		yes	Kills the batchfile, if created.
funct	batRun		yes	Runs a batchfile, and returns output of batchfile

modEncryption:			### Some encryption methods ###
funct	md5		no	Creates md5 hash
funct	sha1		no	Creates sha1 hash
funct	toBase64	no	Encrypts to base64 encryption format
funct	fromBase64	no	Decrypts from base64 encryption format

modEnum:			### Originally made for all enum functions, but only 1 is in there ###
funct	enumProcesses	yes	Enumerates all system processes by EXE name

modExeData:			### Module to add files of 3GB or lower to an EXE, and to read from it ###
funct	edGetData	yes	Returns all data in the backend of the EXE file
funct	edSetData	yes	Puts some data to the backend of an EXE file(3GB limit)
funct	edRemoveData	yes	Removes data from the backend of an EXE file
prop	edPath		yes	Returns the full path to the EXE file of the current process

modFiles:			### I don't use it often, it is used for basic file handling ###
funct	FileAppend	yes	Append data to an file
funct	FileWrite	yes	Deletes the file, and replaces it with the new data
funct	FileDelete	yes	Deletes the file, without returning an error when the file doesn't exists
funct 	FileData	yes	Returns all data in the file
funct	FileMove	yes	Moves the file to another map, or renames it
funct	FileCopy	yes	Copys the file to another file
funct	FileAttributes	yes	Returns the attributes of an file

modInStart:			### So you want your program to start automatically when your computer starts? ###
sub	PutInStart	yes	Makes your program start when the pc starts, value = your file location, var = "some_text_without_spaces_that_is_unique_to_every_program"

modMouse:			### I don't think I have to explain this ###
prop	MouseX		yes	X position of the cursor
prop	MouseY		yes	Y position of the cursor
prop	MouseDoubleClickTime	yes	Time between 2 mouse clicks to see it as an doubleclick.
sub	MouseClip	yes	Lock mouse into an squere
sub	MouseSwapButtonsyes	Swap the buttons of the mouse

modNumtoStr:			### Changes an number (50) to an ascii-string(in hex: "00 00 00 32") ###
funct	nts4T		yes	Change number to an ascii-string
funct	nts4F		yes	Change ascii-string to an number

modStrings:			### All string-related functions ###
funct	StrToArrayStr	yes	Change an array into an array-string
funct	StrFromArrayStr	yes	Change an array-string into an array
funct	ArrayStrLen	yes	Recieve the length of an array-string in arrays
funct	StrByte		yes	Recieve one character from an position in an string.
funct	StrCompare	yes	Compare two strings, and return an number
funct	StrToMix	yes	Add two strings into one string: Str1 + Str2 = Str3
funct	StrFromMix	yes	Get a string from an combined string: Str3 - Str2 = Str1
funct	StrToURL	no	Changes an string to an URL-format string
funct	StrFromURL	no	Changes an URL-format string to an normal string
funct	StrSetLength	yes	Set the length of an string: "115" becomes "00115" for example
funct	StrIn		yes	Checks if Str2 is in Str1
funct	StrReplace	yes	("test by aston", " ", ";", "by", "of") becomes "test;of;aston"
funct	StrLeft		yes	Advanced Left$()
funct	StrRight	yes	Advanced Right$()
funct	StrTo		yes	Returns the string before Str2 in Str1 ("test by aston", " by ") returns "test"
funct	StrFrom		yes	Returns the string after Str2 in Str1 ("test by aston", " by ") returns "aston"
funct	StrBetween	yes	Returns the stringg between Str2 and Str3 in Str1 ("test by aston", " ", " ") returns "by"
funct	StrToHex	yes	Changes an normal string to an hex-string
funct	StrFromHex	yes	Changes an hex string to an normal string

modSysTray:			### Add your program to "that heaven for icons in the right-bottom corner" ###
funct	TrayAdd		?	Add your form to the Tray
funct	TrayModify	?	Modify the icon in the Tray
funct	TrayDelete	?	Delete the tray icon of your form
				* To recieve mouse clicks, to create an menu use the Form_MouseUp

modTreeView:			### I hate treeviews, why? they make it so hard to easily edit them ###
funct	tvnPut		yes	Create an treeview from this treeview format:
				Plusstring = "->"; NodesText = 
				Node1
				->SubNode1.1
				->SubNode1.2
				Node2
				->SubNode2.1
				->->SubSubNode2.1.1

modWebcam			### Add pictures directly from the webcam to your picturebox ###
sub	wbcStart	yes	Start the webcam
sub	wbcStop		yes	Stop the webcam
sub	wbcToPicture	yes	Create an snapshot, and load it in the picturebox


modWindows			### My newest module, it is not finished and I barely beginned ###
funct	wndHide		yes	Hide window
funct	wndShow		yes	Show window
funct	wndFlash	yes	Flash window








SpecialApi:

About specialapi, it was my first module where I put all API functions back then when I was a newbie, 
and didn't understand all those API functions. There are some errors in it, the code is NOT nice
so please don't complain about this. I am to lazy to write all functions and information about them
in here, so sorry about that.


Afterword:
Sorry about the bad english :P I am dutch.
You can use all this code for free it all programs.
AND HAVE FUN!!!





