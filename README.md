# !["|-.-|"](icons8-merge-vertical-64.png) exe2wordsize
Detects Windows-compatible application bitness, **without ever running it.**

[![Download](https://img.shields.io/badge/download-success?style=for-the-badge&logo=github&logoColor=white)](https://github.com/TAbdiukov/exe2wordsize/releases/download/1.40/exe2wordsize.exe)

## Assumptions taken
* **Never runs the executable at any point**
* Host OS can run the application tested (for example, Windows XP, but not Vista, and a DOS app)
* Minimum byte-size applicable is always returned (for example, 32-bit apps on 64-bit host OS are analysed to use 32-bit word size, even if it is emulated)
* WinAPI may be glitchy, hence their output is doubted
* WinAPI should be versatile

## Current state
* Beta+


## Usage
```
exe2wordsize <path_to_app>
exe2wordsize <path_to_app> * <args>
```

### Examples
```
exe2wordsize C:/Projects/idk/Project1.exe
exe2wordsize "C:/Projects/idk/Project1.exe" * M=2 R=8192
```

### Manual
`<path_to_app>` - Path to your executable. `"`-tolerable

`*` - Delimiter. Only required to use optional args.

----

`<args>` - Optional arguments, space-delimited. Supported args below,
* M=(number) - Set analysis mode. 3 modes supported,
	* 0 - Automatic and flexible (Default)
	* 1 - Rely only on WinAPI. 64-bit input may be unreliable
	* 2 - Rely only on raw-reading. 32/64-bit detection only
* R=(number) - In raw-reading mode (M=2), how many bytes to read at most
(*Hint*: Only used in MODE = 2, default 8192)

----

#### Output

In JSON format,

* *path*: path supplied
* *args*: arguments supplied
* *time*: Unix timestamp
* *code*: (error-)code
* *code_desc*: (error-)code description
* *wordsize*: deduced wordsize
* *desc*: analytical description
* *walkthrough*: walkthrough process taken

## How to compile
1. *[Recommended for compatibility]* Get a Windows XP VM
2. Get **Microsoft Visual Basic 6.0** 

	* **Tip:** There is is a portable build, only a few megabytes. Look up <ins>Portable Microsoft Visual Basic 6.0 SP6</ins>

3. Start **Microsoft Visual Basic 6.0**, open up the project.
4. Go to File → Make *.exe → Save
5. Patch the app for CLI use:
	* You can use my [AMC patcher](https://github.com/TAbdiukov/AMC_patcher-CLI). For example,

		```
		amc C:\Projects\exe2wordsize\exe2wordsize.exe 3
		```
		
	* Or you can use the original Nirsoft's [Application Mode Changer](http://www.nirsoft.net/vb/console.zip) ([docs](http://www.nirsoft.net/vb/console.html)), unpack the archive and then run **appmodechange.exe**

6. Done!

## Found WinAPI bugs
### General
* `SHGetFileInfo` IS buggy on 64-bit executables. Hence `MODE=2` had to be implemented
* `GetBinaryType` BinaryType IDs are poorly documented

### File-reading
* For binary reading, `InputB` returns some patterned gibberish, despite documentation online. Use `Input` instead

### Other
* No easy way in Visual Basic 6 to pass args → worked around using `*` delimiter.

## Acknowledgements

* Merge Vertical icon icon by [Icons8](https://icons8.com)
    * Although I have their subscription, better safe than sorry

* Much of useful online documentation chained together!

## See also
*My other small Windows tools,*  

* [AMC_patcher-CLI](https://github.com/TAbdiukov/AMC_patcher-CLI) – (CLI) Patches app's SUBSYSTEM flag to modify app's behavior.
* **<ins>exe2wordsize</ins>** – (CLI) Detects Windows-compatible application bitness, without ever running it.
* [HWZ](https://github.com/TAbdiukov/HWZ) – (CLI) CLI engine to forge / synthesize handwriting.
* [SCAPTURE.EXE](https://github.com/TAbdiukov/SCAPTURE.EXE) – (GUI) Simple screen-capturing tool for embedded systems.
