# !["|-.-|"](icons8-merge-vertical-64.png) exe2wordsize
Detects word size the Windows-compatible application was designed for **without running them at any point**

Uses a bunch of WinAPI functions to determine
* Ironically, some functions are buggy, read below. The tool takes that into consideration!

## Assumptions taken
* **Never run the executable at any point**
* Host OS can run the application tested (for example, Windows XP, but not Vista, and a DOS app)
* Minimum byte-size applicable is always returned (for example, 32-bit apps on 64-bit host OS are analysed to use 32-bit word size, even if it is emulated)
* WinAPI may be glitchy, hence their output is doubted
* WinAPI should be present in all known OSs (-> Debugging versatility)

## Current state
* Beta


## Usage
```
exe2wordsize <path_to_app>
exe2wordsize <path_to_app> * <args>
```

### Example
```
exe2wordsize C:/Projects/idk/Project1.exe
exe2wordsize "C:/Projects/idk/Project1.exe" * M=2 R=8192
```

### Manual
`<path_to_app>` - Path to your executable. `"`-tolerable

`*`- Delimiter required if you use args.
(*Hint*: Don't have to use asterick if no args required)

----

`<args>` - Extra arguments, space-delimited. Supported args below,
* M=(number) - Set analysis mode. Modes supported,
	* 0 - Automatic and flexible (Default)
	* 1 - Rely only on WinAPI. 64-bit input may be unreliable
	* 2 - Rely only on raw-reading. Only 32/64-bit detection, false-positive theoretically possible

* R=(number) - In raw-reading mode (`M=2`), how many bytes to read at most for analysis
(*Hint*: Only applicable in MODE = 2. Unused in other modes)

## How to compile
1. *[Recommended for compatibility]* Get a Windows XP VM
2. Get a **Microsoft Visual Basic 6.0** 

***Tip:** I unofficially recommend a portable version sticking around on BT, as you won't have to mess around with the installation and registry. Plus, it's only a few megabytes. Check out **Portable Microsoft Visual Basic 6.0 SP6***

3. Fire up **Microsoft Visual Basic 6.0**, open up the project.
4. Go to File -> Make *.exe -> Save
5. Patch the app for CLI use:
* You can use my [AMC patcher](https://github.com/TAbdiukov/AMC_patcher-CLI). For example,
	amc C:\Projects\HWZ\hwz.exe 3

* Or you can use the original Nirsoft's [Application Mode Changer](http://www.nirsoft.net/vb/console.zip) ([info](http://www.nirsoft.net/vb/console.html)), unpack the archive and then run the **appmodechange.exe**

6. Done!

## Found WinAPI bugs
### General
* `SHGetFileInfo` IS buggy on 64-bit executables. Hence `MODE=2` had to be implemented
* `GetBinaryType` BinaryType IDs are poorly documented

### File-reading
* In some line, `Open X For Binary Access Read As file_descriptor Len = length_var`; `length_var` argument appear to always be ignored, despite available documentation (is this becaise of `Binary` file reading mode?) In this app, the String length isn't enforced, since doing so would mean effectively undermining detection process through some extra time-consuming operation. If need be, the enforcement takes 5 mins to implement.
* For binary reading, `InputB` returns some patterned gibberish, despite documentation online. Use `Input` instead

### Other
* No easy way to pass args -> had to do that the tricky way (via `*`)

## Acknowledgements

* Merge Vertical icon icon by [Icons8](https://icons8.com)
    * Although I have their subscription, better safe than sorry

* A bunch of useful info online chained together!
