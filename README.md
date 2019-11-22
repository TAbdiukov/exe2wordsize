# Win-exe2wordsize
Detects word size the Windows-compatible application was designed for.

Uses a bunch of WinAPI functions to determine
* Ironically, some functions are buggy! The tool takes that into consideration!

# Current state
* Prealpha-alpha. The tool works, but misses several key features

## TODO

* True CLI interface

* **OOP all other projects, via Type** (woah, is this how it works?)
	* ~~This project~~
	* CLI tool
	* The drawing tool
	* Another tool I may've missed?

* Proper compilation and usage instructions
	* How to pass args? `/P:C:/program.exe, /A:bcdef`?
* List noticed buggy WinAPI functions just in case
	* Blog about them?

* On hint, fix Stackoverflow information and comments? Since some WinAPI calls are ｇｌｉｔｃｈｙ　く俺カ [glitchy]

### 100% Done userstories

* ~~Makeuseof `EXE` struct or rid off of it~~ **There for legacy/justincase**
* ~~Json output~~ **Done! Rudimental for sure, but it works!** Btw, `vbCrLf` used for `newline`
* ~~Reading x64 apps correctly, as its not so easy due to [WinAPI bugs (lol)](https://stackoverflow.com/questions/25063530/why-do-i-get-nonsense-from-getmodulefilenameex-on-64-bit-windows-8)~~ DONE goodie!
	* ~~[Hint](https://superuser.com/questions/358434/how-to-check-if-a-binary-is-32-or-64-bit-on-windows)~~
	* ~~On hint, how much to read on the file (0x100, 0x1000, 0x10000, max?)~~ **0x1000 proven to be fine, but I pick 2 times more** (~ Euler number), that is, 0x2000 => 8192**
	* ~~On hint, what endianness to use?~~

