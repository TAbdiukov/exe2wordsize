# Win-exe2wordsize
Detects word size the Windows-compatible application was designed for.

Uses a bunch of WinAPI functions to determine

# Current state
* Prealpha. It compiles and the tool barely works, but much of improvements required
v0.1

## TODO

* **OOP all other projects, via Type** (woah, is this how it works?)
* Makeuseof `EXE` struct or rid off of it
* True CLI interface
* Reading x64 apps correctly, as its not so easy due to [WinAPI bugs (lol)](https://stackoverflow.com/questions/25063530/why-do-i-get-nonsense-from-getmodulefilenameex-on-64-bit-windows-8)
	* [Hint](https://superuser.com/questions/358434/how-to-check-if-a-binary-is-32-or-64-bit-on-windows)
	* On hint, how much to read on the file (0x100, 0x1000, 0x10000, max?)
	* On hint, what endianness to use?
* Proper compilation and usage instructions
