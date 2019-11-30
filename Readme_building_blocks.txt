#1

exe2wordsize v1.34

USAGE:
```
exe2wordsize <path_to_app>
exe2wordsize <path_to_app> * <args>
```

FOR EXAMPLE:
```
exe2wordsize C:/Projects/idk/Project1.exe
exe2wordsize "C:/Projects/idk/Project1.exe" * M=2 R=8192
```

MANUAL:
<path_to_app> - Path to your executable. "-tolerable

`*`- Delimiter required if you use args.
(Hint: Don't have to use asterick if no args required)

<args> - Extra arguments, space-delimited. Supported args below,
* M=(number) - Set analysis mode. Modes supported,
	* 0 - Automatic and flexible (Default)
	* 1 - Rely only on WinAPI. 64-bit input may be unreliable
	* 2 - Rely only on raw-reading. Only 32/64-bit detection, false-pos theoretical
ly possible
* R=(number) - In raw-reading mode (R=2), how many bytes to read at most for ana
lysis
  (Hint: Only applicable in MODE = 2. Unused in other modes)
OUTPUT:
In JSON format

#2
Merge Vertical icon icon by Icons8

<a target="_blank" href="/icons/set/merge-vertical--v1">Merge Vertical icon</a> icon by <a target="_blank" href="https://icons8.com">Icons8</a>
