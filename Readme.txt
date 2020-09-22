This program places a password on any executable file (must be .exe), whether you have made it or not. The password must be entered every time the program starts.

To make a protected exe, you must compile both projects. When that is done, put the executable you want to protect in the same directory as the other two programs. Do not run LockExe.exe, as it is useless on its own. Run LockExeMake.exe and keep the first box the same. In the second box, enter the name of the file you want to protect, e.g. MyFile.exe. The third field is the password that the program will ask for when the protected executable starts up.

When you finish entering the data, click Combine and a new file will be created called Locked012345.exe. You can rename and move this to whatever and wherever you want, as long as its extension is .exe or .scr. If you run the new program, it will ask you for a password. Enter it correctly and the original program will run.

Please note that the security provided by this application is very low. It will stop anybody that is not prepared to browse the file in a hex viewer as the data is not encrypted, but not crackers, which is a feature I might decide to add.

Thankyou for downloading my code. I hope that it is useful to you.