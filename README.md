# Ricoh-Monitor
Ricoh monitoring counters and more, with OIDs

After running the script it creates a printers_config.json. By default it has 2 IP's but you can change and put what ever, even add more. 
I compiled to do a .exe
Invoke-PS2EXE -InputFile "script_name.ps1" -OutputFile "name.exe" -IconFile "icon_example.ico" -Title "title of software" -Company "Company" -Product "Software name" -Description "Software Description"

After compiling and everything working:
  Put in Task Schedule every first day of the month
  Keeping it in C:/ with printers_config.json in the same folder, so if something breaks, you could troubleshoot

I made this .exe because it has smtp settings, in .exe its encripted.
