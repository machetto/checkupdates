## Windows Updates Checker

The scipt scans remote computers (via COM's call to 'Microsoft.Update.Session') in specified OUs for Windows Updates and reports whether they are compliant.

Computers can be excluded from a scan (use 'exception_list.txt') and/or included in a scan (use 'inclusion_list.txt').

Reports will saved into a '.\reports' subfolder and also optionaly emailed if receipents are specified.
