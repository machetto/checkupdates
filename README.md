## Windows Updates Checker

The script scans remote computers (via COM's call to 'Microsoft.Update.Session') in specified Active Directory OUs for Windows Updates and reports whether they are compliant.

It scans the computers in parallel using PoweShell jobs (ten jobs by default).

Computers can be excluded from a scan (use 'exception_list.txt') and/or included in a scan (use 'inclusion_list.txt').

Compliance reports will saved into a '.\reports' subfolder and also optionaly emailed if receipents are specified.
