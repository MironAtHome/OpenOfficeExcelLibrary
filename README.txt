Open XML Microsoft Office Excel Reporting Library
implementation.

Open XML Microsoft GitHub repository:
https://github.com/OfficeDev/office-content.git

SpreadsheetLight Company URL:
http://spreadsheetlight.com/

Install Visual Studio 2015 community edition ( requires
to create an account and register with the Microsoft Developer
Extensions ). The account, as Visual Studio itself, are free 
of charge.

Download Microsoft Open XML SDK 2.5 with source code.
Download SpreadsheetLight from web site, including source code.
Run diff to see fixes. SpreadsheetLight URL: http://spreadsheetlight.com/

Fixes are, reuse of streams, to fix saving.
Removal of dependency from WindowsBase.dll, noted
by SpreadsheetLight. This dependency removal is prompted by changes in Open XML SDK 2.5.
By effect of unintended consequence, the removal of stale dependency
fixes error "Unable to determine the identity of domain", affecting
SpreadsheetLight distribution on Windows 7 ( verified, potentially
other Microsoft OS platforms affected. ) The two Microsoft Connect
bulletins identify the issue for reporting services:
https://connect.microsoft.com/SQLServer/feedback/details/779932/subscription-fails-on-large-output-files
https://connect.microsoft.com/SQLServer/feedback/details/764356/subscription-fails-due-to-error-system-io-isolatedstorage-isolatedstorageexception-unable-to-determine-the-identity-of-domain
Fixing this was quite nice, we can generate report of any size.

Code compiles without issues on Mono, however, invoking
assembly at the moment fails due to a single dependency on gdiplus 
native library on Mac. Suspect is either 64 bit ( I insist on using ) for
our reports, or library loader requires configuration and / or reference
to binary in order to properly invoke executable.

This code, in conjunction with git clone of Mono and Pash compiled
free of issues on fresh install of Yousemite OS ( Mac OS X flavor ).

And of course, build free of issues on Windows.

The goal is to fix port to X'es ( it could be Linux will work without issues, as opposed to Mac )
as evidence in On-Line discussions suggests Linux's code for UI is both
more flexible with regards to 32 vs. 64 bit choice.

A single remaining issues related to formatting affects numbers in decimal format.
It does not affect underlying number, just cosmetic issue. Ideally it would be nice
to display numbers in standard format - negative parenthesized, two digits to the right after
decimal point.