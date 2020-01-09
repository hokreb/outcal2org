# Outlook Calendar 2 org mode file

I like the way org mode can be used to organize all things which I have to do. Unfortunately 
I have to use three Calenders: google (private), outlook (business) and of course org mode by emacs. 

To let the outlook dates show up in the agenda of org mode, I created this small program, which displays
the calendar entries from the actual running outlook instance in org mode syntax. 

If one pipe this into a file, it can be incorporated into emacs org mode agenda easely.

## Setup for Emacs

- build executable
- place it in emacs/bin folder
- add (async-shell-command "\\path\\to\\emacs\\bin\\outcal2org.exe -30 90 > \\path\\to\\org\\files\\outlook.org") to .emacs

After starting emacs, an updated outlook.org file will be produced, which show up in org agenda

I am sure, there exists a way to set a hook, so that it will be call before agenda is viewed, so that 
is always actual without restarting emacs
