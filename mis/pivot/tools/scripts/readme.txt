Die Datei sendkey.vbs öffnet die in Parameter1 angegebene Anwendung 
und sendet an sie die in der als Parameter2 angebenen Textdatei enthaltenen
Tastaturkommandos.

Aufruf:

sendkey.vbs <Parameter1> <Parameter2>


akzeptierte Tatstaturkommandos (durch CRLF getrennt):

Key		Code
---------------------------------
BACKSPACE	{BACKSPACE}, {BS}, or {BKSP}
BREAK		{BREAK}
CAPS LOCK	{CAPSLOCK}
DEL or DELETE	{DELETE} or {DEL}
DOWN ARROW	{DOWN}
END		{END}
ENTER		{ENTER}or ~
ESC		{ESC}
HELP		{HELP}
HOME		{HOME}
INS or INSERT	{INSERT} or {INS}
LEFT ARROW	{LEFT}
NUM LOCK	{NUMLOCK}
PAGE DOWN	{PGDN}
PAGE UP		{PGUP}
PRINT SCREEN	{PRTSC}
RIGHT ARROW	{RIGHT}
SCROLL LOCK	{SCROLLLOCK}
TAB		{TAB}
UP ARROW	{UP}
F1		{F1}
F2		{F2}
F3		{F3}
F4		{F12}
F13		{F13}
F14		{F14}
F15		{F15}
F16		{F16}

To specify keys combined with any combination of the SHIFT, CTRL, and ALT keys, precede the key code with one or more of the following codes:

Key		Code
---------------------------------
SHIFT		+
CTRL 		^
ALT		%

To specify that any combination of SHIFT, CTRL, and ALT should be held down while several other keys are pressed, enclose the code for those keys in parentheses. For example, to specify to hold down SHIFT while E and C are pressed, use "+(EC)". To specify to hold down SHIFT while E is pressed, followed by C without SHIFT, use "+EC".
To specify repeating keys, use the form {key number}. You must put a space between key and number. For example, {LEFT 42} means press the LEFT ARROW key 42 times; {h 10} means press H 10 times.

Note   You can't use SendKeys to send keystrokes to an application that is not designed to run in Microsoft Windows. Sendkeys also can't send the PRINT SCREEN key {PRTSC} to any application

Zusätzlich wird folgendes Kommando akzeptiert:

{SLEEP x}

x steht für eine Zeitangabe in ms
Der Befehl läßt sendkey die angegebene Zeit warten, bis das nächste 
Tastenkommando abgeschickt wird.
