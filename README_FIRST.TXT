
    Flying Windows V. 2.5.0
 

WHO / WHAT (with Update Info) / HOW / FAQ ?

 
    Last edit       :   10/20/2004

    Started         :   January 2003, VB5 version, Light_Templer, Germany (schwepps_bitterlemon@gmx.de)
    Conversion      :   To VB6 in September/October 2004 with adding of three new features.
    Copyrights      :   All copyrights by Light Templer. Don't sell this as a compiled DLL!
    Risk            :   Use at your own risk. I'm not responsible for anything ;)
 
 
    A freeware addin for MS Visual Basic 6
 
    Many thanks to Carlos J. Quintero (www.mztools.com) for his kick into my ass to do it by myself!
    And much more thanks for his great freeware addin  'MZ-Tools  :
    The BEST thing a developer could happen! Visit his site - you will be surprised!
 
 
    WHAT windows are flying?
 
    Flying Windows claims fully back the valuable work area of your IDE! No unneeded tool or property
    window is open when writing code or designing a form. Now the whole screen belongs to you to design
    your form or write some code! And when you need a tool/property window:  One quick move (no mouseclick)
    of your mouse is enough to get it back / take it away. 
    2nd hilite: Small property windows for controls like the known tooltips are assisting your form design!
    And some more nice gimmicks...

    Advertisement stop. ;) - To the facts:

    Implemented in Flying Windows so far

                            * Hotcorners for most used tool windows (toolbox, properties, project
                              explorer and immediate window)

                            * Empty the immediate window when ctrl key is pressed on opening by hotcorner.

                            * Tooltip windows for controls showing the name and the most important properties.

                            * Click on a property value in the tooltip window with ctrl key pressed puts this
                              value into the clipboard.

                            * Tooltip windows are moveable by caption bar. Leave them open with ctrl key pressed
                              when mouse leaves tooltip window. Close them with a simple mouseclick.

                            * Show mouse pointers absolut screen coordinates in VB IDEs title bar in pixel.

                            * With an open source code pane moving the mouse to the left border of the screen
                              opens the coresponding designer window (Shift-F7).

                            * The size of comboboxes of a code pane are increased to show much more values
                              without a need for boring scrolling.


    --- NEW WITH UPDATE 1 ---
 
                            * Moving the mouse to the right border of the screen increases the topmost code window
                              to full VB IDE client area size (but doesn't maxmize it!).
 
                            * Moving the mouse to the right border of the screen AND hold the <Ctrl> key pressed
                              closes the topmost window which has VB IDE as parent window. This can be a code,
                              a designer or any other window (not dialog!): e.g. Object Browser, Watch Window, ...

    --- NEW WITH UPDATE 2 ---

                            * With an open designer code window moving the mouse to the left border of the screen opens
                              the coresponding code window (same as pressing <F7> ).


    --- NEW WITH UPDATE 3 (V. 2.5.0) ---
 
                            * Now your CODE is flying, too ;) ! I have added a powerfull AutoComplete feature to Flying
                              Windows:  Write 2 or 3 letters, press F-12 (function key 12) and this small keyword will
                              replaced by a longer word or a couple of lines with code - whatever you want. A long list
                              with abbreviations missed in native VB is included. No more writing 'End With', 'Select Case'
                              'Private WithEvents' or a standard header block. Try 'ew' and press F-12, try 'sc' and press
                              F-12 or as an example for the full power type '*!' and press F-12 ...
                              The replacement is done by VBs 'SendKey' command so you can use all of its possiblities.


            _________________________________________________________________________

                  This is my first VB AddIn and my first conversion to VB6.
              (But not my first VB proggy ;) ). Plz be kind to it.
    ___________________________________________________________________________




    UPDATE 1 - 10/7/2004      CHANGES/FIXES

    Thx for all comments from comunity on PSC - here is what i 've changed to get better:

    1 - Mouse pointer will moved from hotcorner to over the window which just appears as before. That cannot
        be changed by design of this function. But now the mouse pointer is much closer to the hotcorner so
        you don't have to align the mouse. The feeling is much better this way. Thx, Alaeddin Hallak.
        For other (own) solutions: All of this calculation is done in 'SetMousePosOverWindow()'

    2 - Tests for docking mode of hotcorner windows are added. The Overflow Error (Raised, when setting a docked
        windows's Left position property) is catched and gives a long msgbox note what to do to avoid this.
        Sorry to all for this problem in first release. I'm not using docked windows and so i didn't get this
        problem on my system earlier.

    3 - Added a configurable value to FlyingWindows options dialog to set a max value for increased combo boxes.

    4 - Changed all msgboxes (error warnings) to msgbox "..." , vbExclamation + vbMsgBoxSetForeground, ... to
        get the error in forground. Thx, Tom Pydeski.

    5 - Added features for 'mouse at RIGHT screen side' event. Please look 40 lines above ;)



    UPDATE 2 - 10/15/2004      CHANGES/FIXES

    1 - Flying Windows Option Dialog is now shown in 'modal' mode.
 
    2 - Switch in option dialog for increasing the combo boxes doesn't work. Fixed.

    3 - A switch is added to option dialog to get the left screen border /right screen border functions off and on.

    4 - Added additional function for 'mouse at LEFT screen side' event. Please look 43 lines above ;)



    UPDATE 3 - CHANGES/FIXES

    No changes - just the new main feature 'AutoText' - plz read above for details.
 



________________________________________________________________________________

 SMALL TUTORIAL : How to use all of the features Flying Windows gives to VB IDE:
________________________________________________________________________________


* Load the project, build the DLL (compile it), put it into your addin directory and
  register it as usual (regsrv32 FlyingWindowsVB6.dll on command line).

* Restart Visual Basic and check in addin manager for FlyingWindows VB6 loaded.
  Your Add-Ins menu now shows the entry 'Flying Windows  with the butterfly logo.
  Here you can set some options (right now: Switch of features you don't want).

* When Flying Windows is running you have the coordinates of your mouse pointer on 
  screen in VB's caption bar displayed in pixel units. Just move arround the mouse
  to see what I mean.



IMPORTANT PREPARATION:

!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  To make your tool and property windows fly, they don't have to be in docked mode!
  In VB IDE's [option] dialog called from <Tools> menu you find the |Docking| tab.
  There you can switch them to 'not docked'. Thank you.
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


* Ensure that VB is maximized to get the largest available screen area for work,
  but please don't work with maximized code/designer windows.

* Now close all of your tools and property windows and move the mouse pointer to the
  left top corner: Bingo! The toolbox apears and the mouse pointer is set within the
  tool box  area. Just move the mouse out of the toolbox window and the toolbox
  disappears!

* Recall the toolbox by the hotcorner and carefully size the bottom and the left border
  of this window to your needs. DON'T leave the window area during this or the window
  will be gone ... ;)

* Repeat this step for the property window (right top), the immediate window (left bottom)
  and the project explorer window (right bottom).

Now you are ready for take off!  ;)



Whats more? A lot - you will see :)

* Some times, you need one of this windows open permanently. To keep it open just hold the
  <Ctrl> key pressed when you are leaving the area of this windows. So the window keeps
  open. To fast close it just move the mouse into the hotcorner and back away from the
  windows. The window is closed.

* A special trick is attached to the immediate window (thats my solution to a very old
  topic ;) ):  Holding the <Ctrl> key pressed when going to the hotcorner (left bottom)
  clears the immediate window. Nothing faster than this ;)


Now to the others

* Load a simple project with a form and open the code window of the form only. Move the
  mouse to the left border of the screen. Here we go: The designer window of this form will
  opened on top.

* The other way round: A designer window is on top, you move the mouse to the left screen
  edge and the coresponding code window appears.

* On the right side of the screen there are two functions,too:  Maximize the topmost code pane
  to VB IDEs client area size (but don't maximize it! Thats because of Flying Windows cannot
  work with maximized code windows. But you will see: You don't need them anymore ;) )

* The 2nd: Go to the right side of the screen with <Ctrl> key pressed and the topmost window
  (not dialog!) will be closed. (Thats much! faster then targeting this small |x| on top/right
  and click ... ;)

* Go back to the code window and open the comboboxes on top: You see larger is better ;)
  Much more values on one look without a need to scroll. This works for open/save dialogs
  combos, too ...


And now to the 2nd highlight of Flying Windows

* Small property windows for controls when designing a form/usercontrol: You get them like
  tooltips, just open a form designer window and stay with the mouse over a control for a
  second. The property window will be shown with the most important values of the control.

* You can keep this property windows even open and move them arround! Just hold the <Ctrl> pressed
  and move the mouse away from the control. The property window stays open and can be moved on
  its little yellow caption an top.

* Open as many of this windows as you like to have all properties you need for screen design
  in one look.

* To close a property window just left click into the blue value area of it.

* You need a copy of a value in clipboard? Just click on this value and hold the <Ctrl> key
  pressed! That works for the controls name on caption, too.



The last added main feature of Flying Windows  AUTOTEXT

* AutoComplete small keywords to full (longer) commands or even code snipets with many lines.

* NEEDED PREPARATION:
  To FlyingWindows V. 2.5.0 a text definitin file 'FW_AutoText.txt' is added. Put this file to
  the same directory where your 'FlyingWindowsVB6.Dll' is installed. Thats all.

* From FW options dialog you reach the definition list by pressing the button <Edit AutoText List>.
  Windows Notepad.Exe opens with the file loaded and offers you a long list with predefined shortcuts
  for you. Have a look to the (simple) definition system and change to your needs. Most of keywords
  are the first digits of the wanted commands, e.g.  ew  for 'End With' ,  in  for 'Is Nothing'  or
  prf  for 'Private Function ()'

* When FWs option dialog is closed with <ok> button, this text file is reloaded/parsed.

* In a code pane try it out - its simple:
	
	- Enter  puf
	- Press F-12

The keyword puf is replaced by  Public Function () . Try some more - its big fun and much faster than
anything else.

* Some notes good to know when editing the definition file

	- There are six special codewords added which are replaced with current values when called

		%AUTHOR%			Please edit the first code lines of Flying Windows - Connect.Dsr
		%DATE%				if you want to change from american date format to european or
		%TIME%				if you want your name put in here.

		%PROJECTDESCRIPTION%		This three values (the only one VB offers, version isn't accessable :( )
		%PROJECTFILENAME%		are taken from your current project. You can use them e.g. to automate
		%PROJECTNAME%			the creation of a modul header. An example is at the end of the file. 

	- The replacement is done by Visual Basics 'SendKey' command. Please read the VB online help to get more
	  information what can be done!  (e.g. the ~ is for a new line, write {(} for a single ( and +{Down} to move
	  the cursor one line down with shift key pressed (means: Select this line).


Let me know if you have any problems using this new feauture - it works fine to me right now in all tests ... ;)



_________________________________________________________________________
Have fun with it. I use it daily all the times since one and a half year
and I will not missing it.

                             Light Templer
_________________________________________________________________________


===========
=   FAQ   =
===========

Q:  I want to work with maximized code windows and get problems with Flying Windows.
A:  Sorry for that. VB's IDE works totaly different in docked and in not-docked mode.
    Right now I don't have good solutions to handle this. The IDE COM support is buggy,
    bad documented and needs lot of time to do even simple things ... 
    But I have implemented this 'increase code window to full IDE client area size' ting
    to overcome this. Hope that helps ;)

Q:  I have an idea for a good new feature.
A:  Please write it to the PSC submission thread or send me an email with a detailed
    description (and maybe code ;) ) to  schwepps_bitterlemon@gmx.de





One word to the background/coding of Flying Windows:

There is lot of hard stuff in this addin: Because of the missing of a lot of strongly needed
features in VB IDE's COM interface I had to dig very deep to get so simple things like a
reference to the control under the mouse even for controls without a window handle and
deeply nested in containers ...

Take a look at the code and forgive me for the style. Most of it is done in early 2003 and
not on current coding design level. (But it works and has comments ;) )


20th October 2004
Light Templer

