## WebbIE 4

A .Net WinForms web browser based on the WebView control, which of course is Internet Explorer. 

Renders web pages as pure text so they work with screenreaders for blind people. Lots of features useful to screenreader users:
* Crop pages - remove the rubbish like advert links and navigation bars so they can navigate the page easily.
* Forms, video, audio all supported with specialised user interfaces.
* Lots of shortcut keys for keyboard-only users.
* Change font size and appearance for low-vision. 

Developed since 2000, so it's been around a while!

Project page:
https://www.webbie.org.uk/webbrowser

Released under the GNU Public Licence V3.

By Alasdair King, https://www.alasdairking.me.uk

# Building WebbIE 4

## WEBBROWSER COMPATIBILITY MODE
I make no attempt to set the FEATURE_BROWSER_EMULATION value so it will render pages in IE7 mode. This is what WebbIE has been doing, and will work on older machines (e.g. XP).

## APPLET, OBJECT
I've abandoned trying to support these. At best I can switch to the IE view and click them, but this doesn't work well - mouse pointer position is hard - and I've abandoned the web view. Finally, if you do render the contents, they usually just say "You need Flash!" or "Your browser does not work!" which is usually a lie and always unhelpful.

## WEB VIEW
I've abandoned the ability to switch to the Internet Explorer view. This is because it was the source of many difficulties with screenreader focus, and because this is serving a slightly different set of users - people with a fair amount of vision - as opposed to screenreader users. So in the spirit of "simple single-purpose apps" I've split this functionality out into a new program, Viewie. You can still open the current page in Internet Explorer if you want: then your screenreader will activate all its HTML web browser hooks and you will be better off than using the Web View in WebbIE. This will also help people who accidentally trigger web view and get stuck there.

## FONTS AND SCALING
I've used Segoe UI (the standard Windows font from Vista onwards) and 16pt sizes for readability, which seems a good size for people with some sight. Except the status bar, which is 14pt because it doesn't size!

## COLOUR AND INVERT
(OR Color!) I've removed the ability to invert the colour of the text area. I'm going to respect the Windows setting for high-visibility, so if you set this then it'll be yellow-on-black or whatever, consistent with other Windows programs. 

## FORWARD
In keeping with modern fashion and general usage, and to keep the UI simple, I've removed the Forward button. You can still go forwards with the Navigate menu.

## PRINTING
You can only now print the text area, not the web page layout. There is clearly a use case for needing to print the page, however, so I'll have to add that back in at a future date.

## SCREENREADER COMPATIBILITY
WebbIE won't do anything if you are holding down CAPSLOCK since this is a common screenreader control key.