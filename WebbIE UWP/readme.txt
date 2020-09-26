UWP Windows 10 version of WebbIE.
19 Oct 2016

It's a start, gets a web page and displays the text content.

Meanwhile, InvokeAsync.htm is a start of JavaScript function that parses the DOM to get the text content we need for the display. 

I guess this would generate some parsed content that we would build a second object model out of, and would include enough information to call back into the DOM of the web page to (1) click buttons (2) fill in text boxes (3) follow links.

Interesting!