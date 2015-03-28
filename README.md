learning-VBScript
=================

This repository includes code that I wrote for *myself* towards learning VBScript(!). 

I *had* to learn VBScript as part of a new project that I've been working on recently and decided to take the fun approach for once and consequently came up with a simple click-based game that runs ***ONLY on Windows*** and another library of functions that can be plugged into QTP/UFT scripts (or not).

Whilst the game is (rudimentarily) complete, the library is a work-in-progress and I intend to extend the library with more functions when I have the time and patience to do so.

***Addendum***: I've also included a simple script that parses CricInfo's live RSS feed and gives me the current score for a match with two teams that I happen to be interested in. I intend to keep adding similar useless utilities and hacks whenever I can.

Details, Installation and Usage
======================
* ***Idle*** - toggles NUMLOCK on and off to prevent your Windows machine from locking up when left idle. In its purest essence, it's a poor man's Caffeine. Change NUMLOCK to SCROLLLOCK and things should still work just fine!

* ***The Fastest Click*** - is a game that tracks how fast you can click on two MsgBox buttons in succession. It turned out to be incredibly successful during a short trialing period amongst my colleagues and subsequently led to people cheating their way to high scores. Please read the source for more details!

* ***Live Scores*** was created to parse the 'Live Scores' RSS feed from CricInfo and search for a cricket match presently running that you happen to be interested in. With caps on bandwidth and me not wanting to open a content rich/heavy page, this serves me really well and live scores are a click away!

* ***The Function Library*** contains a bunch of utilitarian functions that I wrote to learn the basics of VBScript. They make for good learning exercises and also help when you're writing automated scripts with QTP/UFT.

As with any VBS file, all you have to do is download and execute! The function library can be exported as a QFL file for usage inside of QTP/UFT.
