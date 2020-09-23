<div align="center">

## Add Scripting to your apps\!


</div>

### Description

Every wanted to allow users to control certain aspects of your program...to make you programmes shorter and easier...this shows you how to add scripting to your apps using the VBScript control. It mostly focuses on the purpose and usage of it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[IRBMe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/irbme.md)
**Level**          |Beginner
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/irbme-add-scripting-to-your-apps__1-37693/archive/master.zip)





### Source Code

```
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List href="./Scripting_files/filelist.xml">
<title>I’m can’t be bothered doing a long intro so I’ll cut it short</title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Author>CODING GENIUS</o:Author>
 <o:LastAuthor>CODING GENIUS</o:LastAuthor>
 <o:Revision>2</o:Revision>
 <o:TotalTime>37</o:TotalTime>
 <o:Created>2002-08-06T23:23:00Z</o:Created>
 <o:LastSaved>2002-08-07T00:01:00Z</o:LastSaved>
 <o:Pages>2</o:Pages>
 <o:Words>792</o:Words>
 <o:Characters>4520</o:Characters>
 <o:Company>Developement</o:Company>
 <o:Lines>37</o:Lines>
 <o:Paragraphs>9</o:Paragraphs>
 <o:CharactersWithSpaces>5550</o:CharactersWithSpaces>
 <o:Version>9.4402</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;
	mso-font-charset:2;
	mso-generic-font-family:auto;
	mso-font-pitch:variable;
	mso-font-signature:0 268435456 0 0 -2147483648 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
@page Section1
	{size:595.3pt 841.9pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
<!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="1035"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
 <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>
<body lang=EN-GB style='tab-interval:36.0pt'>
<div class=Section1>
<p class=MsoNormal>I’m can’t be bothered doing a long intro so I’ll cut it
short. This will hopefully show you how to add scripting to your programs, what
kind of programs you can add scripting to, and some possible uses of scripting.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The structured programming model is one which should all be
familiar with. Even in Event-Driven languages like VB, we still use A LOT of
structured programming.</p>
<p class=MsoNormal>Structured programming is the type that goes like this:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>INPUT ======&gt; PROCESS ======&gt; OUTPUT</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The input is what the user does. Whether they type a certain
thing, click a button, move the mouse…whatever</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The output is what gets displayed on screen to your users.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The process is the part we are interested in, this is where
most of your programming effort is focused. This is the part where you make
decisions, call functions, carry out certain tasks, and produce a result (not
always the case though) which will be carried on to the output part.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>But, imagine if your users could control certain parts of your
process!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Lets say you are writing a chatter bot…it has to be able to
be added to without re-compiling. You could use databases of phrases, or text
files or whatever. However by themselves, they are very limited in producing
the user’s desired results.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So you search around a bit and you think “Hey, seems that
what my program needs is its own scripting language” but who wants to spend
weeks trying to write a good scripting engine? I’ve tried and believe me, there
is no more extensive usage of the “split” “join” “mid” “left” “right” “instr”
“instrrev” and “len” functions but to name a few. Trust me here when I say even
the best written code is NOT pretty! And it’s a nightmare to debug. “Well
that’s that option out of the window. Back to text files?” I hear you ask. And
I say “NEI my hard-coding amigos! There is a better way!”.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>We aren’t going to create our own scripting language; we are
going to use “VBScript” or “Jscript”. If you have the MSScript control that is included
with Visual Basic Enterprise or Professional Edition (I think) then you are in
luck! The code is simple.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>I’m going to use a real life example just to demonstrate the
use of this brilliant control.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now I go on IRC a lot and I like to amuse my fellow VB
coders in the channels or even help then or provide useful information with my
bots. For those of you who don’t know what they are, they are programs that
automatically connect to a chat room and run without any user intervention. My
bot replies to commands typed by users.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Not a single one of my commands are hard-coded into VB.
Anybody who has a compiled version can write their own commands and customize
them as they please, using a choice of languages! Here’s a simple one…a file
called “Command:Say.VBS”</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Function Say (Data)</p>
<p class=MsoNormal><span style='mso-tab-count:1'>            </span>Say =
Trim(Data)</p>
<p class=MsoNormal>End function</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Obviously my say command is a lot more complex that that, it
trims off spaces, checks for names, can add time and date and…. it does a lot
ok but moving on. This takes us back to the very beginning; you remember the
Input – Process – Output part? Well, we just wrote a function that takes input,
jiggles it about and does what it wants with it, then produces an output.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>We take commands from the user, send them to this “say”
script, and return the output of the function back to the user. I have many
other larger and more complex scripts for my bot, each dealing with its own
error handling, string handling, checking and everything. Some are over 100
lines long.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>That is just one extremely useful…eh…use. But there are
hundreds more; you just got to use your imagination. They are extremely useful
in any AI programs I can tell you that though!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Now, how to actually implement this? Well first you add the
Script control by right clicking somewhere on your toolbox and clicking
Components. Now find it and add it and draw a copy on your form, mine is called
“Script”.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>First thing we have to do is add our code….</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Script.AddCode “Function Say(Data)” &amp; vbcrlf &amp; “Say
= Trim(Data)”<span style="mso-spacerun: yes">  </span>vbcrlf &amp; “End
Function”</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Then we can keep adding functions by using the “AddCode”
method. Note: the previous code is not overwritten.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>We could easily add the code from files (They don’t have to
be named .vbs or .js or whatever, they can be .dat, .txt .script, whatever you
want)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>So how do we input/output data from the function…like so:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Output = Script.Run(ProcedureName,Param1,Param2, Param3,…..)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>You can add as many or as little parameters as your function
requires after the “FunctionName” parameter. It is a ParamArray. Example:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>IRC.SendChat Script.Run(“Say”,ChatMessage)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>And that works nicely for my IRC bots ;)</p>
<p class=MsoNormal>Also a good function to know is the EVAL method. I use it to
impress my buds on IRC. Just imagine:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>“Hey bot, whats sin(1 + 4) * cos(-sin(cos(100)) + 100 /
1000) – 10^2”</p>
<p class=MsoNormal>“That’s easy, the answer is…&lt;ANSWER HERE&gt;” </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Only it gives the answer of course :-P</p>
<p class=MsoNormal>It looks quite impressive. But it can be used in programs
like calculators, or graph plotters (I have an example of one of those here on
PSC)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>It takes a string expression and attempts to solve it
basically. Just experiment to see what it can and can’t do!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Ok, I hope this is a clear tutorial on how to add scripting
to your programs. If you read it at all, please feel free to leave a comment <span
style='font-family:Wingdings;mso-ascii-font-family:"Times New Roman";
mso-hansi-font-family:"Times New Roman";mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>J</span></span>
I like to feel appreciated. Hehe</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Later all. And for those who I talk to on IRC (V*r*sF*e*,
D*A*K and all the others, you know who you are, but others don’t. Nice privacy
thing there <span style='font-family:Wingdings;mso-ascii-font-family:"Times New Roman";
mso-hansi-font-family:"Times New Roman";mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>J</span></span>)
I’ll have my new APIBot up soon! </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
</body>
</html>
```

