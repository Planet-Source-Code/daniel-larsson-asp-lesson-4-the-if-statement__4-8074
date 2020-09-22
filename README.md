<div align="center">

## ASP Lesson \#4 \- The IF statement


</div>

### Description

This article will try to explain to you how to use the IF-statement and why you use it. It will teach you the basics of ASP IF statements, how you write them, what you use them for and why we use them.<p>Also read why this is such a late publication.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Larsson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-larsson.md)
**Level**          |Beginner
**User Rating**    |4.7 (33 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__4-7.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-larsson-asp-lesson-4-the-if-statement__4-8074/archive/master.zip)





### Source Code

<font size="5">Some information from me:</font><br>
First of all I like to say thank you to everyone that have been sending me emails about how you liked my previous tutorials. Again, a very big thanks to everyone!<br>Now why is this article being written so looong time before the previous ones? Simple, I have been *very* busy with my IRL and have been unable to finish my guide.<br><br>
I plan to write an article about once a week from now on so hopefully you will learn something from me like you have been doing to the previous three lessons.<br>
Please don't complain about my english, im writing this in the best english I can but since im not from USA or England I probably do a lot of grammar mistakes. I sure hope you can live with these, else ... send me an updated language article.
<p><hr>
<font size="4">And now lets begin...</font>
<p>
A great advantage of ASP with VBScript is to be able to do different things depended on the relation to data submited by the users. I shall now try to show you how you can make your script do different things depending on information submited by the users.
Lesson number 4 will mainly be about what ASP-Programmers call: If...Then...Else. This is called a <b>statement</b>.
<br>
<p>
<font size="4">On one condition</font>
<p>As I just explained to you I will now talk about what If...Then...Else means within ASP. To start with we have to explain each of these words with the following example: If it is this way you do this, and if not you do this. Pretty simple, ain't it? It is simple!
<p>
If Number > 50 Then
<p>
This is an example of an IF-Statement. This line of code starts with If. This if is then followed by a variable named Number and then followed by a greater then (>) character and then followed by the number 50. At the end of this code is the Then word.<br>
To explain this simple example of code i will now explain in words what it really do. Translated into english this code does the following:<p>
If the variable Number is bigger then 50 you do this. Between the words IF and THEN you can actually put any condition statement that you would like to check for.<br>
What i mean with this is that you could also write: If Number >= 50 Then...or If number <= 50then...or even If Number = 50 Then ...as long as this statement is being set to a value earlier in your script.<p>
I'm trying to teach you not to use a variable that have not been declared. Take this example: If you declare the variable Number like we did in ASP Lesson #2, but without assigning it a value?
What would this script do then? It would check if the value in the variable Number would be greater then 50 and everytime it would genereate the answer: No, which is obvious because it doesnt have a value.<p>
Dim Number<br>
Number = 500 <br>
<br>
If Number > 50 Then
<br><br>
As you can see in this example code the variable Number is greater then 50. Because of this your script will then perform the code that is written after the "Then" word.
<br><br>
If Number > 50 Then
<br><br>
The number was bigger then fifty.
<br><br>
<font size="4">What will happen if this number is not bigger then fifty?</font>
<br><br>
Lets say that the variable would only have one value that would be equal to 40, what would happen then? Then we would also tell the script what will happen. The browser can't write that the number was bigger then 50 because the value only was 40. To achieve this extra bit of code we must add: Else.
<br><br>
&lt;%<br>
Dim Number<br>
Number = 40 <br>
<br>
If Number > 50 Then <br>
%> <br>
<br>
The number was bigger then fifty.
<br>
&lt;%Else%>
<br>
<br>The number was not bigger then fifty.
<br>
&lt;%End If%>
<br><br>I'm now using delimiters on many lines in the script. I have to do this because the two texts("The number was bigger then fifty ... and ... The number was not bigger then fifty.")that will be printed to the user
are part of simple HTML code that will not be needed to be executed on the server. We then tell the server to stop executing ASP-Code and start working with HTML-Code again. I do this because I dont want the two texts to
interfere with the rest of my script (a.k.a ASP code). Everytime I have to add more code I have to write another delimiter so the server will know when I have used some more ASP-Code.
<br>
<br>Now let us review this, yet simple, code deeper. We started with defining a variable and gave it a value of 40. We then make a IF statament.If the number was bigger then
fifty then do this. If this value of the variable number would have been greater then 50 the text "This number was bigger then fifty." would be generated in the browser.
But in this case this statement wasn't real. This is also known as as statment not being TRUE (which I will explain more later on).
<br>
<br>Because thit statement wasn't TRUE the ASP server looked for the "Else" line in your ASP-Code to see what will happen when this number was not bigger then 50.
When the server did find the "Else" line in your code the server generated all the code that was written after the "Else" part. In our case this was the text that told the user
about that the number was not bigger then fifty. Finally the If...Then...Else-statement with "End If".
<br>
<br>This was a simple example. Of course this statement is used for a lot mor sophisticated thinkgs then checking for if a value is bigger then another value. You often use thise
statement if you want to check if the user have enterd a valid e-mail adress, if he entered a number, if he entered anything at all and things liike that. When you code ASP you
will find out that this statement is used very often and its un-replaceable. It's like a car without tires. IF a car does not have tired, you cand drive it. If you can't check
what information your script is currently working with, you can't make a dynamic website.
<br>
<br>To give you all believers out there a very sophisticated use of the If-statement i will let you know that you can use it to generate different code depending on what webbrowser
the user uses. Doing this will let you choose how your page can look like in internet explorer and in netscape in example.
<br>
<br>Before I put an end to this lesson I will like to add one word to this statement: ElseIf. What if neither of the conditions you are checking for are true or false? The variable
Number could actually be exactly 50. What would happen then? Then nothing would have happened since you had not told the script to do something when it was exactly 50. We now have three alternatives:
Greater then, Less then or Equal to.
<br>
<br>
&lt;%<br>
Dim Number<br>
Number = 50<br>
<br>
<br>
If Tal > 50 Then<br>
%> <br>
<br>
The number was bigger then 50.<br>
<br>
&lt;% ElseIf Number = 50 Then %> <br>
<br>
The number was equal to 50.<br>
<br>
&lt;% Else %> <br>
<br>
The number was less then 50.<br>
<br>
&lt;% End If %>
<br>In this case the Number was equal 50 and because of that the user will se that information in his browser. Please note that you may use how many ElseIf words in your code as you want to within your if-statement.
<br>
<br>I would like you to try and get this script to work and see what will happen. This is the key to understanding any programming langage, and especially ASP. Try to experiment with this code and see if you can do something
fun with it. Try exhanging the value of the variable Number for a new value. You should also tryt o add a few more ElseIf satements in your sript. I will supply a script beneath this that uses more then one ElseIf statement so
you will find out how you can use it in a more sophisticated way.
<br>
<br>
<br>
&lt;%<br>
Dim Number<br>
Number = 1000 <br>
<br>
If Number > 1000 Then %> <br>
<br>
The number was bigger then 1000.<br>
<br>
&lt;% ElseIf Number < 1000 Then %> <br>
<br>
The number was less then 1000.<br>
<br>
&lt;% ElseIf Number = 1000 Then %> <br>
<br>
The number is equal to 1000.<br>
<br>
&lt;% Else %> <br>
<br>
The number was something else.
<br>
&lt;%End If%> <br>
<br>As you may notice we've now made a sort of script that covers every value. Even if the numbers would not fit into any of the first options, which would be quite amazing since it's declared do 1000 exactly, the script will generate a correct answer.
<p>I hope you have now learned what this if...then...else sentence is. If you want to know more about ASP im currently working on an article about the DO statement which will be in ASP lesson #5. I hope to get some comments about this one since its
pretty hard to cover the whole if-statement as easy as it can be.

