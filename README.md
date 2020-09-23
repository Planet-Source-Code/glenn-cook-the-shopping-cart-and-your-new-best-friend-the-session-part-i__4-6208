<div align="center">

## The Shopping Cart and your new best friend\.\.\.The Session: Part I


</div>

### Description

This shopping cart programming excersise is designed to help beginning programmers with some common programming concepts as well as provide more experienced programmers information on ASP's powerful programming environment and how to set up global arrays for web applications.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Glenn Cook](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/glenn-cook.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__4-7.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/glenn-cook-the-shopping-cart-and-your-new-best-friend-the-session-part-i__4-6208/archive/master.zip)





### Source Code

<p><font face="Verdana"><strong>This shopping cart programming excersise</strong>
<small>is designed to help beginning programmers with some common programming
concepts as well as provide more experienced programmers information on ASP's
powerful programming environment and how to set up global arrays for web
applications.&nbsp; The heart of a shopping cart is:</small></font>
<ul>
 <li><font face="Verdana"><small>The Session Object</small></font> <img align="right" src="http://www.aspalliance.com/glenncook/images/rude.jpg" width="100" height="128">
 <li><font face="Verdana"><small>Global Variables and Constants</small></font>
 <li><font face="Verdana"><small>Two-dimensional Arrays</small></font>
 <li><font face="Verdana"><small>For.....next.</small></font></li>
</ul>
<p><font face="Verdana"><small>If you are a real newbie, the first thing every
new ASP developer should realize is that ASP compiles your code and treats the
ASP pages within your site like a running program. &nbsp; It's not just feeding
requested HTML files to your browser, it's actually compiling code at the server
before it sends your browser anything.&nbsp; In this way it reacts to user input
and feeds the user pages based on that input- making it dynamic!</small></font></p>
<p><font face="Verdana"><strong><big>The Almighty Session&nbsp;</big></strong></font></p>
<p><font face="Verdana"><small>&nbsp;&nbsp;&nbsp; ASP's dynamic state is created
and maintained by ASP's&nbsp; &quot;Session&quot; object.&nbsp; It's always
watching the user (Microsoft likes this), and in my opinion, is what make ASP so
incredible.&nbsp; It takes the &quot;stateless&quot; HTTP protocol and through
the use of a teeny cookie, makes a session state so the developer can create
global variables.<br>
&nbsp;&nbsp;&nbsp;&nbsp;When a user requests an ASP page the server writes a
cookie to the user's machine and assignes them a unique session ID. To create
global variables all you have to do is to create some Session() variables that
are bound to that session ID.&nbsp; If you're getting confused at this point let
me review the process in another way.&nbsp; Here's what happened when you
requested this page:</small></font></p>
<p><font face="Verdana"><small>&nbsp;&nbsp;&nbsp; You type <a href="http://www.activeserverpages.com/glenncook/"><font color="mediumblue" face>&quot;www.activeserverpages.com/glenncook/index.asp&quot;</font></a>and
Data Return's server says, &quot;Hey, they want an ASP page and it looks like
the person doesn't have a cookie with a session id so I'm going to send a new
unique ID. &nbsp; Before I send this page let me see if the global.asa file
wants me to do anything special, like connect to a database, create some global
variables whatever. &nbsp; Oh, Glenn wants us to create a couple of </small><strong>global
variables</strong><small> which will be unique to that user's Session.SessionID.&nbsp;
The ID I'm giving this user session is:&quot; 729929829 &quot; (By the way
that's really the user ID Data Return's server gave you. If you don't believe me
close your browser and come back here, you'll see a whole new session ID.&nbsp;
When you close your browser your session cookie also expires.&nbsp; Cool, huh?!)</small></font></p>
<p><font face="Verdana"><small>The Session object starts and ends in the
&quot;Global.asa&quot; file. What is the global.asa file?&nbsp; Very simply,
it's the file that allows the developer to create global variables using ASP's
Session object.&nbsp; The global.asa file is kind of like your config.sys and
autoexec.bat files when you load DOS.&nbsp; The first thing DOS does is look to
these files to free up memory space, start drivers, and execute programs.&nbsp;
ASP looks to the global.asa file in the same way for configuration information
for that user session.&nbsp; Ok, let's see what the global.asa's guts looks
like.</small></font></p>
<table border="1" cellPadding="4" width="118%">
 <tbody>
  <tr>
   <td vAlign="top"><font color="#008000" face="Courier New"><small><small>&lt;!--#INCLUDE
    file=&quot;ShoppingCartContants.inc&quot;--&gt;</small><br>
    <small>&lt;SCRIPT LANGUAGE=VBScript RUNAT=Server&gt;<br>
    <br>
    Sub Application_OnStart
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
    End Sub<br>
    <br>
    Sub Application_OnEnd<br>
    End Sub<br>
    <br>
    Sub Session_OnStart<br>
    <br>
    ReDim The Cart(Attributes,TotalItems)<br>
    Session(&quot;Cart&quot;) =TheCart<br>
    Session(&quot;MaximumItems&quot;) = TotalItems</small><br>
    <small>Session(&quot;ItemCount&quot;) = 0<br>
    End Sub<br>
    <br>
    Sub Session_OnEnd<br>
    End Sub<br>
    &lt;/SCRIPT&gt;</small></small></font></td>
   <td vAlign="top" width="60%"><small><font color="#ff0000" face="Arial">Right
    here we're using the global.asa's Session_OnStart event to create some
    of those global variables I was talking about. &nbsp; But you'll notice
    I 'included' some file called cart.inc. INC files are great because ASP
    will stick the code from that file right into your ASP page for you.&nbsp;&nbsp;
    All I'm doing is creating some constants(like VB) that I assign to the
    Session's global variables that I made. &nbsp; I could just as easily
    put cart.inc's code into the Session_OnStart event and it would work
    just as well to create these variables but I want to be able to call the
    array I made in this file in other pages.&nbsp; Oh yeah, the array?!&nbsp;
    Don't worry, I'll get to that in a second and IT IS extremely important.
    I want you to see that the array here is made with this code:</font></small>
    <p><font color="#008000" face="Arial"><small>ReDim
    TheCart(Attributes,TotalItems)<small><br>
    </small></small></font></p>
    <p><small><font color="#008000" face="Arial"><small>Session(&quot;Cart&quot;)
    =The Cart<br>
    Session(&quot;MaximumItems&quot;) = TotalItems</small><br>
    <small>Session(&quot;ItemCount&quot;) = 0<br>
    End Sub</small></font></small></p>
    <p><small><font color="#ff0000" face="Arial">The elements of the array
    are defined by the constants in the &quot;cart.inc&quot; file below.&nbsp;
    Understanding the array here is the key to understanding how to make a
    shopping cart.</font></small></p>
    <p><small><font color="#ff0000" face="Arial">*You don't see anything
    about the SessionID in the code because ASP does that for you
    automatically!</font></small></p>
   </td>
  </tr>
  <tr>
   <td vAlign="top"><font color="#000000" face="Courier New"><strong>&quot;ShoppingCartContants.inc&quot;</strong></font>
    <p><font color="#008000" face="Courier New"><small><small>&lt;SCRIPT
    LANGUAGE=VBScript RUNAT=Server&gt;<br>
    <br>
    const TotalItems = 5<br>
    const Attributes = 5<br>
    <br>
    Const cartProductID = 1<br>
    Const cartProductName = 2<br>
    Const cartDescription&nbsp;&nbsp;&nbsp;&nbsp; = 3<br>
    Const cartItemPrice = 4<br>
    Const cartItemQuantity = 5<br>
    &lt;/SCRIPT&gt;</small></small></font></p>
   </td>
   <td vAlign="top" width="60%"><font face="Arial"><strong>The Array!</strong></font>
    <p><font color="#ff0000" face="Arial"><small>OK, this is your
    two-dimensional array inc file.&nbsp; To understand 2X arrays think
    about Microsoft Excel - it's exactly like a table!&nbsp; This is the
    array I just created- right now all the cells are empty because we
    haven't put anything in there yet:<br>
    </small></font></p>
    <div align="left">
     <table border="1" cellPadding="0" height="180" width="100%">
      <tbody>
       <tr>
        <td height="51" width="16%"><font face="Arial" size="1">Attributes&gt;&gt;&gt;</font></td>
        <td height="51" width="16%"><font face="Arial" size="1">cartProductID</font></td>
        <td height="51" width="17%"><font face="Arial" size="1">cartProductName</font></td>
        <td height="51" width="17%"><font face="Arial" size="1">cartDescription</font></td>
        <td height="51" width="17%"><font face="Arial" size="1">cartItemPrice</font></td>
        <td height="51" width="17%"><font face="Arial" size="1">CartItemQuantity</font></td>
       </tr>
       <tr>
        <td height="21" width="16%"><small><font face="Arial"><small>Item1</small></font></small></td>
        <td height="21" width="16%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
       </tr>
       <tr>
        <td height="21" width="16%"><small><font face="Arial"><small>Item2</small></font></small></td>
        <td height="21" width="16%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
       </tr>
       <tr>
        <td height="21" width="16%"><small><font face="Arial"><small>Item3</small></font></small></td>
        <td height="21" width="16%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
       </tr>
       <tr>
        <td height="21" width="16%"><small><font face="Arial"><small>Item4</small></font></small></td>
        <td height="21" width="16%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
       </tr>
       <tr>
        <td height="21" width="16%"><small><font face="Arial"><small>Item5</small></font></small></td>
        <td height="21" width="16%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
        <td height="21" width="17%"></td>
       </tr>
      </tbody>
     </table>
    </div>
    <p><small><font color="#ff0000" face="Arial">There is space for five
    &quot;Items&quot; because I made the array/table only have slots for
    five items.&nbsp; Most shopping cart programs will resize the
    &quot;TotalItems&quot; variable as they add more items to the table but
    there's no sense in taking up a lot of memory space if it's not
    necessary.&nbsp; Remember, variables are nothing more than spaces in
    memory for you to stick information.&nbsp; If I made 1,000 slots for
    items in this array it would just slow things down unnecessarily.</font></small></p>
    <p><font color="#ff0000" face="Arial"><small>If you are new to
    programming this is one of the first little bumps in your learning
    curve. These types of arrays are very common and you will definitely see
    them again and again. The array is usually defined using a Do While
    statement and it looks like: sideshowbob(i,1), sideshowbob(i,2) . The i
    defines the number of rows(or the number of items) and the number in
    parentheses tells you which column to stick the information into.&nbsp;
    It the same as telling Excel to stick the info into cell 1A, 1B, 1C etc.</small></font></p>
    <p><font color="#ff0000" face="Arial"><small>Now, all we have to do is
    fill those cells with information.&nbsp; A shopping cart allows a user
    to pick items from a database of products/items.&nbsp; Naturally, our
    database will have 5 attributes for each product/item, just like the
    array, and as the user selects these items for their order, they will
    get added to this array. Eventually this array is written to our
    database.</small></font></p>
    <p><small><font color="#ff0000" face="Arial">Remember that this array is
    &quot;alive&quot; only while the user is visting the site because it's
    bound to the user's sessionID.&nbsp; If the user ends the session, the
    cookie and the global variables are destroyed.</font></small></p>
    <p><small><font color="#ff0000" face="Arial">*Note: Sessions are usually
    assigned a 20 minute timout by the server administrator. &nbsp; So if
    you add items to the global array and leave your browser open while you
    go to lunch, after 20 minutes of inactivity ASP will automatically
    destroy your session cookie.</font></small></p>
    <p>&nbsp;</p>
   </td>
  </tr>
 </tbody>
</table>
<p><font face="Verdana"><small>Ok, so now you know the secret to a shopping
cart- global variables and a cookie. &nbsp; But you still don't know how to
implement them in an ASP application. Well, until I get a little more time I'm
going to have to leave you hanging.&nbsp; Here's a mini lesson in the meantime
to help you in the right direction.</small></font>
<ul>
 <li><font face="Verdana"><small>Session(&quot;Cart&quot;), should hold the
  individual productID information (The first part of the array).</small></font>
 <li><font face="Verdana"><small>Session(&quot;ItemCount&quot;),&nbsp; will
  hold the number of items currently in the array (The second part of the
  array). Hey, but you said that TotalItems was the other part of the array!
  Well, yeah I did, but that array is the table that actually holds the
  information you're sticking into the array- the productID(s) and the total
  number of products currently in the cart. It's kind of like an array within
  an array.</small></font></li>
</ul>
<p><font face="Verdana"><small>The other part of this puzzle is sticking these
items into the array and extracting them when necessary.&nbsp;</small></font></p>
<p><font face="Verdana"><small>1.&nbsp; Go read Charles Carol's database
tutorial.</small></font></p>
<p><font face="Verdana"><small>2.&nbsp; I highly recommend Jim Hoffman's SQL
tutorial.&nbsp; It is incredible and free!<strong> <a href="http://w3.one.net/~jhoffman/sqltut.htm"><font color="mediumblue" face>Check
it out.</font></a></strong></small></font></p>
<p><font face="Verdana"><small>Study this scenario very carefully with the info
I've given you and if you beat me to a solution I'll publish it. </small></font></p>

