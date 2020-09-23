<div align="center">

## Active Sessions, Active Logins and Total Site Hits with ASP


</div>

### Description

Active Sessions, Active Logins and Total Site Hits with ASP
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SoftwareMaker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/softwaremaker.md)
**Level**          |Advanced
**User Rating**    |3.9 (35 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__4-33.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/softwaremaker-active-sessions-active-logins-and-total-site-hits-with-asp__4-7422/archive/master.zip)





### Source Code

<P>Many websites use a tracking-component pasted on their homepage to keep track
on the no. of visits to their website. This method is not accurate as it doesnt
really keep track of the genuine page-hits to their site (not that it really
matters...). Everytime the page where the component is refreshed, the counter is
incremented even though the user is the same and if a link takes you to some
other page other than the page where this component is pasted, the tracker
device will lose count of this hit. Sure you can paste this tracker component on
every page on your site, but that would be highly cumbersome, inefficient and
inaccurate as that would be lots of double-counting. You also cannot keep track
of the active sessions on your site which means the number of visitors to your
site at any moment. Morever, once you update your page, or when the server
shuts-down for maintenance, you lose all data and history of your hit
counts.</P>
<P>I am sure there are components out there who can do the job but it all
depends on the price and whether your remote web administrator allows you to
install these components on your virtual server (sharing the server with many
websites).</P>
<P>There is a simple and free way to do what is described above via Microsoft
Active Server Pages (ASP) which is widely used and implemented and free. Most
websites, whether they are running IIS or Apache supports ASP.</P>
<P>All you need to do is to upload this script on this page to your server and
you will be able to see a fairly accurate count of hits on your site, not just
the page alone.</P>
<P>The scripts provided below MUST be pasted on a text-based file called
GLOBAL.ASA. This file is always read by the server first whenever a browser
requests HTTP content from the web server and its therefore accurate as any
pages requested by the server will go through the GLOBAL.ASA file first.</P>
<P>The GLOBAL.ASA file is a fairly narrow topic but I will only highlight the 4
events that the GLOBAL.ASA file implements. You can also declare
application-wide objects and variables on this file but we will only focus on
keeping track of site-hits here.</P>
<p>Many websites use a tracking-component pasted on their homepage to keep track on the no. of visits to their website. This method is not accurate as it doesnt really keep track of the genuine page-hits to their site (not that it really matters...). Everytime the page where the component is refreshed, the counter is incremented even though the user is the same and if a link takes you to some other page other than the page where this component is pasted, the tracker device will lose count of this hit. Sure you can paste this tracker component on every page on your site, but that would be highly cumbersome, inefficient and inaccurate as that would be lots of double-counting. You also cannot keep track of the active sessions on your site which means the number of visitors to your site at any moment. Morever, once you update your page, or when the server shuts-down for maintenance, you lose all data and history of your hit counts.</p>
<p>I am sure there are components out there who can do the job but it all depends on the price and whether your remote web administrator allows you to install these components on your virtual server (sharing the server with many websites).</p>
<p>There is a simple and free way to do what is described above via Microsoft Active Server Pages (ASP) which is widely used and implemented and free. Most websites, whether they are running IIS or Apache supports ASP.</p>
<p>All you need to do is to upload this script on this page to your server and you will be able to see a fairly accurate count of hits on your site, not just the page alone.</p>
<p>The scripts provided below MUST be pasted on a text-based file called GLOBAL.ASA. This file is always read by the server first whenever a browser requests HTTP content from the web server and its therefore accurate as any pages requested by the server will go through the GLOBAL.ASA file first.</p>
<p>The GLOBAL.ASA file is a fairly narrow topic but I will only highlight the 4 events that the GLOBAL.ASA file implements. You can also declare application-wide objects and variables on this file but we will only focus on keeping track of site-hits here.</p>
<p>The events that the GLOBAL.ASA file implements corresponds to the Application and Session Object. They are namely Application_OnStart, Session_OnStart, Session_OnEnd and Application_OnEnd. They are all run in that order.</p>
<p>Application objects are used throughout the application regardless of the users and is started when the web-server starts and ends when the web-server shuts down (or when you copy a newer version of the GLOBAL.ASA file onto it). Session objects, on the other hand, are used per browser session. This means that a new browser, regardless of where it comes from, instantiates a new session. Session objects are cleared and re-initialized when the session times out which in most cases is 10 minutes.</p>
<p>Once you understand the basic concepts of Application and Session objects, you are ready to do wonders with your web application which traditionally takes a lot more coding and programming than your basic standalone ones as the HTTP protocols web applications run on are basically stateless (dont remember anything once they are run).</p>
<p>Lets get our hands dirty by understanding and writing some codes. Because, a Session is instantiated whenever a new browser requests content, a page hit should always be incremented here. To provide for data-persistence when a server shuts down, we will attempt to keep an application count (used throughout the application until the server shuts down) and then write that application count to a text file or a database to keep a "Hard-Copy" record of the total number of site-hits.</p>
<p align="center"> Sub Session_OnStart()<br>
  Dim TotalVisits<br>
  Application.Lock</p>
<p align="left">Application(&quot;aHitCounter&quot;) is the application variable for keeping track of total number of site-hits<br>
Application(&quot;aActiveSess&quot;) is the application variable for keeping the number of active sessions of your site the current moment.<br>
In case this Application variable to keep track of active sessions doesnt exists, initialize it to 0</p>
<p align="center">  If IsEmpty(Application("aActiveSess")) Then<br>
    Application("aActiveSess") = 0<br>
  End If<br></p>
Read from a stored Text File or a Database that keeps a record of the total no. of site hits.<br>
(We wont cover this topic here) and save the result to a variable called TotalVisits<br>
Increment TotalVisits by 1
<p align="center">  TotalVisits = Clng(TotalVisits) + 1</p>
<p>Save this TotalVisits Variable to the Application("aHitCounter") variable<br>
This application variable can then be used throughout the application, regardless of the number of sessions</p>
<p align="center">  Application("aHitCounter") = TotalVisits</p>
<p>Write Back to the same stored Text File or Database that keeps a record of the total no. of site hits the
incremented value of the total number of Hits<br>
To Increment the Active Session Count Whenever a New Session Starts</p>
<p align="center">  Application("aActiveSess") = Application("aActiveSess") + 1<br>
  Application.UnLock<br>
  End Sub</p>
<p>After a new user requests any pages from your site, this GLOBAL.ASA file will always run with the this event above. After that, throughout the application, we can access the total number of hits and the active sessions through the Application("aHitCounter") variable and Application("aActiveSess") variable respectively.<br>
Now, we have to take care of the Application("aActiveSess") to make sure that this counter is decremented properly when the user session ends or times out. This is done via the Session_OnEnd Event</p>
<p>Decrement the Active Session Count After Session Ends</p>
<p align="center">  Sub Session_OnEnd</p>
<p align="center">  Application.Lock<br>
  Application("aActiveSess") = Application("aActiveSess") - 1<br>
  Application.UnLock</p>
<p align="center">
  End Sub</p>
<p>Thats all. When the user session ends or times out, the above event will run and this application("aActiveSess") will be decremented accordingly to better reflect the number of ACTIVE sessions on you site.<br>
After reading above Part one of this article and understanding the concept of Application and Session objects, I will dwell into keeping track of active Logins on your site. This is slightly trickier as we have to take into account that not every user that surfs your site will Login and not every Login user will manually terminate their Login session by
logging out.<br>
To achieve the above objectives, we will make use of two flag variables called Application("aflgLogin") and Session("sflgLogin") that keeps track of whether the user has logged out manually or the session terminated naturally.<br>
If you have a Login page running on ASP, chances are that you will be verifying the username and password passed with the values stored in a database. Once that is verified, you would most likely assign Session variables to the Login user to keep track of his username, etc. All we need to add after the user is verified is to add these 2 lines of code once the user is verified.</p>
<p align="center">  'ON THE LOGIN PAGE<br>
  Session("sflgLogin") = True 'Session Flag Variable<br>
  Application("aLogins") = Application("aLogins") + 1</p>
<p>From there, we know the number of Active Logins by retrieving the value of Application("aLogins") variable and write it to an ASP page.</p>
<p>On the logout page if the user wants to manually logout, we implement these lines behind the ASP Logout page</p>
<p align="center">  'ON THE LOGOUT PAGE, Decrement the Active Logins accordingly<br>
  If Session("sblnLogin") = True then 'Utilize the Session Flag Variable<br>
Application(&quot;aLogins&quot;) = Application("aLogins") - 1</p>
<p>'For Safety measures that the Active Logins cannot fall below ZERO</p>
<p align="center">  If Application("aLogins") <= 0 then<br>
    Application("aLogins") = 0<br>
  End If<br>
  Session("sblnLogin") = False<br>
  Application.Lock</p>
A application Flag variable to be passed to Session_OnEnd when Session.Abandon is called next<br>
Reason why we pass to an Application Flag Variable is because the Session.Abandon method clears all
Session variables and runs the Session_OnEnd event of the GLOBAL.ASA file.<br>
We need to keep track of whether the user has logged out manually or if the session has timed out naturally
and we therefore need to keep track of this Application Flag Variable.<p align="center">
  Application(&quot;ablnLogin&quot;) = False<br>
  Application.Unlock<br>
  End If<br>
  Session.Abandon</p>
<p>On the GLOBAL.ASA file, we need to add these lines in the following events</p>
<p align="center">
  Sub Session_OnStart</p>
<p align="center">'In case this Application variable to keep track of active Logins doesnt exists, initialize it to 0<br>
<br>
  If IsEmpty(Application("aLogins")) Then<br>
    Application("aLogins") = 0<br>
  End If<br>
  End Sub<br>
  Sub Session_OnEnd<br>
  Application.Lock</p>
Application(&quot;aflgLogin&quot;) = False means User has Destroyed Session, therefore
Reset Application("aflgLogin") then Exit Sub<br>
No need to Decrement Application("aLogins") anymore as it had been done so at LogOut ASP Page<p align="center">
  If Application("aflgLogin") = False then<br>
    Application("aflgLogin") = ""<br>
    Application.Unlock<br>
  Exit Sub<br>
  End If</p>
If Program flowed into here, it means that Application("aflgLogin") is not False
which means that the session<br>
terminated naturally, therefore Decrement Application(&quot;aLogins&quot;) here.<p align="center">
  Application("aLogins") = Application("aLogins") - 1</p>
<p>'For Safety measures that the Active Logins cannot fall below ZERO</p>
<p align="center">  If Application("aLogins") <= 0 then<br>
    Application("aiLogins") = 0<br>
  End If<br>
  Application.Unlock<br>
  End Sub</p>
<p>These are just my way of implementing a counter to keep track of Active Logins on my site. It may not be the most efficient method around so if any of you fellow developers have a better way to implement this, please do not hesistate to email me (itnews@Softwaremaker.Net) and let me know.</p>
<p><br>
My website http://www.Softwaremaker.Net has an implementation of this script. Try it out by refreshing my homepage rightaway, then refreshing it after 10 minutes when your session times out or try to open up a new browser window to start a new session to see the active sessions and the number of site hits increment then wait for it to time out and watch the active session count decrement. There is also a counter to keep track the number of active Logins at the current moment so try loggin in and watch the active Login Counter increment and then do two different logouts to see the Active Login Counter decrement, one by manually terminating the session by clicking the Logout button and the other by waiting for the session to end after 10 minutes.</p>
<p>As you can see, once you understand the concept of the Application and the Session Object, we can easily manipulate the variables accordingly to what we want. But be careful, the perfect software engineering concept is often a Holy Grail at best. Overuse of the Application object can have undesirable effects on your web application as it will be harder to keep track of what each application variable is storing. It will also result in slower loading and memory leakage if unused application objects are not cleared properly. (Enter ASP.NET and the Garbage Collector to solve these problems...) The GLOBAL.ASA file has its own quirks too. Each time the web server is shut down, the GLOBAL.ASA file has to be re-copied back up to its remote location or your site will fail to initialize. But these are topics to be discussed on another day in another forum. Enjoy.</p>

