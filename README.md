<div align="center">

## Autocheck www\.mail\.com E\-mail accounts


</div>

### Description

This code runs in your system tray and checks a specifed www.mail.com account at intervals you specify, ie. every 5 minutes, 20 minutes etc etc.

If new e-mail is found, it will play a user specifed wav file and ask you if you want to goto the login page.

The way it works is to recreate the login page into a string, filling in login and password while its doing it. A Javascript to automatically submit the form is also added, which stimulates clicking the submit button onLoad. This string is saved to a file, then opened into a webbrowser, which stimulates logging into your account. It will then use Inet to retrieve your inbox -- which it we be able to retrieve as you were logged in using the webbrowser control -- and do searches on that page for keywords denoting no new message. If theres no new messages - there must be new messages.

Login information is encrypted and stored in the registry for security reasons. Sure it aint pretty, but it does what its meant to do. Hope someone finds it useful...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2000-07-12 17:39:10
**By**             |[Jamie Dillon](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jamie-dillon.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD77167122000\.zip](https://github.com/Planet-Source-Code/jamie-dillon-autocheck-www-mail-com-e-mail-accounts__1-9703/archive/master.zip)








