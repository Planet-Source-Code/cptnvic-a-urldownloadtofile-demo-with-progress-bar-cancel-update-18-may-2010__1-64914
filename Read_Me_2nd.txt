IF YOU THINK YOU CAN JUST UN-ZIP THIS CODE AND HIT F5... YOU ARE NUTZ!!
IF YOU DON'T READ ANYTHING ELSE... READ THE READ_ME_FIRST.txt File!!!!!
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
The credits are at the bottom of this page.
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Updated: 15 June 2006:

I updated this project because Gabe wrote and asked if a cancel could
be implemented in my previous project.  I knew it could because I use
variations of it often, but had never gotten around to submitting the
code to PSC.  Frankly, I didn't think the previous project had enough
interest to warrant spending the time on.  Anyway, I slathered a cancel
function into the previous version and hated it... so NOW, you get the
improved version because I didn't want Gabe to be too ashamed of me!

+ IBindStatusCallback stuff is now in a class module for more portability
  and ease of use... PLUS...
+ 5 (I think) extra events (free of charge too!)
+ Cancel capability for Gabe
+ Other misc. tidiness tasks accomplished.  (But, it's still a demo!)
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
BEFORE you email... the answer is NO.  You don't have to distribute the 
entire type library.  When you compile your project, VB will get what 
it needs from the type library and compile it into your executible as it 
would for a user control.  AND NOT ONLY THAT... The OTHER answer is NO!  
URLDownloadToFile will NOT circumvent a servers password protected 
files/folders.  (Although, I have some code laying around here somewhere
that will offer the username/password with URLDownloadToFile).
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Stuff from version1:
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
I write alot of sports related software, and this project was started 
when I noticed that URLDownloadToFile was returning the same file 
irrespective of how many times I updated the source on the clients web site.
That is not a good thing when you are trying to update scores!  I noticed 
that while the file returned... my internet monitor never even blinked.  
So, I decided to see if anyone on PSC had noticed this and had arrived 
at a fix.  Several people noticed, no one (that I could find) offered a fix.

While I was researching the nature of my problem (which, by the way is 
IE's cache), I came accross some similar code on PSC (using URLDownloadToFile) 
and while reading the comments to one piece of code, someone's comments 
were something like, "If you want to impress me, write a sub that handles the
callbacks needed for a progress bar."... something like that.

I don't know who made those comments, but I spent most of the weekend 
trying to boil down the urlmon.dll, and if the commentor can do that in a 
sub... he'll/she'll have NO PROBLEM finding work!  I finally gave up.  I see 
no practical method of getting this dll to be VB friendly.  I was about to 
return to my original problem when I stumbled accross this type library.  
And, VIOLA, problem semi-solved.  On the plus hand, what makes URLDownloadToFile
attractive is that it only takes a couple of lines of code to implement.  
On the minus hand, it locks up your app until the file is done downloading... 
for big files, that stinks.  This project does away with that.

This olelib.URLDownloadToFile takes a bit more code since you must expose 
a number of events.  In addition, the type library weighs in at just over 1 meg... 
however, the project attached will compile to 40KB (Native code, optimized
for speed) so VB's not stealing that much from the library.  ON the plus 
hand, having events frees you to do other things while downloads are happening.  
In this project, I send the download off on it's own in a different form... but 
you can handle it as you like.  

I don't care what anyone else tells you, the urlmon.dll is one squirrely mother 
grabber!  Twitch after twitch!  I have faithfully tried to find and eliminate 
any bugs... but I'm sure there are still some I haven't had the misfortune to find
yet.  If you find them, let me know.

Having said all of that, the purpose of this hastily thrown together project is to:
1) Retain the relatively light URLDownloadToFile 
2) Use the DeleteUrlCacheEntry (Wininet.dll) to clear a pre-existing file from IE's cache
3) Provide for progress notification of a download's status.

As far as I can tell, I have accomplished those tasks.
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Known Bugs:
Nothing major:

In the IBindStatusCallback_OnProgress sub:
	(This bug(?) repaired with error handling in new version)
	
In the ShowDownLoad Function:
	(STILL an Issue!)
	DeleteUrlCacheEntry does an excellent job of removing cached files from IE's cache...
	particularly for web related items (html, gif's, etc.).  HOWEVER, it does not find 
	all cached file copies.  This happens to me when I download a MP3... delete the file, 
	and re-download... it shows back up in a flash... so it is clearly cached somewhere, 
	I just haven't figured out where yet... could be recent files...
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
KNOWN CONCERNS:
	My antivirus program never blinks when I'm downloading using this code... 
	I'm not so sure that's good.
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
CREDITS:
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
I GOT A LOT OF HELP FROM:  
++ http://msdn.microsoft.com/library/default.asp?url=/workshop/networking/moniker/reference/ifaces/ibindstatuscallback/ibindstatuscallback.asp 
++ AND OF COURSE, Edamo's OLE interfaces & functions v1.81, available freely at:
++ http://www.mvps.org/emorcillo/download/vb6/tl_ole.zip 
++ He has simple re-use requirements... see them at: 
++ http://www.mvps.org/emorcillo/en/index.shtml 
++ He has some other great stuff at: http://www.mvps.org/emorcillo/en/index.shtml, sadly, 
   they have given up on VB long ago.

I spent hours researching this, and if you spend hours doing the same, you'll find a very few 
sites with any help on this subject.  The two, or so, sites that I found that had even a code 
snippet on this were all bug-ridden if they ran at all.  My project may not be pretty... but 
it WILL RUN!

As far as I'm concerned, you may do with this code as you wish as long as what you wish does 
not infringe on my rights to do with my code as I wish.  That seems fair enough!

DO ME A FAVOR WOULD YOU... since most of this project is comprised of api (that I didn't write) 
and relies heavily on Edamo's type library, I'm not expecting any votes.  All I did was glue this 
stuff together.  HOWEVER, I have been trying to read more comments when I visit PSC... and MERCY!  
A less constructive bunch would be hard to find!  I know it's a pain to find a bug and report it 
gracefully, or see some good code and offer some coder some encouragement and maybe a pointer... 
but it sure seems like a good idea to me.  Try it, you may like it!  At least I offer you this 
deal, if you'll leave some constructive critisim... I'll learn to spell cool with a K.
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++	
