Auto Archive Apple Script for Outlook Mac 2011
==============================================

Out the box Microsoft Outlook for Mac 2011 has no equivalent tool to the auto archive functionality found in Outlook for Windows. This script addresses that by offering close to all the functionality available in Windows.

Original author: [Michael Needham](http://blog.7thdomain.com)

Thanks to the community on my blog/github for feedback, ideas and contributions 


To [download](https://github.com/smneedham/OutlookMac-Archive-ExchangeToLocal/tags), select the version you want from the drop down then download the .scpt file from the 'download' folder of that release


  



Background:
----------
If you aren’t in the know: Archiving for Outlook is a process of automatically copying your full email folder structure and/or calendar events from (typically) an exchange server to a local folder structure on your computer. It’s used in scenarios like:

- Your company operates a data limit or ‘delete mail after x days’ policy which means you must move mail out of the exchange account to not cause disruption to your flow of mail (or to not suffer data loss of important mails in your past history)

- You want to keep your exchange account lean for performance reasons, but still be able to search gigs of mail going back many years

- You have a number of attachments in calendar events from years back that accumulate over time to lower the available data quota for email

- You aren’t happy with more manual methods of dragging mail down to your local storage when you happen to remember…etc

The search for a tool:

First stop was to Google for an AppleScript that would do the job. Though I found an simple implementation for Mac Mail there was nothing I could find for Outlook Mac 2011 that matched the features of the Windows version.

Next stop was search for a free tool. Only found a commercial tool called OEAO (Outlook Exchange Account Optimizer). This tool came closer but is trial software and cost $25 to licence. It also didn’t support archiving calendar events or the concept of ‘exclusion’ folders at the time (I didn’t want to archive, junk mail and deleted items for example).

After messing with Outlook Rules and finding them very weak it looked like the only option was to dust off the coding skills and write my first AppleScript. Here’s the end result:

- A script which by default will archive mail from your primary exchange account using the following default archive settings:

  -Archive mail older than 80 days from all folders (except folders like ‘Subscribed Public Folders’, ‘Junk Mail’, ‘Deleted Items’, etc)
  
  -All archived mail folders will be placed under the folder of ‘On My Computer’ called ’Archive Mail’
  
  -All non-recurring calendar items older then 2 years will be archived to a calendar of ‘On My Computer’ called ‘Archive Calendar’

- The script has a number of parameters which you can change if you are not happy with the default settings.  

- Calendar and mail archiving can be configured separately, depending on your needs.

- Fire and forget: … using the ‘Tools’ menu in Outlook and ‘Run Schedule’ you can run the script on a regular basis.

For detail on configuring the script and other instructions see the following blog posts:

http://blog.7thdomain.com/2012/09/03/auto-archive-script-for-outlook-mac-2011/

http://blog.7thdomain.com/2012/10/06/configuring-outlook-mac-2011-auto-archive-script/

Change Log
==========

2012/12/08 - [Ver 2.2.0](https://github.com/smneedham/OutlookMac-Archive-ExchangeToLocal/tree/v2.2.0): 

- If you create a category in Outlook called “Do No Archive” there is a setting in the script that will now ignore any mails or calendar items assigned to this category even if they are candidates for archiving

- You can optionally enable a setting to not archive items flagged as Todo but are not yet marked complete even if they are candidates for archiving

- You can now archive recurring calendar events (but be warned that will remove the entire series even in to present day so use with caution)

- AppleScripts default timeout period that it waits for applications like Outlook to finish processing a task is quite short. The script now overrides this to 2 minutes by default but it can be made longer if you still experience timeouts (especially when you first run the script on a large mailbox and it has to process a long back log of archive items).

 

Script [Ver 2.13]: http://bit.ly/Srh2md

Change log 2012/11/27

- Added in a simulation mode setting which allows you to review the empty archive folder structure created on ‘On My Computer’ without moving all the mail and calendar items (can be run repeatedly without issue). If run in the AppleScript Editor the candidate mail/events that will be moved once simulation is turned off are logged to the events window. This allows you to test out various parameters to the script to see the effect before archiving any items.

- Added in a new setting which allows you to archive sub-folders of excluded folders. Useful if you want exclude your inbox but still archive sub-folders of the inbox. By default this setting is not turned on and any excluded folder that is listed will also have it’s sub-folders ignored.

- Found that attempting to archive “sync errors” folders was causing the script to lockup. Updated the default exclusion folder to ignore this folder

- Timeouts can also occur on slower CPU machines when Outlook can’t move the mail quick enough for the speed of the script. Made this a parameter and increased it to 200ms delay between message moves. This mostly affects the initial archive processing when you have large back log of mail which can take a long time and is most prone to timeouts. When run daily this small speed delay will make very little difference. If you aren’t suffering from timeouts then you can change this parameter to .1 or zero seconds. [UPDATE:  In version 2.2 (to be released shortly) I have added an overall script timeout parameter too that, if increased, will stop the script from timing out when there are large volumes of mail to process (usually only on the first time you run the script on a large mailbox)


Script [Ver 2.12]: http://goo.gl/RtDPy

Change log 2012/11/05

- For most users the script works flawlessly but a small percentage of users have problems detecting the Archive Folder or the Archive Calendar especially if they customise the script parameters. I can’t replicate the problem but this release is an attempt to fix the routine that creates these folder/calendars to see of it makes a difference


Script [Ver 2.11]: http://goo.gl/lYgZK

Change log 2012/10/13

- When using the script for the first time on a large mailbox the script could lock up Outlook due to bug in Outlooks message move command. Outlook can’t handle the speed at which the script is sending move requests. By introducing a slight delay between processing of messages the script can now manage very large mailboxes (though it will take a little longer on first run)


Script [Ver 2.10]: http://goo.gl/lEXh4

Change log 2012/10/06:

- Script now looks for the primary exchange account automatically which means if you are happy with the default archive settings (see blog post above) then it will just run out the box with no need to edit the script file and fill in exchange account name and other parameters (proving difficult for less technical users)

- Wrote a post on configuring the script for those that want to modify the default settings or other more advanced tasks


Script [Ver 2.0]: http://goo.gl/syzIV

Change log 2012/09/20:

- Added in the feature to optionally archive calendar events

- Can independently control the mail and calendar archive settings


Script [Ver 1.01]: http://goo.gl/7nKgv

Change log 2012/09/18:

- The archive folder specified will be created if not found under ‘On My Computer’

- Attempts to fix the problem that on some Mac/Outlook versions this script fails to find the archive folder when it is created manually


Script [Ver 1.00]: http://goo.gl/Xplpt 

- Initial release


Disclaimer: Free to use but 100% at your own risk (works for me)

Feedback/Bugs/Suggestions welcome


