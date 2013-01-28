Auto Archive Apple Script for Outlook Mac 2011
==============================================

Out the box Microsoft Outlook for Mac 2011 has no equivalent tool to the auto archive functionality found in Outlook for Windows. This script addresses that by offering close to all the functionality available in Windows

Original author: Michael Needham (blog.7thdomain.com)

Thanks to the community on my blog/github for feedback, ideas and contributions 


To download the latest release go to: https://github.com/smneedham/OutlookMac-Archive-ExchangeToLocal/tree/master/download


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


