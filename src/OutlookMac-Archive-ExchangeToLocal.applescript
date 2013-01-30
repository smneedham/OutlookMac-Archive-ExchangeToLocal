(*
	-- ========== Outlook Mac 2011 Archive Script to local folders 2.3.0 =============
	
	Author: 		Michael Needham Oct 2012, blog.7thdomain.com/2012/09/03/auto-archive-script-for-outlook-mac-2011/ (feedback/suggestions welcome)
	
	Details:
	 			Mail:
				-----
				- Script to auto archive a full folder structure from the default (or nominated) Mail account to local 'on my computer' root folder
	 			- Set parameteres in 'Global Settings' section below, before running script
				- Script can be run manually from AppleScript Editor which is useful if you want to review the debug event log (click twice on the "Events" button above the logging window to see log output), however...
				- It's also recommended you schedule the script from Outlook's 'Run Schedule' tool to execute on a regular basis (daily recommended)
				
				Calendar:
				---------
				- Script also archives calendar events from the nominated mail account to local 'on my computer' archive calendar
	 			- Set parameteres in global settings section below before running script 


				- Calendar and mail archiving can be separately enabled with different archive periods
	
	Disclaimer: Free to use at your own risk and liability	 
*)

tell application "Microsoft Outlook"
	
	-- Global Settings which you can change if required ---------------------------------------------------------------------------------------------------------------------------
	
	-- General	
	set mailAccountDescription to "<mail account>" -- By default you don't have to change this to your account name as the script will attempt to auto detect the primary account. However, if you have multiple accounts then you set this to the greyed out name of your mail Account in the main Outlook window holding all your folders (Inbox etc)
	set runInSimulationMode to true -- when set to true no mail or calendar events will be archived. It will however create the appropriate folder structures under 'On My Computer' and the candidate items that will be archived will be logged to the events window for you to review. The script can be run repeatedly to test out the effect off different parameters below. You can also optionally delete the empty folders that were created from running in this mode if you want to re-run the simulation
	set minutesBeforeTimeOut to 2 -- When first running this script against a very large mailbox (one user had 150 000 mails in one folder to archive, for example), it can take the script a long time to build the arrays necessary for calculating the items to be archived. By default AppleScript will timeout quite quickly if it feels an application is taking too long to respond. This timeout value overrides that to allow it handle the long processing times. You can make it longer if you still experience timeouts (CPU dependent).  In extreme large cases an alterntive is to manually drag down mail to your archive and then allow the script to keep the archive up to date from there.
	set processingDelay to 0.2 -- The number of milliseconds to wait between moving messages on Outlook. On slower machines Outlook can't handle the speed at which the script requests mail to be moved sometimes causing a lock-up. It also makes Outlook more responsive while running in the background. 
	set doNotArchiveCategoryName to "Do Not Archive" -- If you create an Outlook category that has this exact name (case sensitive) and assign that category to messages or calendar events the archiving process will skip those items indefinitely
	
	
	-- Mail Archive parameters
	set archiveMailItems to true -- no mail archiving will take place if set to false
	set daysBeforeMailArchive to 80 -- number of days to keep mail in your mail account before archiving
	set localMailArchiveRootFolderName to "Archive Mail" -- name of the root archive mail folder to create under 'On My Computer'. If an existing archive mail folder is found it will use it, otherwise it will create the folder for you. Your inbox, sent items etc will appear under this folder
	set excludedMailFoldersList to {"Subscribed Public Folders", "Junk E-mail", "Deleted Items", "Sync Issues", "quarantine", "Conversation History"} -- list of mail folders in your mail account to exclude (sub-folders will also be excluded).
	set processSubFoldersofExcludedFolders to false -- By setting to true subfolders will be archived even though the parent folder is excluded for all excluded folders in above list (e.g. excluding your inbox but allowing it's sub-folders to be archived). Note that in this mode, folders with the repeated same name in your folder tree hierarchy will be all excluded if included in the excluded list.
	set doNotArchiveInCompleteTodoItems to false -- If set to true then archiving will ignore all items that are marked with a todo flag but are not complete (including items with no due date which are by definition always incomplete)
	
	
	
	-- Calendar Archive parameters
	set archiveCalendarItems to true -- no calendar archiving will take place if set to false
	set localArchiveCalendarName to "Archive Calendar" -- name of the archive calendar to create under 'On My Computer'. If an existing calendar is found it will use it, otherwise it will create the calendar for you
	set daysBeforeCalendarArchive to 730 -- number of days to keep non-recurring calendar events in your mail account before archiving
	set archiveReccuringEvents to false -- If you wish to also archive recurring events then set this to true. Warning: if a recurring event is moved to the archive it will remove the entire series from your calendar even if those recurrances are present today 
	
	--End Global Settings  (do not modify parameters or code beyond this line unless you know what you are doing) ---------------------------------------------------
	
	
	
	
	--set mail account (if none specified then use the first account found if it's not a delegated or other users folder account)
	if mailAccountDescription is "<mail account>" then
		set mailAccount to item 1 of exchange accounts
		if exchange type of exchangeAccount is not primary account then
			error "Please set an exchange account which is not delegated or another users folder account"
		end if
		log ("Processing " & name of exchangeAccount as text) & " - the primary exchange account"
	else
		set exchangeAccount to exchange account exchangeAccountDescription
		log "Processing " & exchangeAccountDescription & " - the set exchange account"
	end if
	
	log "==================== Mail ===================="
	
	-- Archive Mail if required
	if archiveMailItems then
		log "Processing mail folders"
		-- Run archive process to local folders
		my archiveMailFolders(mail folders of mailAccount, excludedMailFoldersList, my createMailArchiveFolder(localMailArchiveRootFolderName, on my computer), daysBeforeMailArchive)
	end if
	
	log "================== Calendar ==================="
	
	-- Archive Calendar Events if required
	if archiveCalendarItems then
		-- select default  calendar
		set defaultCalendar to default calendar of mailAccount
		log ("Processing " & name of mailAccount as text) & "'s primary calendar: " & (name of defaultCalendar as text)
		
		-- Move all non-recurring events to archive calendar that exceed the period of days from current date		
		my archiveCalendarEvents(defaultCalendar, my createLocalArchiveCalendar(localArchiveCalendarName), daysBeforeCalendarArchive)
	end if
	
	log "Done!"
end tell


(*================= Mail Archiving ================*)

-- Recursively archive the tree of mail folders (but ignoring the excluded folders)
on archiveMailFolders(mailFolders, excludedFolders, archiveRootFolder, daysBeforeArchive)
	
	tell application "Microsoft Outlook"
		
		-- Calculate the earliest date of mail that must remain on mail server
		set earliestDate to ((current date) - (daysBeforeArchive * days))
		log "Earliest Date - " & earliestDate
		
		
		repeat with mailFolder in mailFolders
			set mailFolderName to name of mailFolder as text
			
			set mailFolderExcluded to (mailFolderName) is in excludedFolders
			set subFoldersExist to my hasSubFolders(mailFolder)
			set currentArchiveFolder to ""
			
			-- Avoid excluded folders unless requested to process their sub-folders regardless
			if not mailFolderExcluded or my processSubFoldersofExcludedFolders then
				
				-- Only create the local folder if archiving will occur or sub-folders exist in the excluded folder
				if subFoldersExist or not mailFolderExcluded then
					-- create the destination folder locally if it doesn't exist already
					set currentArchiveFolder to my createMailArchiveFolder(mailFolderName, archiveRootFolder)
				end if
				
				if not mailFolderExcluded then
					-- archive mail in current folder
					my archiveMail(mailFolder, currentArchiveFolder, earliestDate)
				end if
				
				-- recurse sub-folders
				if subFoldersExist then
					log mailFolderName & " has sub-folders"
					my archiveMailFolders(mail folders in mailFolder, excludedFolders, currentArchiveFolder, daysBeforeArchive)
				end if
			else
				log mailFolderName & " and sub-folders excluded"
			end if
			
		end repeat
		
	end tell
end archiveMailFolders

-- Create Local Mail Archive Folder unless it exists already
-- Returns the created/found folder
on createMailArchiveFolder(mailFolderName, archiveRootFolder)
	tell application "Microsoft Outlook"
		set foundItemList to every mail folder of archiveRootFolder where name is mailFolderName
		set currentArchiveFolder to ""
		
		if (count of foundItemList) is greater than 0 then
			log "Found existing folder " & mailFolderName
			set currentArchiveFolder to mail folder mailFolderName of archiveRootFolder
		else
			log "Creating folder " & mailFolderName
			set currentArchiveFolder to make new mail folder in archiveRootFolder with properties {name:mailFolderName}
		end if
		
		return currentArchiveFolder
	end tell
end createMailArchiveFolder

-- Archive mail from mail folder to Mail Archive folder but only if older than earliestDate
on archiveMail(mailFolder, currentArchiveFolder, earliestDate)
	tell application "Microsoft Outlook"
		with timeout of (my minutesBeforeTimeOut) * 60 seconds
			
			set mailMessages to messages of mailFolder
			repeat with theIncrementValue from 1 to count of mailMessages
				set theMessage to item theIncrementValue of mailMessages
				
				if time sent of theMessage is less than earliestDate then
					
					if my excludedFromArchiving(theMessage) then
						-- Mail has been excluded by assignment to 'Do Not Archive' category
						log "Skipping mail marked for no archiving -  " & (subject of theMessage as text) & " -  " & (time sent of theMessage as text)
					else
						if todo flag of theMessage is not not flagged and todo flag of theMessage is not completed and my doNotArchiveInCompleteTodoItems is true then
							log "Skipping mail marked for todo but not complete -  " & (subject of theMessage as text) & " -  " & (time sent of theMessage as text)
						else
							log "Archiving mail -  " & (subject of theMessage as text) & " -  " & (time sent of theMessage as text)
							if not my runInSimulationMode then
								set todo flag of theMessage to not flagged
								move theMessage to currentArchiveFolder
							end if
							delay my processingDelay
						end if
					end if
				else
					log "Folder archive complete"
					exit repeat
				end if
			end repeat
		end timeout
	end tell
	
end archiveMail

(*================= Calendar Archiving ================*)

-- Archive all non-recurring events
on archiveCalendarEvents(accountCalendar, localArchiveCalendar, daysBeforeCalendarArchive)
	
	tell application "Microsoft Outlook"
		with timeout of (my minutesBeforeTimeOut) * 60 seconds
			
			
			-- Calculate the earliest date of calendar events that must remain on account server
			set earliestDate to ((current date) - (daysBeforeCalendarArchive * days))
			log "Earliest Date - " & earliestDate
			
			repeat with calendarEvent in calendar events of accountCalendar
				-- repeat starts with oldest event
				if end time of calendarEvent is less than earliestDate then
					if is recurring of calendarEvent is false and is occurrence of calendarEvent is false or my archiveReccuringEvents is true then
						
						
						if my excludedFromArchiving(calendarEvent) then
							-- Calendar Event has been excluded by assignment to 'Do Not Archive' category
							log ("Skipping event marked with no archiving - " & subject of calendarEvent as text) & " " & end time of calendarEvent as text
						else
							
							log ("Archiving event - " & subject of calendarEvent as text) & " " & end time of calendarEvent as text
							if not my runInSimulationMode then
								move calendarEvent to localArchiveCalendar
								delay my processingDelay
							end if
						end if
						
						
					else
						log ("Skipping recurring event - " & subject of calendarEvent as text) & " " & end time of calendarEvent as text
					end if
				else
					exit repeat
				end if
			end repeat
		end timeout
	end tell
end archiveCalendarEvents

-- Create local Archive Calendar unless it exists already
-- Returns the created/found folder
on createLocalArchiveCalendar(calendarName)
	tell application "Microsoft Outlook"
		
		log "Number of local calendars: " & (count of calendars of on my computer) as text
		set currentArchiveCalendar to ""
		set foundItemList to every calendar of on my computer where name is calendarName
		if (count of foundItemList) is greater than 0 then
			log "Found existing archive calendar: " & calendarName
			set currentArchiveCalendar to calendar calendarName of on my computer
		else
			log "Creating new calendar: " & calendarName
			set currentArchiveCalendar to make new calendar in on my computer with properties {name:calendarName}
		end if
		
		return currentArchiveCalendar
	end tell
end createLocalArchiveCalendar


(*=======================================================
	-- Utility helper methods
*)

-- Checks a categorizable item (e.g. calendar or mail item) for the presence of the 'Do Not Archive' category
on excludedFromArchiving(anItem)
	tell application "Microsoft Outlook"
		
		-- Check if item's category list for the 'do not archive' category
		-- [Wrote this procedurally due to problems with using 'every syntax' on category list]
		set catList to categories of anItem
		set foundDoNotArchive to false
		repeat with Y from 1 to count of catList
			set currentCat to item Y of catList
			if name of currentCat is equal to my doNotArchiveCategoryName then
				set foundDoNotArchive to true
			end if
		end repeat
		return foundDoNotArchive
		
	end tell
end excludedFromArchiving


-- Determines whether a folder has sub-folders or not
on hasSubFolders(mailFolder)
	tell application "Microsoft Outlook"
		if (count of mail folders in mailFolder) is greater than 0 then
			return true
		else
			return false
		end if
	end tell
end hasSubFolders

-- Determine if passed in folder is a root folder
on isRootFolder(mailFolder)
	tell application "Microsoft Outlook"
		if (name of container of mailFolder) is missing value then
			return true
		else
			return false
		end if
		
	end tell
end isRootFolder

