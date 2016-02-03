-- Downloaded From: http://c-command.com/scripts/spamsieve/outlook-filter-mailboxes
-- Last Modified: 2015-10-29


property pMinutesBetweenChecks : 3
property pGoodCategoryName : "Good"
property pJunkCategoryName : "Junk"
property pEnableDebugLogging : false

on mailboxNamesToFilter()
	return {"Inbox", "INBOX"}
end mailboxNamesToFilter

-- Do not modify below this line.

on run
	-- This is executed when you run the script directly.
	my filterMailboxes()
end run

on idle
	-- This is executed periodically when the script is run as a stay-open application.
	my filterMailboxes()
	return 60 * pMinutesBetweenChecks
end idle

on filterMailboxes()
	tell application "System Events"
		if not (exists process "Microsoft Outlook") then
			my debugLog("Outlook is not running")
			return
		end if
	end tell
	try
		set _mailboxes to my mailboxesToFilter()
		repeat with _mailbox in _mailboxes
			set _messages to my messagesToFilterFromMailbox(_mailbox)
			if pEnableDebugLogging then
				my debugLog((count of _messages) & " messages to filter in " & locationFromMailbox(_mailbox))
			end if
			repeat with _message in _messages
				set _score to scoreMessage(_message)
				if doesMessageHaveCategoryNamed(_message, pJunkCategoryName) then
					my processSpamMessage(_message, 100)
				else if _score ³ 50 then
					my processSpamMessage(_message, _score)
				else
					my processGoodMessage(_message, _score)
				end if
			end repeat
		end repeat
	on error _error
		my logToConsole("Error: " & _error)
	end try
end filterMailboxes

on mailboxesToFilter()
	set _result to {}
	set _names to my mailboxNamesToFilter()
	repeat with _name in _names
		tell application "Microsoft Outlook"
			considering case
				set _matches to (every mail folder whose name is _name)
			end considering
		end tell
	end repeat
	set _result to _result & _matches
	return _result
end mailboxesToFilter

on messagesToFilterFromMailbox(_mailbox)
	tell application "Microsoft Outlook"
		set _messages to my unreadMessagesFromMailbox(_mailbox)
		set _result to {}
		repeat with _message in _messages
			if my shouldFilterMessage(_message) then
				copy _message to end of _result
			end if
		end repeat
		return _result
	end tell
end messagesToFilterFromMailbox

on unreadMessagesFromMailbox(_mailbox)
	tell application "Microsoft Outlook"
		set _startDate to current date
		try
			with timeout of 2 * 60 seconds
				-- changed to have all messages 10/28/15
				-- set _messages to messages of _mailbox whose is read is false
				set _messages to messages of _mailbox
			end timeout
		on error _error number _errorNumber
			my logToConsole("Outlook reported error Ò" & _error & "Ó (number " & _errorNumber & ") getting the messages from " & my locationFromMailbox(_mailbox))
			return {}
		end try
		set _endDate to current date
		set _duration to _endDate - _startDate
		set _statusMessage to "Outlook took " & _duration & " seconds to get unread messages from " & my locationFromMailbox(_mailbox)
		if _duration > 5 then
			my logToConsole(_statusMessage)
		else
			my debugLog(_statusMessage)
		end if
		return _messages
	end tell
end unreadMessagesFromMailbox

on shouldFilterMessage(_message)
	if pGoodCategoryName is "" then
		return true
	end if
	return not my doesMessageHaveCategoryNamed(_message, pGoodCategoryName)
end shouldFilterMessage

on doesMessageHaveCategoryNamed(_message, _categoryName)
	tell application "Microsoft Outlook"
		set _categories to _message's category
		repeat with _category in _categories
			if _category's name is _categoryName then
				return true
			end if
		end repeat
		return false
	end tell
end doesMessageHaveCategoryNamed

on scoreMessage(_message)
	tell application "Microsoft Outlook"
		set _source to _message's source
	end tell
	if _source is missing value then
		my logToConsole("Outlook could not get the source of message: " & _message's subject)
		return 49
	else
		tell application "SpamSieve"
			return score message _source
		end tell
	end if
end scoreMessage

on processSpamMessage(_message, _score)
	my debugLogMessage("Predicted Spam (" & _score & ")", _message)
	if my isSpamScoreUncertain(_score) then
		my applyCategoryNamed(_message, "Uncertain Junk")
	else
		my applyCategoryNamed(_message, pJunkCategoryName)
		tell application "Microsoft Outlook"
			try
				set _pendingFolder to folder "Pending Review" of _message's account
			on error
				set _pendingFolder to folder "Pending Review"
			end try
			move _message to _pendingFolder
		end tell
	end if
	
end processSpamMessage

on processGoodMessage(_message, _score)
	my debugLogMessage("Predicted Good (" & _score & ")", _message)
	my applyCategoryNamed(_message, pGoodCategoryName)
end processGoodMessage

on isSpamScoreUncertain(_score)
	tell application "SpamSieve"
		set _keys to {"Border", "OutlookUncertainJunk"}
		set _defaults to {75, true}
		try
			set {gUncertainThreshold, gUncertainJunk} to lookup keys _keys default values _defaults
		on error
			set {gUncertainThreshold, gUncertainJunk} to _defaults
		end try
	end tell
	return _score < gUncertainThreshold and gUncertainJunk
end isSpamScoreUncertain

-- Categories

on categoryForName(_categoryName)
	tell application "Microsoft Outlook"
		try
			-- "exists category _categoryName" sometimes lies
			return category _categoryName
		on error
			try
				-- getting by name doesn't always work
				repeat with _category in categories
					if _category's name is _categoryName then return _category
				end repeat
			end try
			set _category to make new category with properties {name:_categoryName}
			if _categoryName is pGoodCategoryName then
				set _category's color to {0, 0, 0}
			end if
		end try
		return category _categoryName
	end tell
end categoryForName

on applyCategoryNamed(_message, _categoryName)
	tell application "Microsoft Outlook"
		set _categoryToApply to my categoryForName(_categoryName)
		set _categories to _message's category
		repeat with _category in _categories
			if _category's id is equal to _categoryToApply's id then return
		end repeat
		set category of _message to {_categoryToApply} & category of _message
	end tell
end applyCategoryNamed

-- Logging

on debugLogMessage(_string, _message)
	if not pEnableDebugLogging then return
	tell application "Microsoft Outlook"
		set _location to my locationFromMailbox(_message's folder)
		set _subject to _message's subject
	end tell
	my debugLog(_string & ": [" & _location & "] " & _subject)
end debugLogMessage

on locationFromMailbox(_mailbox)
	tell application "Microsoft Outlook"
		try
			set _accountName to name of _mailbox's account
		on error
			set _accountName to "On My Mac"
		end try
		set _mailboxName to name of _mailbox
		return _accountName & " > " & _mailboxName
	end tell
end locationFromMailbox

on debugLog(_message)
	if pEnableDebugLogging then my logToConsole(_message)
end debugLog

on logToConsole(_message)
	set _logMessage to "SpamSieve [Outlook Filter Mailboxes] " & (characters 1 thru 64 of _message)
	do shell script "/usr/bin/logger -s " & _logMessage's quoted form
end logToConsole
