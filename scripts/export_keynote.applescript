on run argv
	set sourcePath to item 1 of argv
	set targetPath to item 2 of argv
	set sourceFile to POSIX file sourcePath
	set targetFolder to POSIX file targetPath

	tell application "Keynote"
		activate
		set openedDocument to open sourceFile
		delay 1
		try
			export openedDocument to targetFolder as slide images with properties {image format:PNG}
		on error errorMessage number errorNumber
			close openedDocument saving no
			error errorMessage number errorNumber
		end try
		close openedDocument saving no
	end tell
end run
