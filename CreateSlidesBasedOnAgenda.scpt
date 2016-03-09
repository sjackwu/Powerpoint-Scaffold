
	-- copy agenda items to clipboard	
	tell application "System Events"
		-- hit command+a and command+c to copy the agenda
		keystroke "a" using {command down}
		delay 0.1
		keystroke "c" using {command down}
	end tell
	
	--zh-tw strings on MS PowerPoint menu (插入 > 章節)
	set Tinsert to "插入"
	set Tsection to "章節"
	--en-US strings on MS PowerPoint menu (Insert > Section)
	--set Tinsert to "Insert"
	--set Tsection to "Section"
	
	-- get content of clipboard
	set agenda to paragraphs of (the clipboard)
	
	-- for some reason, it always create an extra pair of slides...
	repeat with i from 1 to ((count of agenda) - 1)
		set Tagenda to item i of agenda
		
		-- the actual part that creates slides in powerpoint
		tell application "Microsoft PowerPoint"
			-- create section header slide and give it a title (agenda item)
			tell active presentation
				make new slide at end with properties {layout:slide layout section header}
			end tell
			tell last slide of active presentation
				set content of text range of text frame of shape 1 to Tagenda
			end tell
		end tell
		
		-- create a section for the agenda item, you can not create a section contains no slide. Need to create the section header slide first (see above)
		set the clipboard to item i of agenda
		tell application "Microsoft PowerPoint" to activate
		tell application "System Events"
			tell process "Microsoft PowerPoint"
				click menu item Tsection of menu Tinsert of menu bar 1
			end tell
			-- hit command+v to paste the clipboard and hit return to close it.
			keystroke "v" using {command down}
			keystroke return
		end tell
		
		-- create a blank text slide for the agenda item
		tell application "Microsoft PowerPoint"
			tell active presentation
				make new slide at end with properties {layout:slide layout text slide}
			end tell
			tell last slide of active presentation
				set content of text range of text frame of shape 1 to Tagenda & ":"
			end tell
		end tell
	end repeat
	
	-- create a thank you slide and an "end" section for it.
	tell application "Microsoft PowerPoint"
		tell active presentation
			make new slide at end with properties {layout:slide layout title slide}
		end tell
		tell last slide of active presentation
			set content of text range of text frame of shape 1 to "Thank You!"
		end tell
		
		set the clipboard to "End"
		
		tell application "System Events"
			tell process "Microsoft PowerPoint"
				click menu item Tsection of menu Tinsert of menu bar 1
			end tell
			-- hit command+v to paste the clipboard and hit return to close it.
			keystroke "v" using {command down}
			keystroke return
		end tell
	end tell