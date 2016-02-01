on run {input, parameters}
	-- get content of clipboard (copied agenda)
	set agenda to paragraphs of (the clipboard)	
	-- for some reason, the it always create an extra pair of slides...
	repeat with i from 1 to ((count of agenda) - 1)
		set Tagenda to item i of agenda
		-- the actual part that creates slides in powerpoint
		tell application "Microsoft PowerPoint"
			-- create section header slide
			tell active presentation
				make new slide at end with properties {layout:slide layout section header}
			end tell
			-- put the agenda item on the section header slide
			tell last slide of active presentation
				set content of text range of text frame of shape 1 to Tagenda
			end tell
			-- create a blank text slide after the section header
			tell active presentation
				make new slide at end with properties {layout:slide layout text slide}
			end tell
		end tell
	end repeat
	return input
end run
