# Powerpoint-Scaffold
Create section headers and blank text slides based on an agenda slide using Automator/AppleScript (Mac)

Facing an empty powerpoint window is very discouraging when you try to write a ppt file ("what should I say?"). Even you manage to come up an agenda with items you want to say, creating the required slides is also a repetitive task.

This automator workflow helps you eliminate the troublesome tasks by create the section header slides and blank text slides automatically from the agenda slide (you still need to do you part, think what you want to say!) 

To use it:
  1. Download the workflow and double click it. Allow it to be installed to your Service tolder.
  2. Open PowerPoint, create a blank new ppt file.
  3. After the title slide, create a blank text slide and write your agenda.
  4. Find the service (CreateSlidesBasedOnAgenda) at the Powerpoint menu > Services.
  5. For the fist use, you may see a popup asking you to allow SystemUIServer.app to run in System Preference > Security and Privacy > Privacy
  6. Watch the service works for you.

Release history:
  [0.1] 2016-2-1. Initial release.
  [0.2] 2016-3-8.
	1. Add sections in the slides
	2. Due to lack of support in applescript for sections in powerpoint, we
	use UI controls in insert sections. As a result, different languages
	will need different strings (insert > section)
	3. Add slide and section for “End”
  [0.3] 2016-3-9.
	1. Add key press events to copy agenda to clipboard (cmd+A > cmd+C)