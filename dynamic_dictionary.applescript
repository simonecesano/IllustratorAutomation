-- ---------------------------------------------------------------------------
-- test script for dynamic data structures in Applescript 
-- ---------------------------------------------------------------------------

use framework "Foundation"

set variableListFunderName to {}
set i to 1
repeat while i <= 12
       -- create a dictionary
       set funderNameDict to current application's NSMutableDictionary's alloc()'s init()
       tell funderNameDict
       	    -- set value for key
       	    set t to { a: 1, b: 2 }
       	    its setObject: t forKey: "key " & i
       end tell
       set variableListFunderName to variableListFunderName & ""
       set item i of variableListFunderName to funderNameDict
       set i to i + 1
end repeat
log (length of variableListFunderName as text) & " items"
log "------"


set j to 1
repeat while j <= 12
    log j
    repeat with k in allKeys() of item j of variableListFunderName
	   set i to item j of variableListFunderName
	   set v to i's objectForKey:k

	   log k as text
	   log { (a of v as text), (b of v as text) }
    end repeat
    set j to j + 1
end repeat


-- https://developer.apple.com/library/content/documentation/AppleScript/Conceptual/AppleScriptLangGuide/reference/ASLR_handlers.html#//apple_ref/doc/uid/TP40000983-CH7g-SW4
-- https://developer.apple.com/library/content/documentation/AppleScript/Conceptual/AppleScriptLangGuide/conceptual/ASLR_about_handlers.html#//apple_ref/doc/uid/TP40000983-CH206-SW13
