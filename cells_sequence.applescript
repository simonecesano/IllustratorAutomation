-- ---------------------------------------------------------------------------
-- script to test the order in which an Excel cell selection gets iterated on 
-- ---------------------------------------------------------------------------

set values_list to {}

tell application "Microsoft Excel"
     log count of rows of selection
    repeat with j from 1 to count large of selection
    	   set d to cell j of selection
	   set v to value of d
	   log v
    end repeat
    set values_list to values_list & v
end tell
