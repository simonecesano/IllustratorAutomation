-- ---------------------------------------------------------------------------
-- script to test the order in which an Excel cell selection gets iterated on 
-- ---------------------------------------------------------------------------

set values_list to {}

tell application "Microsoft Excel"
    log ("" & (count of rows of selection as integer) & " rows")
    log ("" & (count of columns of selection as integer) & " columns")
    log "-------------"

    repeat with j from 1 to count of columns of selection
    	   set d to cell j of row 1 of selection
	   set v to value of d
	   log v
    end repeat
    log "-------------"
    repeat with j from (count of columns of selection as integer + 1) to count large of selection
    	   set d to cell j of selection
	   set v to value of d
	   log v
    end repeat
    set values_list to values_list & v
end tell
