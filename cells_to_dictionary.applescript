-- ---------------------------------------------------------------------------
-- script to transform a table into a list of dictionaries
-- ---------------------------------------------------------------------------

use framework "Foundation"

set AppleScript's text item delimiters to {", "}

set all_items to {}
set fields to {}

tell application "Microsoft Excel"
    -- log ("" & (count of rows of selection as integer) & " rows") & ", " & ("" & (count of columns of selection as integer) & " columns")
    -- log "-------------"

    set col_count to count of columns of selection

    repeat with j from 1 to col_count
    	   set d to cell j of row 1 of selection
	   set v to value of d
	   copy v to the end of |fields|
    end repeat
    -- log fields as text
    -- set row_dict to ""
    repeat with j from (col_count + 1) to count large of selection
    	   set d to cell j of selection
	   set v to value of d
	   set i to (1 + (j - 1) mod col_count)
	   set k to item i of fields

	   if i = 1 then
	       -- log "-------------"
	       -- log "create new dict"
	       set row_dict to current application's NSMutableDictionary's alloc()'s init()
	   end if
	   -- log { j, k, v }

	   tell row_dict to setObject: v forKey: k

	   if j mod col_count = 0 then
	       -- log "push dict"
	       copy row_dict to the end of all_items
	   end if
    end repeat
end tell

-- log "-------------"


-- log fields
-- log count of all_items

repeat with dict_item in all_items
   set vs to {}
   repeat with k in fields
        set v to (dict_item's objectForKey:k) as text
	log { k, v }
   end repeat
end repeat


-- 1. copy the first artboard as many times as there are items in array
-- 2. copy every item on the first board and replace the text content of items
--    with content of corresponding dictionary
