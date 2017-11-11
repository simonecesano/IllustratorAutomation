set values_list to {}

tell application "Microsoft Excel"
    repeat with j from 1 to count large of selection
    	   set d to cell j of selection
	   set v to value of d
	   log v
	   set values_list to values_list & v
    end repeat
end tell

log values_list as text

tell application "Adobe Illustrator"
  tell document 1
       set artboards_count to count of artboards
       set AppleScript's text item delimiters to {", "}

       set items_list to {}

       log count of page items
       log items_list

       set artboard_rectangles to {}
       set c to 2       

       repeat while c <= artboards_count 
	      set artboard_rectangles to artboard_rectangles & (artboard rectangle of artboard c)
       	      set c to c + 1
       end repeat
       log count of artboard_rectangles	

       set base to artboard rectangle of artboard 1
       log base as text

       set i to 1 -- page item counter 
       set k to 1 -- cell counter

       set m to count of page items

       set t to page items

       repeat while i <= m
           set current_item to item i of t

	   set c to 1
	   log "page items"
	   log count of page items
	   -- log "index " & (index of i as text)

	   repeat while c <= count of artboard_rectangles

	      set x_off to (item 1 of base) - (item c of artboard_rectangles)
	      set y_off to (item 2 of base) - (item (c + 1) of artboard_rectangles)	      

	      set pos to position of current_item
	      set x to (item 1 of pos) - x_off
	      set y to (item 2 of pos) - y_off
	      
	      log pos as text
	      set new_item to duplicate current_item to end with properties { position: { x, y } }

	      set contents of new_item to "" & (item k of values_list)
	      set k to k + 1

       	      set c to c + 4

	      log "-------"
	      log "page items"
	      log count of page items
	      log "-------"
	      if (count of page items > (4 * 48)) then
	      	 return
	      end if 
       	   end repeat
	   log "######"
	   set i to i + 1
       end repeat

       log count of artboards
       log count of artboard_rectangles	
   end tell
end tell

