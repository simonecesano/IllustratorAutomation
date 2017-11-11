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
       	      -- log index of artboard c as text
	      -- log artboard rectangle of artboard c as text
	      -- log artboard rectangle of artboard 1 as text
	      set artboard_rectangles to artboard_rectangles & (artboard rectangle of artboard c)
	      -- set n to name of artboard c
	      -- log n as text
       	      set c to c + 1
       end repeat
       log count of artboard_rectangles	

       set base to artboard rectangle of artboard 1
       log base as text
       set i to 1
       set m to count of page items

       set t to page items

       repeat while i <= m
           set current_item to item i of t

	   set c to 1
	   log "page items"
	   log count of page items
	   -- log "index " & (index of i as text)

	   repeat while c <= count of artboard_rectangles
       	      -- log c

	      set x_off to (item 1 of base) - (item c of artboard_rectangles)
	      set y_off to (item 2 of base) - (item (c + 1) of artboard_rectangles)	      

	      log x_off
	      log y_off

	      set pos to position of current_item
	      set x to (item 1 of pos) - x_off
	      set y to (item 2 of pos) - y_off
	      
	      log pos as text
	      duplicate current_item to end with properties { position: { x, y } }

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

