-- ---------------------------------------------------------------------------
-- script that gets data from Excel and then 
-- creates more artboards of the same size as the first one
-- in the file filling text frames that begin with '{{ ' with the
-- corresponding value in the Excel data 
-- TODO: make number of per row variable
-- ---------------------------------------------------------------------------

use framework "Foundation"

on run argv
   set AppleScript's text item delimiters to {", "}

   -- ---------------------------------------------------------
   -- START: intro setup
   -- ---------------------------------------------------------

   
   tell application "Microsoft Excel" to set artboard_count to (count of rows in selection - 1)

   -- if count of argv > 0 then
   --    set artboard_count to item 1 of argv
   -- else
   --    set artboard_count to 12
   -- end if

   -- ---------------------------------------------------------
   -- set width and height
   -- height is square root of count of artboards 
   -- ---------------------------------------------------------
   set artboard_h_count to (artboard_count ^ 0.5) as integer
   set artboard_v_count to artboard_count / artboard_h_count

   -- this is the ceiling calculation
   if artboard_v_count > 0 and artboard_v_count mod 1 is not 0 then
      set artboard_v_count to (artboard_v_count div 1) + 1 
   else 
      set artboard_v_count to (artboard_v_count div 1)
   end if

   -- ---------------------------------------------------------
   -- END: intro setup
   -- ---------------------------------------------------------


   -- ---------------------------------------------------------
   -- START: get data from Excel
   -- ---------------------------------------------------------

   set all_items to {}
   set fields to {}

   tell application "Microsoft Excel"
       set col_count to count of columns of selection
   
       repeat with j from 1 to col_count
       	   set d to cell j of row 1 of selection
   	   set v to "{{ " & (value of d) & " }}"
   	   copy v to the end of |fields|
       end repeat

       repeat with j from (col_count + 1) to count large of selection
       	   set d to cell j of selection
   	   set v to string value of d as text
   	   set i to (1 + (j - 1) mod col_count)
   	   set k to item i of fields
   
   	   if i = 1 then
   	       set row_dict to current application's NSMutableDictionary's alloc()'s init()
   	   end if

	   tell row_dict to setObject: v forKey: k
   
   	   if j mod col_count = 0 then
   	       copy row_dict to the end of all_items
   	   end if
       end repeat
   end tell

   set i to 1
   repeat while i <= count of all_items
      set dict_item to item i of all_items      
      repeat with k in fields
   	   set v to (dict_item's objectForKey:k) as text
   	   log { i, k, v }
      end repeat
      set i to i + 1
   end repeat
   -- return

   -- ---------------------------------------------------------
   -- END: get data from Excel
   -- ---------------------------------------------------------
   
   tell application "Adobe Illustrator"
     tell document 1
	  log count of artboards
	  if count of artboards > 1 then return

	  -- ---------------------------------------------------------
	  -- START of create artboards
	  -- ---------------------------------------------------------
	  set w to width   -- should be of artboard
	  set h to height  -- should be of artboard

	  set y to artboard_v_count
	  repeat while y > 0
	      set x to 1
	      repeat while x <= artboard_h_count and (count of artboards < (1 + artboard_count))
		 set art_rect to { w * (x - 1), h * y, w * x, h * (y - 1) }
		 set n to "artboard_" & (x as integer) & "_" & (y as integer)
		 set ab to make new artboard with properties { artboard rectangle: art_rect }
		 set x to x + 1
		 -- if (count of artboards >= (1 + artboard_count)) then return
	      end repeat
	      set y to y - 1
	  end repeat
	  -- ---------------------------------------------------------
	  -- END of create artboards
	  -- ---------------------------------------------------------
	  
	  -- ---------------------------------------------------------
	  -- START of duplicate artboards
	  -- ---------------------------------------------------------

	  set artboards_count to count of artboards
	  set artboard_rectangles to {}
	  set values_list to {}

	  set c to 2 -- start from second artboard      
	  repeat while c <= artboards_count 
		 set artboard_rectangles to artboard_rectangles & (artboard rectangle of artboard c)
		 set c to c + 1
	  end repeat
	  log count of artboard_rectangles	

	  set base to artboard rectangle of artboard 1
	  log base as text

	  set k to 1 -- cell counter
	  set m to count of page items
	  set t to page items
	  set j to 1

	  -- -------------------------------------
	  -- iterate through artboard_rectangles
	  -- -------------------------------------
	  set c to 1
	  repeat while c <= count of artboard_rectangles
	      set i to 1 -- page item counter 

	      repeat while i <= m
		  set current_item to item i of t
		  set pos to position of current_item
		  set x_off to (item 1 of base) - (item c of artboard_rectangles)
		  set y_off to (item 2 of base) - (item (c + 1) of artboard_rectangles)	      

		  set x to (item 1 of pos) - x_off
		  set y to (item 2 of pos) - y_off

		  if (class of container of current_item is not "group item") then
		     try 
		     	 set new_item to duplicate current_item to end with properties { position: { x, y } }
		     on error e
		     	 log "----------- ERROR ------------"
		     	 log e
		  	 log class of current_item as text
		  	 log class of container of current_item as text
			 log name of current_item as text
		     	 set new_item to duplicate current_item to end with properties { position: { x, y } }
		     end try

		     if ((class of new_item as text) is "text frame") and (contents of new_item begins with "{{ ") and (true) then
			set item_key to (contents of new_item)
			set dict_item to item j of all_items
		     	try
			   tell application "Microsoft Excel" to set v to (dict_item's objectForKey:item_key) as text
			   log { item_key, v }
			   if (v does not equal "missing value") then set contents of new_item to v
		     	on error e
		     	    log e 
		     	end try
		     end if
		  end if
		  
		  set k to k + 1
		  set i to i + 1
	      end repeat
	      set c to c + 4
	      set j to j + 1
	  end repeat

	  -- ---------------------------------------------------------
	  -- END of duplicate artboards
	  -- ---------------------------------------------------------
     end tell
   end tell
end run
