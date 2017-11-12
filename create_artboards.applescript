-- ---------------------------------------------------------------------------
-- script that creates more artboards of the same size as the first one
-- in the file
-- TODO: make number of artboards variable
-- ---------------------------------------------------------------------------

on run argv
   set AppleScript's text item delimiters to {", "}

   -- get arguments
   if count of argv > 0 then
      set artboard_count to item 1 of argv
   else
      set artboard_count to 12
   end if

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
   
   -- set artboard_v_count to artboard_v_count as integer
   -- log (artboard_count, artboard_h_count, artboard_v_count)
   -- log (artboard_count, artboard_h_count, artboard_v_count as integer)
   -- would be cool if width and height could be switched around when the artboard is portrait 

   -- return
   
   tell application "Adobe Illustrator"
     tell document 1
	  log visible bounds as list
	  log count of artboards
	  log artboard rectangle of artboard 1 as list

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

		  set new_item to duplicate current_item to end with properties { position: { x, y } }

		  if (k <= count of values_list) then set contents of new_item to "" & (item k of values_list)

		  set k to k + 1
		  set i to i + 1
	      end repeat
	      set c to c + 4
	  end repeat
	  

	  -- ---------------------------------------------------------
	  -- END of duplicate artboards
	  -- ---------------------------------------------------------
     end tell
   end tell
end run
