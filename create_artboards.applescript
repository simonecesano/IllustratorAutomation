-- ---------------------------------------------------------------------------
-- script that creates more artboards of the same size as the first one
-- in the file
-- TODO: make number of artboards variable
-- ---------------------------------------------------------------------------

tell application "Adobe Illustrator"
  tell document 1
       log visible bounds as list
       log count of artboards
       log artboard rectangle of artboard 1 as list

       set w to width
       set h to height

       set y to 4
       repeat while y > 0
       	   set x to 1
       	   repeat while x <= 3
	      set art_rect to { w * (x - 1), h * y, w * x, h * (y - 1) }
	      set n to "artboard_" & (x as integer) & "_" & (y as integer)
              set ab to make new artboard with properties { artboard rectangle: art_rect }
	      set x to x + 1
	   end repeat
	   set y to y - 1
       end repeat
       set artboards_count to count of artboards
       set c to 1       
       -- delete artboard 1
       set AppleScript's text item delimiters to {", "}
       repeat while c <= artboards_count 
       	      log index of artboard c as text
	      log artboard rectangle of artboard c as text
       	      set c to c + 1
       end repeat

       log count of artboards
   end tell
end tell

