-- ---------------------------------------------------------------------------
-- script to create an Illustrator text box for each cell in the
-- Excel selection
-- ---------------------------------------------------------------------------

set values_list to {}

tell application "Microsoft Excel"
    repeat with j from 1 to count large of selection
    	   set c to cell j of selection
	   set v to value of c as string
	   log v
	   set values_list to values_list & v
    end repeat
end tell

log values_list as text

tell application "Adobe Illustrator"
    set adihaus to text font "Adihaus-Regular"
    set bf to text font "Adihaus-Bold"
    set x to 80
    set y to -160
    tell document 1
       repeat with a from 1 to length of values_list
       	   set t to item a of values_list
	   set t_text to make new text frame with properties {contents: t, position: { x, y }}
	   
	   set size of text range of t_text to 16
	   set text font of text range of t_text to adihaus
	   set y to y + 24
	   set name of t_text to t
	   log name of t_text as text
       end repeat
    end tell
end tell

