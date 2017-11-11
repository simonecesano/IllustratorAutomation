set values_list to {}

tell application "Microsoft Excel"
    repeat with j from 1 to count large of selection
    	   set d to cell j of selection

	   set c to get offset d row offset 0 column offset -1
	   set v to value of c as string

	   set orig_val to value of d 
	   set value of d to v
	   replace d what " " replacement "_"

	   set name of d to value of d as text

	   log v
	   set values_list to values_list & v
    end repeat
end tell
