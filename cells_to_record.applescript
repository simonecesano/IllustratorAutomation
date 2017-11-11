set values_list to {}

tell application "Microsoft Excel"
    repeat with j from 1 to count large of selection
    	   set d to cell j of selection

	   set c to get offset d row offset 0 column offset 1
	   
	   set k to value of d as string
	   set v to value of c
	   set k of values_list to v 
	   log v
	   -- set values_list to values_list & v
    end repeat
end tell
