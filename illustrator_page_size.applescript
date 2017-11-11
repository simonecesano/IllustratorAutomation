set style_count to 6
set max_width to 4
set pages_per_style to 2

set actual_height to pages_per_style * (style_count / max_width as integer)
if style_count > max_width then
   set actual_width to max_width
else 
   set actual_width to style_count
end if

tell application "Adobe Illustrator"
  set t to text font "Adihaus-Regular"
  set bf to text font "Adihaus-Bold"
  tell document 1
       set w to width
       set h to height
       log w
       log h
       -- set width to w * actual_width
       -- set height to 
       set artboard rectangle of artboard 1 to { 0.0, 0.0, w * actual_width, -h * actual_height }
  end tell
  log actual_width
  log actual_height	
end tell
