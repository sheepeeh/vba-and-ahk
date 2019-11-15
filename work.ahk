; Text-only paste (strips all formatting)
#v::                       
   Clip0 = %ClipBoardAll%
   ClipBoard = %ClipBoard% 
   Send ^v                 
   Sleep 150                
   ClipBoard = %Clip0%     
   VarSetCapacity(Clip0, 0)
Return

; Join lines when pasting (useful for pasting text from PDFs)
#j::                                                              
   StringReplace, ClipBoard, ClipBoard, `r`n, %a_space%, All 
   ClipWait
   Send ^v                                                     
Return

; Replace backslashes with forward slashes
#z::
	Clipboard := StrReplace(Clipboard, "\", "/")
	ClipWait
	Send ^v  
Return

::betterps1::PS1="[\[\e[38;5;239m\]\T\[\e[m\]-\u\[\e[01\e[38;5;22m\]@\[\e[m\]\h:\[\e[38;5;100m\]\w\[\e[m\]]\\$ "

::shrugem::¯\_(ツ)_/¯

::knoci::keep new one, copy indexing from

::knofi::keep new one (final)