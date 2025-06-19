-- AppleScript to convert DOCX files to PDF using Microsoft Word
-- Usage: osascript docx_to_pdf.applescript "/path/to/input.docx" "/path/to/output.pdf"

on run argv
    -- Check if we have the required arguments
    if (count of argv) < 2 then
        display dialog "Usage: osascript docx_to_pdf.applescript input_file output_file" buttons {"OK"} default button "OK"
        return "ERROR: Missing arguments"
    end if
    
    set inputFile to item 1 of argv
    set outputFile to item 2 of argv
    
    try
        -- Launch Microsoft Word if it's not already running
        tell application "Microsoft Word"
            activate
            
            -- Open the DOCX file
            open inputFile
            
            -- Get the active document
            set theDoc to active document
            
            -- Wait a moment for the document to fully load
            delay 0.5
            
            -- Save as PDF
            save as theDoc file name outputFile file format format PDF
            
            -- Close the document without saving changes
            close theDoc saving no
            
        end tell
        
        return "SUCCESS: Converted " & inputFile & " to " & outputFile
        
    on error errMsg number errNum
        -- Error handling
        return "ERROR: " & errMsg
    end try
end run
