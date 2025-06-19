-- Batch AppleScript to convert multiple DOCX files to PDF using Microsoft Word
-- This approach opens Word once and processes all files, which should reduce permission prompts

on run argv
    -- Check if we have arguments
    if (count of argv) < 1 then
        display dialog "Usage: osascript batch_docx_to_pdf.applescript /path/to/directory" buttons {"OK"} default button "OK"
        return "ERROR: Missing directory argument"
    end if
    
    set inputDir to item 1 of argv
    
    try
        -- Get list of DOCX files using shell command (more reliable)
        set shellCommand to "find " & quoted form of inputDir & " -name '*.docx' -type f"
        set docxFilesList to do shell script shellCommand
        
        if docxFilesList = "" then
            return "ERROR: No DOCX files found in directory"
        end if
        
        -- Split the file list
        set docxFiles to paragraphs of docxFilesList
        set fileCount to count of docxFiles
        
        -- Launch Microsoft Word once
        tell application "Microsoft Word"
            activate
            
            -- Process each file
            repeat with i from 1 to fileCount
                try
                    set docxFilePath to item i of docxFiles
                    
                    -- Use shell commands for file name manipulation (more reliable)
                    set fileName to do shell script "basename " & quoted form of docxFilePath
                    set baseNameCmd to "basename " & quoted form of docxFilePath & " .docx"
                    set baseName to do shell script baseNameCmd
                    set pdfPath to (inputDir & "/" & baseName & ".pdf")
                    
                    -- Open the DOCX file
                    open docxFilePath
                    
                    -- Get the active document
                    set theDoc to active document
                    
                    -- Wait a moment for the document to load
                    delay 0.3
                    
                    -- Save as PDF
                    save as theDoc file name pdfPath file format format PDF
                    
                    -- Close the document
                    close theDoc saving no
                    
                    -- Log success
                    log "Converted: " & fileName & " to " & baseName & ".pdf"
                    
                on error errMsg
                    log "Error converting file " & i & ": " & errMsg
                end try
            end repeat
            
        end tell
        
        return "SUCCESS: Processed " & fileCount & " files"
        
    on error errMsg number errNum
        return "ERROR: " & errMsg
    end try
end run
