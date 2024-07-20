nano orchestrate_pdf_processing.sh


#!/bin/bash

LOG_FILE="/Users/orkravitz/dev/orchestrate_pdf_processing.log"

{
    echo "$(date) - Starting word files to PDF script"
    /Library/Frameworks/Python.framework/Versions/3.10/bin/python3 /Users/orkravitz/dev/wordFilesToPdf/wordFilesToPdfFromDrive.py
    echo "$(date) - Completed word files to PDF script"

    echo "$(date) - Starting split PDF script"
    /Library/Frameworks/Python.framework/Versions/3.10/bin/python3 /Users/orkravitz/dev/Pdf_Split_Merge_Encryp/PDF_Split_and_Merge/SplitPdfFiles.py
    echo "$(date) - Completed split PDF script"

    echo "$(date) - Starting encrypt PDF script"
    /usr/bin/java -cp /Users/orkravitz/dev/Pdf_Encrypt_Watermarks/PDF_Watermarks/lib/Spire.Pdf.jar /Users/orkravitz/dev/Pdf_Encrypt_Watermarks/PDF_Watermarks/src/EncryptPDF.java
    echo "$(date) - Completed encrypt PDF script"

    echo "$(date) - Starting merge PDF script"
    /Library/Frameworks/Python.framework/Versions/3.10/bin/python3 /Users/orkravitz/dev/Pdf_Split_Merge_Encryp/PDF_Split_and_Merge/MergePdfFiles_Encrypt.py
    echo "$(date) - Completed merge PDF script"


    echo "$(date) - Completed all tasks"
} >> "$LOG_FILE" 2>&1
