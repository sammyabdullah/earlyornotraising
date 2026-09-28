# Daily founder-notes Excel

When the user uploads an Excel file, run:

    pip install openpyxl  # if missing
    python3 fill_recall.py <uploaded.xlsx> <output.xlsx>

Put the output in the primary working directory or scratchpad, send it back to the user with
SendUserFile, and reply with the script's one-line summary (rows filled, rows flagged).
The rules (date parsing, keyword precedence, sentence templates, REVIEW flags) live in
`fill_recall.py`; change them there, not ad hoc.
