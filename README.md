# mail-merge-tex
Python script to create multiple PDF files from a single template.tex and a csv file.

## Requirements:
(pdflatex) or (latex, dvipdf) or (latex, dvips, ps2pdf) or (xelatex), latexmk

Choice of compiler can be changed by changing the 1st line of the tex file
as follows:

    %! TeX program = pdflatex
    %! TeX program = latex+dvipdf
    %! TeX program = latex+dvips+ps2pdf
    %! TeX program = xelatex

## Compilation:

    mail-merge-tex.py template1.tex candidates.csv -o "<ApplicationID>_<Firstname>_<Lastname>.pdf"
    
where `<ApplicationID>`, `<Firstname>`, and `<Lastname>` are csv headers. Do not forget to enclose the output with quotes.
