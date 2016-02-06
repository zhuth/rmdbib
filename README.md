# rmdbib
R Markdown with customizable support for bibliography, and fix problems for East-Asian language support

Usage:

`` cat sample.md | python bib2md.py | pandoc -f markdown --reference-docx=template.docx -o sample.docx; python bib2md.py sample.docx ``