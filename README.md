# rmdbib
R Markdown with customizable support for bibliography, and fix problems for East-Asian language support

Usage:

`` cat sample.md | python rmdbib.py | pandoc -f markdown --reference-docx=template.docx -o sample.docx; python rmdbib.py sample.docx ``