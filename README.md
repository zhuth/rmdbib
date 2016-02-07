# rmdbib
Customizable Sphinx-styled citations for R Markdown.

## Usage

BibTex database content should be put line by line between `<!--@bibtex` and `@-->`.

To use customizable templates, please put templates between `<!--@templates` and `@-->`.

For example:

	<!--@bibtex
	@book{bapat1997nonnegative,
	author = {Bapat, Ravi B and Raghavan, Tirukkannamangai ES},
	volume = {64},
	title = {Nonnegative matrices and applications},
	publisher = {Cambridge University Press},
	year = {1997}}
	@article{hinton2006reducing,
	author = {Hinton, Geoffrey E and Salakhutdinov, Ruslan R},
	number = {5786},
	volume = {313},
	journal = {Science},
	pages = {504--507},
	title = {Reducing the dimensionality of data with neural networks},
	year = {2006},
	publisher = {American Association for the Advancement of Science}}
	@-->

	<!--@templates
		default	{author}. *{title}*. {publisher}, {year}.
		article	{author}. {title}. *{journal}*, {year} ({number}): {pages}.
		default_zh	{author}：《{title}》，{publisher}，{year}。
		article_zh	{author}：{title}，《{journal}》，{year}（{number}）：{pages}。
	@-->

To cite with footnote, please use: ``:cite:`BibTexKey` `` or ``:cite:`BibTexKey`:(extra text appending to the footnote.)``. Entry type and its template text should be separated by a tab (\t) in `templates`. 若需对中文参考文献（根据标题判定）使用不同的引文格式，请在引用项类型名称之后加上 `_zh` 。

Shell usage (with pandoc):

	cat sample.md | python rmdbib.py | pandoc -f markdown --reference-docx=template.docx -o sample.docx; python rmdbib.py sample.docx

## 解决 Word 2013/2016 中文引号（“”、‘’）、中点（·）和省略号（……）的字体问题

Word 2013/2016 将上述中文标点符号使用西文字体渲染。本脚本可以解决这一问题。用法：

	python rmdbib.py sample.docx

