#!env python
#coding: utf-8

import sys, re
import bibtexparser

def bib2md(input, output):
    templates = {
        'default': u"{author}. *{title}*. {publisher}, {year}.",
        'article': u"{author}. {title}. *{journal}*, {year} ({number}): {pages}.",
        'default_zh': u"{author}：《{title}》，{publisher}，{year}。",
        'article_zh': u"{author}：{title}，《{journal}》，{year}（{number}）：{pages}。",
    }
    
    citation_pattern = re.compile(ur'\:cite\:\`(.+?)\`(:\(.*?\))?')
    template_pattern = re.compile(ur'\{([a-zA-Z]+)\}')
    zh_pattern = re.compile(u'[\u4e00-\u9fa5]+')
    
    
    fulltext = u''
    lines = [_.decode('utf-8') for _ in input.readlines()]
    
    bib = u''
    inbib = False
    intemp = False
    
    for _ in lines:
        _s = _.strip()
        if _s == u'<!--@bibtex':
            inbib = True
        elif _s == u'<!--@templates':
            intemp = True
        elif _s == u'@-->':
            inbib = False
            intemp = False
            continue

        if inbib:
            bib += _ + u'\n'
        elif intemp:
            kv = _s.split('\t')
            if len(kv) == 2:
                templates[kv[0]] = kv[1]
        else:
            fulltext += _
        
    b = bibtexparser.loads(bib)
    entries = {}
    for _ in b.entries:
        if 'pages' in _:
            _['pages'] = _['pages'].replace(u'--', u'–')
        _['ZH'] = True if zh_pattern.search(_['title']) else False
        entries[_['ID'].lower()] = _
    
    citeNum = 0
    citeNotes = ''
    last = 0
    
    for _ in citation_pattern.finditer(fulltext):
        output.write(fulltext[last:_.start()].encode('utf-8'))
        entryKey = _.group(1)
        noteplus = _.group(2)
        if noteplus and noteplus != u'':
            noteplus = u' ' + noteplus[2:-1]
        else:
            noteplus = u''
        entry = entries[entryKey]
            
        citeId = u'^cite%d' % citeNum
        citeNum += 1
        output.write(u'[%s] ' % citeId)
        zh = '_zh' if entry['ZH'] else ''
        note = templates.get(entry['ENTRYTYPE'] + zh, templates['default' + zh])
        note = template_pattern.sub(lambda mo: entry.get(mo.group(1), u''), note) + noteplus
        last = _.end()
        citeNotes += u'\n[%s]: %s' % (citeId, note)
        
    output.write(fulltext[last:].encode('utf-8'))
    output.write(citeNotes.encode('utf-8'))
    
def docx_deal(docx):
    import zipfile
    z = zipfile.ZipFile(docx, 'r')
    zf = {}
    for _ in z.namelist():
        zf[_] = z.read(_)
    z.close()

    doc = zf['word/document.xml'].decode('utf-8')
    doc = re.sub(ur'([…·“”‘’])', u'</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia"/></w:rPr><w:t>\\1</w:t></w:r><w:r><w:t>', doc)
    doc = re.sub(ur'<w:p>(<w:pPr><w:pStyle w:val="(FirstParagraph|BodyText)" />)', u'<w:p>\\1<w:ind w:firstLineChars="200" w:firstLine="420"/>', doc)
    # print doc
    zf['word/document.xml'] = doc.encode('utf-8')
    
    z = zipfile.ZipFile(docx, 'w')
    for _ in zf:
        z.writestr(_, zf[_], compress_type=zipfile.ZIP_DEFLATED)
    z.close()

if __name__ == '__main__':
    if len(sys.argv) == 1:
        bib2md(sys.stdin, sys.stdout)
    else:
        docx_deal(sys.argv[1])
