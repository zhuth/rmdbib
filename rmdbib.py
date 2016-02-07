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
    
    import lxml.etree as ET
    import copy
    
    nss = {
        'w':"http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        'm':"http://schemas.openxmlformats.org/officeDocument/2006/math",
        'r':"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        'o':"urn:schemas-microsoft-com:office:office",
        'v':"urn:schemas-microsoft-com:vml",
        'w10':"urn:schemas-microsoft-com:office:word",
        'a':"http://schemas.openxmlformats.org/drawingml/2006/main",
        'pic':"http://schemas.openxmlformats.org/drawingml/2006/picture",
        'wp':"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    }

    doc = zf['word/document.xml'].decode('utf-8')
    #doc = re.sub(ur'([…·“”‘’])', u'</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="majorEastAsia" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorEastAsia"/></w:rPr><w:t>\\1</w:t></w:r><w:r><w:t>', doc)
    #doc = re.sub(ur'<w:p>(<w:pPr><w:pStyle w:val="(FirstParagraph|BodyText)" />)', u'<w:p>\\1<w:ind w:firstLineChars="200" w:firstLine="420"/>', doc)
    doc = doc.encode('utf-8')

    #doc = zf['word/document.xml']
    xt = ET.fromstring(doc)
    
    def _deal_p(_p):
        ps = _p.xpath('.//w:pStyle', namespaces=nss)
        if len(ps) == 0 or ps[0].attrib.get('{%s}val' % nss['w'], 'BodyText') in ['BodyText', 'FirstParagraph', '']:
            ppr = _p.xpath('./w:pPr', namespaces=nss)
            if len(ppr) == 0:
                ppr = ET.Element('{%s}pPr' % nss['w'], nsmap=nss)
                _p.insert(0, ppr)
            else:
                ppr = ppr[0]
                
            ppr.append(ET.Element('{%s}ind' % nss['w'], {
                '{%s}firstLineChars' % nss['w']: "200",
                '{%s}firstLine'      % nss['w']: "420"
            }, nsmap=nss))
    
    def _deal_t(_t):
        if not _t.text:
            return None
        m = puncs.search(_t.text)
        if not m:
            return None
        _tp = _t.getparent()
        nwr = copy.deepcopy(_tp)
        nwr2 = copy.deepcopy(_tp)
        nwr.xpath('./w:t', namespaces=nss)[0].text = m.group(0)
        nwrrpr = nwr.xpath('./w:rPr', namespaces=nss)
        if len(nwrrpr) == 0:
            nwrrpr = ET.Element('{%s}rPr' % nss['w'], nsmap=nss)
            nwr.insert(0, nwrrpr)
        else:
            nwrrpr = nwrrpr[0]
        nwrrpr.append(ET.Element('{%s}rFonts' % nss['w'], {
            '{%s}asciiTheme'    % nss['w']:"majorEastAsia", 
            '{%s}eastAsiaTheme' % nss['w']:"majorEastAsia", 
            '{%s}hAnsiTheme'    % nss['w']:"majorEastAsia"}, nsmap=nss))
        nwr2.xpath('./w:t', namespaces=nss)[0].text = _t.text[m.end():]
        _t.text = _t.text[:m.start()]
        _tpp = _tp.getparent()
        _tpp.insert(_tpp.index(_tp)+1, nwr)
        _tpp.insert(_tpp.index(_tp)+2, nwr2)
        return nwr2.xpath('./w:t', namespaces=nss)[0]
    
    # deal with indent
    for _p in xt.xpath('//w:p', namespaces=nss):
        _deal_p(_p)
    
    # deal with quotation marks
    puncs = re.compile(ur'[…·“”‘’]')
    for __t in xt.xpath('//w:t', namespaces=nss):
        _t = __t
        while _t is not None:
            _t = _deal_t(_t)
    
    doc = '<?xml version="1.0" encoding="UTF-8"?>' + ET.tostring(xt, encoding=ET.ElementTree(xt).docinfo.encoding)
    #print doc; return
    
    zf['word/document.xml'] = doc
    
    z = zipfile.ZipFile(docx, 'w')
    for _ in zf:
        z.writestr(_, zf[_], compress_type=zipfile.ZIP_DEFLATED)
    z.close()

if __name__ == '__main__':
    if len(sys.argv) == 1:
        bib2md(sys.stdin, sys.stdout)
    else:
        docx_deal(sys.argv[1])
