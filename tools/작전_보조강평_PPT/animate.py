"""Add slide transitions and click-step entrance animations to a pptxgenjs deck.

Shapes whose name starts with a<step><f|w>_ appear together on click <step>
(f = fade, w = wipe from left). Slides using the DARK_TITLE layout get a push transition,
all others a fade.
"""
import re
import sys
import zipfile
from collections import defaultdict

src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
items = {n: zin.read(n) for n in zin.namelist()}


def layout_of(slide_name):
    rel = items[slide_name.replace('slides/', 'slides/_rels/') + '.rels'].decode()
    m = re.search(r'slideLayouts/(slideLayout\d+)\.xml', rel)
    lay = items['ppt/slideLayouts/%s.xml' % m.group(1)].decode()
    return re.search(r'<p:cSld name="([^"]*)"', lay).group(1)


class Ids:
    def __init__(self):
        self.n = 2

    def __call__(self):
        self.n += 1
        return self.n


def effect(ids, spid, kind, node):
    if kind == 'w':
        preset, sub, filt, dur = 22, 8, 'wipe(left)', 600
    else:
        preset, sub, filt, dur = 10, 0, 'fade', 500
    return (f'<p:par><p:cTn id="{ids()}" presetID="{preset}" presetClass="entr" presetSubtype="{sub}" fill="hold" grpId="0" nodeType="{node}">'
            f'<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            f'<p:set><p:cBhvr><p:cTn id="{ids()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:to><p:strVal val="visible"/></p:to></p:set>'
            f'<p:animEffect transition="in" filter="{filt}"><p:cBhvr><p:cTn id="{ids()}" dur="{dur}"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
            f'</p:childTnLst></p:cTn></p:par>')


report = []
for name in sorted(n for n in items if re.match(r'ppt/slides/slide\d+\.xml$', n)):
    xml = items[name].decode()
    groups = defaultdict(list)
    for m in re.finditer(r'<p:cNvPr id="(\d+)" name="a(\d+)([fw])_[^"]*"', xml):
        groups[int(m.group(2))].append((m.group(1), m.group(3)))
    lay = layout_of(name)
    trans = ('<p:transition spd="slow"><p:push dir="u"/></p:transition>' if lay == 'DARK_TITLE'
             else '<p:transition spd="med"><p:fade/></p:transition>')
    timing = ''
    if groups:
        ids = Ids()
        steps = ''
        bld = ''
        for st in sorted(groups):
            inner = ''
            for k, (spid, kind) in enumerate(groups[st]):
                inner += effect(ids, spid, kind, 'clickEffect' if k == 0 else 'withEffect')
                # pictures cannot take bldP; only shapes with text bodies/geometry
                bld += f'<p:bldP spid="{spid}" grpId="0" animBg="1"/>'
            steps += (f'<p:par><p:cTn id="{ids()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                      f'<p:par><p:cTn id="{ids()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                      f'{inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>')
        # picture shapes must not appear in bldLst
        pics = set(re.findall(r'<p:pic>.*?<p:cNvPr id="(\d+)"', xml, re.S))
        bld = ''.join(b for b in re.findall(r'<p:bldP [^>]*/>', bld) if re.search(r'spid="(\d+)"', b).group(1) not in pics)
        timing = ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
                  '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
                  f'{steps}</p:childTnLst></p:cTn>'
                  '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
                  '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
                  '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>'
                  + (f'<p:bldLst>{bld}</p:bldLst>' if bld else '') + '</p:timing>')
    xml = re.sub(r'<p:transition.*?</p:transition>|<p:transition[^>]*/>', '', xml, flags=re.S)
    xml = re.sub(r'<p:timing>.*?</p:timing>', '', xml, flags=re.S)
    assert '</p:clrMapOvr>' in xml, name
    xml = xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + trans + timing, 1)
    items[name] = xml.encode()
    report.append((name.split('/')[-1], lay, len(groups), sum(len(v) for v in groups.values())))

with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
    for n in zin.namelist():
        z.writestr(zin.getinfo(n), items[n])
for r in sorted(report, key=lambda r: int(re.search(r'\d+', r[0]).group())):
    print(r)
