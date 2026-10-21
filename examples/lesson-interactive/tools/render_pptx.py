#!/usr/bin/env python3
"""Approximate PPTX -> PNG renderer (layout proofing only)."""
import sys, os, io, zipfile
from PIL import Image, ImageDraw, ImageFont
from lxml import etree

NS = {
 'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
 'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
 'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A = '{%s}' % NS['a']; P = '{%s}' % NS['p']; R = '{%s}' % NS['r']

FONT_REG = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
FONT_BOLD = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
FONT_OBL = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
_fcache = {}
def font(sz, b=False, i=False):
    k = (round(sz), b, i)
    if k in _fcache: return _fcache[k]
    p = FONT_BOLD if b else (FONT_OBL if i else FONT_REG)
    f = ImageFont.truetype(p, max(6, int(round(sz))))
    _fcache[k] = f
    return f

def emu2px(v, scale): return v * scale

def parse_color(el, ctx):
    """a:solidFill child -> (r,g,b,a) or None"""
    if el is None: return None
    sf = el.find('a:solidFill', NS)
    if sf is None: return None
    return color_from_spec(sf, ctx)

def color_from_spec(sf, ctx):
    srgb = sf.find('a:srgbClr', NS)
    if srgb is not None:
        v = srgb.get('val')
        a = 255
        al = srgb.find('a:alpha', NS)
        if al is not None: a = int(int(al.get('val'))/100000*255)
        return tuple(int(v[i:i+2],16) for i in (0,2,4)) + (a,)
    sch = sf.find('a:schemeClr', NS)
    if sch is not None:
        v = sch.get('val')
        m = ctx.get('theme', {})
        base = m.get(v, (0,0,0))
        lummod = sch.find('a:lumMod', NS); lumoff = sch.find('a:lumOff', NS)
        r,g,b = base[:3]
        import colorsys
        h,l,s = colorsys.rgb_to_hls(r/255,g/255,b/255)
        if lummod is not None: l *= int(lummod.get('val'))/100000
        if lumoff is not None: l += int(lumoff.get('val'))/100000
        r,g,b = colorsys.hls_to_rgb(h, max(0,min(1,l)), s)
        return (int(r*255),int(g*255),int(b*255),255)
    prst = sf.find('a:prstClr', NS)
    if prst is not None:
        return (0,0,0,255)
    return None

def theme_colors(z):
    try: x = z.read('ppt/theme/theme1.xml')
    except KeyError: return {}
    root = etree.fromstring(x)
    scheme = root.find('.//a:clrScheme', NS)
    out = {}
    for child in scheme:
        tag = etree.QName(child).localname
        srgb = child.find('a:srgbClr', NS)
        val = srgb.get('val') if srgb is not None else '000000'
        out[tag] = tuple(int(val[i:i+2],16) for i in (0,2,4))+(255,)
    # aliases
    out['tx1'] = out.get('dk1'); out['bg1'] = out.get('lt1')
    out['tx2'] = out.get('dk2'); out['bg2'] = out.get('lt2')
    return out

def rels_for(z, part):
    d, b = os.path.split(part)
    rp = f"{d}/_rels/{b}.rels"
    out = {}
    if rp in z.namelist():
        root = etree.fromstring(z.read(rp))
        for rel in root:
            tgt = rel.get('Target')
            if tgt.startswith('../'):
                full = os.path.normpath(os.path.join(d, tgt)).replace('\\','/')
            elif tgt.startswith('/'):
                full = tgt[1:]
            else:
                full = os.path.normpath(os.path.join(d, tgt)).replace('\\','/')
            out[rel.get('Id')] = full
    return out

class Ctx(dict): pass

def resolve_image(z, part, rid):
    tgt = rels_for(z, part).get(rid)
    if not tgt or tgt not in z.namelist(): return None
    return tgt

def load_im(z, tgt):
    try:
        return Image.open(io.BytesIO(z.read(tgt))).convert('RGBA')
    except Exception:
        return None

def apply_crop(im, srcRect):
    if srcRect is None: return im
    def f(k): 
        v = srcRect.get(k)
        return int(v)/100000 if v else 0
    l,r,t,b = f('l'),f('r'),f('t'),f('b')
    W,H = im.size
    box = (int(l*W), int(t*H), int(W-(r*W)), int(H-(b*H)))
    if box[2]<=box[0] or box[3]<=box[1]: return im
    return im.crop(box)

def draw_text_block(draw, body, box, ctx, scale, base_font_px=None):
    """Very approximate word-wrap renderer."""
    x0,y0,x1,y1 = box
    # anchor
    bodyPr = body.find('a:bodyPr', NS)
    anchor = 't'
    lIns = rIns = 91440; tIns = bIns = 45720
    if bodyPr is not None:
        anchor = bodyPr.get('anchor','t')
        if bodyPr.get('lIns') is not None: lIns = int(bodyPr.get('lIns'))
        if bodyPr.get('rIns') is not None: rIns = int(bodyPr.get('rIns'))
        if bodyPr.get('tIns') is not None: tIns = int(bodyPr.get('tIns'))
        if bodyPr.get('bIns') is not None: bIns = int(bodyPr.get('bIns'))
    pad_l = lIns*scale  # EMU -> px
    pad_t = tIns*scale
    pad_r = rIns*scale
    pad_b = bIns*scale
    bx0 = x0+pad_l; by0=y0+pad_t; bx1=x1-pad_r; by1=y1-pad_b
    # Build lines: list of (list of (text,font,color), align, lh)
    lines = []
    for p in body.findall('a:p', NS):
        pPr = p.find('a:pPr', NS)
        algn = pPr.get('algn','l') if pPr is not None else 'l'
        lvl = int(pPr.get('lvl',0)) if pPr is not None else 0
        lnSpc = 1.15
        if pPr is not None:
            ls = pPr.find('a:lnSpc/a:spcPct', NS)
            if ls is not None: lnSpc = int(ls.get('val'))/100000
            else:
                ls2 = pPr.find('a:lnSpc/a:spcPts', NS)
                if ls2 is not None: lnSpc = int(ls2.get('val'))/100/20
        runs = []
        for r in p.findall('a:r', NS):
            rPr = r.find('a:rPr', NS)
            t = r.find('a:t', NS)
            txt = t.text if t is not None and t.text else ''
            sz = 1800; b=False; i=False; col=(0,0,0,255); und=False
            if rPr is not None:
                if rPr.get('sz'): sz = int(rPr.get('sz'))
                b = rPr.get('b')=='1'; i = rPr.get('i')=='1'; und = rPr.get('u')=='sng'
                col = parse_color(rPr, ctx) or col
            runs.append({'t':txt,'sz':sz/100,'b':b,'i':i,'c':col,'u':und})
        if not runs:
            epr = p.find('a:endParaRPr', NS)
            sz = int(epr.get('sz'))/100 if (epr is not None and epr.get('sz')) else 18
            runs = [{'t':'','sz':sz,'b':False,'i':False,'c':(0,0,0,255),'u':False}]
        lines.append((runs, algn, lnSpc))
    # layout
    y = by0
    rendered = []
    for runs, algn, lnSpc in lines:
        # split into words with intrinsic fonts
        words = []  # (w, font, color, und)
        maxsz = 0
        for r_ in runs:
            szpx = r_['sz']*scale/72.0*96/72.0  # emu->px for pt: pt*96/72
            szpx = r_['sz']*96/72.0 * (scale/ (914400/96) / (96/72))  # simplify below
            # scale: px per EMU. pt -> EMU = pt*12700. px = EMU*scale
            szpx = r_['sz']*12700*scale
            maxsz = max(maxsz, r_['sz'])
            f = font(szpx, r_['b'], r_['i'])
            for w in r_['t'].replace('\xa0',' ').split(' '):
                words.append((w, f, r_['c'], r_['u'], r_['t'].startswith(' ') is False))
        lineh = maxsz*12700*scale*lnSpc*1.2
        # wrap
        cur=[]; curw=0
        avail = bx1-bx0
        for witem in words:
            wd = draw.textlength(witem[0]+' ', font=witem[1])
            if cur and curw+wd > avail:
                rendered.append((cur, curw, algn, lineh, maxsz))
                cur=[]; curw=0
            cur.append(witem); curw += wd
        rendered.append((cur, curw, algn, lineh, maxsz))
    total_h = sum(r[3] for r in rendered)
    if anchor == 'ctr':
        y = by0 + max(0, (by1 - by0 - total_h) / 2)
    elif anchor == 'b':
        y = by1 - total_h
    for cur, curw, algn, lineh, maxsz in rendered:
        if algn=='ctr': x = bx0 + (avail-curw)/2
        elif algn=='r': x = bx1-curw
        else: x = bx0
        for w, f, c, und, _ in cur:
            draw.text((x, y), w, font=f, fill=c)
            tw = draw.textlength(w, font=f)
            if und: draw.line([(x, y+lineh*0.85),(x+tw,y+lineh*0.85)], fill=c, width=1)
            x += draw.textlength(w+' ', font=f)
        y += lineh
    return y

def xfrm_of(el):
    x = el.find('a:xfrm', NS)
    if x is None:
        # check for p:spPr-ish nested
        x = el.find('.//a:xfrm', NS)
    return x

def get_xfrm(spPr):
    x = spPr.find('a:xfrm', NS) if spPr is not None else None
    if x is None: return None
    off = x.find('a:off', NS); ext = x.find('a:ext', NS)
    if off is None or ext is None: return None
    rot = int(x.get('rot',0))/60000
    return dict(x=int(off.get('x')), y=int(off.get('y')), cx=int(ext.get('cx')), cy=int(ext.get('cy')), rot=rot)

def render_shape(sh, img, draw, ctx, scale, part, z, tx=(0,0,1,1), depth=0):
    tag = etree.QName(sh).localname
    if tag == 'grpSp':
        gsp = sh.find('p:grpSpPr', NS)
        x = get_xfrm(gsp)
        if x is None: 
            for c in sh: 
                if etree.QName(c).localname in ('sp','pic','grpSp','cxnSp'): render_shape(c, img, draw, ctx, scale, part, z, tx, depth+1)
            return
        chOff = gsp.find('a:xfrm/a:chOff', NS); chExt = gsp.find('a:xfrm/a:chExt', NS)
        ox = int(chOff.get('x')) if chOff is not None else 0
        oy = int(chOff.get('y')) if chOff is not None else 0
        cx = int(chExt.get('cx')) if chExt is not None else x['cx']
        cy = int(chExt.get('cy')) if chExt is not None else x['cy']
        sx = x['cx']/cx if cx else 1; sy = x['cy']/cy if cy else 1
        # compose with parent tx: tx maps child-emu -> parent-emu
        ptx, pty, psx, psy = tx
        def child_map(px, py):
            X = ox + (px - ox)  # coords in child space
            return ((x['x'] + (px-ox)*sx) * psx + ptx, (x['y'] + (py-oy)*sy) * psy + pty)
        ntx = (x['x']*psx+ptx - ox*sx*psx, x['y']*psy+pty - oy*sy*psy, sx*psx, sy*psy)
        for c in sh:
            if etree.QName(c).localname in ('sp','pic','grpSp','cxnSp'):
                render_shape(c, img, draw, ctx, scale, part, z, ntx, depth+1)
        return
    ptx, pty, psx, psy = tx
    spPr = sh.find('p:spPr', NS) or sh.find('p:grpSpPr', NS)
    x = get_xfrm(spPr) if spPr is not None else None
    if x is None: return
    if abs(x['cx']*psx) > 1e8 or abs(x['cy']*psy) > 1e8: return
    X0 = (x['x']*psx + ptx)*scale; Y0 = (x['y']*psy + pty)*scale
    W = x['cx']*psx*scale; H = x['cy']*psy*scale
    box = [X0, Y0, X0+W, Y0+H]

    if tag == 'pic':
        blip = sh.find('.//a:blip', NS)
        if blip is not None:
            rid = blip.get(R+'embed')
            tgt = resolve_image(z, part, rid) if rid else None
            if tgt:
                im = load_im(z, tgt)
                if im:
                    im = apply_crop(im, sh.find('.//a:srcRect', NS))
                    if W>0 and H>0 and W<4000 and H<4000:
                        im2 = im.resize((max(1,int(W)), max(1,int(H))), Image.LANCZOS)
                        # duotone skip
                        img.alpha_composite(im2, (int(X0), int(Y0)))
        return
    if tag == 'cxnSp':
        # line
        ln = spPr.find('a:ln', NS) if spPr is not None else None
        col = (0,0,0,255); wd = 1
        if ln is not None:
            col = parse_color(ln, ctx) or col
            wd = max(1,int(int(ln.get('w',12700))/12700*scale*96/72))
        flipH = spPr.find('a:xfrm', NS).get('flipH')=='1' if spPr.find('a:xfrm', NS) is not None else False
        draw.line([(X0,Y0),(X0+W,Y0+H)], fill=col, width=wd)
        return
    # sp
    geom = spPr.find('a:prstGeom', NS) if spPr is not None else None
    prst = geom.get('prst') if geom is not None else 'rect'
    fill = parse_color(spPr, ctx) if spPr is not None else None
    ln = spPr.find('a:ln', NS) if spPr is not None else None
    lcol = parse_color(ln, ctx) if ln is not None else None
    lw = max(1, int(int(ln.get('w',12700))/12700*1.2)) if ln is not None and ln.get('w') else 1
    if prst in ('ellipse','oval'):
        if fill: draw.ellipse(box, fill=fill, outline=lcol, width=lw if lcol else 0)
        elif lcol: draw.ellipse(box, outline=lcol, width=lw)
    elif prst in ('roundRect','round1Rect','round2SameRect'):
        if fill: draw.rounded_rectangle(box, radius=min(W,H)*0.12, fill=fill, outline=lcol, width=lw if lcol else 0)
        elif lcol: draw.rounded_rectangle(box, radius=min(W,H)*0.12, outline=lcol, width=lw)
    else:
        if fill: draw.rectangle(box, fill=fill, outline=lcol, width=lw if lcol else 0)
        elif lcol: draw.rectangle(box, outline=lcol, width=lw)
    body = sh.find('p:txBody', NS)
    if body is not None:
        draw_text_block(draw, body, box, ctx, scale)

def render(path, outdir, scale=0.09, only=None):
    z = zipfile.ZipFile(path)
    pres = etree.fromstring(z.read('ppt/presentation.xml'))
    sldSz = pres.find('p:sldSz', NS)
    W = int(sldSz.get('cx')); H = int(sldSz.get('cy'))
    ctx = ctxif = Ctx(); ctx['theme'] = theme_colors(z)
    order = []
    for rel in etree.fromstring(z.read('ppt/_rels/presentation.xml.rels')):
        order.append((rel.get('Id'), rel.get('Target')))
    rid_to_slide = {rid: os.path.normpath(os.path.join('ppt', t)).replace('\\','/') for rid,t in order if t.startswith('slides/')}
    slide_parts = []
    for sldId in pres.find('p:sldIdLst', NS):
        rid = sldId.get(R+'id')
        if rid in rid_to_slide: slide_parts.append(rid_to_slide[rid])
    os.makedirs(outdir, exist_ok=True)
    outs = []
    for i, part in enumerate(slide_parts, 1):
        if only and i not in only: continue
        root = etree.fromstring(z.read(part))
        bgel = root.find('p:cSld/p:bg/p:bgPr/a:solidFill', NS)
        bgcol = color_from_spec(bgel, ctx) if bgel is not None else (255,255,255,255)
        img = Image.new('RGBA', (int(W*scale), int(H*scale)), bgcol)
        draw = ImageDraw.Draw(img)
        tree = root.find('p:cSld/p:spTree', NS)
        for c in tree:
            if etree.QName(c).localname in ('sp','pic','grpSp','cxnSp'):
                render_shape(c, img, draw, ctx, scale, part, z)
        out = os.path.join(outdir, f'slide{i:02d}.png')
        img.convert('RGB').save(out)
        outs.append(out)
    return outs

if __name__ == '__main__':
    p = sys.argv[1]; o = sys.argv[2]
    only = set(int(x) for x in sys.argv[3].split(',')) if len(sys.argv)>3 else None
    sc = float(sys.argv[4]) if len(sys.argv)>4 else 0.09
    for f in render(p, o, sc, only): print(f)
