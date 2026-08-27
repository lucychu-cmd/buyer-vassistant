#!/usr/bin/env python3
"""
Lovers & Friends — Merch Board builder (LF Femme).

ONE deterministic script. Run it instead of re-deriving the layout from prose, so the
output always matches the approved Cowork result.

Usage:
    python merch_board.py "Buy_Meeting_Prep_List (LF Nov).xls" "2026 Nov"
        arg1 = source .xls (must have the 'Line Guide View' sheet with embedded CAD images)
        arg2 = cover period label shown on the cover (e.g. "2026 Nov", "2026 Q2"); confirm with user
        arg3 = (optional) output .pptx path; default "Lovers and Friends MERCH BOARD <period>.pptx"

Deps: python-pptx, openpyxl, pillow, and LibreOffice (`soffice`) for .xls->.xlsx.
The month divider(s) always show the actual R.Month from the data; the cover shows <period>.

Hard rules baked in (do NOT change without updating the context doc):
  * Canvas 10 x 5.625 in.  Category mixed-case Playfair 25 #666666 top-left.
    Brand top-right "Lovers and Friends" Playfair 20pt #3A3A3A, one row.
  * Fabric headers Playfair 5.5 BLACK; card text Lato 4 BLACK, regular (NOT bold), centered.
  * ALL text boxes wrap="square" (word_wrap=True) and text is PRE-BROKEN into one paragraph
    per visual line in a SINGLE box -- never wrap="none" (that overflows/overlaps).
  * Images sized by HEIGHT (1.39 full / 1.20 half-body AR>=0.55), width by ratio, sharpened.
"""
import sys, os, re, json, subprocess, shutil, collections
import openpyxl
from PIL import Image, ImageChops, ImageFilter, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BRAND  = os.environ.get("MB_BRAND", "Lovers and Friends")   # e.g. MB_BRAND="NBD"
SRC    = sys.argv[1] if len(sys.argv) > 1 else "source.xls"
PERIOD = sys.argv[2] if len(sys.argv) > 2 else "2026"
OUT    = sys.argv[3] if len(sys.argv) > 3 else f"{BRAND} MERCH BOARD {PERIOD}.pptx"
WORK   = "_mb_work"; os.makedirs(WORK, exist_ok=True)

# ---------------------------------------------------------------- 1. load data
xlsx = os.path.join(WORK, "src.xlsx")
subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", WORK, SRC],
               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
shutil.move(os.path.join(WORK, os.path.splitext(os.path.basename(SRC))[0] + ".xlsx"), xlsx)
wb = openpyxl.load_workbook(xlsx, data_only=True); ws = wb["Line Guide View"]
H = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}
COL = {"status":"Status","stylecode":"Stylecode","name":"Style Name","class":"Class",
       "rmonth":"R.Month","lead":"SKU LeadTime","vendor":"Vendor Name","price":"Retail Price",
       "notes":"Design Notes","buyernote":"Buyer Note","material":"Material Name"}
rows = []
for r in range(2, ws.max_row + 1):
    d = {k: ws.cell(r, H[v]).value for k, v in COL.items() if v in H}
    if not d.get("name") and not d.get("stylecode"): continue
    d["row"] = r; rows.append(d)

# ---------------------------------------------------------------- 2. extract CAD images
import zipfile
zf = zipfile.ZipFile(xlsx)
draw = zf.read("xl/drawings/drawing1.xml").decode("utf-8")
rels = zf.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
rid2media = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="\.\./media/([^"]+)"', rels))
imgdir = os.path.join(WORK, "img"); os.makedirs(imgdir, exist_ok=True)
row2img = {}
for a in re.findall(r"<xdr:(?:twoCellAnchor|oneCellAnchor)\b.*?</xdr:(?:twoCellAnchor|oneCellAnchor)>", draw, re.S):
    fr = re.search(r"<xdr:from>.*?<xdr:row>(\d+)</xdr:row>", a, re.S)
    rid = re.search(r'r:embed="(rId\d+)"', a)
    if not (fr and rid and rid.group(1) in rid2media): continue
    media = rid2media[rid.group(1)]; er = int(fr.group(1)) + 1
    dst = os.path.join(imgdir, f"row{er}.{media.split('.')[-1]}")
    with open(dst, "wb") as f: f.write(zf.read("xl/media/" + media))
    row2img[er] = dst

# ---------------------------------------------------------------- 3. derive + categorize
ACC=['handbag','necklace','clutch','tote','earring','bracelet','anklet','belt','bag','jewelry','jewellery','ring','hat','scarf','sunglass','choker','headband','hair','glove','purse','wallet','keychain']
# Accessories are matched as WHOLE WORDS (optional plural), never bare substrings.
# A substring 'ring' wrongly excluded "ELARA FRINGE MIDI DRESS", and 'belt' would catch
# "belted"; whole-word matching still catches real accessories (ring, belt, bag) while
# letting fringe/belted/jewel-neckline through. (Residual edge: a literal "O-RING" detail
# in a dress name would still match 'ring' — rare enough to accept.)
_ACC_RE=re.compile(r'\b(?:'+'|'.join(re.escape(w) for w in ACC)+r')(?:s|es)?\b', re.I)
def is_accessory(name): return bool(_ACC_RE.search(name or ''))
OUTER=['jacket','coat','blazer','cardigan','bolero','kimono','trench','parka','cape','poncho','duster']
RJ=['romper','jumpsuit','one piece','onesie','playsuit','catsuit']
TOPS=['top','tank','tee','bralette','bodysuit','bustier','cami','corset','shirt','blouse','halter','crop']
BOT=['skirt','skort','short','pant','bottom','legging','trouser','jean','culotte']
def color_of(sc):
    if not sc: return ''
    return re.sub(r'^[A-Z]\d{2}', '', str(sc).split('-')[-1]).strip()
def price_of(v):
    try: f=float(v); return '' if f<=0.01 else f'${int(round(f))}'
    except: return ''
def tier_of(r):
    m=re.search(r'TIER\s*([ABC])', f"{r.get('notes') or ''} {r.get('buyernote') or ''}", re.I)
    return f'TIER {m.group(1).upper()}' if m else ''
def annot_of(cls):
    c=str(cls or '')
    return '' if (c.lower()=='new' or not c) else '*'+c.replace('_',' ').upper()+'*'
def has(name,w): n=(name or '').lower(); return any(x in n for x in w)
def categorize(r):
    sc=str(r['stylecode'] or ''); name=str(r['name'] or ''); vend=str(r['vendor'] or '').upper()
    third=sc[2:3].upper() if len(sc)>=3 else ''; nl=name.lower()
    if third=='X' or 'SOVSK' in vend or 'SOVEREIGN' in vend: return None      # swim
    if is_accessory(name): return None                                         # accessories (whole-word)
    if str(r['status']).upper()=='DROPPED': return 'DROP'
    if str(r['lead']).upper()=='CHASE': return 'Chase'
    if str(r['lead']).upper().replace('-',' ').strip()=='LONG LEAD': return 'Long Lead'
    if 'EEFG' in vend or 'NEWVA' in vend or 'NEW VANTAGE' in vend or 'ELEVEN ELEVEN' in vend: return 'Knits'
    if 'sweater' in nl and 'dress' not in nl and 'gown' not in nl: return 'Knits'
    if has(name,OUTER): return 'Outerwear'
    if has(name,['dress','gown']): return 'Dresses'
    if has(name,RJ): return 'Rompers & Jumpsuits'
    if has(name,TOPS): return 'Tops'
    if has(name,BOT): return 'Bottoms'
    if re.search(r'\b(MAXI|MIDI|MINI)\b', name, re.I): return 'Dresses'
    return 'Tops'
items=[]
for r in rows:
    cat=categorize(r)
    if cat is None: continue
    items.append({'row':r['row'],'name':str(r['name'] or '').strip(),'color':color_of(r['stylecode']),
        'price':price_of(r['price']),'tier':tier_of(r),'fabric':str(r['material'] or '').strip() or 'OTHER',
        'annot':annot_of(r['class']),'recolor':str(r['class']).lower()=='recolor',
        'month':str(r['rmonth'] or '').strip(),'cat':cat,'img':row2img.get(r['row'])})
print('kept', len(items), 'of', len(rows), '|', dict(collections.Counter(x['cat'] for x in items)))

# ---------------------------------------------------------------- 4. crop + sharpen images
imgp=os.path.join(WORK,"imgp"); os.makedirs(imgp, exist_ok=True)
def trim(im,thresh=10):
    im=im.convert('RGB'); diff=ImageChops.difference(im,Image.new('RGB',im.size,(255,255,255)))
    bbox=diff.convert('L').point(lambda p:255 if p>thresh else 0).getbbox()
    return im.crop(bbox) if bbox else im
def enhance(im,scale=4):
    im=im.resize((im.size[0]*scale,im.size[1]*scale), Image.LANCZOS)
    return im.filter(ImageFilter.UnsharpMask(radius=1.2,percent=130,threshold=1))
def n_top_figures(im):
    # count separated figures in the top strip; 2 => real front+back CAD (safe to crop to left).
    # A single figure that merely carries a colour swatch returns 1 -> must NOT be sliced.
    g=im.convert('L'); w,h=g.size
    import numpy as _np
    strip=_np.asarray(g.crop((0,0,w,max(2,int(h*0.16)))))
    on=((strip<235).mean(axis=0))>0.06
    segs=[]; i=0; n=len(on)
    while i<n:
        if on[i]:
            j=i
            while j<n and on[j]: j+=1
            segs.append([i,j]); i=j
        else: i+=1
    merged=[]
    for s in segs:
        if merged and s[0]-merged[-1][1]<0.06*w: merged[-1][1]=s[1]
        else: merged.append(s)
    return len([s for s in merged if (s[1]-s[0])>=0.08*w])
for it in items:
    p=it.get('img')
    if not p or not os.path.exists(p): it['imgp']=None; it['ar_final']=0.4; continue
    im=Image.open(p).convert('RGB'); w,h=im.size; ar=w/h
    twofig = (ar>=0.45 and n_top_figures(im)>=2)        # only slice genuine front+back CADs
    out = trim(im.crop((0,0,w//2,h))) if ((not it['recolor']) and twofig) else trim(im)
    it['ar_final']=out.size[0]/out.size[1]
    dst=os.path.join(imgp,f"r{it['row']}.jpg"); enhance(out).save(dst,quality=95); it['imgp']=dst

# ---------------------------------------------------------------- 5. build deck (10 x 5.625)
BLACK=RGBColor(0,0,0); DKGREY3=RGBColor(0x66,0x66,0x66); TITLE=RGBColor(0x3A,0x3A,0x3A)
SAGE=RGBColor(0xD9,0xE8,0xCD); GREY=RGBColor(0xEC,0xEB,0xE9)
CW,CH,MX=10.0,5.625,0.48; CARD,GIN,GGRP=0.59,0.04,0.30
HDR_H,IMG_H,GAP_IMG,TXT_H,GAP_AFTER=0.28,1.39,0.02,0.45,0.06
HALF_AR,HALF_H,CELL_PAD=0.55,1.20,0.10
GRID_TOP=0.60; ROWPITCH=HDR_H+IMG_H+GAP_IMG+TXT_H+GAP_AFTER; ROWS_PER=2; RIGHT=CW-MX
PLAYFAIR,PLAYFAIR_SB,LATO='Playfair Display','Playfair Display SemiBold','Lato'
SANS=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',1000)
SER =ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf',1000)
def wrap(text,sz,serif,maxw_pt):           # pre-break into lines that fit the cell
    f=SER if serif else SANS; out,cur=[],''
    for w in str(text).split(' '):
        t=w if not cur else cur+' '+w
        if f.getlength(t)/1000.0*sz<=maxw_pt or not cur: cur=t
        else: out.append(cur); cur=w
    if cur: out.append(cur)
    return out
for it in items:
    ar=it.get('ar_final') or 0.4; ih=HALF_H if ar>=HALF_AR else IMG_H
    it['iw']=ih*ar; it['ih']=ih; it['cw']=max(CARD, ih*ar+CELL_PAD)

prs=Presentation(); prs.slide_width=Inches(CW); prs.slide_height=Inches(CH); BLANK=prs.slide_layouts[6]
def bg(slide,color):
    r=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb=color; r.line.fill.background(); r.shadow.inherit=False
    slide.shapes._spTree.remove(r._element); slide.shapes._spTree.insert(2,r._element)
def tbox(slide,x,y,w,h,anchor=MSO_ANCHOR.TOP,wrap_=True):
    tf=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)).text_frame
    tf.word_wrap=wrap_           # ALWAYS True for multi-word text (square wrap) -- never "none"
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0; tf.vertical_anchor=anchor
    return tf
def addline(tf,text,font,sz,color,align=PP_ALIGN.CENTER,first=False):
    p=tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment=align; p.line_spacing=Pt(round(sz*1.15,2))
    r=p.add_run(); r.text=text; r.font.name=font; r.font.size=Pt(sz); r.font.bold=False
    r.font.italic=False; r.font.color.rgb=color
def brand_tr(slide):
    addline(tbox(slide,CW-0.3-4.2,0.12,4.2,0.5,wrap_=False),BRAND,PLAYFAIR,20,TITLE,align=PP_ALIGN.RIGHT,first=True)
def category_title(slide,txt):              # mixed case, never uppercased
    addline(tbox(slide,0.30,0.06,6.0,0.6),txt,PLAYFAIR,25,DKGREY3,align=PP_ALIGN.LEFT,first=True)

# cover
s=prs.slides.add_slide(BLANK)
tf=tbox(s,0.8,1.9,CW-1.6,1.9,MSO_ANCHOR.MIDDLE)
addline(tf,BRAND,PLAYFAIR,44,TITLE,first=True)
addline(tf,'Merchandising Assortment',PLAYFAIR,35,TITLE); addline(tf,PERIOD,PLAYFAIR,35,TITLE)
brand_tr(s)

def add_card(slide,cell_x,row_top,it):
    iw,ih,cw=it['iw'],it['ih'],it['cw']
    if it.get('imgp'):
        slide.shapes.add_picture(it['imgp'],Inches(cell_x+(cw-iw)/2),Inches(row_top+HDR_H+(IMG_H-ih)),Inches(iw),Inches(ih))
    tf=tbox(slide,cell_x,row_top+HDR_H+IMG_H+GAP_IMG,cw,TXT_H); maxw=cw*72; first=True
    for txt in [it['name'],it['color'],it['price'],it['tier']]+([it['annot']] if it['annot'] else []):
        if not txt: continue
        for ln in wrap(txt,4,False,maxw): addline(tf,ln,LATO,4,BLACK,first=first); first=False
def add_fabric_header(slide,x,w,row_top,fabric):
    tf=tbox(slide,x,row_top,w,HDR_H,MSO_ANCHOR.MIDDLE); first=True
    for ln in wrap(fabric,5.5,True,w*72): addline(tf,ln,PLAYFAIR,5.5,BLACK,first=first); first=False
def layout_rows(its):
    groups=[]
    for it in its:
        if groups and groups[-1][0]==it['fabric']: groups[-1][1].append(it)
        else: groups.append((it['fabric'],[it]))
    rows_,row,cursor,first=[],[],MX,True; ROWW=RIGHT-MX
    for fab,git in groups:
        gw=sum(it['cw'] for it in git)+(len(git)-1)*GIN; gap_before=0 if first else GGRP
        if (not first) and gw<=ROWW and (cursor+gap_before+gw)>RIGHT:
            rows_.append(row); row=[]; cursor=MX; first=True
        for i,it in enumerate(git):
            gap=0 if first else (GGRP if i==0 else GIN); cell_x=cursor+gap
            if (not first) and cell_x+it['cw']>RIGHT:
                rows_.append(row); row=[]; cell_x=MX; first=True
            row.append((cell_x,it)); cursor=cell_x+it['cw']; first=False
    if row: rows_.append(row)
    return rows_
def render_category(cat,its,droppage=False):
    its=sorted(its,key=lambda r:(r['fabric'].upper(),r['name'].upper()))
    rows_=layout_rows(its)
    for i in range(0,len(rows_),ROWS_PER):
        s=prs.slides.add_slide(BLANK)
        if droppage: bg(s,GREY)
        category_title(s,cat); brand_tr(s)
        for ri,row in enumerate(rows_[i:i+ROWS_PER]):
            row_top=GRID_TOP+ri*ROWPITCH; j=0
            while j<len(row):
                k=j
                while k+1<len(row) and row[k+1][1]['fabric']==row[j][1]['fabric']: k+=1
                add_fabric_header(s,row[j][0],row[k][0]+row[k][1]['cw']-row[j][0],row_top,row[j][1]['fabric']); j=k+1
            for cell_x,it in row: add_card(s,cell_x,row_top,it)

CAL=['January','February','March','April','May','June','July','August','September','October','November','December']
MONTHS=sorted({it['month'] for it in items if it['month']}, key=lambda m: CAL.index(m) if m in CAL else 99)
CATS=['Dresses','Tops','Bottoms','Outerwear','Rompers & Jumpsuits','Knits','Chase','Long Lead','DROP']
for m in MONTHS:
    s=prs.slides.add_slide(BLANK); bg(s,SAGE)
    addline(tbox(s,0.8,2.0,CW-1.6,1.6,MSO_ANCHOR.MIDDLE),m.upper(),PLAYFAIR_SB,48,TITLE,first=True); brand_tr(s)
    for cat in CATS:
        its=[it for it in items if it['month']==m and it['cat']==cat]
        if its: render_category(cat,its,droppage=(cat=='DROP'))
prs.save(OUT); print('saved', OUT, '|', len(prs.slides._sldIdLst), 'slides')
