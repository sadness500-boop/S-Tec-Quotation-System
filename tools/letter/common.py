from lxml import etree
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
HP='{http://www.hancom.co.kr/hwpml/2011/paragraph}'
class Fmt:
    def __init__(s,kind):
        s.kind=kind
        if kind=='docx': s.P,s.R,s.T,s.TR,s.TC,s.TBL=W+'p',W+'r',W+'t',W+'tr',W+'tc',W+'tbl'
        else: s.P,s.R,s.T,s.TR,s.TC,s.TBL=HP+'p',HP+'run',HP+'t',HP+'tr',HP+'tc',HP+'tbl'
    def texts(s,p):
        out=[]
        for r in p.findall(s.R):
            for t in r.findall(s.T):
                if len(t):  # 형광펜 등 하위 요소 제거
                    tx=''.join(t.itertext())
                    for c in list(t): t.remove(c)
                    t.text=tx
                out.append(t)
        return out
    def ptext(s,p): return ''.join(t.text or '' for t in s.texts(p))
    def settext(s,t,v):
        t.text=v
        if s.kind=='docx': t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    def rep(s,p,old,new,must=True):
        ts=s.texts(p); full=''.join(t.text or '' for t in ts)
        i=full.find(old)
        if i<0:
            if must: raise Exception(f'not found: {old!r} in {full!r}')
            return False
        j=i+len(old); pos=0; first=True
        for t in ts:
            tx=t.text or ''; a,b=pos,pos+len(tx); pos=b
            if b<=i or a>=j: continue
            lo,hi=max(i,a)-a,min(j,b)-a
            if first: s.settext(t,tx[:lo]+new+tx[hi:]); first=False
            else: s.settext(t,tx[:lo]+tx[hi:])
        return True
    def setp(s,p,v):
        ts=s.texts(p)
        if not ts: raise Exception('no text el')
        s.settext(ts[0],v)
        for t in ts[1:]: s.settext(t,'')
    def cells(s,tr): return tr.findall(s.TC)
    def cellps(s,tc): return [p for p in tc.iter(s.P)]
    def setcell(s,tc,v):
        ps=s.cellps(tc)
        withtext=[p for p in ps if s.texts(p)]
        p=withtext[0] if withtext else ps[0]
        if not s.texts(p): s.addtext(p,tc)
        s.setp(p,v)
        for q in ps:
            if q is not p and s.texts(q): s.setp(q,'')
    def rows(s,tbl): return tbl.findall(s.TR)

    def addtext(s,p,ctx):
        import copy
        if s.kind=='docx':
            src=None
            tr=ctx.getparent()
            for r in tr.iter(s.R):
                if r.find(W+'t') is not None and r.find(W+'rPr') is not None: src=r; break
            r=etree.SubElement(p,s.R)
            if src is not None: r.append(copy.deepcopy(src.find(W+'rPr')))
            etree.SubElement(r,s.T)
        else:
            r=p.find(s.R)
            if r is None:
                r=etree.SubElement(p,s.R); r.set('charPrIDRef','0')
            # hp:t는 linesegarray보다 앞(런 안)에 둔다
            etree.SubElement(r,s.T)
