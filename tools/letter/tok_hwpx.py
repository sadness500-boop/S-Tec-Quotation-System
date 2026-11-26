import copy, re
from lxml import etree
from common import *
F=Fmt('hwpx')
t=etree.fromstring(open('../hx/Contents/section0.xml','rb').read())
E=list(t)
rep,setp,setcell=F.rep,F.setp,F.setcell
def tbl_of(p): return next(p.iter(F.TBL))
def mk_marker(ref,text,before=False):
    p=copy.deepcopy(E[3])  # 빈 문단 복제
    for x in list(p):
        if etree.QName(x).localname!='run': p.remove(x)
    r=p.find(F.R)
    if r is None: r=etree.SubElement(p,F.R); r.set('charPrIDRef','0')
    for x in list(r): r.remove(x)
    tt=etree.SubElement(r,F.T); tt.text=text
    (ref.addprevious if before else ref.addnext)(p); return p
rep(E[1],'주식회사 에스텍베스트','{{coFull}}'); rep(E[1],'부산광역시 부산진구 중앙대로 920 설비건설회관 2F','{{coAddr}}')
rep(E[2],'[TEL : 051)897-5280 / FAX : 051)897-5284]','{{coTel}}')
p=E[5]; rep(p,'베스트','{{docPfx}}'); rep(p,'2026-1126','{{docNum}}'); rep(p,'2026.11.26','{{date}}')
rep(E[6],'CTR 모빌리티','{{client}}'); rep(E[7],'재계약 담당','{{ref}}')
rep(E[8],'2027년도 경비·미화 용역 재계약 갱신 협조의 件','{{title}}')
p=E[13]; rep(p,' 2027년 최저임금',' {{Y}}년 최저임금'); rep(p,' 2027년도 재계약',' {{Y}}년도 재계약')
rep(E[17],'CTR 모빌리티 (울산, 서산, 대구)','{{target}}')
rep(E[19],'경비·미화 4.25% 인상 반영 (서산공장 4.93%)','{{incTxt}}')
rep(E[22],'에스텍베스트 (051-897-5280) 또는 ','{{contactA}} 또는 ')
rep(E[23],' 최민영 지사장(010-6560-5599)',' {{contactB}}')
rep(E[24],'2027년 개정법령','{{Y}}년 개정법령')
rep(E[28],'2027년 도급비용 산출내역서(울산·서산·대구)','{{att2}}')
rep(E[31],'주식회사 에스텍베스트','{{coFull}}')
rep(E[32],'정   규   상','{{ceoSp}}')
rep(E[38],'2027년 인상','{{Y}}년 인상')
setp(E[47],'         - {{mwLine}}')
tb=tbl_of(E[48]); R=F.rows(tb)
c=F.cells(R[0]); [setcell(c[i],v) for i,v in [(1,'{{Ym1}}년'),(2,'{{Y0}}년'),(3,'{{Y}}년')]]
for k,key in [(1,'mw'),(2,'mi'),(3,'mr')]:
    c=F.cells(R[k]); [setcell(c[i],'{{%s_%s}}'%(key,n)) for i,n in [(1,'m1'),(2,'0'),(3,'1')]]
setcell(F.cells(R[1])[4],'{{mwNoteH}}')
rep(E[50],'(2027년 확정','({{Y}}년 확정')
setp(E[51],'         - {{insLine}}')
tb=tbl_of(E[52]); R=F.rows(tb)
c=F.cells(R[0]); setcell(c[1],'{{Y0}}년'); setcell(c[2],'{{Y}}년'); setcell(c[5],'{{Y}}년')
# 장기요양 행(두 줄 셀)을 견본으로 쓰면 줄 수가 달라지므로 국민연금 행을 견본으로
c=F.cells(R[1]); [setcell(c[i],v) for i,v in enumerate(['{{@ins:i}}{{i.l}}','{{i.pH}}','{{i.cH}}','{{i.dp}}','{{i.dr}}','{{i.st}}'])]
for r in R[2:]: tb.remove(r)
rep(E[62],'2027년 인상률 산정 (경비·미화 : 4.25%)','{{Y}}년 인상률 산정 ({{svc}} : {{totRate}}%)')
setp(E[63],'{{@fac:f}}         {{f.no}}) {{f.l}}{{f.dH}} : {{f.v}}')
for k in (64,65,66,67): t.remove(E[k])
setp(E[68],'         {{@notes:n}}※ {{n.t}}')
setp(E[72],'     가. 계약기간 : {{period}}')
# 사업장 블록
rep(E[85],' 1) 울산공장',' {{s.no}}) {{s.name}}')
tb=tbl_of(E[86]); R=F.rows(tb)
c=F.cells(R[0]); setcell(c[1],'{{Y0}}년 계약현황'); setcell(c[2],'{{Y}}년 계약갱신(안)')
c=F.cells(R[2]); [setcell(c[i],v) for i,v in enumerate(['{{@s.cats:c}}{{c.cat}}','{{c.n0}}','{{c.u0}}','{{c.a0}}','{{c.n1}}','{{c.u1}}','{{c.a1}}','{{c.note}}'])]
tb.remove(R[3])
c=F.cells(R[4]); [setcell(c[i],v) for i,v in enumerate(['합계','{{s.n0}}','{{s.u0}}','{{s.a0}}','{{s.n1}}','{{s.u1}}','{{s.a1}}','{{s.sumNote}}'])]
c=F.cells(R[5]); [setcell(c[i],v) for i,v in enumerate(['{{Y}}년 계약금액','{{s.n1}}','','{{s.k1}}','{{s.rateTxt}}'])]
for k in (88,89,90,91,92): t.remove(E[k])
for k in range(74,84): t.remove(E[k])
mk_marker(E[85],'{{#sites:s}}',before=True); mk_marker(E[87],'{{/sites}}')
rep(E[99],'이에 2027년','이에 {{Y}}년')
out=etree.tostring(t,xml_declaration=True,encoding='UTF-8',standalone=True)
open('hwpx_section0.xml','wb').write(out)
print('ok',len(out)); print(sorted(set(re.findall(r'\{\{[^}]+\}\}',out.decode()))))
