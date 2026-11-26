import copy, re
from lxml import etree
from common import *
F=Fmt('docx')
t=etree.fromstring(open('../dx/word/document.xml','rb').read())
b=t.find(W+'body'); E=list(b)
rep,setp,setcell=F.rep,F.setp,F.setcell
def mk_marker(after_el, text, before=False):
    p=etree.Element(W+'p'); r=etree.SubElement(p,W+'r'); tt=etree.SubElement(r,W+'t'); tt.text=text
    (after_el.addprevious if before else after_el.addnext)(p); return p
# 0 머리 표: 회사명 · 주소 · TEL
hdr=E[0]
for p in hdr.iter(W+'p'):
    tx=F.ptext(p)
    if '주식회사' in tx: setp(p,'{{coFull}}')
    elif '부산광역시' in tx: setp(p,'{{coAddr}}')
    elif 'TEL' in tx: setp(p,'{{coTel}}')
# 2 문서번호·날짜
p=E[2]; rep(p,'베스트','{{docPfx}}'); rep(p,'2026-1126','{{docNum}}'); rep(p,'2026. 11. 26.','{{date}}')
rep(E[3],'CTR 모빌리티','{{client}}'); rep(E[4],'재계약 담당','{{ref}}')
rep(E[5],'2027년도 경비·미화 용역 재계약 갱신 협조의 件','{{title}}')
p=E[8]; rep(p,' 2027년 최저임금',' {{Y}}년 최저임금'); rep(p,' 2027년도 재계약',' {{Y}}년도 재계약')
rep(E[10],'CTR 모빌리티 (울산, 서산, 대구)','{{target}}')
rep(E[11],'경비·미화 인상률 반영','{{svc}} 인상률 반영')
# 12 사업장 요약 표
tb=E[12]; R=F.rows(tb)
c=F.cells(R[0]); setcell(c[1],'{{Y0}}년 월 도급비'); setcell(c[2],'{{Y}}년 월 도급비(안)')
c=F.cells(R[1]); setcell(c[0],'{{@sites:s}}{{s.name}}'); setcell(c[1],'{{s.m0}}원'); setcell(c[2],'{{s.m1}}원'); setcell(c[3],'{{s.rateTxt}}')
tb.remove(R[2]); tb.remove(R[3])
c=F.cells(R[4]); setcell(c[1],'{{totM0}}원'); setcell(c[2],'{{totM1}}원'); setcell(c[3],'{{totRateTxt}}')
rep(E[15],'공장별','사업장별')
rep(E[17],'에스텍베스트(051-897-5280) 또는 최민영 지사장(010-6560-5599)','{{contact}}')
rep(E[20],'2027년 도급비용 산출내역서(울산·서산·대구)','{{att2}}')
setp(E[21],'{{coFull}}')
p=E[22]; rep(p,'대표이사  정   규   상','대표이사  {{ceoSp}}')
rep(E[26],' 2027년 인상',' {{Y}}년 인상')
p=E[29]; setp(p,'- {{mwLine}}')
tb=E[31]; R=F.rows(tb)
c=F.cells(R[0]); [setcell(c[i],v) for i,v in [(1,'{{Ym1}}년'),(2,'{{Y0}}년'),(3,'{{Y}}년')]]
for k,key in [(1,'mw'),(2,'mi'),(3,'mr')]:
    c=F.cells(R[k]); [setcell(c[i],'{{%s_%s}}'%(key,n)) for i,n in [(1,'m1'),(2,'0'),(3,'1')]]
setcell(F.cells(R[1])[4],'{{mwNote}}')
setp(E[33],'- {{insLine}}')
tb=E[34]; R=F.rows(tb)
c=F.cells(R[0]); setcell(c[1],'{{Y0}}년'); setcell(c[2],'{{Y}}년')
c=F.cells(R[1]); [setcell(c[i],v) for i,v in enumerate(['{{@ins:i}}{{i.l}}','{{i.p}}','{{i.c}}','{{i.d}}','{{i.note}}'])]
for r in R[2:]: tb.remove(r)
rep(E[38],'. 2027년 인상률 산정 (경비·미화)','. {{Y}}년 인상률 산정 ({{svc}})')
rep(E[39],' 4.25% ',' {{totRate}}% ')
tb=E[40]; R=F.rows(tb)
c=F.cells(R[1]); [setcell(c[i],v) for i,v in enumerate(['{{@fac:f}}{{f.no}}','{{f.l}}','{{f.d}}','{{f.v}}'])]
for r in R[2:6]: tb.remove(r)
c=F.cells(R[6]); setcell(c[0],'합  계 ({{svc}} 인상률)'); setcell(c[1],'{{totRate}}%')
setp(E[42],'가. 계약기간 : {{period}}')
# 사업장 블록: E44(제목) E45(표) + 특이사항 E49 / 나머지 삭제
rep(E[44],'1) ','{{s.no}}) '); rep(E[44],'울산공장','{{s.name}}')
tb=E[45]; R=F.rows(tb)
c=F.cells(R[0]); setcell(c[1],'{{Y0}}년 계약현황'); setcell(c[2],'{{Y}}년 계약갱신(안)')
c=F.cells(R[2]); [setcell(c[i],v) for i,v in enumerate(['{{@s.cats:c}}{{c.cat}}','{{c.n0}}','{{c.u0}}','{{c.a0}}','{{c.n1}}','{{c.u1}}','{{c.a1}}','{{c.note}}'])]
tb.remove(R[3])
c=F.cells(R[4]); [setcell(c[i],v) for i,v in enumerate(['합계','{{s.n0}}','{{s.u0}}','{{s.a0}}','{{s.n1}}','{{s.u1}}','{{s.a1}}','{{s.sumNote}}'])]
c=F.cells(R[5]); [setcell(c[i],v) for i,v in enumerate(['계약금액','{{s.n0}}','','{{s.k0}}','{{s.n1}}','','{{s.k1}}','{{s.rateTxt}}'])]
note=E[49]; setp(note,'※ {{?s.note}}')
b.remove(E[46]); b.remove(E[47]); b.remove(E[48]); b.remove(E[50]); b.remove(E[51])
E[45].addnext(note)
mk_marker(E[44],'{{#sites:s}}',before=True); mk_marker(note,'{{/sites}}')
rep(E[52],'「2027년 도급비용 산출내역서」','「{{Y}}년 도급비용 산출내역서」')
rep(E[55],'이에 2027년','이에 {{Y}}년')
out=etree.tostring(t,xml_declaration=True,encoding='UTF-8',standalone=True)
open('docx_document.xml','wb').write(out)
print('ok',len(out))
print(sorted(set(re.findall(r'\{\{[^}]+\}\}',out.decode()))))
