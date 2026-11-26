# 사용자가 준 원본 공문(워드·한글)에서 만든 {{토큰}} 양식을 묶습니다. 직인 그림은 공개 저장소에 넣지 않도록 빈 그림으로 바꿉니다.
import zipfile, io
from PIL import Image
def blank(fmt,size):
    b=io.BytesIO(); Image.new('RGBA' if fmt=='PNG' else 'RGB',size,(255,255,255,0) if fmt=='PNG' else (255,255,255)).save(b,fmt); return b.getvalue()
def rebuild(src,dst,repl):
    zi=zipfile.ZipFile(src); zo=zipfile.ZipFile(dst,'w')
    for it in zi.infolist():
        b=repl.get(it.filename, zi.read(it.filename))
        zo.writestr(it.filename,b,compress_type=zipfile.ZIP_STORED if it.filename=='mimetype' else zipfile.ZIP_DEFLATED)
    zo.close()
rebuild('../t.docx','tpl.docx',{'word/document.xml':open('docx_document.xml','rb').read(),'word/media/image2.png':blank('PNG',(75,79))})
rebuild('../t.hwpx','tpl.hwpx',{'Contents/section0.xml':open('hwpx_section0.xml','rb').read(),'BinData/image2.bmp':blank('BMP',(75,79)),'Preview/PrvText.txt':'계약갱신 공문'.encode('utf-8')})
