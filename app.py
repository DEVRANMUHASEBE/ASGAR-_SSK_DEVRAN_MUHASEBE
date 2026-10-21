import io, os, json, calendar
from datetime import date
from pathlib import Path
import streamlit as st
import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(page_title="DEVRAN | Asgari İşçilik",page_icon="🏗️",layout="wide")
st.markdown('''<style>.stApp{background:#f7f9fc} .block-container{padding-top:1.4rem;max-width:1320px}h1,h2,h3{color:#172d4f}.stButton button[kind=primary]{background:#16345e;border-color:#16345e}</style>''',unsafe_allow_html=True)
LOGO = Path(__file__).with_name('logo.png')
a,b=st.columns([1,7])
with a:
    if LOGO.exists(): st.image(str(LOGO),width=110)
    else: st.markdown('## 🏛️')
with b:
    st.title('DEVRAN MÂLİ MÜŞAVİRLİK')
    st.caption('İNŞAAT ASGARİ İŞÇİLİK HESAPLAMA • v1.0')
st.divider()

def tr(n, unit=' TL'):
    return f'{n:,.2f}'.replace(',','X').replace('.',',').replace('X','.')+unit

def money(v):
    try:
        if isinstance(v,str): return float(v.replace('.','').replace(',','.'))
        return float(v or 0)
    except (ValueError, TypeError): return 0.0

def build_xlsx(rows):
    wb=Workbook()
    sh=wb.active
    sh.title='Aylik_SGK'
    sh.append(['Yıl','Ay','Prime Esas Kazanç (TL)','Ödenen SGK Primi (TL)'])
    for cell in sh[1]:
        cell.fill=PatternFill('solid',fgColor='16345E')
        cell.font=Font(bold=True,color='FFFFFF')
        cell.alignment=Alignment(vertical='center')
    sh.row_dimensions[1].height=26
    for letter,width in {'A':14,'B':20,'C':30,'D':30}.items():
        sh.column_dimensions[letter].width=width
    for r in rows:
        sh.append([int(r[0]),str(r[1]),float(r[2]),float(r[3])])
    for row in sh.iter_rows(min_row=2,min_col=3,max_col=4):
        for cell in row: cell.number_format='#,##0.00'
    summary=wb.create_sheet('Aciklama')
    summary.append(['DEVRAN MALİ MÜŞAVİRLİK - AYLIK SGK ŞABLONU'])
    summary.append(['Her ayın prime esas kazancını C sütununa, fiilen ödenen primi D sütununa yazınız.'])
    summary.append(['Asgari işçilik karşılaştırmasında C sütununun toplamı esas alınır.'])
    summary.column_dimensions['A'].width=110
    summary['A1'].font=Font(bold=True,color='16345E',size=14)
    buf=io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def read_xlsx(data):
    wb=load_workbook(io.BytesIO(data),read_only=True,data_only=True)
    sh=wb['Aylik_SGK']
    result=[]
    for r in sh.iter_rows(min_row=2,max_col=4,values_only=True):
        if not r or r[0] in (None,''): continue
        try:
            y=int(float(r[0]))
            result.append([y,str(r[1]),money(r[2]),money(r[3])])
        except (ValueError,TypeError):continue
    wb.close()
    return result

def pdf_report(project,details,monthly,total,required,base,missing,estimated,rate,prime):
    buf=io.BytesIO(); doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=34,leftMargin=34,topMargin=32,bottomMargin=32)
    fonts=['/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf','/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf']
    font=next((x for x in fonts if os.path.exists(x)),None)
    fname='Helvetica'
    if font:
        pdfmetrics.registerFont(TTFont('Devran',font));fname='Devran'
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='DevTitle',fontName=fname,fontSize=15,leading=20,alignment=TA_CENTER,spaceAfter=14))
    styles.add(ParagraphStyle(name='DevTxt',fontName=fname,fontSize=9,leading=14))
    story=[Paragraph('DEVRAN MÂLİ MÜŞAVİRLİK',styles['DevTitle']),Paragraph('İNŞAAT ASGARİ İŞÇİLİK - ÖZET HESAP RAPORU',styles['DevTxt']),Spacer(1,12)]
    info=[['Proje / Mükellef',project['ad']],['Ada / Parsel',project['ada']+' / '+project['parsel']],['İnşaat Dönemi',project['dates']],['Yapı / Blok Sayısı',str(len(details))]]
    def tbl(rows,widths):
        t=Table([[Paragraph(str(v),styles['DevTxt']) for v in row] for row in rows],colWidths=widths,repeatRows=1)
        t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.35,colors.lightgrey),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E8EDF5')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
        return t
    story += [tbl(info,[145,365]),Spacer(1,12)]
    summary=[['Kalem','Tutar'],['Toplam inşaat maliyeti',tr(total)],['Gerekli asgari işçilik matrahı',tr(required)],['Bildirilen prime esas kazanç',tr(base)],['Eksik işçilik matrahı',tr(missing)],['Tahmini ek prim (%'+tr(prime,'')+')',tr(estimated)],['Fiilen ödenen SGK primi (bilgi)',tr(sum(x[3] for x in monthly))]]
    story += [tbl(summary,[320,190]),Spacer(1,12)]
    detailrows=[['Blok / Yapı','Alan (m²)','Birim maliyet','Yapı maliyeti']]
    for r in details: detailrows.append([r['name'],tr(r['area'],''),tr(r['cost'],''),tr(r['total'],'')])
    story += [tbl(detailrows,[146,99,125,140]),Spacer(1,12)]
    explanation=f"Asgari işçilik oranı %{rate:g}, araştırma indirimi %{project['discount']:g}; tahmini prim oranı %{prime:g}. Birim maliyet ve tarihler kontrol edilmelidir. Prim tahmini gecikme zammı, ceza ve teşvikleri içermez."
    story.append(Paragraph(explanation,styles['DevTxt']))
    doc.build(story);return buf.getvalue()

with st.sidebar:
    st.header('⚙️ Hesaplama ayarları')
    work_rate=st.number_input('Asgari işçilik oranı (%)',min_value=0.,max_value=100.,value=9.0,step=.25,help='Yapı türüne ait resmî SGK tebliğ oranını kontrol ederek giriniz. %38,75 burada kullanılmaz.')
    discount=st.number_input('Araştırma indirimi (%)',min_value=0.,max_value=100.,value=25.,step=1.,help='SGK araştırmasının niteliğine göre uygulanabilmesi kontrol edilmelidir.')
    premium_rate=st.number_input('SGK toplam prim oranı (%)',min_value=0.,max_value=100.,value=38.75,step=.25)
    st.warning('2026 %38,75 SGK toplam prim oranıdır; asgari işçilik oranı değildir.')
    st.caption('Bu hesap ön değerlendirmedir; SGK’nın nihai tahakkuku değildir.')

st.subheader('1 · Proje bilgileri')
c1,c2,c3=st.columns([2,1,1]);name=c1.text_input('Proje / mükellef adı',placeholder='Örn. 34 Villa Projesi');ada=c2.text_input('Ada No');parsel=c3.text_input('Parsel No')
c1,c2,c3=st.columns(3)
start=c1.date_input('İnşaat başlangıç tarihi',value=date(2024,12,1),format='DD.MM.YYYY')
end=c2.date_input('İnşaat bitiş / hesap tarihi',value=date.today(),format='DD.MM.YYYY')
method=c3.selectbox('Hesaplama mevzuatı', ['16.05.2026 sonrası','16.05.2026 öncesi'],help='İşlemin tabi olduğu düzenlemeyi kontrol ediniz.')
if end<start:st.error('Bitiş tarihi başlangıçtan önce olamaz.');st.stop()

st.subheader('2 · Yapı bilgileri')
classes=['I-A','I-B','II-A','II-B','II-C','III-A','III-B','III-C','IV-A','IV-B','IV-C','V-A','V-B','V-C','V-D']
if 'blocks' not in st.session_state:st.session_state.blocks=[]
if 'next_id' not in st.session_state:st.session_state.next_id=1

def add_block():
    i=st.session_state.next_id;st.session_state.next_id+=1;st.session_state.blocks.append(i)

def remove_block(i):st.session_state.blocks.remove(i)

st.caption('Blok yoksa doğrudan tek yapı alanını doldur. Varsa aşağıdaki + butonuyla istediğin kadar blok ekle.')
if not st.session_state.blocks:
    x,y=st.columns(2);area=x.number_input('Toplam inşaat alanı (m²)',min_value=0.0,step=1.,format='%.2f');klass=y.selectbox('Yapı sınıfı',classes,index=classes.index('IV-B'))
    entries=[{'name':'Tek yapı','area':area,'klass':klass,'start':start,'end':end}]
else:
    entries=[]
    for i in st.session_state.blocks:
        with st.container(border=True):
            a,b,c,d=st.columns([2,2,2,.7]); n=a.text_input('Blok adı',value=f'{i}. Blok',key=f'n{i}');ar=b.number_input('Alan (m²)',min_value=0.,format='%.2f',key=f'a{i}');cl=c.selectbox('Yapı sınıfı',classes,index=classes.index('IV-B'),key=f'c{i}');
            if d.button('🗑️',key=f'd{i}',help='Bloğu sil'):remove_block(i);st.rerun()
            x,y=st.columns(2);bs=x.date_input('Başlangıç',value=start,key=f's{i}',format='DD.MM.YYYY');be=y.date_input('Bitiş',value=end,key=f'e{i}',format='DD.MM.YYYY')
            entries.append({'name':n,'area':ar,'klass':cl,'start':bs,'end':be})
if st.button('➕ Blok Ekle',on_click=add_block):st.rerun()

st.subheader('3 · Yıllara göre birim maliyet tarifesi')
st.info('Tarife tablosunu bir kez doldurunca program başlangıç/bitiş tarihleri ve sınıfa göre bedeli otomatik seçer. İlk sürümde resmî tutarları doğrulanmamış yıllar için rakam uydurulmuyor.')
needed=sorted({(year,e['klass']) for e in entries for year in (range(e['start'].year,e['end'].year) if method.startswith('16.05.2026 sonrası') and e['end'].year>e['start'].year else ([e['end'].year] if e['start'].year==e['end'].year else [e['end'].year-1]))})
if 'costs' not in st.session_state:st.session_state.costs={}
for year,cl in needed:
    key=f'{year}|{cl}'
    a,b=st.columns([2,3]); a.write(f'**{year} / {cl}**')
    st.session_state.costs[key]=b.number_input(f'{year} / {cl} resmî birim maliyet (TL/m²)',min_value=0.,value=float(st.session_state.costs.get(key,0)),step=100.,key=f'cost_{key}',label_visibility='collapsed')

def unit_cost(e):
    if e['start'].year==e['end'].year:yrs=[e['end'].year]
    elif method.startswith('16.05.2026 sonrası'):yrs=list(range(e['start'].year,e['end'].year))
    else:yrs=[e['end'].year-1]
    values=[st.session_state.costs.get(f'{y}|{e["klass"]}',0) for y in yrs]
    if not all(values):return None,yrs
    return sum(values)/len(values),yrs

st.subheader('4 · Aylık SGK bildirimi / Excel')
months=['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran','Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
monthly_base=[[y,m,0.,0.] for y in range(start.year,end.year+1) for mi,m in enumerate(months,1) if date(y,mi,1)<=end and date(y,mi,calendar.monthrange(y,mi)[1])>=start]
if 'monthly_data' not in st.session_state:st.session_state.monthly_data=monthly_base
if 'monthly_sig' not in st.session_state:st.session_state.monthly_sig=None
sig=(start,end)
if st.session_state.monthly_sig!=sig:
    prev={(r[0],r[1]):r for r in st.session_state.monthly_data}
    st.session_state.monthly_data=[prev.get((r[0],r[1]),r) for r in monthly_base]
    st.session_state.monthly_sig=sig
x,y=st.columns(2)
with x:
    st.download_button('⬇️ EXCEL ŞABLON İNDİR',data=build_xlsx(st.session_state.monthly_data),file_name=f'DEVRAN_SGK_SABLON_{start.year}_{end.year}.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
with y:
    upload=st.file_uploader('Doldurulmuş Excel’i yükle',type=['xlsx'],label_visibility='collapsed')
if upload is not None:
    filebytes=upload.getvalue();finger=(upload.name,len(filebytes),__import__('hashlib').md5(filebytes).hexdigest())
    if st.session_state.get('last_upload')!=finger:
        try:
            loaded=read_xlsx(filebytes)
            known={(r[0],r[1]):r for r in loaded}
            st.session_state.monthly_data=[known.get((r[0],r[1]),r) for r in monthly_base]
            st.session_state.last_upload=finger
            st.success(f'{len(loaded)} aylık kayıt okundu.')
        except Exception as exc:st.error(f'Excel okunamadı: {exc}')
frame=pd.DataFrame(st.session_state.monthly_data,columns=['Yıl','Ay','Prime Esas Kazanç (TL)','Ödenen SGK Primi (TL)'])
edited=st.data_editor(frame,hide_index=True,use_container_width=True,disabled=['Yıl','Ay'],column_config={'Prime Esas Kazanç (TL)':st.column_config.NumberColumn(min_value=0,format='%.2f'),'Ödenen SGK Primi (TL)':st.column_config.NumberColumn(min_value=0,format='%.2f')},key='sgk_editor')
monthly=[[int(r['Yıl']),str(r['Ay']),money(r['Prime Esas Kazanç (TL)']),money(r['Ödenen SGK Primi (TL)'])] for _,r in edited.iterrows()]
st.session_state.monthly_data=monthly

st.subheader('5 · ANLIK HESAPLAMA SONUCU')
details=[];warnings=[]
for e in entries:
    c,yrs=unit_cost(e)
    if e['end']<e['start']:warnings.append(f'{e["name"]}: tarih aralığı hatalı')
    if c is None:warnings.append(f'{e["name"]}: {e["klass"]} için {", ".join(map(str,yrs))} yılı/yılları birim maliyeti eksik')
    details.append({'name':e['name'],'area':e['area'],'klass':e['klass'],'cost':c or 0,'total':e['area']*(c or 0),'years':yrs})
for w in warnings:st.warning(w)
all_area=sum(r['area'] for r in details);total=sum(r['total'] for r in details)
required=total*(work_rate/100)*(1-discount/100)
base=sum(r[2] for r in monthly);paid=sum(r[3] for r in monthly)
missing=max(0,required-base);surplus=max(0,base-required)
estimated=missing*premium_rate/100
p,q,r,s=st.columns(4)
p.metric('Toplam m²',tr(all_area,' m²'));q.metric('İnşaat maliyeti',tr(total));r.metric('Asgari işçilik matrahı',tr(required));s.metric('Bildirilen PEK',tr(base))
p,q,r=st.columns(3)
p.metric('EKSİK İŞÇİLİK',tr(missing));q.metric('FAZLA BİLDİRİM',tr(surplus));r.metric('TAHMİNİ EK PRİM',tr(estimated))
st.caption('Ödenen SGK primi: '+tr(paid)+' | Asgari işçilik kıyası ödenen primle değil, prime esas kazançla yapılır.')
if warnings:st.error('Birim maliyetler tamamlanmadığı için hesap ve PDF sonucu kesinleştirilmemiştir. Önce eksikleri doldurunuz.')
st.dataframe(pd.DataFrame(details).rename(columns={'name':'Yapı','area':'m²','klass':'Sınıf','cost':'Birim TL/m²','total':'Toplam TL','years':'Tarife yılları'}),hide_index=True,use_container_width=True)
project={'ad':name or 'Belirtilmedi','ada':ada or '-','parsel':parsel or '-','dates':f'{start:%d.%m.%Y} - {end:%d.%m.%Y}','discount':discount}
try:
    pdf=pdf_report(project,details,monthly,total,required,base,missing,estimated,work_rate,premium_rate)
    st.download_button('📄 PDF İNDİR (ÖZET SONUÇ)',data=pdf,file_name=f'DEVRAN_ASGARI_ISCILIK_{ada or "ADA"}_{parsel or "PARSEL"}.pdf',mime='application/pdf',disabled=bool(warnings) or all_area<=0,type='primary',use_container_width=True)
except Exception as e:st.error(f'PDF oluşturulamadı: {e}')
st.caption('Kaynak: SGK Asgari İşçilik açıklamaları ve 16.05.2026 tarihli SSİY değişikliği. Kısmi ilişiksizlik dönem ayrımı sonraki sürüme eklenecektir.')
