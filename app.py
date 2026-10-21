import io, os
from datetime import date
from pathlib import Path
import streamlit as st
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Resmi Gazete yapi yaklasik birim maliyetleri (TL/m2).
# 2024 RG 20.02.2024/32466; 2025 RG 31.01.2025/32799;
# 2026 RG 03.02.2026/33157.
OFFICIAL_COSTS = {
    2024: {'I-A':1500,'I-B':2100,'II-A':3500,'II-B':5250,'II-C':7750,
           'III-A':12250,'III-B':14400,'III-C':None,'IV-A':15300,
           'IV-B':17400,'IV-C':18700,'V-A':21300,'V-B':22250,
           'V-C':24300,'V-D':26800},
    2025: {'I-A':2100,'I-B':3050,'I-C':3300,'I-D':3900,
           'II-A':6600,'II-B':10200,'II-C':12400,'III-A':17100,
           'III-B':18200,'III-C':19150,'IV-A':21500,'IV-B':27500,
           'IV-C':32600,'V-A':34400,'V-B':35600,'V-C':39500,
           'V-D':43400,'V-E':86250},
    2026: {'I-A':2600,'I-B':3900,'I-C':4200,'I-D':4800,
           'II-A':8100,'II-B':12500,'II-C':15100,'III-A':19800,
           'III-B':21050,'III-C':23400,'IV-A':26450,'IV-B':33900,
           'IV-C':40500,'V-A':42350,'V-B':43850,'V-C':48750,
           'V-D':53500,'V-E':103500},
}

def tariff_years(entry):
    """Bitiş yılı hariç: 1-2 takvim yılında başlangıç yılı; 3+ yılda bitişten önceki yılların ortalaması."""
    first, last = entry['start'].year, entry['end'].year
    if last < first:
        return []
    if last - first <= 1:
        return [first]
    return list(range(first, last))

def tariff_cost(year, klass):
    return OFFICIAL_COSTS.get(year, {}).get(klass)

st.set_page_config(page_title="DEVRAN | Asgari İşçilik",page_icon="🏗️",layout="wide")
st.markdown('''<style>.stApp{background:#f7f9fc} .block-container{padding-top:1.4rem;max-width:1320px}h1,h2,h3{color:#172d4f}.stButton button[kind=primary]{background:#16345e;border-color:#16345e}</style>''',unsafe_allow_html=True)
LOGO = Path(__file__).with_name('logo.png')
a,b=st.columns([1,7])
with a:
    if LOGO.exists(): st.image(str(LOGO),width=110)
    else: st.markdown('## 🏛️')
with b:
    st.title('DEVRAN MÂLİ MÜŞAVİRLİK')
    st.caption('İNŞAAT ASGARİ İŞÇİLİK HESAPLAMA • v1.5')
st.divider()

def tr(n, unit=' TL'):
    return f'{n:,.2f}'.replace(',','X').replace('.',',').replace('X','.')+unit

def pdf_report(project,details,total,required,rate):
    buf=io.BytesIO(); doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=34,leftMargin=34,topMargin=32,bottomMargin=32)
    fonts=[str(Path(__file__).with_name('DejaVuSans.ttf')),'/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf','/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf']
    font=next((x for x in fonts if os.path.exists(x)),None)
    fname='Helvetica'
    if font:
        pdfmetrics.registerFont(TTFont('Devran',font));fname='Devran'
    styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='DevTitle',fontName=fname,fontSize=15,leading=20,alignment=TA_CENTER,spaceAfter=14))
    styles.add(ParagraphStyle(name='DevTxt',fontName=fname,fontSize=9,leading=14))
    story=[]
    if LOGO.exists():
        try:
            logo=Image(str(LOGO),width=68,height=68,kind='proportional');logo.hAlign='CENTER';story.append(logo);story.append(Spacer(1,5))
        except Exception:pass
    story += [Paragraph('DEVRAN MÂLİ MÜŞAVİRLİK',styles['DevTitle']),Paragraph('İNŞAAT ASGARİ İŞÇİLİK - ÖZET HESAP RAPORU',styles['DevTxt']),Spacer(1,12)]
    info=[['Proje / Mükellef',project['ad']],['Ada / Parsel',project['ada']+' / '+project['parsel']],['İnşaat Dönemi',project['dates']],['İnşaat Durumu',project['status']],['Yapı / Blok Sayısı',str(len(details))]]
    def tbl(rows,widths):
        t=Table([[Paragraph(str(v),styles['DevTxt']) for v in row] for row in rows],colWidths=widths,repeatRows=1)
        t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),.35,colors.lightgrey),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E8EDF5')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
        return t
    story += [tbl(info,[145,365]),Spacer(1,12)]
    summary=[['Kalem','Tutar'],['Toplam inşaat alanı',tr(sum(d['area'] for d in details),' m²')],['Toplam inşaat maliyeti',tr(total)],['Asgari işçilik oranı',f'%{rate:g}'],['Araştırma indirimi',f"%{project['discount']:g}"],['Hesaplanan asgari işçilik matrahı',tr(required)]]
    story += [tbl(summary,[320,190]),Spacer(1,12)]
    detailrows=[['Blok / Yapı','Alan (m²)','Birim maliyet','Yapı maliyeti']]
    for r in details: detailrows.append([r['name'],tr(r['area'],''),tr(r['cost'],''),tr(r['total'],'')])
    story += [tbl(detailrows,[146,99,125,140]),Spacer(1,12)]
    explanation=f"Asgari işçilik oranı %{rate:g}, araştırma indirimi %{project['discount']:g}. Bu rapor yalnızca hesaplanan asgari işçilik matrahını gösterir; bildirilen prime esas kazanç ve ödenen primler karşılaştırılmamıştır. Birim maliyet ve tarihler kontrol edilmelidir."
    story.append(Paragraph(explanation,styles['DevTxt']))
    doc.build(story);return buf.getvalue()

with st.sidebar:
    st.header('⚙️ Hesaplama ayarları')
    work_rate=st.number_input('Asgari işçilik oranı (%)',min_value=0.,max_value=100.,value=9.0,step=.25,help='Yapı türüne ait resmî SGK tebliğ oranını kontrol ederek giriniz. %38,75 burada kullanılmaz.')
    discount=st.number_input('Araştırma indirimi (%)',min_value=0.,max_value=100.,value=25.,step=1.,help='SGK araştırmasının niteliğine göre uygulanabilmesi kontrol edilmelidir.')
    st.caption('Bu hesap ön değerlendirmedir; SGK’nın nihai tahakkuku değildir.')

st.subheader('1 · Proje bilgileri')
c1,c2,c3=st.columns([2,1,1]);name=c1.text_input('Proje / mükellef adı',placeholder='Örn. 34 Villa Projesi');ada=c2.text_input('Ada No');parsel=c3.text_input('Parsel No')
c1,c2,c3=st.columns(3)
start=c1.date_input('İnşaat başlangıç tarihi',value=date(2024,12,1),format='DD.MM.YYYY')
ongoing=c2.checkbox('İnşaat henüz bitmedi',value=False,help='Devam eden inşaat için bitiş yerine hesaplama tarihi kullanılır. Sonuç kesin ilişiksizlik hesabı değildir.')
end=c2.date_input('İnşaat bitiş / hesap tarihi',value=date.today(),format='DD.MM.YYYY',help='İnşaat devam ediyorsa hesaplama tarihini giriniz.')
c3.info('İnşaat devam ediyorsa girilen tarih yalnızca geçici hesaplamada kullanılır.' if ongoing else 'Tarife yılı, başlangıç ve bitiş tarihlerine göre otomatik seçilir.')
if end<start:st.error('Bitiş tarihi başlangıçtan önce olamaz.');st.stop()

st.subheader('2 · Yapı bilgileri')
classes=['I-A','I-B','I-C','I-D','II-A','II-B','II-C','III-A','III-B','III-C','IV-A','IV-B','IV-C','V-A','V-B','V-C','V-D','V-E']
if 'blocks' not in st.session_state:st.session_state.blocks=[]
if 'next_id' not in st.session_state:st.session_state.next_id=1

def add_block():
    i=st.session_state.next_id;st.session_state.next_id+=1;st.session_state.blocks.append(i)

def remove_block(i):st.session_state.blocks.remove(i)

st.caption('Blok yoksa doğrudan tek yapı alanını doldur. Varsa aşağıdaki + butonuyla istediğin kadar blok ekle.')
if not st.session_state.blocks:
    x,y=st.columns(2);area=x.number_input('Toplam inşaat alanı (m²)',min_value=0.0,step=1.,format='%.2f');klass=y.selectbox('Yapı sınıfı',classes,index=classes.index('IV-B'))
    entries=[{'name':'','area':area,'klass':klass,'start':start,'end':end}]
else:
    entries=[]
    for i in st.session_state.blocks:
        with st.container(border=True):
            a,b,c,d=st.columns([2,2,2,.7]); n=a.text_input('Blok adı',value='' if len(st.session_state.blocks)==1 else f'{i}. Blok',key=f'n{i}',placeholder='Tek blok ise boş bırakabilirsiniz');ar=b.number_input('Alan (m²)',min_value=0.,format='%.2f',key=f'a{i}');cl=c.selectbox('Yapı sınıfı',classes,index=classes.index('IV-B'),key=f'c{i}');
            if d.button('🗑️',key=f'd{i}',help='Bloğu sil'):remove_block(i);st.rerun()
            x,y=st.columns(2);bs=x.date_input('Başlangıç',value=start,key=f's{i}',format='DD.MM.YYYY');be=y.date_input('Bitiş',value=end,key=f'e{i}',format='DD.MM.YYYY')
            entries.append({'name':n,'area':ar,'klass':cl,'start':bs,'end':be})
if st.button('➕ Blok Ekle',on_click=add_block):st.rerun()

st.subheader('3 · Otomatik birim maliyet tarifesi')
st.caption('Aynı yıl veya takip eden yıl bitişinde başlangıç yılı; 3 yılda aradaki 2 yıl; 4 yılda aradaki 3 yılın ortalaması uygulanır. Tanımlı tarifeler: 2024–2026.')

# Kullanıcının seçtiği mevzuat yöntemi tarife yıllarını belirler.
needed = sorted({(y, e['klass']) for e in entries for y in tariff_years(e)})
missing_rates = [(y, cl) for y, cl in needed if tariff_cost(y, cl) is None]
overrides = {}
if missing_rates:
    st.warning('Tarifede bulunamayan yıl/sınıf kombinasyonları için resmî tutarı kontrol ederek giriniz.')
    for year, cl in missing_rates:
        overrides[(year, cl)] = st.number_input(f'{year} / {cl} — TL/m² (elle doğrulanacak)', min_value=0., step=100., key=f'override_{year}_{cl}')

for year, cl in needed:
    amount = tariff_cost(year, cl)
    if amount is not None:
        st.write(f'✅ **{year} / {cl}: {tr(amount, " TL/m²")}** (resmî tarife)')

def unit_cost(e):
    years = tariff_years(e)
    values = [tariff_cost(y, e['klass']) if tariff_cost(y, e['klass']) is not None else overrides.get((y, e['klass']), 0) for y in years]
    if not all(values):
        return None, years
    return sum(values)/len(values), years

st.subheader('ANLIK HESAPLAMA SONUCU')
details=[];warnings=[]
for e in entries:
    c,yrs=unit_cost(e)
    if e['end']<e['start']:warnings.append(f'{e["name"] or "Yapı"}: tarih aralığı hatalı')
    if c is None:warnings.append(f'{e["name"] or "Yapı"}: {e["klass"]} için {", ".join(map(str,yrs))} yılı/yılları birim maliyeti eksik')
    details.append({'name':e['name'],'area':e['area'],'klass':e['klass'],'cost':c or 0,'total':e['area']*(c or 0),'years':yrs})
for w in warnings:st.warning(w)
all_area=sum(r['area'] for r in details);total=sum(r['total'] for r in details)
required=total*(work_rate/100)*(1-discount/100)
p,q,r=st.columns(3)
p.metric('Toplam m²',tr(all_area,' m²'))
q.metric('İnşaat maliyeti',tr(total))
r.metric('Hesaplanan asgari işçilik matrahı',tr(required))
if ongoing: st.info('İNŞAAT DEVAM EDİYOR: Aşağıdaki sonuçlar seçilen hesap tarihi itibarıyla geçici tahmindir. Nihai asgari işçilik/ilişiksizlik sonucu olarak kullanılmamalıdır.')
if warnings:st.error('Birim maliyetler tamamlanmadığı için hesap ve PDF sonucu kesinleştirilmemiştir. Önce eksikleri doldurunuz.')
st.dataframe([{'Blok / Yapı':d['name'],'m²':d['area'],'Sınıf':d['klass'],'Birim TL/m²':d['cost'],'Toplam TL':d['total'],'Tarife yılları':', '.join(map(str,d['years']))} for d in details],hide_index=True,use_container_width=True)
project={'ad':name or 'Belirtilmedi','ada':ada or '-','parsel':parsel or '-','dates':f'{start:%d.%m.%Y} - {end:%d.%m.%Y}' + (' (hesap tarihi)' if ongoing else ''),'status':'Devam ediyor (geçici hesap)' if ongoing else 'Tamamlandı','discount':discount}
try:
    pdf=pdf_report(project,details,total,required,work_rate)
    st.download_button('📄 PDF İNDİR (ÖZET SONUÇ)',data=pdf,file_name=f'DEVRAN_ASGARI_ISCILIK_{ada or "ADA"}_{parsel or "PARSEL"}.pdf',mime='application/pdf',disabled=bool(warnings) or all_area<=0,type='primary',use_container_width=True)
except Exception as e:st.error(f'PDF oluşturulamadı: {e}')
st.caption('Hesaplama notu: Bu program bir ön hesap yapar. SGK kuralları, inşaatın durumu ve varsa kısmi ilişiksizlik ayrı değerlendirilmelidir.')
