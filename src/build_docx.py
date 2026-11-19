from content import *
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
FONT='Arial'; BLUE=RGBColor(0x1F,0x4E,0x79)
doc=Document()
sec=doc.sections[0]; sec.page_width=Cm(21); sec.page_height=Cm(29.7)
for m in ('left_margin','right_margin'): setattr(sec,m,Cm(2))
sec.top_margin=Cm(2); sec.bottom_margin=Cm(2)
def rfonts(el,size=None,bold=None):
    rpr=el.get_or_add_rPr()
    f=rpr.find(qn('w:rFonts'))
    if f is None: f=OxmlElement('w:rFonts'); rpr.insert(0,f)
    for a in ('w:ascii','w:hAnsi','w:cs','w:eastAsia'): f.set(qn(a),FONT)
    if size:
        for tag in ('w:sz','w:szCs'):
            e=OxmlElement(tag); e.set(qn('w:val'),str(int(size*2))); rpr.append(e)
    if bold:
        rpr.append(OxmlElement('w:bCs'))
for sname in ['Normal','Heading 1','Heading 2','Heading 3','Title','List Bullet','List Number']:
    st=doc.styles[sname]; st.font.name=FONT; rfonts(st.element)
    ppr=st.element.get_or_add_pPr(); b=OxmlElement('w:bidi'); ppr.append(b)
    if st.element.rPr is not None: st.element.rPr.append(OxmlElement('w:rtl'))
doc.styles['Normal'].font.size=Pt(11)
for s,sz in (('Heading 1',16),('Heading 2',13.5),('Heading 3',12)):
    doc.styles[s].font.size=Pt(sz); doc.styles[s].font.color.rgb=BLUE; rfonts(doc.styles[s].element,sz,True)
rfonts(doc.styles['Normal'].element,11)
def setrtl(p):
    ppr=p._p.get_or_add_pPr()
    if ppr.find(qn('w:bidi')) is None: ppr.append(OxmlElement('w:bidi'))
def run(p,text,bold=False,size=None,color=None,italic=False):
    r=p.add_run(text); r.bold=bold; r.italic=italic
    if size: r.font.size=Pt(size)
    if color: r.font.color.rgb=color
    rpr=r._r.get_or_add_rPr(); rpr.append(OxmlElement('w:rtl'))
    if bold: rpr.append(OxmlElement('w:bCs'))
    if size:
        e=OxmlElement('w:szCs'); e.set(qn('w:val'),str(int(size*2))); rpr.append(e)
    return r
def P(text='',bold=False,size=None,align=None,color=None,style=None,italic=False):
    p=doc.add_paragraph(style=style); setrtl(p)
    p.alignment=align if align is not None else WD_ALIGN_PARAGRAPH.RIGHT
    if text: run(p,text,bold,size,color,italic)
    return p
def RP(parts,style=None):
    p=doc.add_paragraph(style=style); setrtl(p); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    for t,b in parts: run(p,t,b)
    return p
def H(text,lvl):
    p=doc.add_heading(level=lvl); setrtl(p); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; run(p,text); return p
def B(text,bold_lead=None):
    p=doc.add_paragraph(style='List Bullet'); setrtl(p); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    if bold_lead: run(p,bold_lead,True)
    run(p,text); return p
def shade(cell,hexc):
    tcPr=cell._tc.get_or_add_tcPr(); s=OxmlElement('w:shd'); s.set(qn('w:val'),'clear'); s.set(qn('w:color'),'auto'); s.set(qn('w:fill'),hexc); tcPr.append(s)
def T(header,rows,widths=None,size=9,fills=None,hdr='1F4E79',ltr_cols=()):
    t=doc.add_table(rows=1,cols=len(header)); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    tblPr=t._tbl.tblPr; bv=OxmlElement('w:bidiVisual'); tblPr.append(bv)
    for j,h in enumerate(header):
        c=t.rows[0].cells[j]; c.text=''; p=c.paragraphs[0]; setrtl(p); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        run(p,str(h),True,size,RGBColor(255,255,255)); shade(c,hdr)
    trPr=t.rows[0]._tr.get_or_add_trPr(); th=OxmlElement('w:tblHeader'); trPr.append(th)
    for i,r in enumerate(rows):
        cells=t.add_row().cells
        for j,v in enumerate(r):
            c=cells[j]; c.text=''; p=c.paragraphs[0]; setrtl(p)
            sv='' if v is None else str(v)
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER if (j in ltr_cols or len(sv)<14) else WD_ALIGN_PARAGRAPH.RIGHT
            run(p,sv,False,size, RGBColor(0xC0,0,0) if sv=='(فارغ)' else None)
            if fills and fills(i,j,v): shade(c,fills(i,j,v))
    if widths:
        t.autofit=False
        lay=OxmlElement('w:tblLayout'); lay.set(qn('w:type'),'fixed'); tblPr.append(lay)
        for j,w in enumerate(widths): t.columns[j].width=Cm(w)
        for row in t.rows:
            for j,w in enumerate(widths): row.cells[j].width=Cm(w)
    doc.add_paragraph()
    return t
def fmtn(x): return f'{x:,}'
SEVC={'عالية':'F8CBAD','متوسطة':'FFE699','منخفضة':'C6EFCE'}
def status(score,t): return 'محقق' if score>=t else ('يحتاج تحسين' if score>=95 else 'حرج')
STC={'محقق':'C6EFCE','يحتاج تحسين':'FFE699','حرج':'F8CBAD'}
TG={'C':99.9,'A':99.99,'T':100,'S':99.9,'V':100,'Y':100}
OVERALL=sum(SCORES.values())/6

# ---------- cover
for _ in range(6): P()
P('مشروع جودة البيانات',True,26,WD_ALIGN_PARAGRAPH.CENTER,BLUE)
P('خطة الجودة ودراسة وتقييم جودة بيانات الرحلات الجوية',True,16,WD_ALIGN_PARAGRAPH.CENTER)
P('التقرير التنفيذي، وقواعد الأعمال وقواعد الجودة، وخطة الإصلاح',False,13,WD_ALIGN_PARAGRAPH.CENTER)
P(); P('مجموعة البيانات: flights.csv (رحلات المغادرة من مطارات نيويورك لعام 2013)',False,11,WD_ALIGN_PARAGRAPH.CENTER)
P('إعداد: هشام',True,12,WD_ALIGN_PARAGRAPH.CENTER); P('التاريخ: 18 نوفمبر 2026',False,11,WD_ALIGN_PARAGRAPH.CENTER)
doc.add_page_break()

# ---------- contents
H('المحتويات',1)
for t in ['1. الملخص التنفيذي','2. وصف مجموعة البيانات','3. خطة الجودة','   3.1 نطاق الخطة','   3.2 قائمة أولويات عناصر البيانات','   3.3 الأنشطة والأهداف المرحلية','   3.4 الموارد والميزانية','   3.5 الفريق ومصفوفة المهام والمسؤوليات (RACI)','4. دراسة جودة البيانات وفق الأبعاد الستة','5. التقرير التنفيذي لمشاكل الجودة','6. قواعد الأعمال وربطها بقواعد جودة البيانات','7. توزيع مهام إصلاح الجودة على الفريق','8. التوصيات والخطوات التالية','ملحق: المنهجية والملف الداعم']:
    P(t)
doc.add_page_break()

# ---------- 1 exec summary
H('1. الملخص التنفيذي',1)
P(f'قُيِّمت جودة ملف الرحلات الجوية flights.csv، وفيه {fmtn(N)} سجلاً و21 حقلاً لرحلات المغادرة من مطارات نيويورك الثلاثة (EWR و JFK و LGA) خلال عام 2013، لـ 16 ناقلاً جوياً إلى 105 وجهات. طُبقت 18 قاعدة جودة على مستوى السجلات و8 ضوابط على مستوى الملف، موزعة على أبعاد الجودة الستة.')
P(f'المؤشر العام للجودة {OVERALL:.1f}%. اكتمال الخلايا مرتفع ({CELL_COMPLETE:.2f}%)، لكن الضعف الأكبر في الحداثة والتوفر والاتساق. رُصدت {len(ISSUES)} مشكلة جودة: {sum(i[4]=="عالية" for i in ISSUES)} عالية الخطورة، و{sum(i[4]=="متوسطة" for i in ISSUES)} متوسطة، و{sum(i[4]=="منخفضة" for i in ISSUES)} منخفضة.')
T(['بُعد الجودة','النتيجة الحالية','المستهدف','الحالة'],
  [[DIM[d],f'{SCORES[d]:.2f}%',f'{TG[d]}%',status(SCORES[d],TG[d])] for d in 'CAVSTY'],[6,3.5,3,3.5],10,
  fills=lambda i,j,v: STC.get(v) if j==3 else None)
P('أبرز النتائج:',True)
B(' الأوقات مسجلة بالتوقيت المحلي دون منطقة زمنية، فتظهر مدة الطيران أطول من الزمن المنقضي بين الإقلاع والهبوط في 130,280 سجلاً (38.7%).','الاتساق:')
B(' 22,754 سجلاً (6.76%) فيها أرقام تسجيل طائرات لا تطابق صيغة FAA الرسمية.','الصلاحية:')
B(' 8,255 رحلة ملغاة (2.45%) لا تظهر إلا كقيم فارغة، لعدم وجود حقل لحالة الرحلة، و1,175 رحلة بيانات وصولها ناقصة.','الاكتمال:')
B(' آخر سجل في 2013-12-31، ولا يوجد ختم زمني للتحميل.','الحداثة:')
B(' لا يوجد قاموس بيانات ولا جداول مرجعية للطائرات والمطارات ولا قناة وصول محكومة.','التوفر:')
B(' 356 رحلة بسرعات غير منطقية، و35 حالة لطائرة تُقلع من مطارين خلال أقل من ساعة، و42 تأخيراً يتجاوز 10 ساعات.','الدقة:')
P(f'الخطة المقترحة مدتها 12 أسبوعاً على 6 مراحل، بفريق من {len(TEAM)} أدوار وميزانية تقديرية {fmtn(int(sum(b[2]*b[3]*b[4]*b[5] for b in BUDGET)*(1+CONTINGENCY)))} ريال (شاملة احتياطي 10%). تُعالَج المشاكل عبر 14 قاعدة أعمال مرتبطة بـ 26 قاعدة وضابط جودة، و{len(TASKS)} مهمة إصلاح موزعة على الفريق.')
doc.add_page_break()

# ---------- 2 data description
H('2. وصف مجموعة البيانات',1)
P('الملف مسطح (CSV) بحجم 42.9 ميجابايت. يمثل كل سجل رحلة مغادرة مجدولة واحدة. الجدول التالي يصف الحقول كما فُهمت من التحليل لعدم وجود قاموس بيانات:')
FD={'id':'معرف تسلسلي فني للسجل','year':'سنة الرحلة','month':'شهر الرحلة','day':'يوم الرحلة','dep_time':'وقت الإقلاع الفعلي (HHMM محلي)','sched_dep_time':'وقت الإقلاع المجدول (HHMM)','dep_delay':'تأخير الإقلاع بالدقائق (السالب = مبكر)','arr_time':'وقت الوصول الفعلي (HHMM بتوقيت الوجهة)','sched_arr_time':'وقت الوصول المجدول (HHMM)','arr_delay':'تأخير الوصول بالدقائق','carrier':'رمز الناقل الجوي (IATA)','flight':'رقم الرحلة','tailnum':'رقم تسجيل الطائرة','origin':'مطار المغادرة','dest':'مطار الوجهة','air_time':'مدة الطيران بالدقائق','distance':'المسافة بالميل','hour':'ساعة الإقلاع المجدول','minute':'دقيقة الإقلاع المجدول','time_hour':'تاريخ وساعة الإقلاع المجدول','name':'اسم الناقل الجوي'}
rows=[]
for c in ORIG.columns:
    s=ORIG[c]; rows.append([c,FD[c],'نص' if s.dtype==object or str(s.dtype)=='str' else 'رقمي',fmtn(int(s.isna().sum())),f'{s.isna().mean()*100:.2f}%',fmtn(int(s.nunique()))])
T(['الحقل','الوصف','النوع','الفارغ','نسبة الفراغ','القيم المميزة'],rows,[3,6.5,1.8,2,2,2.2],9,ltr_cols=(0,))

# ---------- 3 plan
doc.add_page_break(); H('3. خطة الجودة',1)
H('3.1 نطاق الخطة',2)
P('الهدف العام: رفع جودة بيانات الرحلات الجوية إلى مستوى يمكّن من قياس مؤشرات الانضباط في المواعيد وكفاءة الأسطول وأداء الناقلين بموثوقية، وبناء آلية مستدامة لمراقبة الجودة.')
P('الأهداف التفصيلية:',True)
for t in ['قياس خط الأساس لجودة البيانات على الأبعاد الستة.','حصر مشاكل الجودة وتوثيق أثرها وأسبابها الجذرية مع عينات من السجلات.','صياغة قواعد الأعمال وترجمتها إلى قواعد جودة قابلة للقياس الآلي.','معالجة المشاكل عالية الخطورة وتحقيق المستهدفات المعتمدة.','تشغيل مراقبة يومية للجودة وتسليمها لمكتب إدارة البيانات.']: B(t)
T(['البند','داخل النطاق','خارج النطاق'],[
 ['البيانات','ملف flights.csv كاملاً: 336,776 سجلاً و21 حقلاً','بيانات الطقس والركاب والحجوزات والتذاكر'],
 ['الفترة الزمنية','1 يناير – 31 ديسمبر 2013','الفترات الأخرى (إلى أن تُفعَّل آلية التحديث)'],
 ['الجهات','مطارات المغادرة EWR و JFK و LGA، و16 ناقلاً، و105 وجهات','رحلات الوصول إلى نيويورك'],
 ['الأنشطة','التحليل والتقييم، والقواعد، والمعالجة، والجداول المرجعية، والمراقبة','إعادة تصميم أنظمة الحجز والتشغيل لدى الناقلين (توصيات فقط)'],
 ['الأبعاد','الاكتمال والدقة والحداثة والاتساق والصلاحية والتوفر','بُعدا الأمن والخصوصية (لا توجد بيانات شخصية)'],
],[3,7,7],9.5)
P('الافتراضات والقيود:',True)
B('الجهة المالكة افتراضية، وهي "إدارة عمليات الرحلات"، ويُفترض توفر صلاحية الوصول إلى سجلات المصدر للتحقق.')
B('لعدم وجود قاموس بيانات، فُسرت معاني الحقول بالتحليل والمعرفة التشغيلية.')
B('التحقق من الدقة يعتمد على قواعد المعقولية لعدم وجود مصدر مرجعي خارجي في الملف.')

H('3.2 قائمة أولويات عناصر البيانات',2)
P('رُتبت العناصر بأربعة معايير مرجحة على مقياس من 1 إلى 5: الأثر على الأعمال (35%)، والاستخدام في التقارير والمؤشرات (25%)، والمتطلبات التنظيمية (15%)، وحجم المشاكل المرصودة (25%). عناصر الفئة "حرج" هي محور المعالجة في المرحلة الأولى.')
er=sorted(ELEMENTS,key=lambda e:-escore(e)); rows=[]
prev=None; rank=0
for k,e in enumerate(er,1):
    s=escore(e)
    if s!=prev: rank=k; prev=s
    rows.append([rank,e[0],e[1],e[2],e[3],e[4],e[5],f'{s:.2f}',etier(s)])
TC={'حرج (أولوية 1)':'F8CBAD','مهم (أولوية 2)':'FFE699','داعم (أولوية 3)':'E2EFDA'}
T(['الترتيب','العنصر','الوصف','الأثر','الاستخدام','التنظيمي','المشاكل','الدرجة','الفئة'],rows,[1.4,2.8,3.8,1.3,1.5,1.5,1.4,1.4,2.9],8.5,fills=lambda i,j,v: TC.get(v) if j==8 else None)

H('3.3 الأنشطة والأهداف المرحلية',2)
P('تُنفَّذ الخطة على 6 مراحل خلال 12 أسبوعاً. تنتهي كل مرحلة بمعلم (Milestone) يُعتمد قبل الانتقال إلى المرحلة التالية:')
T(['المرحلة','الأسابيع','الأنشطة الرئيسية','المعلم','معيار النجاح'],[list(p) for p in PHASES],[3.2,1.4,6,3.5,3.4],8.5)
P('المستهدفات المرحلية لمؤشرات الجودة:',True)
T(['البُعد','المؤشر','خط الأساس','المستهدف'],[[t[0],t[1],f'{t[2]:.2f}%',t[3]] for t in TARGETS],[2.5,8,3,3.5],9.5)

H('3.4 تقدير الموارد والميزانية وتخصيصها',2)
P('قُدرت التكلفة على أساس: العدد × عدد الأشهر × نسبة التخصيص × التكلفة الشهرية (بالريال السعودي). الأرقام تقديرية لأغراض التخطيط، وتُحسب تلقائياً في ورقة "الميزانية" بالملف الداعم.')
rows=[]; tot=0; cat={}
for b in BUDGET:
    v=b[2]*b[3]*b[4]*b[5]; tot+=v; cat[b[1]]=cat.get(b[1],0)+v
    rows.append([b[0],b[1],b[2],b[3],f'{int(b[4]*100)}%',fmtn(b[5]),fmtn(int(v))])
cont=tot*CONTINGENCY
rows+= [['المجموع الفرعي','','','','','',fmtn(int(tot))],['احتياطي الطوارئ (10%)','','','','','',fmtn(int(cont))],['الإجمالي الكلي','','','','','',fmtn(int(tot+cont))]]
T(['البند','الفئة','العدد','الأشهر','التخصيص','الشهري','الإجمالي'],rows,[5.5,2.6,1.2,1.2,1.6,2,2.4],8.5,fills=lambda i,j,v: 'DDEBF7' if i>=len(BUDGET) else None)
T(['الفئة','المبلغ (ريال)','النسبة'],[[k,fmtn(int(v)),f'{v/tot*100:.1f}%'] for k,v in cat.items()],[5,4,3],9.5)
P('تخصيص الميزانية على المراحل: تذهب نحو 15% إلى التخطيط والتقييم (المرحلتان 1 و2)، و15% إلى تعريف القواعد، و45% إلى المعالجة، و20% إلى المراقبة والأتمتة، و5% إلى الإغلاق والتدريب.')

H('3.5 بناء الفريق ومصفوفة المهام والمسؤوليات (RACI)',2)
T(['الرمز','الدور','الجهة','المسؤوليات الرئيسية'],[list(t) for t in TEAM],[1.5,4,4,7.5],9,ltr_cols=(0,))
P('مصفوفة RACI: الحرف R يعني المنفذ، و A المساءل والمعتمد (شخص واحد لكل نشاط)، و C المستشار، و I المُطَّلع.',True)
RF={'R':'BDD7EE','A':'F8CBAD','C':'FFE699','I':'F2F2F2'}
T(['النشاط']+ROLES,[[a]+[m[r] for r in ROLES] for a,m in RACI],[4.6]+[12.4/len(ROLES)]*len(ROLES),8,fills=lambda i,j,v: RF.get(v) if j>0 else None)

# ---------- 4 assessment
doc.add_page_break(); H('4. دراسة جودة البيانات وفق الأبعاد الستة',1)
P('طُبق التحليل على الملف كاملاً، أي على جميع السجلات وليس عينة منها، بلغة Python (مكتبة pandas). لكل بُعد عُرّفت قواعد قابلة للقياس، ونتيجة البُعد هي متوسط نسب النجاح لقواعده. التفاصيل الكاملة والسجلات المخالفة في الملف الداعم.')
DEF={'C':'مدى توفر القيم المطلوبة وعدم وجود فراغات في الحقول الإلزامية.','A':'مدى تمثيل القيم للواقع الفعلي، ويُقاس هنا بقواعد المعقولية التشغيلية.','T':'مدى حداثة البيانات وتوفرها في الوقت المطلوب لاتخاذ القرار.','S':'مدى انسجام القيم فيما بينها داخل السجل وبين السجلات، ومع البيانات المشتقة.','V':'مدى مطابقة القيم للصيغ والنطاقات والقوائم المرجعية المعتمدة.','Y':'مدى سهولة الوصول إلى البيانات وفهمها واستخدامها، مع البيانات الوصفية والمرجعية.'}
FIND={
'C':[f'اكتمال الخلايا {CELL_COMPLETE:.2f}%، والسجلات المكتملة بالكامل {ROWS_COMPLETE:.2f}%.','الحقول الأساسية (carrier و flight و origin و dest والتاريخ والأوقات المجدولة) مكتملة 100%.','نمط الفراغ منتظم: الرحلة الملغاة يغيب فيها dep_time و dep_delay معاً (8,255 رحلة)، والرحلات المحوّلة يغيب فيها arr_delay و air_time (1,175 رحلة).','رقم الطائرة مفقود في 2,512 سجلاً، كلها رحلات ملغاة.'],
'A':['قيم التأخير متسقة رياضياً مع الأوقات، لكن هناك 42 تأخيراً يتجاوز 600 دقيقة (أعلاها 1,301 دقيقة) يحتاج تحققاً.','4 رحلات بسرعة تفوق 600 ميل/ساعة (حتى 703 ميل/ساعة)، و352 رحلة أقل من 150 ميل/ساعة، معظمها إلى PHL.','35 حالة تعارض في دوران الطائرة، منها N12564 مجدولة للإقلاع من LGA و EWR في 20:00 من اليوم نفسه.'],
'T':['البيانات لقطة تاريخية لعام 2013 فقط، وعمرها نحو 13 سنة، فلا تصلح للقرارات التشغيلية الحالية.','التغطية الزمنية كاملة (365 يوماً) بمتوسط 923 رحلة يومياً، وانخفاض 28 نوفمبر (634 رحلة) منطقي لأنه يوم عطلة عيد الشكر.','لا يوجد ختم زمني لتحميل أو تحديث السجلات، فلا يمكن قياس زمن وصول البيانات.'],
'S':['الحقول المشتقة متسقة 100%: hour و minute و time_hour مع sched_dep_time، و name مع carrier.','قيمة dep_delay تطابق الفرق بين dep_time و sched_dep_time في جميع السجلات، لكن في 1,207 حالة تجاوز للمنتصف لا يُخزَّن التاريخ الفعلي.','تعارض في 38.7% من السجلات بين air_time والزمن المنقضي، بسبب التوقيت المحلي.','مسافتان مختلفتان لمسارَي EGE، و17 طائرة مسجلة لدى ناقلين. لا يوجد تكرار في مفتاح الرحلة.'],
'V':['رموز الناقلين والمطارات والتواريخ صالحة 100%.','179 قيمة وقت مسجلة 2400 بدلاً من 0000.','22,754 سجلاً (381 طائرة) بأرقام تسجيل لا تطابق صيغة FAA، مثل N0EGMQ و N5ETAA.'],
'Y':['الملف بصيغة مفتوحة (CSV) ويُقرأ كاملاً دون أخطاء.','لا يوجد قاموس بيانات يوضح الوحدات والمنطقة الزمنية ومعنى الفراغ.','لا توجد جداول مرجعية مرتبطة (الطائرات، المطارات ومناطقها الزمنية، المسارات).','لا توجد قناة وصول محكومة (قاعدة بيانات أو API) ولا اتفاقية مستوى خدمة.'],
}
for k,d in enumerate('CAVSTY',1):
    H(f'4.{k} {DIM[d]}: النتيجة {SCORES[d]:.2f}% ({status(SCORES[d],TG[d])})',2)
    RP([('التعريف: ',True),(DEF[d],False)])
    rr=[r for r in rule_rows if r['dim']==d]
    if rr: T(['القاعدة','الوصف','العنصر','المخالف','نسبة النجاح'],[[r['id'],r['rule'],r['el'],fmtn(r['fail']),f"{r['passrate']:.2f}%"] for r in rr],[1.5,8,3.5,1.8,2.2],8.5,fills=lambda i,j,v: ('F8CBAD' if v!='0' else 'C6EFCE') if j==3 else None)
    cc=[c for c in check_rows if c['dim']==d]
    if cc: T(['الضابط','الوصف','النتيجة','الدليل'],[[c['id'],c['rule'],'ناجح' if c['ok'] else 'مخالف',c['ev']] for c in cc],[1.5,6,1.8,7.7],8.5,fills=lambda i,j,v: {'ناجح':'C6EFCE','مخالف':'F8CBAD'}.get(v) if j==2 else None)
    P('النتائج:',True)
    for f in FIND[d]: B(f)

# ---------- 5 executive report
doc.add_page_break(); H('5. التقرير التنفيذي لمشاكل جودة البيانات',1)
P(f'يعرض هذا التقرير {len(ISSUES)} مشكلة جودة رُصدت في الملف، مرتبة حسب الخطورة. لكل مشكلة أبعادها وحجمها وأثرها على الأعمال وسببها الجذري المحتمل، مع عينة من السجلات الفعلية. معايير الخطورة: "عالية" إذا أثرت على مؤشر رئيسي أو على أكثر من 2% من السجلات، و"متوسطة" إذا أثرت على تحليلات تشغيلية محددة، و"منخفضة" إذا كان أثرها محدوداً.')
iss_sorted=sorted(ISSUES,key=lambda i:SEV_ORDER[i[4]])
rows=[]
for i in iss_sorted:
    c=issue_count(i); rows.append([i[0],i[1],'، '.join(DIM[d].split(' (')[0] for d in i[2]),i[4],fmtn(c) if c is not None else 'كامل الملف',f'{c/N*100:.2f}%' if c is not None else '—'])
T(['الرقم','المشكلة','الأبعاد','الخطورة','السجلات','النسبة'],rows,[1.6,7,3,1.7,2,1.7],8.5,fills=lambda i,j,v: SEVC.get(v) if j==3 else None)
for i in iss_sorted:
    H(f'{i[0]}: {i[1]}',3)
    c=issue_count(i)
    RP([('البُعد: ',True),('، '.join(DIM[d] for d in i[2]),False),('   |   الخطورة: ',True),(i[4],False),('   |   القواعد: ',True),('، '.join(i[3]),False)])
    RP([('الوصف: ',True),(i[5],False)])
    RP([('الأثر: ',True),(i[6],False)])
    RP([('السبب الجذري المحتمل: ',True),(i[7],False)])
    P('عينة من السجلات:',True)
    smp=samples(i)
    vals=[[('(فارغ)' if (isinstance(v,float) and pd.isna(v)) else (str(int(v)) if isinstance(v,float) and float(v).is_integer() and smp.columns[j]!='speed_mph' else str(v))) for j,v in enumerate(r)] for r in smp.itertuples(index=False)]
    ncol=len(smp.columns); w=17/ncol
    T(list(smp.columns),vals,[w]*ncol,7.5,hdr='595959',ltr_cols=tuple(range(ncol)))

# ---------- 6 BR
doc.add_page_break(); H('6. قواعد الأعمال وربطها بقواعد جودة البيانات',1)
P('صيغت قواعد الأعمال بلغة الأعمال بالتعاون المفترض مع مالك البيانات وأمين البيانات، ثم رُبطت كل قاعدة بقواعد جودة قابلة للقياس الآلي، مع إجراء الإصلاح الذي يعالج المخالفات. تُستخدم هذه القواعد للتصحيح الآن، وللمراقبة المستمرة بعد الإصلاح.')
rmap={x['id']:x for x in rule_rows}; cmap={x['id']:x for x in check_rows}
rows=[]
for b in BR:
    dims=sorted({(rmap.get(r) or cmap.get(r))['dim'] for r in b[2]},key='CAVSTY'.index)
    m=None
    for rr in b[2]:
        if rr in F: m=F[rr] if m is None else m|F[rr]
    cnt=fmtn(int(m.sum())) if m is not None else f"{sum(not cmap[r]['ok'] for r in b[2])} من {len(b[2])} ضوابط"
    rows.append([b[0],b[1],'، '.join(b[2]),'، '.join(DIM[d].split(' (')[0] for d in dims),cnt,b[3]])
T(['القاعدة','قاعدة العمل','قواعد الجودة','الأبعاد','المخالف حالياً','إجراء الإصلاح'],rows,[1.5,4.8,2,2,1.8,4.9],8)
P('سجل قواعد جودة البيانات التفصيلي:',True)
allr=[[r['id'],DIM[r['dim']].split(' (')[0],r['check'],r['thr'],fmtn(r['fail'])] for r in rule_rows]+[[c['id'],DIM[c['dim']].split(' (')[0],c['check'],c['thr'],'ناجح' if c['ok'] else 'مخالف'] for c in check_rows]
T(['القاعدة','البُعد','منطق الفحص','حد القبول','المخالف'],allr,[1.5,2,8.5,3,2],8,ltr_cols=(2,))

# ---------- 7 tasks
doc.add_page_break(); H('7. توزيع مهام إصلاح الجودة على الفريق',1)
P('وُزعت مهام الإصلاح على أعضاء الفريق وفق مصفوفة RACI، مع تحديد المنفذ والمعتمد والداعم والأسابيع والأولوية. الأولوية 1 تُنفذ أولاً لأنها تعالج المشاكل عالية الخطورة وعناصر الفئة الحرجة.')
rn={t[0]:t[1] for t in TEAM}
def nm(c): return '، '.join(rn.get(x.strip(),x.strip()) for x in c.split('،'))
PC={'1':'F8CBAD','2':'FFE699','3':'C6EFCE'}
T(['المشكلة','مهمة الإصلاح','المنفذ (R)','المعتمد (A)','الدعم (C)','الأسابيع','الأولوية'],[[t[0],t[1],nm(t[2]),nm(t[3]),nm(t[4]),t[5],t[6]] for t in TASKS],[1.5,6,2.4,2.2,2.6,1.4,1.2],8,fills=lambda i,j,v: PC.get(v) if j==6 else None)
P('حمل العمل حسب الدور:',True)
load={}
for t in TASKS: load.setdefault(t[2],[]).append(t[0])
T(['الدور','عدد المهام (منفذ)','المشاكل'],[[rn[k],len(v),'، '.join(v)] for k,v in sorted(load.items(),key=lambda x:-len(x[1]))],[5,3,9],9)

# ---------- 8 recs
H('8. التوصيات والخطوات التالية',1)
for t,l in [('إضافة حقل حالة الرحلة ',' وجعله إلزامياً في نظام المصدر، فهو أسرع إصلاح أثراً على مؤشرات الأداء.'),('تحويل جميع الأوقات إلى تاريخ ووقت كامل بتوقيت UTC',' مع الاحتفاظ بالتوقيت المحلي، لمعالجة مشكلتي الاتساق الأكبر (DQ-06 و DQ-07).'),('بناء جداول البيانات المرجعية',' (الطائرات بفترات سريان، والمطارات بمناطقها الزمنية، والمسارات) وربطها بالملف.'),('فرض ضوابط الإدخال في المصدر',' (صيغة رقم الطائرة، ونطاق الأوقات، والقوائم المرجعية) لمنع تكرار المشاكل بدلاً من تصحيحها لاحقاً.'),('تفعيل التحديث اليومي الآلي',' مع ختم زمني ولوحة مراقبة يومية لقواعد الجودة الـ 26 وحدود القبول.'),('اعتماد قاموس البيانات',' ونشر البيانات عبر قناة محكومة الصلاحيات، مع تحديد مالك وأمين لكل عنصر.')]:
    B(l,t)

H('ملحق: المنهجية والملف الداعم',1)
B('طُبقت جميع القواعد على كامل السجلات (336,776 سجلاً) باستخدام Python و pandas.')
B('الحالة: "محقق" إذا بلغت النتيجة المستهدف، و"يحتاج تحسين" إذا كانت بين 95% والمستهدف، و"حرج" إذا قلت عن 95%.')
B('الملف الداعم "تحليل_جودة_بيانات_الرحلات.xlsx" يتضمن: الملخص، والتحليل الوصفي، وقواعد الجودة ونتائجها، وسجل المشاكل، والعينات، والسجلات المخالفة (حتى 300 لكل قاعدة)، وأولويات العناصر، وقواعد الأعمال، والفريق و RACI، والأنشطة والمراحل، والميزانية (بمعادلات)، ومهام الإصلاح.')

# footer page numbers
for s in doc.sections:
    p=s.footer.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(); f1=OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'),'begin'); it=OxmlElement('w:instrText'); it.text='PAGE'; f2=OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'),'end')
    r._r.append(f1); r._r.append(it); r._r.append(f2)
from analysis import REPORTS_DIR
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
out = REPORTS_DIR / 'flights_dq_project_ar.docx'
doc.save(out); print('saved', out)
