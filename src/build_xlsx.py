from content import *
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
FONT='Arial'
HDR=PatternFill('solid',fgColor='1F4E79'); SUB=PatternFill('solid',fgColor='DDEBF7')
thin=Side(style='thin',color='BFBFBF'); B=Border(left=thin,right=thin,top=thin,bottom=thin)
SEVF={'عالية':'F8CBAD','متوسطة':'FFE699','منخفضة':'C6EFCE'}
wb=Workbook()
def sheet(title, first=False):
    ws=wb.active if first else wb.create_sheet()
    ws.title=title; ws.sheet_view.rightToLeft=True; return ws
def header(ws,row,cols,widths=None):
    for j,c in enumerate(cols,1):
        cell=ws.cell(row=row,column=j,value=c); cell.font=Font(name=FONT,bold=True,color='FFFFFF'); cell.fill=HDR
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); cell.border=B
    if widths:
        for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes=ws.cell(row=row+1,column=1)
def put(ws,row,vals,fmt=None,fill=None,wrap=True):
    for j,v in enumerate(vals,1):
        c=ws.cell(row=row,column=j,value=v); c.font=Font(name=FONT); c.border=B
        c.alignment=Alignment(vertical='top',wrap_text=wrap,horizontal='right' if isinstance(v,str) else 'center')
        if fmt and j in fmt: c.number_format=fmt[j]
        if fill: c.fill=PatternFill('solid',fgColor=fill)
def title(ws,text,sub=None):
    ws['A1']=text; ws['A1'].font=Font(name=FONT,bold=True,size=14,color='1F4E79')
    if sub: ws['A2']=sub; ws['A2'].font=Font(name=FONT,italic=True,color='595959')

# 1 summary
ws=sheet('الملخص',True); title(ws,'ملخص تقييم جودة بيانات الرحلات (flights.csv)','مصدر البيانات: رحلات المغادرة من مطارات نيويورك (EWR, JFK, LGA) لعام 2013. تاريخ التقييم: 2026-11-18')
kp=[('إجمالي السجلات',N,'#,##0'),('عدد الحقول',21,'0'),('عدد الناقلين',df.carrier.nunique(),'0'),('عدد الوجهات',df.dest.nunique(),'0'),('عدد الطائرات (tailnum)',df.tailnum.nunique(),'#,##0'),('اكتمال الخلايا',CELL_COMPLETE/100,'0.00%'),('السجلات المكتملة بالكامل',ROWS_COMPLETE/100,'0.00%'),('عدد مشاكل الجودة المرصودة',len(ISSUES),'0')]
header(ws,4,['المؤشر','القيمة'],[34,20,16,16,60])
for i,(k,v,f) in enumerate(kp,5): put(ws,i,[k,v],fmt={2:f})
r0=5+len(kp)+1
header(ws,r0,['بُعد الجودة','النتيجة الحالية','المستهدف','الحالة','طريقة القياس'])
ws.freeze_panes=None
tgt={'C':0.999,'A':0.9999,'T':1,'S':0.999,'V':1,'Y':1}
meth={'C':'متوسط نسب النجاح لقواعد R01–R04','A':'متوسط نسب النجاح لقواعد R15–R17','T':'نسبة ضوابط R19–R21 المحققة','S':'متوسط نسب النجاح لقواعد R09–R14 و R18','V':'متوسط نسب النجاح لقواعد R05–R08','Y':'نسبة ضوابط R22–R26 المحققة'}
for i,d in enumerate(['C','A','T','S','V','Y'],r0+1):
    put(ws,i,[DIM[d],round(SCORES[d]/100,6),tgt[d],None,meth[d]],fmt={2:'0.00%',3:'0.00%'})
    ws.cell(row=i,column=4,value=f'=IF(B{i}>=C{i},"محقق",IF(B{i}>=0.95,"يحتاج تحسين","حرج"))').font=Font(name=FONT,bold=True)
    ws.cell(row=i,column=4).border=B; ws.cell(row=i,column=4).alignment=Alignment(horizontal='center')
rng=f'D{r0+1}:D{r0+6}'
ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"حرج"'],fill=PatternFill('solid',fgColor='F8CBAD')))
ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"يحتاج تحسين"'],fill=PatternFill('solid',fgColor='FFE699')))
ws.conditional_formatting.add(rng,CellIsRule(operator='equal',formula=['"محقق"'],fill=PatternFill('solid',fgColor='C6EFCE')))
i=r0+8; ws.cell(row=i,column=1,value='المؤشر العام (متوسط الأبعاد الستة)').font=Font(name=FONT,bold=True)
ws.cell(row=i,column=2,value=f'=AVERAGE(B{r0+1}:B{r0+6})').number_format='0.00%'
SUMROW=(r0+1,r0+6)

# 2 profiling
ws=sheet('التحليل الوصفي'); title(ws,'التحليل الوصفي للحقول (Data Profiling)')
cols=['الحقل','النوع','القيم الموجودة','القيم الفارغة','نسبة الفراغ','القيم المميزة','أصغر قيمة','أكبر قيمة','القيمة الأكثر تكراراً','ملاحظة']
header(ws,3,cols,[16,10,14,12,11,12,22,22,26,46])
notes={'dep_time':'فراغ = رحلة ملغاة ضمنياً؛ يحتوي 2400','arr_time':'يحتوي 2400؛ بالتوقيت المحلي للوجهة','arr_delay':'فراغ يشمل الملغاة والمحوّلة','air_time':'مطابق لفراغ arr_delay','tailnum':'2,512 فارغ؛ 22,754 بصيغة غير صالحة','dest':'105 وجهة','distance':'مسافة غير موحدة لمسار EGE','time_hour':'نص وليس تاريخاً','id':'معرف فني تسلسلي','name':'مكرر من carrier (اشتقاقي)'}
for i,c in enumerate(ORIG.columns,4):
    s=ORIG[c]; vc=s.value_counts()
    put(ws,i,[c,str(s.dtype).replace('str','object'),int(s.notna().sum()),int(s.isna().sum()),s.isna().mean(),int(s.nunique()),
              str(s.min()) if s.dtype==object or str(s.dtype)=='str' else float(s.min()),str(s.max()) if s.dtype==object or str(s.dtype)=='str' else float(s.max()),
              f'{vc.index[0]} ({vc.iloc[0]:,})',notes.get(c,'')],fmt={3:'#,##0',4:'#,##0',5:'0.00%',6:'#,##0'})

# 3 rules
ws=sheet('قواعد الجودة'); title(ws,'قواعد جودة البيانات ونتائج القياس','عدد السجلات الكلي في الخلية B2 ويُستخدم في حساب نسب النجاح')
ws['B2']=N; ws['B2'].number_format='#,##0'; ws['A2']='إجمالي السجلات'
header(ws,4,['رقم القاعدة','البُعد','العنصر','وصف القاعدة','منطق الفحص','حد القبول','السجلات المخالفة','نسبة النجاح','الحالة'],[11,20,24,50,46,16,14,12,12])
r=5
for x in rule_rows:
    put(ws,r,[x['id'],DIM[x['dim']],x['el'],x['rule'],x['check'],x['thr'],x['fail'],None,None],fmt={7:'#,##0'})
    ws.cell(row=r,column=8,value=f'=1-G{r}/$B$2').number_format='0.00%'; ws.cell(row=r,column=8).border=B
    ws.cell(row=r,column=9,value=f'=IF(G{r}=0,"ناجح","مخالف")').border=B; r+=1
for x in check_rows:
    put(ws,r,[x['id'],DIM[x['dim']],x['el'],x['rule'],x['check']+'  |  الدليل: '+x['ev'],x['thr'],'—',1 if x['ok'] else 0,None],fmt={8:'0%'})
    ws.cell(row=r,column=9,value=f'=IF(H{r}=1,"ناجح","مخالف")').border=B; r+=1
ws.conditional_formatting.add(f'I5:I{r}',CellIsRule(operator='equal',formula=['"مخالف"'],fill=PatternFill('solid',fgColor='F8CBAD')))
ws.conditional_formatting.add(f'I5:I{r}',CellIsRule(operator='equal',formula=['"ناجح"'],fill=PatternFill('solid',fgColor='C6EFCE')))

# 4 issues
ws=sheet('سجل المشاكل'); title(ws,'سجل مشاكل جودة البيانات (Issue Log)')
header(ws,3,['رقم المشكلة','المشكلة','الأبعاد','القواعد المرتبطة','الخطورة','عدد السجلات المتأثرة','النسبة','الوصف','الأثر على الأعمال','السبب الجذري المحتمل'],[11,34,26,14,10,14,10,55,55,40])
for i,iss in enumerate(ISSUES,4):
    cnt=issue_count(iss)
    put(ws,i,[iss[0],iss[1],'، '.join(DIM[d] for d in iss[2]),'، '.join(iss[3]),iss[4],cnt if cnt is not None else 'كامل الملف',(cnt/N) if cnt is not None else None,iss[5],iss[6],iss[7]],fmt={6:'#,##0',7:'0.00%'})
    ws.cell(row=i,column=5).fill=PatternFill('solid',fgColor=SEVF[iss[4]])

# 5 samples
ws=sheet('عينات المشاكل'); title(ws,'عينات حقيقية من الملف لكل مشكلة')
r=3
for iss in ISSUES:
    smp=samples(iss)
    ws.cell(row=r,column=1,value=f'{iss[0]}: {iss[1]}').font=Font(name=FONT,bold=True,color='1F4E79',size=12); r+=1
    for j,c in enumerate(smp.columns,1):
        cell=ws.cell(row=r,column=j,value=c); cell.font=Font(name=FONT,bold=True); cell.fill=SUB; cell.border=B
    r+=1
    for _,row in smp.iterrows():
        for j,v in enumerate(row.values,1):
            v=None if (isinstance(v,float) and pd.isna(v)) or v is pd.NA else (v.item() if hasattr(v,'item') else v)
            c=ws.cell(row=r,column=j,value=v if v is not None else '(فارغ)'); c.font=Font(name=FONT,color='C00000' if v is None else '000000'); c.border=B
        r+=1
    r+=1
for j in range(1,13): ws.column_dimensions[get_column_letter(j)].width=15

# 6 failing records
ws=sheet('السجلات المخالفة'); title(ws,'السجلات المخالفة لكل قاعدة (أول 300 سجل لكل قاعدة)')
kc=['id','year','month','day','carrier','flight','tailnum','origin','dest','sched_dep_time','dep_time','dep_delay','sched_arr_time','arr_time','arr_delay','air_time','distance','speed_mph']
header(ws,3,['رقم القاعدة']+kc,[11]+[11]*len(kc))
r=4
for k in [x['id'] for x in rule_rows]:
    sub=df[F[k]].head(300)
    for _,row in sub[kc].iterrows():
        vals=[k]+[None if pd.isna(v) else (v.item() if hasattr(v,'item') else v) for v in row.values]
        for j,v in enumerate(vals,1): ws.cell(row=r,column=j,value=v)
        r+=1
ws.auto_filter.ref=f'A3:{get_column_letter(len(kc)+1)}{r-1}'

# 7 priorities
ws=sheet('أولويات العناصر'); title(ws,'قائمة أولويات عناصر البيانات','الدرجة = الأثر×35% + الاستخدام×25% + المتطلبات التنظيمية×15% + حجم المشاكل×25% (مقياس 1–5)')
ws['G1']='الأوزان'; ws['G2']=W[0]; ws['H2']=W[1]; ws['I2']=W[2]; ws['J2']=W[3]
header(ws,4,['العنصر','الوصف','الأثر على الأعمال','الاستخدام في المؤشرات','المتطلبات التنظيمية','حجم المشاكل المرصودة','الدرجة المرجحة','الترتيب','الفئة'],[18,28,14,14,14,14,14,10,18])
n=len(ELEMENTS)
for i,e in enumerate(sorted(ELEMENTS,key=lambda e:-escore(e)),5):
    put(ws,i,list(e[:6])+[None,None,None])
    ws.cell(row=i,column=7,value=f'=C{i}*$G$2+D{i}*$H$2+E{i}*$I$2+F{i}*$J$2').number_format='0.00'
    ws.cell(row=i,column=8,value=f'=RANK(G{i},$G$5:$G${4+n})')
    ws.cell(row=i,column=9,value=f'=IF(G{i}>=4,"حرج (أولوية 1)",IF(G{i}>=3,"مهم (أولوية 2)","داعم (أولوية 3)"))')
    for j in (7,8,9): ws.cell(row=i,column=j).border=B

# 8 BR mapping
ws=sheet('قواعد الأعمال'); title(ws,'ربط قواعد الأعمال بقواعد جودة البيانات وإجراءات الإصلاح')
header(ws,3,['قاعدة العمل','نص قاعدة العمل','قواعد الجودة المرتبطة','الأبعاد','السجلات المخالفة حالياً','إجراء الإصلاح'],[10,55,22,30,14,60])
rmap={x['id']:x for x in rule_rows}; cmap={x['id']:x for x in check_rows}
for i,b in enumerate(BR,4):
    dims=sorted({(rmap.get(r) or cmap.get(r))['dim'] for r in b[2]},key='CAVSTY'.index)
    m=None
    for rr in b[2]:
        if rr in F: m=F[rr] if m is None else m|F[rr]
    cnt=int(m.sum()) if m is not None else '%d من %d ضوابط مخالفة'%(sum(not cmap[r]['ok'] for r in b[2]),len(b[2]))
    put(ws,i,[b[0],b[1],'، '.join(b[2]),'، '.join(DIM[d] for d in dims),cnt,b[3]],fmt={5:'#,##0'})

# 9 team & RACI
ws=sheet('الفريق و RACI'); title(ws,'الفريق المقترح ومصفوفة المهام والمسؤوليات (RACI)','R = منفذ، A = مساءل/معتمد، C = مستشار، I = مُطّلع')
header(ws,4,['الرمز','الدور','الجهة','المسؤوليات الرئيسية'],[10,28,26,60])
ws.freeze_panes=None
for i,t in enumerate(TEAM,5): put(ws,i,list(t))
r0=6+len(TEAM)
for j,c in enumerate(['النشاط']+ROLES,1):
    cell=ws.cell(row=r0,column=j,value=c); cell.font=Font(name=FONT,bold=True,color='FFFFFF'); cell.fill=HDR; cell.alignment=Alignment(horizontal='center'); cell.border=B
RF={'R':'BDD7EE','A':'F8CBAD','C':'FFE699','I':'EDEDED'}
for i,(act,m) in enumerate(RACI,r0+1):
    ws.cell(row=i,column=1,value=act).border=B
    for j,role in enumerate(ROLES,2):
        c=ws.cell(row=i,column=j,value=m[role]); c.alignment=Alignment(horizontal='center'); c.fill=PatternFill('solid',fgColor=RF[m[role]]); c.border=B; c.font=Font(name=FONT,bold=m[role] in 'RA')
for j in range(5,14): ws.column_dimensions[get_column_letter(j)].width=8

# 10 phases
ws=sheet('الأنشطة والمراحل'); title(ws,'الأنشطة والأهداف المرحلية (12 أسبوعاً)')
header(ws,3,['المرحلة','الأسابيع','الأنشطة الرئيسية','المعلم (Milestone)','معيار النجاح'],[30,10,60,40,34])
for i,p in enumerate(PHASES,4): put(ws,i,list(p))
r=5+len(PHASES)
header(ws,r,['البُعد','المؤشر','خط الأساس الحالي','المستهدف نهاية المشروع'])
ws.freeze_panes=ws['A4']
for i,t in enumerate(TARGETS,r+1): put(ws,i,[t[0],t[1],t[2]/100,t[3]],fmt={3:'0.00%'})

# 11 budget
ws=sheet('الميزانية'); title(ws,'تقدير الموارد والميزانية (ريال سعودي)','التكلفة = العدد × الأشهر × نسبة التخصيص × التكلفة الشهرية')
header(ws,4,['البند','الفئة','العدد','الأشهر','نسبة التخصيص','التكلفة الشهرية','الإجمالي'],[44,16,8,8,12,16,16])
for i,b in enumerate(BUDGET,5):
    put(ws,i,list(b)+[None],fmt={5:'0%',6:'#,##0'})
    ws.cell(row=i,column=7,value=f'=C{i}*D{i}*E{i}*F{i}').number_format='#,##0'; ws.cell(row=i,column=7).border=B
e=4+len(BUDGET)
rows=[('المجموع الفرعي',f'=SUM(G5:G{e})'),('احتياطي الطوارئ (10%)',f'=G{e+1}*0.1'),('الإجمالي الكلي',f'=G{e+1}+G{e+2}')]
for k,(lab,f) in enumerate(rows,e+1):
    ws.cell(row=k,column=1,value=lab).font=Font(name=FONT,bold=True); c=ws.cell(row=k,column=7,value=f); c.number_format='#,##0'; c.font=Font(name=FONT,bold=True)
k=e+5; header(ws,k,['التوزيع حسب الفئة','','','','','','الإجمالي']); ws.freeze_panes=ws['A5']
for q,cat in enumerate(['موارد بشرية','أدوات وتقنية','بيانات مرجعية','تدريب'],k+1):
    ws.cell(row=q,column=1,value=cat); c=ws.cell(row=q,column=7,value=f'=SUMIF(B5:B{e},A{q},G5:G{e})'); c.number_format='#,##0'
    c2=ws.cell(row=q,column=8,value=f'=G{q}/G${e+1}'); c2.number_format='0.0%'

# 12 tasks
ws=sheet('مهام الإصلاح'); title(ws,'توزيع مهام إصلاح الجودة على الفريق')
header(ws,3,['المشكلة','مهمة الإصلاح','المنفذ (R)','المعتمد (A)','الدعم (C)','الأسابيع','الأولوية','الحالة'],[10,70,22,22,22,10,9,12])
rn={t[0]:t[1] for t in TEAM}
def nm(c): return '، '.join(rn.get(x.strip(),x.strip()) for x in c.split('،'))
for i,t in enumerate(TASKS,4): put(ws,i,[t[0],t[1],nm(t[2]),nm(t[3]),nm(t[4]),t[5],int(t[6]),'لم تبدأ'])
from analysis import REPORTS_DIR
REPORTS_DIR.mkdir(parents=True, exist_ok=True)
out = REPORTS_DIR / 'flights_dq_analysis_ar.xlsx'
wb.save(out); print('saved', out)
print('ok', SUMROW)
