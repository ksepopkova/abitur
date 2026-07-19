
import pandas as pd, re, io, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══ ПАРАМЕТРЫ (менять здесь) ═══
EMAIL_FROM = "result@vuzline.ru"
EMAIL_PASSWORD = "kpamwwoldfbqaooh"
TO_EMAIL = "ksepopkova@gmail.com"
subjects = {"Русский язык": 94, "Математика": 94, "Иностранный язык": 94, "Информатика": 88}
show_dvi = False
selected_cities = ["Москва и Московская область"]
gto_val = None
attestat = True
dvi_score = None
selected_areas = []
# ════════════════════════════════

OBL = {"Русский язык": 0, "Математика": 1, "Обществознание": 2, "История": 3, "Иностранный язык": 4, "Биология": 5, "Химия": 6, "Физика": 7, "Информатика": 8, "География": 9, "Литература": 10, "ДВИ": 11}
VYB = {"Математика": 12, "Обществознание": 13, "История": 14, "Иностранный язык": 15, "Биология": 16, "Химия": 17, "Физика": 18, "Информатика": 19, "География": 20, "Литература": 21}

AREA_GROUPS = {
    "IT и программирование": ["02", "09", "10"],
    "Математика и физика": ["01", "03"],
    "Инженерия и технологии": ["11", "12", "13", "14", "15", "16", "17", "22", "27", "28", "29"],
    "Химия и материалы": ["04", "18"],
    "Архитектура, строительство и дизайн": ["07", "08", "54"],
    "Транспорт и авиация": ["23", "24", "25", "26"],
    "Науки о земле и экология": ["05", "20", "21"],
    "Биология": ["06"],
    "Медицина и фармация": ["30", "31", "32", "33", "34"],
    "Сельское хозяйство и ветеринария": ["35", "36"],
    "Экономика и управление": ["38"],
    "Юриспруденция": ["40"],
    "Психология и социология": ["37", "39"],
    "Политология и медиа": ["41", "42"],
    "Педагогика": ["44"],
    "Лингвистика и филология": ["45"],
    "История и гуманитарные науки": ["46", "47", "48"],
    "Сервис и туризм": ["43"],
    "Физическая культура и спорт": ["49"],
    "Искусство и творчество": ["50", "51", "52", "53", "55"],
}

CANONICAL_COLS = ["Русский язык","Математика","Обществознание","История","Иностранный язык","Биология","Химия","Физика","Информатика","География","Литература","ДВИ","Математика.1","Обществознание.1","История.1","Иностранный язык.1","Биология.1","Химия.1","Физика.1","Информатика.1","География.1","Литература.1","Город","Вуз","Факультет (институт/школа)","Код и специальность \n(направление подготовки)","Образовательная программа (профиль)","Количество бюджетных мест 2026 (Всего)","Проходной балл бюджет 2025","Средний балл бюджет 2025","Число зачисленных 2025","Число зачисленных БВИ","ГТО Золото","ГТО Серебро","ГТО Бронза","Аттестат с отличием","Все индивидуальные достижения, за которые вуз добавляет баллы (ссылка)","Стоимость обучения в год, тыс \n(собираем после 1 июня)","Комментарий"]

def get_city_group(s):
    s = str(s).strip()
    if any(x in s for x in ["Москва","МО","Московская"]): return "Москва и Московская область"
    if any(x in s for x in ["Петербург","Ленинград"]): return "Санкт-Петербург и Ленинградская область"
    if "Новосибирск" in s: return "Новосибирская область"
    return s

def clean_str(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    return "" if s.lower()=="nan" else s

def to_num(val):
    s = str(val).strip()
    if s.upper()=="NEW": return "NEW"
    if s=="-": return "-"
    try:
        f=float(s); return int(f) if f==int(f) else f
    except: return None

def cell_has_value(row,col_idx):
    val=row.iloc[col_idx]
    if pd.isna(val): return False
    s=str(val).strip()
    if s in ("","nan"): return False
    try: return float(s)>0
    except: return True

def cell_value(row,col_idx):
    try: return float(row.iloc[col_idx])
    except: return 0

def has_budget_places(val):
    s=str(val).strip()
    if s in ("-","nan",""): return False
    m=re.match(r"(\d+)\s*\((\d+)\)",s)
    if m and int(m.group(2))==0: return True
    try: return float(s)>0
    except: return True

def is_quota_bvi(val):
    s=str(val).strip()
    m=re.match(r"(\d+)\s*\((\d+)\)",s)
    return bool(m and int(m.group(2))==0)

def is_no_competition(row):
    try: return float(row.iloc[28])==1 and float(row.iloc[29])==1 and float(row.iloc[30])==0
    except: return False

def calc_achievements(row,gto,attestat):
    total=0
    if gto=="Золото": total+=cell_value(row,32)
    elif gto=="Серебро": total+=cell_value(row,33)
    elif gto=="Бронза": total+=cell_value(row,34)
    if attestat: total+=cell_value(row,35)
    return min(total,10)

def calc_student_score(row,subjects):
    score=subjects.get("Русский язык",0)
    obl_set=set()
    for subj,col in OBL.items():
        if subj in ("Русский язык","ДВИ"): continue
        if cell_has_value(row,col):
            score+=subjects.get(subj,0); obl_set.add(subj)
    best_vyb=0
    for subj,col in VYB.items():
        if subj in obl_set: continue
        if cell_has_value(row,col):
            s=subjects.get(subj,0)
            if s>best_vyb: best_vyb=s
    return score+best_vyb

def check_row(row,subjects):
    rus=subjects.get("Русский язык",0)
    if not cell_has_value(row,OBL["Русский язык"]): return None
    if rus<cell_value(row,OBL["Русский язык"]): return None
    obl_required=[(s,c) for s,c in OBL.items() if s not in ("Русский язык","ДВИ") and cell_has_value(row,c)]
    for subj,col in obl_required:
        thresh=cell_value(row,col)
        if thresh>0 and subjects.get(subj,0)<thresh: return None
    dvi_req=cell_has_value(row,OBL["ДВИ"])
    obl_set={s for s,_ in obl_required}
    vyb_req=[(s,c) for s,c in VYB.items() if s not in obl_set and cell_has_value(row,c)]
    if vyb_req:
        if not any(subjects.get(s,0)>0 and subjects.get(s,0)>=cell_value(row,c) for s,c in vyb_req): return None
    return "with_dvi" if dvi_req else "no_dvi"

def is_valid(v):
    try: return float(v)>1
    except: return False

def get_chance(score,pb,sb):
    pb_ok=is_valid(pb); sb_ok=is_valid(sb)
    if not pb_ok and not sb_ok: return "new"
    if pb_ok and sb_ok:
        pb_f,sb_f=float(pb),float(sb)
        if score>=sb_f+10: return "podstrahovka"
        if score>=sb_f: return "realistic"
        if score>pb_f+5: return "probable"
        if score>=pb_f-15: return "risky"
        return "unlikely"
    pb_f=float(pb)
    if score>=pb_f+25: return "podstrahovka"
    if score>=pb_f+10: return "realistic"
    if score>pb_f+5: return "probable"
    if score>=pb_f-15: return "risky"
    return "unlikely"

CHANCE_ORDER={"probable":0,"realistic":1,"podstrahovka":2,"risky":3,"unlikely":4,"quota_bvi":5,"no_competition":6,"new":7,"no_dvi_score":8}
CHANCE_LABEL={"podstrahovka":"🟢 Уверенно","realistic":"🔵 Реалистично","probable":"🟡 Вероятно","risky":"🔴 Рискованно","unlikely":"⚫ Маловероятно","quota_bvi":"🔹 Квоты и БВИ","no_competition":"◾ Общего конкурса не было","new":"⬜ Нет данных","no_dvi_score":"⬜ Нет оценки — не указан балл за ДВИ"}
PRIORITY_LABEL={"podstrahovka":"3–5","realistic":"2–3","probable":"1–2","risky":"1*","unlikely":"—","quota_bvi":"—","no_competition":"1*","new":"1*","no_dvi_score":"—"}

def build_row(row,subjects,gto,attestat,dvi_score=None):
    pb,sb=row.iloc[28],row.iloc[29]
    stu=calc_student_score(row,subjects)
    ach=calc_achievements(row,gto,attestat)
    tot=stu+ach
    dvi_req=cell_has_value(row,OBL["ДВИ"])
    if dvi_req and dvi_score and dvi_score>0: tot+=dvi_score
    if is_quota_bvi(row.iloc[27]): chance="quota_bvi"
    elif is_no_competition(row): chance="no_competition"
    elif dvi_req: chance="no_dvi_score" if not dvi_score else get_chance(tot,pb,sb)
    else: chance=get_chance(tot,pb,sb)
    return {"Город":clean_str(row.iloc[22]),"Вуз":clean_str(row.iloc[23]),"Факультет":clean_str(row.iloc[24]),"Код и специальность":clean_str(row.iloc[25]),"Профиль":clean_str(row.iloc[26]),"Мест":to_num(row.iloc[27]),"Проходной балл":to_num(pb),"Средний балл":to_num(sb),"Ваш балл (ЕГЭ)":stu,"Достижения":ach if ach>0 else "","Конкурсный балл":tot,"Шансы":chance,"Рек. приоритет":PRIORITY_LABEL[chance],"Стоимость обучения (Москва и СПб), тыс руб":clean_str(row.iloc[37]) or "—"}

VUZ_RATING={"МГУ им. Ломоносова":1,"МГТУ им. Баумана":2,"МФТИ":3,"СПБГУ":4,"МИФИ":5,"ВШЭ":6,"МГИМО":7,"РАНХиГС":8,"Политех им. Петра Великого":9,"ИТМО":15,"НГУ":16,"МИСиС":17,"МАИ":22,"МЭИ":24,"РУТ (МИИТ)":75}

def get_rating(vuz):
    for k,v in VUZ_RATING.items():
        if k in vuz: return v
    return 999

print("Загружаем базу...")
dfs=[]
for sheet_name in ["Москва","Питер","Регионы"]:
    df=pd.read_excel("База вузов 2026.xlsx",sheet_name=sheet_name,header=1)
    df=df.iloc[:,:len(CANONICAL_COLS)].copy()
    df.columns=CANONICAL_COLS
    dfs.append(df)
full=pd.concat(dfs,ignore_index=True)
print(f"Строк: {len(full)}")

results=[]
for _,row in full.iterrows():
    cg=get_city_group(row.iloc[22])
    if cg not in selected_cities: continue
    if not has_budget_places(row.iloc[27]): continue
    st=check_row(row,subjects)
    if st is None: continue
    if st=="with_dvi" and not show_dvi: continue
    results.append(build_row(row,subjects,gto_val,attestat,dvi_score))

result=pd.DataFrame(results)
print(f"Найдено: {len(result)}")

result["_co"]=result["Шансы"].map(CHANCE_ORDER)
result=result.sort_values(["Город","Вуз","_co"]).drop("_co",axis=1)
result["Шансы"]=result["Шансы"].map(CHANCE_LABEL)

CP={"🟡 Вероятно":0,"🔵 Реалистично":1,"🟢 Уверенно":2,"🔴 Рискованно":3,"⚫ Маловероятно":4,"🔹 Квоты и БВИ":5,"◾ Общего конкурса не было":6,"⬜ Нет данных":7,"⬜ Нет оценки — не указан балл за ДВИ":8}
result["_cp"]=result["Шансы"].map(CP)
result["_pb"]=pd.to_numeric(result["Проходной балл"],errors="coerce").fillna(0)
result=result.sort_values(["Город","Вуз","_cp","_pb"],ascending=[True,True,True,False]).drop(columns=["_pb"])

good_zones={"🟡 Вероятно","🔵 Реалистично","🟢 Уверенно"}
rc=result[~((pd.to_numeric(result["Проходной балл"],errors="coerce")<=1)&(pd.to_numeric(result["Средний балл"],errors="coerce")<=1))]
vgc=rc[rc["Шансы"].isin(good_zones)].groupby("Вуз").size()
v3=vgc[vgc>=3]
min_g=1 if len(v3)<3 else 3
main_vuz=vgc[vgc>=min_g].sort_values(ascending=False)
few_vuz=vgc[(vgc>=1)&(vgc<min_g)].sort_values(ascending=False)

def sort_by_rating(vlist):
    rated=sorted([v for v in vlist if get_rating(v)<999 and main_vuz.get(v,0)>=4],key=get_rating)
    unrated=[v for v in vlist if v not in rated]
    return rated+unrated

def build_block(vlist):
    rows=[]
    for vuz in vlist:
        vdf=result[result["Вуз"]==vuz].copy()
        for _,r in vdf.iterrows():
            rows.append(r)
    df_out=pd.DataFrame(rows).reset_index(drop=True) if rows else pd.DataFrame()
    if "_cp" in df_out.columns: df_out=df_out.drop(columns=["_cp"])
    return df_out

main_vuz_list=list(main_vuz.index)
if selected_areas:
    area_prefixes=set()
    for area in selected_areas:
        for prefix in AREA_GROUPS.get(area,[]):
            area_prefixes.add(prefix)
    good_result=rc[rc["Шансы"].isin(good_zones)]
    def vuz_in_area(vuz):
        vuz_codes=good_result[good_result["Вуз"]==vuz]["Код и специальность"]
        return any(str(c).split(".")[0] in area_prefixes for c in vuz_codes)
    area_vuz=[v for v in main_vuz_list if vuz_in_area(v)]
    backup_vuz=[v for v in main_vuz_list if not vuz_in_area(v)]
    top_area=sort_by_rating(area_vuz)[:7]
    top_backup=sort_by_rating(backup_vuz)[:7]
else:
    top_area=sort_by_rating(main_vuz_list)[:7]
    top_backup=[]

result_main=build_block(top_area)
result_backup=build_block(top_backup) if top_backup else pd.DataFrame()

few_rows=[]
few_statuses = good_zones | {"🔴 Рискованно"}
for vuz in list(few_vuz.index):
    vdf=result[(result["Вуз"]==vuz)&(result["Шансы"].isin(few_statuses))].copy()
    for _,r in vdf.iterrows():
        few_rows.append(r)
result_few=pd.DataFrame(few_rows).reset_index(drop=True) if few_rows else pd.DataFrame()
if "_cp" in result_few.columns: result_few=result_few.drop(columns=["_cp"])

shown=set(result_main["Вуз"].unique()) if len(result_main)>0 else set()
shown|=set(result_backup["Вуз"].unique()) if len(result_backup)>0 else set()
shown|=set(result_few["Вуз"].unique()) if len(result_few)>0 else set()
dvi_vuz=set(result[result["Шансы"]=="⬜ Нет оценки — не указан балл за ДВИ"]["Вуз"].unique())
result_dvi=result[result["Вуз"].isin(dvi_vuz)&~result["Вуз"].isin(shown)].copy()
if "_cp" in result_dvi.columns: result_dvi=result_dvi.drop(columns=["_cp"])
shown_all=shown|set(result_dvi["Вуз"].unique() if len(result_dvi)>0 else [])
result_risky_only=result[(result["Шансы"]=="🔴 Рискованно")&~result["Вуз"].isin(shown_all)].copy()
if "_cp" in result_risky_only.columns: result_risky_only=result_risky_only.drop(columns=["_cp"])

good_in_main=result_main["Шансы"].isin(good_zones).sum() if len(result_main)>0 else 0
good_in_backup=result_backup["Шансы"].isin(good_zones).sum() if len(result_backup)>0 else 0
if good_in_main+good_in_backup==0:
    print("Нет хороших вариантов — показываем всё")
    result_main=build_block(list(result["Вуз"].unique()))
    result_backup=pd.DataFrame(); result_few=pd.DataFrame(); result_dvi=pd.DataFrame()

frames=[d for d in [result_main,result_backup,result_few,result_dvi,result_risky_only] if len(d)>0]
processed=pd.concat(frames,ignore_index=True) if frames else pd.DataFrame()
print(f"Итого в таблице: {len(processed)}")

e2r={"🟢 Уверенно":"podstrahovka","🔵 Реалистично":"realistic","🟡 Вероятно":"probable","🔴 Рискованно":"risky","⚫ Маловероятно":"unlikely","🔹 Квоты и БВИ":"quota_bvi","◾ Общего конкурса не было":"no_competition","⬜ Нет данных":"new","⬜ Нет оценки — не указан балл за ДВИ":"no_dvi_score"}
processed["Шансы"]=processed["Шансы"].map(lambda x:e2r.get(x,x))

CC={"Уверенно":"D6EED2","Реалистично":"D0E8F5","Вероятно":"FFF0D6","Рискованно":"FAE0E0","Маловероятно":"EBEBEB","Квоты и БВИ":"D7EAF3","Общего конкурса не было":"D8D8D8","Нет данных":"F5F5F5","Нет оценки — не указан балл за ДВИ":"F5F5F5"}
CT={"Уверенно":"1E6B14","Реалистично":"0D5A8A","Вероятно":"8A5A00","Рискованно":"8A1A1A","Маловероятно":"555555","Квоты и БВИ":"1A5C7A","Общего конкурса не было":"4A4A4A","Нет данных":"888888","Нет оценки — не указан балл за ДВИ":"888888"}
HB="379FFC"; HF="FFFFFF"; RO="FFFFFF"; RE="F4F7FB"; AC="379FFC"
thin_g=Side(style="thin",color="DDDDDD"); no_s=Side(style=None); dashed=Side(style="dashDot",color="8BB8E8")

def gborder(ci,bottom=True):
    b=thin_g if bottom else no_s
    l=dashed if ci==6 else no_s
    r=dashed if ci in [1,6,8] else no_s
    return Border(bottom=b,left=l,right=r)

wb=Workbook()
ws3=wb.active; ws3.title="Важная информация"; ws3.column_dimensions["A"].width=90
info=[("📋 Результаты во вкладке Результаты","title"),("Ниже — как читать таблицу и расставлять приоритеты.","subtitle"),("","gap"),("О СЕРВИСЕ","header"),("Данные актуальны на 2026 год. Носят рекомендательный характер.","text"),("Только для поступающих на общих основаниях на основном этапе.","text"),("","gap"),("РАСШИФРОВКА ШАНСОВ","header"),("Уверенно — балл выше среднего на 10+. Рек. приоритет: 3–5+","D6EED2"),("Реалистично — балл выше среднего. Рек. приоритет: 2–3","D0E8F5"),("Вероятно — балл ниже среднего, но выше проходного на 5+. Рек. приоритет: 1–2","FFF0D6"),("Рискованно — балл близко к проходному. Рек. приоритет: 1*","FAE0E0"),("Маловероятно — балл ниже проходного на 15+.","EBEBEB"),("Квоты и БВИ — места для общего конкурса могут не остаться. Известно 3 августа.","D7EAF3"),("Нет данных — новая специальность.","F5F5F5"),("","gap"),("СРОКИ","header"),("До 25 июля — подача документов.","text"),("27 июля — конкурсные списки.","text"),("5 августа 12:00 мск — последний срок согласия на зачисление.","text"),("","gap"),("Консультация: https://vuzline.ru/consultation | result@vuzline.ru","contact")]
for rn,(text,kind) in enumerate(info,1):
    c=ws3.cell(row=rn,column=1,value=text)
    c.alignment=Alignment(wrap_text=True,vertical="center")
    if kind=="title": c.font=Font(name="Montserrat",size=16,bold=True,color=HB); ws3.row_dimensions[rn].height=42
    elif kind=="subtitle": c.font=Font(name="Montserrat",size=10,color="555555"); ws3.row_dimensions[rn].height=22
    elif kind=="gap": ws3.row_dimensions[rn].height=10
    elif kind=="header": c.font=Font(name="Montserrat",size=10,bold=True,color=HB); c.fill=PatternFill("solid",fgColor="EBF4FF"); ws3.row_dimensions[rn].height=24
    elif kind=="contact": c.font=Font(name="Montserrat",size=9,bold=True,color=AC); ws3.row_dimensions[rn].height=36
    elif len(kind)==6: c.font=Font(name="Montserrat",size=9); c.fill=PatternFill("solid",fgColor=kind); ws3.row_dimensions[rn].height=20
    else: c.font=Font(name="Montserrat",size=9); ws3.row_dimensions[rn].height=18

ws=wb.create_sheet("Результаты")
df_out=processed.copy()
cl={"podstrahovka":"Уверенно","realistic":"Реалистично","probable":"Вероятно","risky":"Рискованно","unlikely":"Маловероятно","quota_bvi":"Квоты и БВИ","no_competition":"Общего конкурса не было","new":"Нет данных","no_dvi_score":"Нет оценки — не указан балл за ДВИ"}
if "Шансы" in df_out.columns: df_out["Шансы"]=df_out["Шансы"].map(lambda x:cl.get(x,x))
cols=list(df_out.columns)
chance_i=cols.index("Шансы")+1 if "Шансы" in cols else -1
prio_i=cols.index("Рек. приоритет")+1 if "Рек. приоритет" in cols else -1
for ci,cn in enumerate(cols,1):
    c=ws.cell(row=1,column=ci,value=cn)
    c.font=Font(bold=True,color=HF,name="Montserrat",size=9)
    c.fill=PatternFill("solid",fgColor=HB)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    c.border=gborder(ci,bottom=False)
ws.row_dimensions[1].height=32
for ri,row_ in enumerate(df_out.itertuples(index=False),start=2):
    row_=pd.Series(row_._asdict())
    bg=RO if ri%2==0 else RE
    for ci,(cn,value) in enumerate(row_.items(),1):
        if value is None or (isinstance(value,float) and pd.isna(value)): cv=""
        elif isinstance(value,(int,float)) and not isinstance(value,bool): cv=value
        else:
            s=str(value); cv="" if s in ("None","nan") else s
        c=ws.cell(row=ri,column=ci,value=cv)
        c.fill=PatternFill("solid",fgColor=bg)
        c.border=gborder(ci)
        if ci in [1,2,3,4,5]: c.font=Font(name="Montserrat",size=9,bold=(ci==1)); c.alignment=Alignment(horizontal="left",vertical="center")
        elif ci==chance_i:
            cc=CC.get(str(cv),bg); tt=CT.get(str(cv),"1A1A1A")
            c.fill=PatternFill("solid",fgColor=cc); c.font=Font(name="Montserrat",size=9,color=tt,bold=True); c.alignment=Alignment(horizontal="center",vertical="center")
        elif ci==prio_i: c.font=Font(name="Montserrat",size=9,color=AC,bold=True); c.alignment=Alignment(horizontal="center",vertical="center")
        else: c.font=Font(name="Montserrat",size=9); c.alignment=Alignment(horizontal="center",vertical="center")
ws.freeze_panes="A2"
cw={1:14,2:18,3:22,4:32,5:30,6:7,7:13,8:13,9:13,10:10,11:13,12:18,13:11}
for ci,w in cw.items(): ws.column_dimensions[get_column_letter(ci)].width=w

ws2=wb.create_sheet("Запрос")
sp={"Email":TO_EMAIL,"Предметы":str(subjects),"Города":str(selected_cities),"ГТО":str(gto_val),"Аттестат":str(attestat),"Области":str(selected_areas),"Примечание":"Переотправка с полной выдачей"}
for ci,(k,v) in enumerate(sp.items(),1):
    ws2.cell(row=1,column=ci,value=k).font=Font(bold=True,color=HF,name="Montserrat",size=9)
    ws2.cell(row=1,column=ci).fill=PatternFill("solid",fgColor=HB)
    ws2.cell(row=2,column=ci,value=str(v)).font=Font(name="Montserrat",size=9)
    ws2.column_dimensions[get_column_letter(ci)].width=25

buf=io.BytesIO(); wb.save(buf); buf.seek(0)
msg=MIMEMultipart()
msg["From"]=EMAIL_FROM; msg["To"]=TO_EMAIL; msg["Subject"]="Ваша таблица подбора вузов — Vuzline (обновлённая)"
body="Здравствуйте!\n\nНаправляем обновлённую таблицу — в неё вошли все доступные варианты по вашему запросу, включая программы, которые не попали в предыдущую версию.\n\nСпасибо, что обратили внимание на неполноту — благодаря вам мы исправили ошибку.\n\nУдачи с поступлением!\nКоманда Vuzline\nresult@vuzline.ru"
msg.attach(MIMEText(body,"plain","utf-8"))
att=MIMEBase("application","octet-stream"); att.set_payload(buf.read()); encoders.encode_base64(att)
att.add_header("Content-Disposition","attachment; filename=\"vuzline_results.xlsx\"")
msg.attach(att)
print(f"Отправляем на {TO_EMAIL}...")
with smtplib.SMTP_SSL("smtp.yandex.ru",465) as s:
    s.login(EMAIL_FROM,EMAIL_PASSWORD); s.sendmail(EMAIL_FROM,TO_EMAIL,msg.as_string())
print("Готово!")
