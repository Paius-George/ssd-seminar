import os,sys,copy,zipfile,collections,datetime,json,posixpath,warnings
from lxml import etree as E
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.chart import LineChart,BarChart,PieChart,Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.data_source import NumData,NumVal,StrData,StrVal
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.layout import Layout,ManualLayout
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph,ParagraphProperties,CharacterProperties
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D,XDRPositiveSize2D
from openpyxl.utils import get_column_letter as L
warnings.filterwarnings('ignore')
ROOT=Path= os.getcwd()
SRC='S6_Dashborad_demo (2).xlsx';OUT=sys.argv[1] if len(sys.argv)>1 else 'S6_Dashborad_demo (2) - conform PDF.xlsx'
orig=zipfile.ZipFile(SRC)
wc=openpyxl.load_workbook(SRC,data_only=True)
w=openpyxl.load_workbook(SRC)
rows=[list(r[:12]) for r in wc['Vanzari_02'].values if isinstance(r[0],int)]
headers=list(next(wc['Vanzari_02'].values))[:12]+['Categorie_grupata','Ani','Trimestre']
for r in rows:r.extend(['Tonere' if r[4]=='Tonere' else 'Imprimante',r[1].year,'Qtr'+str((r[1].month-1)//3+1)])
TOTAL=sum(r[5] for r in rows)
BLUE='4F81BD';RED='C0504D';PALE='B8CCE4';DARK='1F497D'
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships';pkg='http://schemas.openxmlformats.org/package/2006/relationships';x14='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
def el(tag,attrs=None,parent=None,uri=ns,**kw):
 a={k:str(v) for k,v in (attrs or {}).items()};a.update({k:str(v) for k,v in kw.items()});n=E.Element('{%s}%s'%(uri,tag),a) if parent is None else E.SubElement(parent,'{%s}%s'%(uri,tag),a);return n
def xml(n):return E.tostring(n,xml_declaration=True,encoding='UTF-8',standalone=True)
values=[]
for f in range(15):
 vals=sorted(set(r[f] for r in rows),key=lambda x:x)
 values.append(vals)
# Native grouped date members include boundary buckets.
values[1]=['<07.01.2016','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','>23.12.2018']
values[13]=['<07.01.2016',2016,2017,2018,'>23.12.2018']
values[14]=['<07.01.2016','Qtr1','Qtr2','Qtr3','Qtr4','>23.12.2018']
def ix(f,v):return values[f].index(v)
def fieldval(r,f):return values[1][r[1].month] if f==1 else r[f]
def setup(s):
 s.sheet_view.showGridLines=True;s.sheet_view.zoomScale=100;s.sheet_properties.pageSetUpPr.fitToPage=True
 s.column_dimensions['A'].width=27
 for c in range(2,22):s.column_dimensions[L(c)].width=13
def title(s,text,end=10):
 s.merge_cells(start_row=1,start_column=1,end_row=1,end_column=end);c=s.cell(1,1,text);c.fill=PatternFill('solid',fgColor=DARK);c.font=Font(name='Calibri',size=14,bold=True,color='FFFFFF');c.alignment=Alignment(vertical='center');s.row_dimensions[1].height=31
def style_range(s,r1,c1,r2,c2,headerrows=1,total=True):
 for rr in range(r1,r2+1):
  for cc in range(c1,c2+1):
   c=s.cell(rr,cc);c.font=Font(name='Calibri',size=11,color='000000');c.alignment=Alignment(vertical='center');c.fill=PatternFill()
   if rr<r1+headerrows or (total and rr==r2):c.fill=PatternFill('solid',fgColor=BLUE if rr<r1+headerrows else 'DCE6F1');c.font=Font(name='Calibri',size=11,bold=True,color='FFFFFF' if rr<r1+headerrows else '000000')
   if isinstance(c.value,(int,float)):c.number_format='0'
specs=[]
def makepivot(sheet,name,rf,cf,dfs,pages=(),start=None,sort=None,top=None,subtotals=False,showheaders=True):
 s=w[sheet] if sheet in w.sheetnames else w.create_sheet(sheet)
 if s.max_row==1 and s['A1'].value is None:setup(s)
 if start is None:start=(5 if len(pages)==3 else 4 if len(pages)==2 else 3,1)
 sr,sc=start
 for p,f in enumerate(pages):s.cell(sr-len(pages)-1+p,sc,headers[f]);s.cell(sr-len(pages)-1+p,sc+1,'(All)')
 keys=sorted(set(tuple(fieldval(r,f) for f in rf) for r in rows)) if rf else [()]
 if 1 in rf:keys.sort(key=lambda k:tuple(values[f].index(v) if f==1 else v for f,v in zip(rf,k)))
 if sort:keys.sort(key=lambda k:sum(r[5] for r in rows if tuple(r[f] for f in rf)==k),reverse=sort=='descending')
 if top:
  keys=sorted(keys,key=lambda k:sum(r[5] for r in rows if tuple(r[f] for f in rf)==k),reverse=True)[:top]
  keys.sort(key=lambda k:sum(r[5] for r in rows if tuple(r[f] for f in rf)==k),reverse=sort=='descending')
 if sheet=='Chart_jud_prod':keys.sort(key=lambda k:(['NV','BI','NE','V','SV','C','SE','S'].index(k[0]),k[1]))
 ckeys=sorted(set(tuple(r[f] for f in cf) for r in rows)) if cf else [()]
 # row tuple + subtotal depth (None means detail)
 rkeys=[]
 if subtotals and len(rf)>1:
  for k in keys:
   if not rkeys or rkeys[-1][0][:1]!=k[:1]:rkeys.append((k[:1],1))
   rkeys.append((k,None))
 else:rkeys=[(k,None) for k in keys]
 if rf==[13,14,1]:
  rkeys=[]
  for k in sorted(set(k[:2] for k in keys)):
   if not rkeys or rkeys[-1][0][0]!=k[0]:rkeys.append((k[:1],1))
   rkeys.append((k,None))
 rkeys.append(((),0))
 colkeys=[(k,di) for k in ckeys for di in range(len(dfs))]
 if cf:colkeys.extend(((),di) for di in range(len(dfs)))
 nh=2 if cf else 1
 nr=1
 s.cell(sr+nh-1,sc,'Row Labels' if showheaders else None)
 for j,(k,di) in enumerate(colkeys):
  c=sc+nr+j
  if cf and len(dfs)==1:
   s.cell(sr+nh-1,c,' / '.join(map(str,k)) if k else 'Grand Total')
  else:
   if cf:s.cell(sr,c,' / '.join(map(str,k)) if k else 'Total '+dfs[di][0])
   s.cell(sr+nh-1,c,dfs[di][0])
 if cf and len(dfs)==1:
  s.cell(sr,sc,dfs[0][0])
  if showheaders:s.cell(sr,sc+1,'Column Labels')
 grand=sum(r[5] for r in rows if not top or tuple(r[f] for f in rf) in keys)
 for i,(k,depth) in enumerate(rkeys):
  rr=sr+nh+i
  if depth==0:s.cell(rr,sc,'Grand Total')
  else:
   s.cell(rr,sc,k[-1]);s.cell(rr,sc).alignment=Alignment(indent=max(0,len(k)-1))
  subset=[r for r in rows if (not k or tuple(fieldval(r,f) for f in rf[:len(k)])==k) and (not top or tuple(r[f] for f in rf) in keys)]
  for j,(ck,di) in enumerate(colkeys):
   df,field,agg,percent=dfs[di];rrs=[r for r in subset if not ck or tuple(r[f] for f in cf)==ck]
   v=sum(r[field] for r in rrs)
   if agg=='average':v=v/len(rrs) if rrs else 0
   if percent:v=v/grand if grand else 0
   c=s.cell(rr,sc+nr+j,v);c.number_format='0.0%' if percent or field==11 else '0'
  if depth:
   for cc in range(sc,sc+nr+len(colkeys)):s.cell(rr,cc).fill=PatternFill('solid',fgColor='DEE8F4');s.cell(rr,cc).font=Font(bold=True,color=DARK,size=10)
 endr=sr+nh+len(rkeys)-1;endc=sc+nr+len(colkeys)-1
 style_range(s,sr,sc,endr,endc,nh)
 for i,(k,depth) in enumerate(rkeys):
  rr=sr+nh+i
  if k:s.cell(rr,sc).alignment=Alignment(indent=max(0,len(k)-1),vertical='center')
  if depth:
   for cc in range(sc,endc+1):s.cell(rr,cc).fill=PatternFill('solid',fgColor='DCE6F1');s.cell(rr,cc).font=Font(name='Calibri',bold=True,size=11)
 for rr in range(sr+nh,endr+1):
  for j,(_,di) in enumerate(colkeys):
   if dfs[di][3] or dfs[di][1]==11:s.cell(rr,sc+nr+j).number_format='0.00%' if sheet in ['Chart_reg_prod','Chart_pie'] else '0.0%'
 spec=dict(sheet=sheet,name=name,rf=list(rf),cf=list(cf),dfs=dfs,pages=list(pages),start=start,nh=nh,nr=nr,keys=keys,rkeys=rkeys,ckeys=ckeys,colkeys=colkeys,endr=endr,endc=endc,sort=sort,top=top,subtotals=subtotals,showheaders=showheaders)
 specs.append(spec);return spec
money=[('Vanzari(lei)',5,'sum',False)]
percent=[('Procent',5,'sum',True)]
sumcaption=[('Sum of Valoarea tranzactiei',5,'sum',False)]
p0=makepivot('PivotTable0','PivotTable0',[2,12,4],[8],sumcaption,[7],subtotals=True)
p1=makepivot('PivotTable1','PivotTable1',[13,14,1],[8],sumcaption,[],subtotals=True)
# Month data use the grouped base date field rather than individual dates.
# Rebuild calendar analysis separately below from monthly source records.
p2=makepivot('PivotTable2','PivotTable2',[8,9],[12],money+percent,[7],subtotals=True)
ps=makepivot('Pivot_sortat','Pivot_sortat',[8],[12],money+percent,[7],sort='descending')
pd=makepivot('Chart_data','Pivot_data',[13,14],[12],sumcaption,[8],showheaders=True)
pr=makepivot('Chart_reg_prod','Pivot_reg_prod',[8],[12],percent,[2],sort='ascending',showheaders=True)
pj=makepivot('Chart_jud_prod','Pivot_jud_prod',[8,9],[12],money,[2],showheaders=True)
pp=makepivot('Chart_pie','Pivot_pie',[12],[],percent,[2,8],showheaders=True)
pa=makepivot('Chart_ag10','Pivot_ag10',[7],[],[('Valoarea tranzactiilor',5,'sum',False),('Comision mediu',11,'average',False)],[2,8,9],sort='ascending',top=10,showheaders=True)
py=makepivot('Chart_prod_ani','Pivot_prod_ani',[4],[2],money,[],showheaders=False)
# Show Details for the printer sales in 2016, as in the grouping demonstration.
s=w.create_sheet('Detalii_Imprimante_2016');setup(s)
for c,v in enumerate(headers[:12],1):s.cell(3,c,v)
detail=[r for r in rows if r[2]==2016 and r[12]=='Imprimante']
for i,r in enumerate(detail,4):
 for c,v in enumerate(r[:12],1):s.cell(i,c,v)
 s.cell(i,2).number_format='dd.mm.yyyy';s.cell(i,12).number_format='0%'
style_range(s,3,1,3+len(detail),12,total=False);s.auto_filter.ref=f'A3:L{3+len(detail)}'
def chart(spec,kind,title_,width=20,height=11):
 s=w[spec['sheet']];sr,sc=spec['start'];nr=spec['nr'];nh=spec['nh'];end=spec['endr']-1
 ch={'line':LineChart,'bar':BarChart,'pie':PieChart}[kind]();ch.title=title_ if spec['sheet'] in ['Chart_reg_prod','Chart_jud_prod'] else None;ch.width=width;ch.height=height;ch.style=None
 ch.graphical_properties=GraphicalProperties(solidFill='FFFFFF',ln=openpyxl.drawing.line.LineProperties(solidFill='7F7F7F',w=6350))
 data_cols=2 if spec['cf'] else 1
 if kind=='pie':data_cols=1
 ch.add_data(Reference(s,min_col=sc+nr,max_col=sc+nr+data_cols-1,min_row=sr+nh-1,max_row=end),titles_from_data=True)
 ch.set_categories(Reference(s,min_col=sc+nr-1,min_row=sr+nh,max_row=end))
 if kind!='pie':
  ch.legend.position='t';ch.y_axis.title=None;ch.x_axis.title=None
  ch.y_axis.majorGridlines=None;ch.x_axis.majorGridlines=None
  for ax in [ch.x_axis,ch.y_axis]:
   ax.txPr=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=850)),endParaRPr=CharacterProperties(sz=850))]);ax.majorTickMark='none';ax.minorTickMark='none'
   ax.spPr=GraphicalProperties(ln=openpyxl.drawing.line.LineProperties(solidFill='868686',w=6350))
  if spec['dfs'][0][3]:ch.y_axis.numFmt='0.00%';ch.y_axis.scaling.min=0;ch.y_axis.scaling.max=0.18;ch.y_axis.majorUnit=0.02
  if kind=='line':ch.y_axis.scaling.min=0;ch.y_axis.scaling.max=35000;ch.y_axis.majorUnit=5000;ch.y_axis.numFmt='0'
  for i,se in enumerate(ch.series):
   se.graphicalProperties.solidFill=[BLUE,RED][i%2];se.graphicalProperties.line.solidFill=[BLUE,RED][i%2]
   if spec['cf']:se.tx=SeriesLabel(v=str(spec['ckeys'][i][0]))
   if kind=='line':se.marker.symbol=['diamond','square'][i%2];se.marker.size=5;se.marker.graphicalProperties.solidFill=[BLUE,RED][i%2];se.marker.graphicalProperties.line.solidFill=[BLUE,RED][i%2];se.graphicalProperties.line.width=19050
 else:
  from openpyxl.chart.series import DataPoint
  ch.series[0].data_points=[DataPoint(idx=0,spPr=openpyxl.chart.shapes.GraphicalProperties(solidFill=BLUE)),DataPoint(idx=1,spPr=openpyxl.chart.shapes.GraphicalProperties(solidFill=RED))]
  ch.dataLabels=DataLabelList();ch.dataLabels.showPercent=True;ch.dataLabels.showVal=False;ch.dataLabels.numFmt='0.0%';ch.dataLabels.dLblPos='ctr';ch.legend=None;ch.firstSliceAng=0
 if kind=='bar':
  ch.type='bar' if spec['sheet'] in ['Chart_jud_prod','Chart_ag10'] else 'col';ch.grouping='clustered' if ch.type=='bar' else 'stacked';ch.overlap=0 if ch.type=='bar' else 100
  if ch.type=='bar':
   ch.x_axis.scaling.orientation='minMax';ch.y_axis.crosses='autoZero';ch.y_axis.numFmt='0';ch.gapWidth=100
  else:ch.gapWidth=150
  if spec['sheet']=='Chart_ag10':ch.legend=None
 ch._pivot_spec=spec['name']
 return ch
charts={}
for sp,kind,ttl in [(pd,'line',None),(pr,'bar','Procentul vanzarilor - analiza regionala'),(pj,'bar','Evolutia teritoriala a vanzarilor'),(pp,'pie',None),(pa,'bar',None)]:
 ch=chart(sp,kind,ttl,height=29 if sp==pj else 11);w[sp['sheet']].add_chart(ch,'F3' if sp!=pa else 'E2');charts[sp['sheet']]=ch
ds=w.create_sheet('Dashboard',0);setup(ds)
SCALE=1.6
def pxwidth(px):return (px-5)/7
for c,px in zip('ABCDEFGH',[180,6,148,62,62,62,70,412]):ds.column_dimensions[c].width=pxwidth(px)
pdf_heights={1:32,**{r:12.1 for r in range(2,12)},12:15,**{r:142/12 for r in range(13,25)},25:15,26:12,27:12,**{r:12 for r in range(28,34)},34:15,**{r:12 for r in range(35,61)}}
for r,h in pdf_heights.items():ds.row_dimensions[r].height=h*SCALE*0.75
for rr in range(1,61):
 for cc in range(1,9):ds.cell(rr,cc).font=Font(name='Calibri',size=11,color='000000')
ds.merge_cells('A1:H1');ds['A1']='RAPORT PRIVIND VANZARILE FIRMEI Export3D.SRL';ds['A1'].font=Font(name='Calibri',size=11,bold=True);ds['A1'].alignment=Alignment(horizontal='center',vertical='center')
for c in ds[1][:8]:c.fill=PatternFill('solid',fgColor=PALE)
for rng,txt in [('C12:G12','Evolutia regionala a vanzarilor'),('H12:H12','Evolutia teritoriala a vanzarilor'),('C25:G25','Evolutia vanzarilor pe categorii de produse'),('C34:G34','Top 10 agenti vanzari')]:
 if rng.split(':')[0]!=rng.split(':')[1]:ds.merge_cells(rng)
 c=ds[rng.split(':')[0]];c.value=txt;c.fill=PatternFill('solid',fgColor=PALE);c.font=Font(name='Calibri',bold=True,color='000000',size=11);c.alignment=Alignment(vertical='center')
def absanchor(x,y,width,height):return AbsoluteAnchor(pos=XDRPoint2D(round(x*SCALE*9525),round(y*SCALE*9525)),ext=XDRPositiveSize2D(round(width*SCALE*9525),round(height*SCALE*9525)))
def dashchart(key,x,y,width,height,plot):
 ch=copy.deepcopy(charts[key]);ch.title=None;ch.legend=None;ch.anchor=absanchor(x,y,width,height);ch.width=width*SCALE*0.0264583333;ch.height=height*SCALE*0.0264583333
 ch.layout=Layout(manualLayout=ManualLayout(layoutTarget='inner',xMode='factor',yMode='factor',x=plot[0],y=plot[1],w=plot[2],h=plot[3]))
 if key in ['Chart_jud_prod','Chart_ag10']:ch.y_axis.delete=True
 if key=='Chart_ag10':ch.x_axis.txPr=RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=750)),endParaRPr=CharacterProperties(sz=750))])
 ds.add_chart(ch)
dashchart('Chart_pie',116,32,107,119,(0.03,0.05,0.94,0.90))
dashchart('Chart_data',224,32,400,119,(0.10,0.04,0.88,0.66))
dashchart('Chart_reg_prod',116,168,252,140,(0.15,0.03,0.83,0.86))
dashchart('Chart_jud_prod',369,168,255,570,(0.31,0.01,0.67,0.98))
dashchart('Chart_ag10',116,436,252,195,(0.29,0.03,0.69,0.95))
pyd=makepivot('Dashboard','Pivot_dashboard_prod_ani',[4],[2],money,[],start=(26,3),showheaders=False)
for row in ds.iter_rows(min_row=26,max_row=33,min_col=3,max_col=7):
 for c in row:c.font=Font(name='Calibri',size=10,bold=c.row in [26,27,33],color='FFFFFF' if c.row<28 else '000000');c.alignment=Alignment(horizontal='right' if c.column>3 else 'left',vertical='center')
ds['C33'].alignment=Alignment(horizontal='left',vertical='center');ds['C33'].border=Border(top=Side(style='thin',color=BLUE))
ds.sheet_view.showGridLines=False;ds.sheet_view.zoomScale=100;ds.freeze_panes=None;ds.print_area='A1:H60';ds.sheet_properties.pageSetUpPr.fitToPage=True;ds.page_setup.orientation='portrait';ds.page_setup.paperSize=ds.PAPERSIZE_A3;ds.page_setup.fitToWidth=1;ds.page_setup.fitToHeight=1
guide=w.create_sheet('Ghid_si_verificari');setup(guide);title(guide,'Verificarea rezolvarii din PDF',8)
notes=[
('1.1 / pag. 1–2','PivotTable0','Ani, grupari/categorii produse, regiuni; filtru Nume_complet_agent; suma tranzactiilor.'),
('1.2 / pag. 3','PivotTable0','Grupare nativa: Imprimanta 3D, laser, InkJet si Multifunctionale -> Imprimante; Tonere separat.'),
('1.2 / pag. 3','Detalii_Imprimante_2016','Show Details pentru vanzarile de imprimante din 2016: '+str(sum(r[5] for r in detail))+' lei.'),
('1.2 / pag. 3–4','PivotTable1','Gruparea nativa a datei pe ani, trimestre si luni; regiuni pe coloane.'),
('1.3 / pag. 4–5','PivotTable2','Regiune si judet; Imprimante/Tonere; suma si procent din totalul general; formatare numerica.'),
('1.4 / pag. 5','Pivot_sortat','Sortare descrescatoare a regiunilor dupa vanzarile totale.'),
('2 / pag. 6','Chart_data / Pivot_data','Grafic liniar pe ani si trimestre, doua grupe principale de produse.'),
('2 / pag. 6','Chart_reg_prod / Pivot_reg_prod','Grafic regional, procente din total, sortat crescator dupa total; filtru an.'),
('2 / pag. 6–7','Chart_jud_prod / Pivot_jud_prod','Vanzari pe regiuni si judete, grupe principale de produse; filtru an.'),
('2 / pag. 7','Chart_pie / Pivot_pie','Structura procentuala Imprimante/Tonere; filtre an si regiune.'),
('2 / pag. 7','Chart_ag10 / Pivot_ag10','Top 10 dinamic dupa valoarea vanzarilor; media ratelor de comision; filtre an, regiune, judet.'),
('2 / pag. 8','Chart_prod_ani / Pivot_prod_ani','Tabel vanzari pe categorii originale si ani; fara filtre de raport initiale.'),
('3 / pag. 8–9','Dashboard','Cinci grafice si tabel Pivot de sinteza, aliniate; patru slicere native Excel.'),
('3 / pag. 8–9','Conexiuni slicere','An, Regiune, Categorie_produs -> toate cele sase rapoarte si tabelul din Dashboard. Agent -> aceleasi, exceptand Top 10.'),
('Pag. 9','Actualizare','PivotCache refreshOnLoad activ. In Excel: Date > Reimprospatare totala, sau click dreapta pe Pivot > Refresh.'),
('Clarificare perioada','2016–2018','PDF mentioneaza 2016–2028 in introducere; datele si toate capturile sunt din 2016–2018.'),
('Clarificare categorii','5 categorii originale','PDF spune patru categorii in paranteza; capturile si datele contin cinci: patru tipuri de imprimante plus Tonere.'),
('Clarificare comision','Media ratelor','Comisionul mediu este media procentelor din coloana Comision, conform capturii din PDF, nu o suma in lei.')]
for j,v in enumerate(['Referinta PDF','Foaie / element','Implementare'],1):guide.cell(3,j,v)
for i,n in enumerate(notes,4):
 for j,v in enumerate(n,1):guide.cell(i,j,v).alignment=Alignment(wrap_text=True,vertical='top')
 guide.row_dimensions[i].height=43
guide.column_dimensions['A'].width=23;guide.column_dimensions['B'].width=43;guide.column_dimensions['C'].width=110
style_range(guide,3,1,3+len(notes),3,total=False)
for rr in guide.iter_rows(min_row=4,max_col=3):
 for c in rr:c.alignment=Alignment(wrap_text=True,vertical='top')
guide.cell(24,1,'Total vanzari');guide.cell(24,2,TOTAL);guide.cell(25,1,'Numar tranzactii');guide.cell(25,2,len(rows))
guide.cell(26,1,'Total Imprimante');guide.cell(26,2,sum(r[5] for r in rows if r[12]=='Imprimante'))
guide.cell(27,1,'Total Tonere');guide.cell(27,2,sum(r[5] for r in rows if r[12]=='Tonere'))
guide.cell(29,1,'Validare efectuata');guide.cell(29,3,'Structura verificata cu Microsoft Open XML SDK pentru Office 2016. Calculele, relatiile, pastrarea foilor sursa si conexiunile slicerelor sunt verificate separat. Interactiunea in aplicatia Microsoft Excel nu a putut fi testata in acest mediu.').alignment=Alignment(wrap_text=True,vertical='top');guide.row_dimensions[29].height=48
w.remove(guide)
w.active=0;w.calculation=openpyxl.workbook.properties.CalcProperties(calcId=191029,fullCalcOnLoad=True)
w.save('lucru_dashboard/base.xlsx')
z=zipfile.ZipFile('lucru_dashboard/base.xlsx');parts={n:z.read(n) for n in z.namelist()}
# Restore source cells, formula results, validation, and original charts byte for byte.
# Source sheets moved to indices 2 and 3 after inserting Dashboard.
for old,new in [(1,2),(2,3)]:
 parts[f'xl/worksheets/sheet{new}.xml']=orig.read(f'xl/worksheets/sheet{old}.xml')
 parts[f'xl/worksheets/_rels/sheet{new}.xml.rels']=orig.read(f'xl/worksheets/_rels/sheet{old}.xml.rels')
for n in orig.namelist():
 if n.startswith(('xl/drawings/','xl/charts/')) or n=='xl/sharedStrings.xml':
  # Dashboard was first in output, so avoid collisions by relocating original drawing/chart assets.
  pass
# Original drawing/worksheet relation targets are relocated to a unique preserved namespace.
for old,new in [(1,2),(2,3)]:
 rp=f'xl/worksheets/_rels/sheet{new}.xml.rels';rr=E.fromstring(parts[rp])
 for e in rr:
  target=e.get('Target');e.set('Target',target.replace('drawing','originalDrawing'))
 parts[rp]=xml(rr)
for n in orig.namelist():
 if n.startswith('xl/drawings/'):
  nn=n.replace('drawing','originalDrawing');d=orig.read(n)
  if n.endswith('.rels'):d=d.replace(b'chart',b'originalChart')
  parts[nn]=d
 elif n.startswith('xl/charts/'):parts[n.replace('chart','originalChart')]=orig.read(n)
parts['xl/sharedStrings.xml']=orig.read('xl/sharedStrings.xml')
ct=E.fromstring(parts['[Content_Types].xml']);ctns='http://schemas.openxmlformats.org/package/2006/content-types'
def content(path,typ):
 if not any(e.get('PartName')=='/'+path for e in ct):el('Override',{'PartName':'/'+path,'ContentType':typ},ct,uri=ctns)
content('xl/sharedStrings.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')
for n in parts:
 if n.startswith('xl/originalDrawings/') and n.endswith('.xml'):content(n,'application/vnd.openxmlformats-officedocument.drawing+xml')
 if n.startswith('xl/originalCharts/') and n.endswith('.xml'):content(n,'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
# The replacement above changes the folder as well as the filenames.
wb=E.fromstring(parts['xl/workbook.xml']);wr=E.fromstring(parts['xl/_rels/workbook.xml.rels'])
origwb=E.fromstring(orig.read('xl/workbook.xml'))
for d in origwb.findall('.//{%s}definedName'%ns):
 if d.get('name')=='_xlnm.Criteria':
  d=copy.deepcopy(d);d.set('localSheetId',str(int(d.get('localSheetId'))+1));wb.find('{%s}definedNames'%ns).append(d)
def addrel(root,typ,target,rid=None):
 rid=rid or 'rId'+str(max([int(e.get('Id')[3:]) for e in root if e.get('Id','').startswith('rId')]+[0])+1)
 el('Relationship',{'Id':rid,'Type':typ,'Target':target},root,uri=pkg);return rid
addrel(wr,rel+'/sharedStrings','sharedStrings.xml')
rid=addrel(wr,rel+'/pivotCacheDefinition','pivotCache/pivotCacheDefinition1.xml')
pcaches=el('pivotCaches');el('pivotCache',{'cacheId':1,'{%s}id'%rel:rid},pcaches)
calc=wb.find('{%s}calcPr'%ns);wb.insert(list(wb).index(calc)+1,pcaches)
cache=el('pivotCacheDefinition',{'{%s}id'%rel:'rId1','saveData':1,'refreshOnLoad':1,'enableRefresh':1,'createdVersion':6,'refreshedVersion':6,'minRefreshableVersion':3,'recordCount':len(rows)})
cs=el('cacheSource',{'type':'worksheet'},cache);el('worksheetSource',{'ref':'A1:L191','sheet':'Vanzari_02'},cs)
cfs=el('cacheFields',{'count':15},cache)
actual_dates=sorted(set(r[1] for r in rows))
for f,h in enumerate(headers):
 cf=el('cacheField',{'name':h,'numFmtId':14 if f==1 else 0,'databaseField':0 if f>=12 else 1},cfs)
 vals=actual_dates if f==1 else values[f]
 if f<12:
  typ='d' if f==1 else ('n' if isinstance(vals[0],(int,float)) else 's')
  si=el('sharedItems',{'count':len(vals),'containsSemiMixedTypes':0,'containsNonDate':int(f!=1),'containsDate':int(f==1),'containsString':int(typ=='s'),'containsNumber':int(typ=='n'),'containsInteger':int(typ=='n' and f!=11)},cf)
  for v in vals:el(typ,{'v':v.isoformat() if typ=='d' else v},si)
 if f==4:el('fieldGroup',{'par':12},cf)
 if f in [1,12,13,14]:
  fg=el('fieldGroup',{'base':4 if f==12 else 1,**({'par':14} if f==1 else {'par':13} if f==14 else {})},cf)
  if f==12:
   dp=el('discretePr',{'count':len(values[4])},fg)
   for v in values[4]:el('x',{'v':ix(12,'Tonere' if v=='Tonere' else 'Imprimante')},dp)
  else:el('rangePr',{'autoStart':1,'autoEnd':1,'groupBy':{1:'months',13:'years',14:'quarters'}[f],'startDate':min(actual_dates).isoformat(),'endDate':max(actual_dates).isoformat(),'groupInterval':1},fg)
  gi=el('groupItems',{'count':len(values[f])},fg)
  for v in values[f]:el('s',{'v':v},gi)
ext=el('extLst',parent=cache);ex=el('ext',{'uri':'{725AE2AE-9491-48be-B2B4-4EB974FC3084}'},ext);el('pivotCacheDefinition',{'pivotCacheId':1,'slicerData':1},ex,uri=x14)
parts['xl/pivotCache/pivotCacheDefinition1.xml']=xml(cache)
cr=el('pivotCacheRecords',{'count':len(rows)})
for r in rows:
 er=el('r',parent=cr)
 for f in range(12):el('x',{'v':actual_dates.index(r[f]) if f==1 else ix(f,r[f])},er)
parts['xl/pivotCache/pivotCacheRecords1.xml']=xml(cr)
rels=E.Element('{%s}Relationships'%pkg);addrel(rels,rel+'/pivotCacheRecords','pivotCacheRecords1.xml','rId1');parts['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels']=xml(rels)
content('xl/pivotCache/pivotCacheDefinition1.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml');content('xl/pivotCache/pivotCacheRecords1.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml')
sheetnum={s.title:i+1 for i,s in enumerate(w)}
def sheetrel(si,typ,target):
 path=f'xl/worksheets/_rels/sheet{si}.xml.rels';rr=E.fromstring(parts[path]) if path in parts else E.Element('{%s}Relationships'%pkg)
 ri=addrel(rr,typ,target);parts[path]=xml(rr);return ri
for pi,sp in enumerate(specs,1):
 sr,sc=sp['start'];rf=sp['rf'];cf=sp['cf'];dfs=sp['dfs'];pages=sp['pages'];multi=len(dfs)>1
 pdx=el('pivotTableDefinition',{'name':sp['name'],'cacheId':1,'dataCaption':'Values','grandTotalCaption':'Grand Total','updatedVersion':6,'minRefreshableVersion':3,'createdVersion':6,'useAutoFormatting':1,'preserveFormatting':1,'showHeaders':int(sp['showheaders']),'compact':1,'compactData':1,'outline':0,'outlineData':0,'gridDropZones':0,'rowGrandTotals':1,'colGrandTotals':1,'enableDrill':1,'showDrill':1,'multipleFieldFilters':1})
 el('location',{'ref':f'{L(sc)}{sr}:{L(sp["endc"])}{sp["endr"]}','firstHeaderRow':1 if cf else 0,'firstDataRow':sp['nh'],'firstDataCol':sp['nr'],**({'rowPageCount':len(pages),'colPageCount':1} if pages else {})},pdx)
 pfs=el('pivotFields',{'count':15},pdx)
 for f in range(15):
  axis='axisRow' if f in rf else 'axisCol' if f in cf else 'axisPage' if f in pages else None
  at={'showAll':0,'defaultSubtotal':int(sp['subtotals'] and f==rf[0]),'compact':1,'outline':0,'subtotalTop':1,'dataField':int(any(d[1]==f for d in dfs)),'multipleItemSelectionAllowed':1}
  if axis:at['axis']=axis
  if sp['sort'] and rf and f==rf[-1]:at['sortType']=sp['sort']
  pf=el('pivotField',at,pfs)
  if axis or f in [2,4,7,8]:
   ordered=list(range(len(values[f])))
   if f in rf:
    present=list(dict.fromkeys(k[rf.index(f)] for k in sp['keys']))
    ordered=[ix(f,k) for k in present]+[i for i in ordered if values[f][i] not in present]
   its=el('items',{'count':len(ordered)+(1 if at['defaultSubtotal'] else 0)},pf)
   for i in ordered:
    attrs={'x':i}
    if sp['top'] and f==rf[-1] and (values[f][i],) not in sp['keys']:attrs['h']=1
    if sp['name']=='PivotTable1' and f==14:attrs['sd']=0
    el('item',attrs,its)
   if at['defaultSubtotal']:el('item',{'t':'default'},its)
  if sp['sort'] and rf and f==rf[-1]:
   au=el('autoSortScope',parent=pf);ar=el('pivotArea',{'dataOnly':0,'outline':0},au);refs=el('references',{'count':1},ar);ref=el('reference',{'field':4294967294,'count':1,'selected':0},refs);el('x',{'v':0},ref)
 # Field item position maps (sorted field order can differ from shared cache order).
 maps={}
 for f,pf in enumerate(pfs):
  its=pf.find('{%s}items'%ns)
  if its is not None:maps[f]={int(it.get('x')):i for i,it in enumerate(its) if it.get('x') is not None}
 if rf:
  rr=el('rowFields',{'count':len(rf)},pdx)
  for f in rf:el('field',{'x':f},rr)
  ri=el('rowItems',{'count':len(sp['rkeys'])},pdx)
  for k,depth in sp['rkeys']:
   item=el('i',{'t':'grand' if depth==0 else 'default' if depth else 'data','r':0,'i':0},ri)
   if not k:el('x',{'v':0},item)
   for f,v in zip(rf,k):
    if f==1:v=values[1][v.month] if isinstance(v,datetime.datetime) else v
    el('x',{'v':maps[f][ix(f,v)]},item)
 if cf or multi:
  cc=el('colFields',{'count':len(cf)+int(multi)},pdx)
  for f in cf:el('field',{'x':f},cc)
  if multi:el('field',{'x':-2},cc)
  ci=el('colItems',{'count':len(sp['colkeys'])},pdx)
  for k,di in sp['colkeys']:
   it=el('i',{'t':'grand' if cf and not k else 'data','r':0,'i':di},ci)
   for f,v in zip(cf,k):el('x',{'v':maps[f][ix(f,v)]},it)
   if multi:el('x',{'v':di},it)
   elif not k:el('x',{'v':0},it)
 if pages:
  pg=el('pageFields',{'count':len(pages)},pdx)
  for f in pages:el('pageField',{'fld':f,'hier':-1},pg)
 dfel=el('dataFields',{'count':len(dfs)},pdx)
 for cap,f,agg,pct in dfs:el('dataField',{'name':cap,'fld':f,'subtotal':agg,'showDataAs':'percentOfTotal' if pct else 'normal','numFmtId':10 if pct or f==11 else 3,'baseField':0,'baseItem':0},dfel)
 el('pivotTableStyleInfo',{'name':'PivotStyleMedium9','showRowHeaders':1,'showColHeaders':1,'showRowStripes':0,'showColStripes':0,'showLastColumn':1},pdx)
 if sp['top']:
  fl=el('filters',{'count':1},pdx);fi=el('filter',{'fld':rf[-1],'type':'count','id':0,'iMeasureFld':0},fl);af=el('autoFilter',{'ref':'A1'},fi);fc=el('filterColumn',{'colId':0},af);el('top10',{'top':1,'percent':0,'val':10,'filterVal':min(sum(r[5] for r in rows if r[7]==k[0]) for k in sp['keys'])},fc)
 path=f'xl/pivotTables/pivotTable{pi}.xml';parts[path]=xml(pdx);content(path,'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml')
 rr=E.Element('{%s}Relationships'%pkg);addrel(rr,rel+'/pivotCacheDefinition','../pivotCache/pivotCacheDefinition1.xml','rId1');parts[f'xl/pivotTables/_rels/pivotTable{pi}.xml.rels']=xml(rr)
 si=sheetnum[sp['sheet']];ri=sheetrel(si,rel+'/pivotTable',f'../pivotTables/pivotTable{pi}.xml')

# PivotChart association plus explicit series caches for immediate display.
spmap={s['name']:s for s in specs}
for sh in w:
 for ch in sh._charts:
  if not hasattr(ch,'_pivot_spec'):continue
  sp=spmap[ch._pivot_spec]
  # Each chart is assigned its physical part number by openpyxl.
  path=f'xl/charts/chart{ch._id}.xml'
  if path not in parts:continue
  cx=E.fromstring(parts[path]);cns='http://schemas.openxmlformats.org/drawingml/2006/chart'
  ps=el('pivotSource',parent=None,uri=cns);el('name',parent=ps,uri=cns).text='['+OUT+']'+sp['sheet']+'!'+sp['name'];el('fmtId',{'val':0},ps,uri=cns)
  chartel=cx.find('{%s}chart'%cns);cx.insert(list(cx).index(chartel),ps)
  for nr in cx.findall('.//{%s}numRef'%cns):
   ff=nr.find('{%s}f'%cns)
   try:
    ref=ff.text;sn,rg=ref.split('!');sn=sn.strip("'");cells=w[sn][rg.replace('$','')];vals=[c.value for row in cells for c in row]
    ca=nr.find('{%s}numCache'%cns)
    if ca is not None:nr.remove(ca)
    ca=el('numCache',parent=nr,uri=cns);el('formatCode',parent=ca,uri=cns).text='General';el('ptCount',{'val':len(vals)},ca,uri=cns)
    for i,v in enumerate(vals):
     if isinstance(v,(int,float)):pt=el('pt',{'idx':i},ca,uri=cns);el('v',parent=pt,uri=cns).text=str(v)
   except Exception:pass
  for cat in cx.findall('.//{%s}cat'%cns):
   for child in list(cat):cat.remove(child)
   if sp['sheet'] in ['Chart_data','Chart_jud_prod']:
    ml=el('multiLvlStrRef',parent=cat,uri=cns);sr,sc=sp['start'];first=sr+sp['nh'];last=sp['endr']-1
    el('f',parent=ml,uri=cns).text="'"+sp['sheet']+"'!$A$"+str(first)+':$A$'+str(last)
    ca=el('multiLvlStrCache',parent=ml,uri=cns);el('ptCount',{'val':len(sp['keys'])},ca,uri=cns)
    for level in [1,0]:
     lvl=el('lvl',parent=ca,uri=cns);previous=None
     for i,k in enumerate(sp['keys']):
      if level==0 and k[level]==previous:continue
      previous=k[level];pt=el('pt',{'idx':i},lvl,uri=cns);el('v',parent=pt,uri=cns).text=str(k[level])
    continue
   sref=el('strRef',parent=cat,uri=cns)
   srcsheet=sp['sheet'];sr,sc=sp['start'];first=sr+sp['nh'];last=sp['endr']-1;cl=L(sc+sp['nr']-1)
   el('f',parent=sref,uri=cns).text="'"+srcsheet+"'!$"+cl+'$'+str(first)+':$'+cl+'$'+str(last)
   ca=el('strCache',parent=sref,uri=cns);el('ptCount',{'val':last-first+1},ca,uri=cns)
   for i,rr in enumerate(range(first,last+1)):
    txt=str(w[srcsheet].cell(rr,sc+sp['nr']-1).value)
    if srcsheet=='Chart_data':txt=str(w[srcsheet].cell(rr,sc).value)+' '+txt
    pt=el('pt',{'idx':i},ca,uri=cns);el('v',parent=pt,uri=cns).text=txt
  parts[path]=xml(cx)

# Native Excel slicers and their report connections.
wext=el('extLst',parent=wb);we=el('ext',{'uri':'{BBE1A952-AA13-448E-AADC-164F8A28A991}'},wext);scl=el('slicerCaches',parent=we,uri=x14)
sl=el('slicers',uri=x14)
slicer_layout=[(2,'An vanzare',32,73),(8,'Regiune',116,157),(4,'Categorie_produs',276,110),(7,'Nume_complet_agent',388,315)]
connections=[]
for i,(f,caption,row,height) in enumerate(slicer_layout,1):
 name='Filtru_'+headers[f];cache_name='Slicer_'+headers[f].replace(' ','_')
 el('slicer',{'name':name,'cache':cache_name,'caption':caption,'startItem':ix(7,'Grecu Vlad') if f==7 else 0,'columnCount':1,'showCaption':1,'style':'SlicerStyleLight2','rowHeight':round(16*SCALE*9525)},sl,uri=x14)
 sc=el('slicerCacheDefinition',{'name':cache_name,'sourceName':headers[f]},uri=x14);pts=el('pivotTables',parent=sc,uri=x14)
 linked=[sp for sp in specs if (sp['name'].startswith('Pivot_') and sp['name']!='Pivot_sortat') and not(f==7 and sp['name']=='Pivot_ag10')]
 for sp in linked:el('pivotTable',{'tabId':sheetnum[sp['sheet']],'name':sp['name']},pts,uri=x14)
 da=el('data',parent=sc,uri=x14);ta=el('tabular',{'pivotCacheId':1,'sortOrder':'ascending','customListSort':0,'showMissing':0,'crossFilter':'showItemsWithDataAtTop'},da,uri=x14);its=el('items',{'count':len(values[f])},ta,uri=x14)
 for j in range(len(values[f])):el('i',{'x':j,'s':1},its,uri=x14)
 path=f'xl/slicerCaches/slicerCache{i}.xml';parts[path]=xml(sc);content(path,'application/vnd.ms-excel.slicerCache+xml')
 rid=addrel(wr,'http://schemas.microsoft.com/office/2007/relationships/slicerCache',f'slicerCaches/slicerCache{i}.xml');el('slicerCache',{'{%s}id'%rel:rid},scl,uri=x14)
 dn=wb.find('{%s}definedNames'%ns)
 if dn is None:dn=el('definedNames');wb.insert(list(wb).index(pcaches),dn)
 el('definedName',{'name':cache_name},dn).text='#N/A'
 connections.append({'slicer':headers[f],'pivots':[s['name'] for s in linked]})
parts['xl/slicers/slicer1.xml']=xml(sl);content('xl/slicers/slicer1.xml','application/vnd.ms-excel.slicer+xml')
si=sheetnum['Dashboard'];rid=sheetrel(si,'http://schemas.microsoft.com/office/2007/relationships/slicer','../slicers/slicer1.xml');sx=E.fromstring(parts[f'xl/worksheets/sheet{si}.xml']);exl=el('extLst',parent=sx);ex=el('ext',{'uri':'{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}'},exl);ls=el('slicerList',parent=ex,uri=x14);el('slicer',{'{%s}id'%rel:rid},ls,uri=x14)
parts[f'xl/worksheets/sheet{si}.xml']=xml(sx)
# Slicer graphic frames in Dashboard's existing DrawingML part.
sr=E.fromstring(parts[f'xl/worksheets/_rels/sheet{si}.xml.rels']);dt=next(e.get('Target') for e in sr if e.get('Type').endswith('/drawing'));dp=posixpath.normpath(posixpath.join('xl/worksheets',dt)) if not dt.startswith('/') else dt[1:];drawing=E.fromstring(parts[dp]);xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';a='http://schemas.openxmlformats.org/drawingml/2006/main';mc='http://schemas.openxmlformats.org/markup-compatibility/2006';sle='http://schemas.microsoft.com/office/drawing/2010/slicer';a14='http://schemas.microsoft.com/office/drawing/2010/main'
for i,(f,cap,row,height) in enumerate(slicer_layout,1):
 an=el('absoluteAnchor',parent=drawing,uri=xdr);el('pos',{'x':round(4*SCALE*9525),'y':round(row*SCALE*9525)},an,uri=xdr)
 el('ext',{'cx':round(110*SCALE*9525),'cy':round(height*SCALE*9525)},an,uri=xdr)
 ac=E.SubElement(an,'{%s}AlternateContent'%mc,nsmap={'mc':mc,'a14':a14});choice=el('Choice',{'Requires':'a14'},ac,uri=mc);gf=el('graphicFrame',{'macro':''},choice,uri=xdr);nv=el('nvGraphicFramePr',parent=gf,uri=xdr);el('cNvPr',{'id':100+i,'name':'Filtru_'+headers[f]},nv,uri=xdr);cnv=el('cNvGraphicFramePr',parent=nv,uri=xdr);el('graphicFrameLocks',{'noGrp':1},cnv,uri=a);tr=el('xfrm',parent=gf,uri=xdr);el('off',{'x':0,'y':0},tr,uri=a);el('ext',{'cx':0,'cy':0},tr,uri=a);gr=el('graphic',parent=gf,uri=a);gd=el('graphicData',{'uri':sle},gr,uri=a);el('slicer',{'name':'Filtru_'+headers[f]},gd,uri=sle);el('Fallback',parent=ac,uri=mc);el('clientData',parent=an,uri=xdr)
for i,fb in enumerate(drawing.findall('.//{%s}Fallback'%mc),1):
 sp=el('sp',parent=fb,uri=xdr);nv=el('nvSpPr',parent=sp,uri=xdr);el('cNvPr',{'id':100+i,'name':'Slicer fallback '+str(i)},nv,uri=xdr);el('cNvSpPr',parent=nv,uri=xdr)
 pr=el('spPr',parent=sp,uri=xdr);geom=el('prstGeom',{'prst':'rect'},pr,uri=a);el('avLst',parent=geom,uri=a);fill=el('solidFill',parent=pr,uri=a);el('srgbClr',{'val':'FFFFFF'},fill,uri=a)
def legendshape(identity,x,y,width,height,color=None,text=None):
 an=el('absoluteAnchor',parent=drawing,uri=xdr);el('pos',{'x':round(x*SCALE*9525),'y':round(y*SCALE*9525)},an,uri=xdr);el('ext',{'cx':round(width*SCALE*9525),'cy':round(height*SCALE*9525)},an,uri=xdr)
 sp=el('sp',parent=an,uri=xdr);nv=el('nvSpPr',parent=sp,uri=xdr);el('cNvPr',{'id':identity,'name':'Legenda '+str(identity)},nv,uri=xdr);el('cNvSpPr',{'txBox':1 if text else 0},nv,uri=xdr);pr=el('spPr',parent=sp,uri=xdr)
 geom=el('prstGeom',{'prst':'rect'},pr,uri=a);el('avLst',parent=geom,uri=a)
 if color:fill=el('solidFill',parent=pr,uri=a);el('srgbClr',{'val':color},fill,uri=a)
 else:el('noFill',parent=pr,uri=a)
 line=el('ln',parent=pr,uri=a);el('noFill',parent=line,uri=a)
 if text:
  tx=el('txBody',parent=sp,uri=xdr);el('bodyPr',{'lIns':0,'tIns':0,'rIns':0,'bIns':0,'anchor':'ctr'},tx,uri=a);el('lstStyle',parent=tx,uri=a);p=el('p',parent=tx,uri=a);r=el('r',parent=p,uri=a);rp=el('rPr',{'lang':'ro-RO','sz':800,'i':1},r,uri=a);el('latin',{'typeface':'Calibri'},rp,uri=a);el('t',parent=r,uri=a).text=text
 el('clientData',parent=an,uri=xdr)
legendshape(201,541,7,27,4,color=BLUE);legendshape(202,574,2,49,12,text='- imprimante')
legendshape(203,541,22,27,4,color=RED);legendshape(204,574,17,49,12,text='- tonere')
parts[dp]=xml(drawing)
# Reassigning frozen panes through openpyxl can leave duplicate selections.
sx=E.fromstring(parts['xl/worksheets/sheet1.xml'])
for sv in sx.findall('.//{%s}sheetView'%ns):
 seen=set()
 for selection in list(sv.findall('{%s}selection'%ns)):
  pane=selection.get('pane','topLeft')
  if pane in seen:sv.remove(selection)
  else:seen.add(pane)
parts['xl/worksheets/sheet1.xml']=xml(sx)
styles=E.fromstring(parts['xl/styles.xml'])
font_order=['b','i','strike','condense','extend','outline','shadow','u','vertAlign','sz','color','name','family','charset','scheme']
for font in styles.findall('.//{%s}fonts/{%s}font'%(ns,ns)):
 font[:]=sorted(font,key=lambda e:font_order.index(E.QName(e).localname))
parts['xl/styles.xml']=xml(styles)
parts['xl/workbook.xml']=xml(wb);parts['xl/_rels/workbook.xml.rels']=xml(wr);parts['[Content_Types].xml']=xml(ct)
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as zz:
 for n,d in parts.items():zz.writestr(n,d)
json.dump({'total':TOTAL,'transactions':len(rows),'connections':connections,'pivots':[(s['sheet'],s['name']) for s in specs]},open('lucru_dashboard/verificari.json','w'),ensure_ascii=False,indent=2)
print(OUT)
