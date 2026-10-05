from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader
import matplotlib.pyplot as plt
import pandas as pd, numpy as np, os, re, json, zipfile, shutil

# Fixed for clean workspace execution using relative paths
FIG = 'figures'
DATA = 'data'
os.makedirs(FIG, exist_ok=True)
os.makedirs(DATA, exist_ok=True)
PDF = 'From_Accounts_to_Capital_Aditya_Makan_FINAL.pdf'
ZIP = 'FINAL_PUBLISH_AUDIT_PACKAGE.zip'


# ---------- Fonts ----------
for n,p in [('DV','/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf'),('DVB','/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf'),('DS','/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'),('DSB','/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf')]:
    pdfmetrics.registerFont(TTFont(n,p))

# ---------- Verified data ----------
holders=pd.DataFrame({'period':['Mar-2019','Mar-2020','Mar-2024'],'value':[282.2,317.4,895.8]})
accounts=pd.DataFrame({'period':['Mar-2019','Mar-2020','Mar-2024'],'value':[354.3,403.9,1508.4]})
broker=pd.DataFrame({'fy':['FY17','FY18','FY19','FY20','FY21','FY22','FY23','FY24'],'share':[24.2,41.0,55.2,60.6,87.6,89.8,90.9,88.4]})
mf_unique=pd.DataFrame({'period':['Mar-2019','Mar-2024'],'value':[191.7,442.5]})
mf_aum=pd.DataFrame({'period':['Aug-2016','Aug-2021','Aug-2026'],'value':[15.63,36.59,87.08]})
sip_annual=pd.DataFrame({'fy':['FY17','FY18','FY19','FY20','FY21','FY22','FY23','FY24','FY25','FY26'],'value':[43921,67190,92693,100084,96080,124566,155972,199219,289352,349589]})
sip_month=pd.DataFrame({'month':['Apr-26','May-26','Jun-26','Jul-26','Aug-26'],'value':[31115,30954,31781,31961,32297]})
participation=pd.DataFrame({'group':['All households','Urban','Rural','Top 9 metros','Maharashtra'],'value':[9.5,15,6,23,17]})
ownership=pd.DataFrame({'measure':['Direct individuals','Individuals: direct + through DMFs','DMFs','FPIs'],'value':[9.1,18.7,11.4,15.8]})
wealth=pd.DataFrame({'measure':['Individuals’ equity exposure','Cumulative household equity wealth creation since Apr-2020'],'value':[76.5,44.0]})
primary=pd.DataFrame({'year':['2022-23','2023-24','2024-25'],'total':[221243,357539,631510],'equity':[32486,46879,95139],'debt':[22090,24882,22400],'mf':[166515,285052,512765],'reit_invit':[152,727,1208]})

ledger=[
['Unique individual demat holders','Mar-2019','282.2 lakh','SEBI 2025'],['Unique individual demat holders','Mar-2020','317.4 lakh','SEBI 2025'],['Unique individual demat holders','Mar-2024','895.8 lakh','SEBI 2025'],
['Individual demat accounts','Mar-2019','354.3 lakh','SEBI 2025'],['Individual demat accounts','Mar-2020','403.9 lakh','SEBI 2025'],['Individual demat accounts','Mar-2024','1,508.4 lakh','SEBI 2025'],['Total demat accounts','Mar-2024','1,513.8 lakh','SEBI 2025'],
['Discount-broker share among accounts added by top 10 CDSL DPs','FY2017','24.2%','SEBI/CDSL'],['Discount-broker share among accounts added by top 10 CDSL DPs','FY2024','88.4%','SEBI/CDSL'],
['Unique individual mutual-fund investors','Mar-2019','191.7 lakh','SEBI 2025'],['Unique individual mutual-fund investors','Mar-2024','442.5 lakh','SEBI 2025'],
['MF AUM','Aug-2016','₹15.63 trillion','AMFI 2026'],['MF AUM','Aug-2026','₹87.08 trillion','AMFI 2026'],['MF folios','Aug-2026','28.35 crore','AMFI 2026'],
['SIP contribution','FY2025-26','₹3,49,589 crore','AMFI 2026'],['SIP contribution','Aug-2026','₹32,297 crore','AMFI 2026'],['SIP contributing accounts','Apr-Aug 2026','1,001.79 lakh','AMFI 2026'],['SIP outstanding accounts','Apr-Aug 2026','1,075.32 lakh','AMFI 2026'],
['Households invested in securities-market products','Survey 2025','9.5%','SEBI Investor Survey 2025'],['Households aware of at least one securities product','Survey 2025','63%','SEBI Investor Survey 2025'],
['Direct individual ownership','Mar-2026','9.1%','NSE Ownership Tracker Q4 FY26'],['Direct + indirect individual ownership','Mar-2026','18.7%','NSE Ownership Tracker Q4 FY26'],['Individual equity exposure','Mar-2026','₹76.5 lakh crore','NSE Ownership Tracker Q4 FY26'],['Cumulative household equity wealth creation','Apr-2020 to Mar-2026','≈₹44 lakh crore','NSE Market Pulse May 2026'],
['Household-sector savings including NPISHs via securities-market primary channels','FY2024-25','₹6,31,510 crore','SEBI 2026'],['Household-sector equity savings including NPISHs via primary market','FY2024-25','₹95,139 crore','SEBI 2026'],['Household-sector debt savings including NPISHs via primary market','FY2024-25','₹22,400 crore','SEBI 2026'],
]
ldf=pd.DataFrame(ledger,columns=['measure','period','value','source']); ldf.to_csv(os.path.join(DATA,'verified_data_ledger.csv'),index=False)
with pd.ExcelWriter(os.path.join(DATA,'verified_data_tables.xlsx'),engine='openpyxl') as w:
    for name,df in [('holders',holders),('accounts',accounts),('broker',broker),('mf_unique',mf_unique),('mf_aum',mf_aum),('sip_annual',sip_annual),('sip_month',sip_month),('participation',participation),('ownership',ownership),('wealth',wealth),('primary',primary)]: df.to_excel(w,name,index=False)
    ldf.to_excel(w,'core_ledger',index=False)

# ---------- Figures ----------
plt.rcParams['font.family']='DejaVu Sans'; plt.rcParams['figure.dpi']=180

def sf(name):
    p=os.path.join(FIG,name); plt.tight_layout(); plt.savefig(p,bbox_inches='tight'); plt.close(); return p
# 1 framework
fig,ax=plt.subplots(figsize=(10,2.6)); ax.axis('off'); xs=np.linspace(.08,.92,5); labs=['Formal\naccess','Investor\nentry','Digital\nintermediation','Recurring\nallocation','Ownership,\nwealth & capital']
for i,x in enumerate(xs):
    ax.text(x,.5,labs[i],ha='center',va='center',fontsize=9.5,bbox=dict(boxstyle='round,pad=.55',fill=False,lw=1.1))
    if i<4: ax.annotate('',xy=(xs[i+1]-.06,.5),xytext=(x+.06,.5),arrowprops=dict(arrowstyle='->',lw=1))
ax.text(.5,.88,'From access to household demand: the paper’s evidence chain',ha='center',fontsize=12,fontweight='bold'); F1=sf('fig01_framework.png')
# 2 stock-flow identity
fig,ax=plt.subplots(figsize=(9.2,3.2)); ax.axis('off')
ax.text(.16,.50,'Previous household\nequity wealth\nHₜ₋₁',ha='center',va='center',fontsize=10,bbox=dict(boxstyle='round,pad=.55',fill=False,lw=1.1))
ax.text(.50,.78,'Net fresh investment\nIₜ',ha='center',va='center',fontsize=9.5,bbox=dict(boxstyle='round,pad=.45',fill=False,lw=1.0))
ax.text(.50,.22,'Valuation + other changes\nVₜ',ha='center',va='center',fontsize=9.5,bbox=dict(boxstyle='round,pad=.45',fill=False,lw=1.0))
ax.text(.84,.50,'Current household\nequity wealth\nHₜ',ha='center',va='center',fontsize=10,bbox=dict(boxstyle='round,pad=.55',fill=False,lw=1.1))
ax.annotate('',xy=(.70,.50),xytext=(.30,.50),arrowprops=dict(arrowstyle='->',lw=1.1))
ax.annotate('',xy=(.76,.57),xytext=(.62,.72),arrowprops=dict(arrowstyle='->',lw=1.1))
ax.annotate('',xy=(.76,.43),xytext=(.62,.28),arrowprops=dict(arrowstyle='->',lw=1.1))
ax.text(.50,.04,'Hₜ − Hₜ₋₁ = Iₜ + Vₜ',ha='center',va='center',fontsize=11,fontweight='bold')
ax.set_title('Stock-flow discipline: a change in wealth is not automatically a cash investment',fontsize=11.5,fontweight='bold'); F2=sf('fig02_stock_flow.png')
# 3 accounts vs unique holders
fig,ax=plt.subplots(figsize=(7.3,3.9)); x=np.arange(3); w=.35; ax.bar(x-w/2,accounts.value,w,label='Individual accounts'); ax.bar(x+w/2,holders.value,w,label='Unique holders'); ax.set_xticks(x,accounts.period); ax.set_ylabel('Lakh'); ax.set_title('Accounts grew faster than distinct individual holders'); ax.legend(frameon=False,fontsize=8); F3=sf('fig03_accounts_holders.png')
# 3 discount brokers
fig,ax=plt.subplots(figsize=(7.3,3.9)); ax.plot(broker.fy,broker.share,marker='o',lw=2); ax.set_ylim(0,100); ax.set_ylabel('% of accounts added'); ax.set_title('Discount-broker share among accounts added by top 10 CDSL DPs'); F4=sf('fig04_broker_share.png')
# 4 MF unique
fig,ax=plt.subplots(figsize=(6.8,3.8)); ax.bar(mf_unique.period,mf_unique.value); ax.set_ylabel('Lakh unique individual investors'); ax.set_title('Unique individual mutual-fund investors'); F5=sf('fig05_mf_unique.png')
# 5 MF AUM
fig,ax=plt.subplots(figsize=(6.8,3.8)); ax.bar(mf_aum.period,mf_aum.value); ax.set_ylabel('₹ trillion'); ax.set_title('Mutual-fund industry AUM'); F6=sf('fig06_mf_aum.png')
# 6 annual SIP
fig,ax=plt.subplots(figsize=(7.6,4)); ax.plot(sip_annual.fy,sip_annual.value,marker='o',lw=2); ax.set_ylabel('₹ crore'); ax.set_title('Annual SIP contributions'); ax.tick_params(axis='x',rotation=35); F7=sf('fig07_sip_annual.png')
# 7 monthly SIP
fig,ax=plt.subplots(figsize=(6.8,3.8)); ax.bar(sip_month.month,sip_month.value); ax.set_ylabel('₹ crore'); ax.set_title('SIP contributions, April–August 2026'); F8=sf('fig08_sip_monthly.png')
# 8 awareness vs participation
fig,ax=plt.subplots(figsize=(6.8,3.8)); ax.bar(['Aware of ≥1\nsecurities product','Invested in\nsecurities products'],[63,9.5]); ax.set_ylim(0,70); ax.set_ylabel('% of households'); ax.set_title('Awareness and participation are distinct measures'); F9=sf('fig09_awareness_participation.png')
# 9b participation geography; retained as separate supporting graphic for the participation section
# 10 ownership will follow; the participation chart is embedded in the section as Figure 10

fig,ax=plt.subplots(figsize=(7.4,3.9)); ax.bar(participation.group,participation.value); ax.set_ylim(0,26); ax.set_ylabel('% of households'); ax.set_title('Participation remains uneven across household groups'); ax.tick_params(axis='x',rotation=25); F10=sf('fig10_participation.png')
# Composite participation figure used in the manuscript as one graphic, preserving both awareness and geographic penetration.
from PIL import Image as PILImage
a=PILImage.open(os.path.join(FIG,'fig09_awareness_participation.png')).convert('RGB'); b=PILImage.open(os.path.join(FIG,'fig10_participation.png')).convert('RGB')
target_h=min(a.height,b.height,420); a=a.resize((int(a.width*target_h/a.height),target_h)); b=b.resize((int(b.width*target_h/b.height),target_h)); W=a.width+b.width+20; H=target_h
combo=PILImage.new('RGB',(W,H),'white'); combo.paste(a,(0,0)); combo.paste(b,(a.width+20,0)); combo.save(os.path.join(FIG,'fig09_participation.png')); os.remove(os.path.join(FIG,'fig09_awareness_participation.png')); os.remove(os.path.join(FIG,'fig10_participation.png'))
# 10 ownership comparison
fig,ax=plt.subplots(figsize=(7.4,3.9)); ax.bar(ownership.measure,ownership.value); ax.set_ylim(0,22); ax.set_ylabel('% of NSE-listed market capitalisation'); ax.set_title('Ownership structure, March 2026'); ax.tick_params(axis='x',rotation=18); F11=sf('fig10_ownership.png')
# 11 wealth
fig,ax=plt.subplots(figsize=(6.9,3.8)); ax.bar(['Equity exposure','Cumulative wealth\ncreation since Apr-20'],wealth.value); ax.set_ylabel('₹ lakh crore'); ax.set_title('Equity exposure and wealth creation are different measures'); F12=sf('fig11_wealth.png')
# 12 primary savings composition FY25
fig,ax=plt.subplots(figsize=(7.6,4.1)); vals=[512765,95139,22400,425,783]; labs=['Mutual funds','Equity','Debt','REITs','InvITs']; ax.bar(labs,vals); ax.set_ylabel('₹ crore'); ax.set_title('Household savings through primary securities-market channels, FY2024–25'); ax.tick_params(axis='x',rotation=25); F13=sf('fig12_primary_savings.png')
# 13 domestic vs foreign
fig,ax=plt.subplots(figsize=(7.2,3.9)); ax.bar(['Individuals\ndirect + indirect','DMFs','FPIs'],[18.7,11.4,15.8]); ax.set_ylim(0,22); ax.set_ylabel('% of market capitalisation'); ax.set_title('Selected ownership categories, March 2026'); F14=sf('fig13_domestic_foreign.png')
# 14 counter-checks
fig,ax=plt.subplots(figsize=(9,4.5)); ax.axis('off'); pairs=[('Accounts ↑','not unique people'),('SIPs ↑','not unique investors'),('AUM ↑','not household wealth'),('Wealth ↑','can include price effects'),('Ownership ↑','not automatically issuer financing'),('Digital access ↑','not proof of long-term investing')]
for i,(a,b) in enumerate(pairs):
 y=.9-i*.15; ax.text(.22,y,a,ha='center',fontweight='bold',fontsize=9.5); ax.annotate('',xy=(.56,y),xytext=(.32,y),arrowprops=dict(arrowstyle='->')); ax.text(.77,y,b,ha='center',fontsize=9.3)
ax.set_title('Counter-checks applied to the headline statistics',fontsize=12,fontweight='bold'); F15=sf('fig15_counterchecks.png')
# 15 identification map
fig,ax=plt.subplots(figsize=(9.2,2.8)); ax.axis('off'); labs=['Pre-period\ntrend','COVID-era\nbreak','Alternative\nexplanations','Robustness\nchecks','Interpretation']; xs=np.linspace(.08,.92,5)
for i,(x,l) in enumerate(zip(xs,labs)):
 ax.text(x,.5,l,ha='center',va='center',fontsize=9.2,bbox=dict(boxstyle='round,pad=.5',fill=False,lw=1));
 if i<4: ax.annotate('',xy=(xs[i+1]-.055,.5),xytext=(x+.055,.5),arrowprops=dict(arrowstyle='->'))
ax.set_title('Identification logic: a break is evidence of a change, not proof of a cause',fontsize=11.5,fontweight='bold'); F16=sf('fig14_identification.png')
# 16 reproducibility
fig,ax=plt.subplots(figsize=(9.2,2.6)); ax.axis('off'); steps=['Official\nsources','Variable\nledger','Raw + transformed\ndata','Code-generated\nfigures','Model\nvalidation','Final\naudit']; xs=np.linspace(.07,.93,6)
for i,(x,l) in enumerate(zip(xs,steps)):
 ax.text(x,.5,l,ha='center',va='center',fontsize=9,bbox=dict(boxstyle='round,pad=.5',fill=False,lw=1));
 if i<5: ax.annotate('',xy=(xs[i+1]-.05,.5),xytext=(x+.05,.5),arrowprops=dict(arrowstyle='->'))
ax.set_title('Reproducible research workflow',fontsize=12,fontweight='bold'); F17=sf('fig16_reproducibility.png')

# ---------- Styles ----------
from reportlab.lib.styles import ParagraphStyle
S={
'body':ParagraphStyle('body',fontName='DV',fontSize=10.05,leading=14.45,spaceAfter=7,textColor=HexColor('#202020')),
'h1':ParagraphStyle('h1',fontName='DSB',fontSize=14.7,leading=18,spaceBefore=5,spaceAfter=7,textColor=HexColor('#172B4D'),keepWithNext=True),
'h2':ParagraphStyle('h2',fontName='DSB',fontSize=11.0,leading=14,spaceBefore=6,spaceAfter=4,textColor=HexColor('#263238'),keepWithNext=True),
'cap':ParagraphStyle('cap',fontName='DS',fontSize=7.5,leading=9.6,spaceBefore=3,spaceAfter=7,textColor=HexColor('#444')),
'small':ParagraphStyle('small',fontName='DS',fontSize=7.6,leading=10,textColor=HexColor('#444')),
'title':ParagraphStyle('title',fontName='DSB',fontSize=24,leading=28,alignment=TA_CENTER),
'subtitle':ParagraphStyle('subtitle',fontName='DV',fontSize=13.2,leading=17,alignment=TA_CENTER,textColor=HexColor('#333')),
'center':ParagraphStyle('center',fontName='DS',fontSize=8.6,leading=11,alignment=TA_CENTER,textColor=HexColor('#444')),
'quote':ParagraphStyle('quote',fontName='DV',fontSize=10.6,leading=15,leftIndent=15,rightIndent=15,spaceBefore=5,spaceAfter=8,textColor=HexColor('#333'))}
def P(x,s='body'): return Paragraph(x,S[s])
def IMG(path,w=150*mm):
 from PIL import Image as PI
 im=PI.open(path); W,H=im.size; return Image(path,width=w,height=w*H/W)
def CAP(x): return P('<b>Source note:</b> '+x,'cap')
def T(data,widths,fs=7.4):
 from reportlab.lib.styles import ParagraphStyle
 from reportlab.platypus import Paragraph
 avail=170*mm
 total=sum(widths)
 if total>avail:
  scale=avail/total
  widths=[w*scale for w in widths]
 cell=ParagraphStyle('tablecell',fontName='DS',fontSize=fs,leading=fs+1.8,textColor=HexColor('#202020'))
 head=ParagraphStyle('tablehead',fontName='DSB',fontSize=fs,leading=fs+1.8,textColor=HexColor('#202020'))
 pdata=[]
 for r,row in enumerate(data):
  pdata.append([x if hasattr(x,'wrap') else Paragraph(str(x).replace('&','&amp;'), head if r==0 else cell) for x in row])
 t=Table(pdata,colWidths=widths,repeatRows=1,hAlign='LEFT')
 t.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('GRID',(0,0),(-1,-1),.35,HexColor('#B8B8B8')),('BACKGROUND',(0,0),(-1,0),HexColor('#ECECEC')),('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),3.5),('BOTTOMPADDING',(0,0),(-1,-1),3.5)])); return t

def SEC(title,paras,fig=None,cap=None,tab=None,tab_source=None):
 out=[P(title,'h1')]+[P(x) for x in paras]
 if tab:
  out += [Spacer(1,2),tab]
  if tab_source: out += [P('<b>Source note:</b> '+tab_source,'cap')]
  out += [Spacer(1,5)]
 if fig: out += [IMG(fig),P(cap,'cap')]
 return out

# ---------- Build ----------
D=SimpleDocTemplate(PDF,pagesize=A4,leftMargin=18*mm,rightMargin=18*mm,topMargin=17*mm,bottomMargin=16*mm,title='From Accounts to Capital',author='Aditya Makan')
st=[]
st += [Spacer(1,31*mm),P('FROM ACCOUNTS TO CAPITAL','title'),Spacer(1,5*mm),P("The Rise of India's Retail-Demand Economy After COVID-19",'subtitle'),Spacer(1,14*mm),P('Aditya Makan','center'),Spacer(1,2*mm),P('B.Sc. Finance, City Premier College, Nagpur<br/>Affiliated to Rashtrasant Tukadoji Maharaj Nagpur University, Maharashtra, India','center'),Spacer(1,8*mm),P('<b>Independent Undergraduate Economics Research Paper</b>','center'),Spacer(1,4*mm),P('Research focus: household participation, digital intermediation, recurring allocation, domestic ownership, and the demand side of India’s capital market.','center'),Spacer(1,18*mm),P('Manuscript date: October 2026','center'),Spacer(1,4*mm),P('Acknowledgements: None.','center'),PageBreak()]
st += [P('Abstract','h1'),P("India’s household participation in securities markets expanded sharply during and after the COVID-19 period. But what exactly expanded? This paper studies the change from the demand side of the capital market, separating account access, distinct investors, digital intermediation, recurring allocation, ownership, wealth and issuer financing. An account is not a person; a SIP contribution is a flow; mutual-fund AUM is not household wealth; and a rise in asset values is not automatically new capital supplied to a company."),P("Using official evidence from SEBI, AMFI and NSE, the paper documents a large increase in distinct individual participation alongside an even larger increase in account counts. PAN-de-duplicated unique individual demat holders rose from 282.2 lakh in March 2019 to 895.8 lakh in March 2024, while individual demat accounts reached 1,508.4 lakh. Among the top ten CDSL depository participants, the share of individual accounts added by discount-broker DPs rose from 24.2% in FY2017 to 88.4% in FY2024. Mutual-fund AUM reached ₹87.08 trillion in August 2026, while SIP contributions reached ₹3,49,589 crore in FY2025–26 and ₹32,297 crore in August 2026. NSE reports that, in March 2026, individuals held 18.7% of NSE-listed market capitalisation when direct holdings and holdings through domestic mutual funds are combined, with total individual equity exposure of ₹76.5 lakh crore."),P("The paper does not assume that COVID alone caused these changes. Digital adoption, broker competition, market returns, household saving and investor attention moved together. It therefore separates post-COVID expansion from causal claims and distinguishes secondary-market ownership from primary-market financing. SEBI’s revised methodology for household-sector savings, which includes NPISHs, reports ₹6,31,510 crore through securities-market primary channels in FY2024–25, including ₹95,139 crore through equity and ₹22,400 crore through debt. These figures strengthen the capital-formation discussion but are broader than direct retail subscriptions."),P("The main conclusion is deliberately limited: India developed a substantially larger and more digitally mediated household demand base for financial assets after the COVID period. The evidence is strongest on participation, intermediation, recurring allocation and ownership. The exact contribution of household demand to productive corporate capital formation remains a separate empirical question requiring more granular flow and issue-level evidence."),P('<b>Keywords:</b> retail investors; household finance; demat accounts; mutual funds; SIP; digital intermediation; equity markets; domestic ownership; capital formation; COVID-19'),P('<b>JEL Classification:</b> D14; E21; G11; G23; G41; O16'),PageBreak()]
# contents
st += [P('Contents and Research Architecture','h1')]
contents=['1. Introduction: What Actually Changed?','2. Related Literature and the Research Gap','3. Research Questions and Contribution','4. Conceptual Framework: From Access to Demand','5. Institutional Setting and Measurement Rules','6. Data, Definitions and Time Windows','7. The Account Expansion and the People Behind It','8. Multiple Accounts: Why the Gap Matters','9. Digital Intermediation and the Entry Channel','10. Mutual Funds as a Second Entry Channel','11. Mutual-Fund Scale and AUM','12. SIPs and Recurring Allocation','13. The Current SIP Run-Rate','14. Participation: A Larger Market Is Not a Universal Market','15. Household Ownership and the Scale of Financial Demand','16. From Wealth to Flows: The Stock-Flow Problem','17. From Household Demand to Corporate Financing','18. Domestic Investors, Foreign Flows and Market Structure','19. Empirical Strategy: What Can Actually Be Estimated?','20. Identification and Robustness','21. Countervailing Interpretations','22. Discussion: What the Evidence Supports','23. Limitations and Reproducible Research Agenda','24. Conclusion','Appendix A. Core Data Ledger','Appendix B. Measurement Definitions','Appendix C. Empirical Specifications','References']
for c in contents: st.append(P(c,'small'))
st += [Spacer(1,6),P('<b>Central research question:</b> How did the expansion of household participation change the demand side of India’s capital market after COVID-19, once accounts, people, flows, wealth and ownership are kept separate?','quote'),P('The manuscript uses an observation → question → evidence → counter-check → interpretation structure. It does not report regression estimates that have not actually been estimated.','small'),PageBreak()]

sections=[]
sections.append(SEC('1. Introduction: What Actually Changed?',[
"India’s securities market became much more accessible to households during the period surrounding COVID-19. Demat accounts multiplied, digital brokers became a major distribution channel, mutual-fund assets expanded, and SIP contributions became a large recurring allocation mechanism. Those facts are easy to observe. The harder question is what they mean together.",
"A rise in accounts can mean more people entered. It can also mean existing investors opened additional accounts. A rise in SIP contributions can mean that more households are allocating regularly, but SIP accounts are not unique people. A rise in household equity wealth can reflect fresh saving, price appreciation, or both. These distinctions determine what the headline numbers can actually prove.",
"The central question is therefore: <b>How did the expansion of household participation change the demand side of India’s capital market after COVID-19?</b> The answer requires a sequence from formal access to actual investor entry, from entry to digital intermediation, from intermediation to recurring allocation, and from allocation to ownership and possible capital-formation effects.",
"The paper deliberately avoids treating COVID as a single-cause explanation. The pandemic coincided with digital adoption, broker competition, market volatility, changes in household saving, low-friction onboarding and a rapid increase in financial attention. An interrupted trend around 2020 can identify a break in the data, but not automatically identify which force caused it.",
"The phrase <i>retail-demand economy</i> is used narrowly here. It means an economy in which households form an increasingly important demand base for financial assets and market-linked investment products. It does not refer to consumer demand, retail sales or aggregate consumption."
],F1,'Figure 1. Five-stage evidence architecture. The stages organize the evidence; they do not imply that every household follows the same sequence. Source: author’s conceptual framework.'))
sections.append(SEC('2. Related Literature and the Research Gap',[
"The paper sits at the intersection of household finance, financial intermediation and the economics of market participation. A recurring theme in household-finance research is that access to financial markets does not automatically translate into identical investment behaviour. Investors differ in participation, portfolio choice, trading intensity and persistence. This matters here because a national increase in accounts is only the first layer of the story.",
"Indian research provides related evidence. Sourirajan and Natarajan (2021), using fund-level panel data, document performance-chasing and timing differences among Indian retail mutual-fund investors. That result is useful here not because it is a post-COVID account measure, but because it shows why the amount of retail money entering a product does not by itself establish the quality or persistence of participation. More recent work on rural India also shows that participation can depend on local professional networks and social learning (Chauhan and Sam, 2026), suggesting that digital access does not eliminate all non-financial barriers to participation.",
"The contribution of this paper is narrower and more empirical than a general study of investor behaviour. It brings together official measures that are usually discussed separately: PAN-de-duplicated individual holders, demat accounts, discount-broker distribution, unique mutual-fund investors, SIP flows, household participation, individual ownership, household equity wealth and revised household-savings flows.",
"The paper does not claim to be the first study of any one of these measures. Its purpose is to connect them while preserving their definitions. The research gap is therefore a measurement-and-transmission gap: <i>when household participation expands, how much of the change represents more people, more recurring allocation, more ownership, and potentially more capital supplied through primary markets?</i>"
],tab=T([['Literature idea','What it tells us','Role in this paper'],['Household finance','Participation and behaviour differ across households','Motivates separation of access, entry and persistence'],['Retail fund-flow research','Retail flows can reflect behaviour beyond simple saving','Supports caution about interpreting flows'],['Digital/rural participation research','Access interacts with social and geographic constraints','Supports participation-gap analysis'],['This paper','Connects institutional aggregate measures','Builds a household-demand transmission framework']], [42*mm,62*mm,64*mm],7.2),tab_source='SEBI, AMFI, NSE, SEBI Investor Survey 2025 and SEBI household-savings study; source definitions are kept separate in the manuscript.'))
sections.append(SEC('3. Research Questions and Contribution',[
"The paper asks six linked questions. First, how much of the increase in demat accounts represents additional distinct investors rather than additional accounts? Second, how did the entry mechanism change, particularly with the rise of discount-broker distribution? Third, did mutual funds and SIPs create a more persistent mechanism for household allocation? Fourth, how large did household-linked ownership and financial-asset exposure become? Fifth, how uneven is participation across households? Sixth, did the larger household demand base reach corporate financing through primary-market channels?",
"The contribution is mainly measurement and synthesis. SEBI is used for investor-entry and account structure. AMFI is used for mutual-fund AUM and SIP flows. NSE is used for individual ownership and household equity exposure. The SEBI Investor Survey is used to measure household participation and its uneven distribution. SEBI’s household-savings study is used to identify a separate flow measure for securities-market saving.",
"Putting these sources together is useful precisely because they do not measure the same thing. A single number called ‘retail participation’ would hide the differences between accounts, people, flows, stocks and ownership. The paper instead asks what each source can answer and then connects the answers without changing their definitions.",
"The approach is intentionally question-first. When the account count rises, the paper asks how many people are behind it. When SIP contributions rise, it asks whether the flow represents repeated allocation. When household wealth rises, it asks how much can be attributed to new investment rather than valuation. When ownership rises, it asks whether the ownership is direct or intermediated. This is the recurring reasoning pattern of the paper."
],tab=T([['Question','Primary evidence','What it can establish'],['People vs accounts','SEBI holders + demat accounts','Distinct participation and account intensity'],['Entry mechanism','SEBI/CDSL broker series','Change in distribution channel'],['Recurring allocation','AMFI SIP series','Scale of repeated allocation'],['Participation gaps','SEBI Investor Survey','Household penetration'],['Ownership','NSE Ownership Tracker','Direct and indirect ownership'],['Capital formation','SEBI household-savings study','Primary and secondary savings channels']], [43*mm,61*mm,64*mm]),tab_source='SEBI, AMFI, NSE and SEBI Investor Survey 2025; the table maps each research question to the source capable of answering it.'))
sections.append(SEC('4. Conceptual Framework: From Access to Demand',[
"The framework has five stages: formal access, investor entry, digital intermediation, recurring allocation, and ownership/wealth/capital effects. The stages are not a mechanical ladder. A person can open multiple accounts without becoming a long-term investor. A mutual-fund investor can participate without using a discount broker. A direct-equity investor can trade without using a SIP.",
"Each measure therefore has a specific job. Account counts describe the infrastructure and access layer. PAN-de-duplicated holders describe distinct individual participation. Broker shares describe the intermediation layer. SIP contributions describe a recurring flow. AUM describes the stock of assets managed by an industry. Ownership describes who holds listed securities. Household equity exposure describes the value of individual-linked holdings. Primary-market flows are needed before the paper can make a direct statement about financing new securities.",
"The stock-flow identity is <b>H<sub>t</sub> − H<sub>t−1</sub> = I<sub>t</sub> + V<sub>t</sub></b>, where H is household equity wealth, I is net fresh investment and V captures valuation and other changes. The identity prevents a common mistake: if prices rise, H can increase even when households add little new cash.",
"The same logic applies to mutual funds. A contribution is a flow into a scheme. AUM is the stock of assets under management. A folio is an account record. A unique investor is a person under a specific definition. These quantities can be related, but they cannot be substituted for one another."
],F2,'Figure 2. Stock-flow identity used to separate household wealth changes from fresh investment. Source: author’s conceptual framework.'))
sections.append(SEC('5. Institutional Setting and Measurement Rules',[
"India’s household securities-market system has several layers. Depositories maintain electronic securities records. Depository participants provide account access. Brokers execute transactions. Mutual funds pool savings and invest across securities. Exchanges provide trading venues. SEBI regulates the securities market, while AMFI publishes industry-level mutual-fund statistics.",
"SEBI’s investor study uses PAN-wise de-duplication across depositories to estimate unique individual demat holders. SEBI’s individual-account series is different: its individual category includes individuals and Hindu Undivided Families. That is why the individual-account count and unique-holder count should not be treated as interchangeable series.",
"AMFI’s folio statistics are also different from unique-person counts. A folio is an account record in the mutual-fund system. AUM can rise through net contributions, market movements and other effects. The paper therefore uses AMFI for industry scale and recurring allocation rather than claiming that every folio is one household.",
"The final rule is temporal. The PAN-de-duplicated holder series used here ends in March 2024. AMFI provides monthly SIP data through August 2026. NSE ownership and household-equity data reach March 2026. SEBI household-savings data reach FY2024–25. The paper reports each source at its valid endpoint rather than creating a false common date."
],tab=T([['Measure','Unit','Measures','Does not measure'],['Individual demat accounts','Lakh accounts','Administrative account stock','Unique people'],['PAN-de-duplicated holders','Lakh persons','Distinct individual holders','Households'],['SIP contribution','₹ crore','Recurring flow','Unique investors or wealth'],['MF AUM','₹ trillion','Industry asset stock','Household-only wealth'],['Individual ownership','% market cap','Ownership stock','New issuer financing'],['Household equity exposure','₹ lakh crore','Value of individual-linked holdings','Cumulative cash invested']], [35*mm,29*mm,58*mm,48*mm]),tab_source='SEBI, AMFI and NSE definitions; the table records the interpretation limits applied throughout the paper.'))
sections.append(SEC('6. Data, Definitions and Time Windows',[
"The core investor evidence comes from SEBI’s ‘Growth of Individual Investors in Indian Securities Market’. AMFI provides mutual-fund AUM, unique-investor references and SIP statistics. NSE provides the March 2026 ownership tracker and May 2026 Market Pulse. The SEBI Investor Survey 2025 supplies household participation. SEBI’s May 2026 household-savings study supplies annual flows through primary and secondary securities-market channels.",
"The data are not all the same frequency. Investor-entry evidence is mainly annual or year-end. SIP contributions are monthly. Ownership is quarterly. Household securities-market savings are annual. This matters because a quarterly ownership change cannot be compared directly with a monthly SIP contribution as if both were the same type of flow.",
"Where the paper calculates a growth rate, the calculation can be reconstructed from the ledger. Unique individual demat holders increased from 282.2 lakh to 895.8 lakh between March 2019 and March 2024, a 217.4% increase reported by SEBI. Individual accounts increased from 354.3 lakh to 1,508.4 lakh, an increase of about 325.7% calculated from the reported levels.",
"The source ledger records the observation period, unit and institutional source for every headline figure. This is important because many errors in financial research are definition errors rather than arithmetic errors. The paper therefore treats definitions as part of the data, not as footnotes added after the analysis."
],tab=T([['Source','Main variables','Latest endpoint used'],['SEBI investor-growth study','Accounts, unique holders, unique MF investors, broker share','Mar-2024 / FY2024'],['AMFI','AUM, folios, SIP contributions and SIP accounts','Aug-2026'],['NSE Ownership Tracker','Individual, DMF and FPI ownership; equity exposure','Mar-2026'],['NSE Market Pulse','Household wealth creation','May-2026 report, Mar-2026 data'],['SEBI Investor Survey','Household participation and awareness','2025 survey'],['SEBI household-savings study','Primary/secondary household savings','FY2024–25']], [43*mm,78*mm,49*mm]),tab_source='SEBI 2025; AMFI 2026; NSE Q4 FY26; SEBI Investor Survey 2025; SEBI household-savings study 2026.'))
sections.append(SEC('7. The Account Expansion and the People Behind It',[
"The account boom is the most visible statistic in India’s retail-market expansion. SEBI reports 354.3 lakh individual demat accounts in March 2019, 403.9 lakh in March 2020 and 1,508.4 lakh in March 2024. Total demat accounts reached 1,513.8 lakh in March 2024.",
"But SEBI’s PAN-de-duplicated series gives a different scale. Unique individual demat holders rose from 282.2 lakh in March 2019 to 317.4 lakh in March 2020 and 895.8 lakh in March 2024. SEBI reports a 217.4% increase over March 2019–March 2024. The account series grew faster than the distinct-holder series.",
"This distinction changes the question rather than eliminating the account evidence. Accounts measure the size of the access layer and the infrastructure through which investors can hold securities. Unique holders measure the number of distinct individual participants under SEBI’s PAN-based definition. Both are economically useful. The mistake would be to use the first as if it were the second.",
"The post-COVID increase is therefore real on both measures. The number of distinct individual holders more than tripled between March 2019 and March 2024, while the number of individual accounts increased even more. Two processes occurred together: more people entered, and account intensity also changed."
],F3,'Figure 3. Individual demat accounts versus PAN-de-duplicated unique individual holders. Source: SEBI 2025; author’s calculation of account growth.'))
sections.append(SEC('8. Multiple Accounts: Why the Gap Matters',[
"SEBI’s analysis goes beyond the two headline levels. It notes that individuals can hold multiple demat accounts and that multiple-account ownership became more important after COVID. This matters because the gap between account growth and unique-holder growth is not random noise. It tells us something about how market access is being used.",
"Multiple-account ownership can have several explanations. Investors may use different brokers, separate long-term holdings from active trading, move between platforms, or maintain accounts across depositories. The aggregate data do not allow the paper to assign a single motive to this behaviour. It is therefore safer to interpret the widening gap as evidence of higher account intensity rather than as proof of any particular investor strategy.",
"This is also why the paper avoids creating an ‘average investor’ by dividing total accounts by unique holders and then interpreting the result as a behavioural statistic. The numerator and denominator have different institutional definitions, and legacy accounts without PAN details are excluded from the unique-holder study. The ratio can be descriptive, but it should not be over-interpreted.",
"The broader lesson is that measurement becomes more important as the market gets larger. When participation is small, an account may look like a reasonable proxy for an investor. When multiple-account ownership becomes widespread, that shortcut becomes increasingly misleading."
],tab=T([['Interpretation','Supported by evidence?','Reason'],['Accounts increased','Yes','SEBI reports large account growth'],['Distinct holders increased','Yes','PAN-de-duplicated series rose strongly'],['Multiple-account ownership became more important','Yes','SEBI explicitly identifies the post-COVID trend'],['Each additional account equals a new investor','No','The source shows multiple accounts per PAN'],['Each extra account reflects active trading','No','Account data do not establish behaviour']], [55*mm,34*mm,81*mm]),tab_source='SEBI 2025; author’s interpretation of the source definitions.'))
sections.append(SEC('9. Digital Intermediation and the Entry Channel',[
"SEBI’s top-ten CDSL DP comparison shows a major shift in how new individual accounts were added. Discount-broker DPs accounted for 24.2% of individual accounts added in FY2017, 60.6% in FY2020, 87.6% in FY2021, 89.8% in FY2022, 90.9% in FY2023 and 88.4% in FY2024. The number of discount-broker DPs among the top ten rose from two to six.",
"The evidence is strong on intermediation. Discount brokers became dominant among new individual accounts in the specific comparison used by SEBI. It is not evidence that discount brokers caused the entire national expansion. Smartphone adoption, digital payments, KYC processes, market returns, financial content and investor attention changed at the same time.",
"The same technology can support different behaviours. Lower onboarding friction can make long-term investing easier. It can also make experimentation, multiple-account opening and short-term trading easier. The broker data therefore identify a change in the entry mechanism, not the quality or persistence of the participation that followed.",
"The paper’s next step is therefore deliberate: move from the broker interface to the allocation behaviour that follows access. That is where mutual funds and SIPs become important."
],F4,'Figure 4. Discount-broker share among individual accounts added by the top ten CDSL DPs. Source: SEBI 2025, based on CDSL data.'))
sections.append(SEC('10. Mutual Funds as a Second Entry Channel',[
"Direct equity is only one route into market-linked assets. Mutual funds provide another. A household can have equity exposure through a mutual fund without directly selecting and holding the underlying shares. The mutual-fund channel therefore expands the definition of household demand beyond direct demat-account ownership.",
"SEBI’s investor-growth study reports that unique individual mutual-fund investors increased from 191.7 lakh in March 2019 to 442.5 lakh in March 2024. The same study notes that mutual funds have played an important role in bringing investors into the securities market, partly because of ease of investment and systematic investment routes.",
"This evidence is useful because it gives the paper a second unique-investor series. The direct-equity series tells us about distinct demat holders. The mutual-fund series tells us about distinct individual mutual-fund investors under its own methodology. The two should not be added mechanically, because an individual can participate through both channels.",
"The correct interpretation is therefore one of overlapping participation channels. The household demand base can expand through direct securities ownership, through pooled investment, or through both. The paper treats mutual funds as a second route into the market rather than as a substitute for direct equity."
],F5,'Figure 5. Unique individual mutual-fund investors rose from 191.7 lakh in March 2019 to 442.5 lakh in March 2024. Source: SEBI 2025.'))
sections.append(SEC('11. Mutual-Fund Scale and AUM',[
"AMFI reports mutual-fund industry AUM of ₹15.63 trillion in August 2016, ₹36.59 trillion in August 2021 and ₹87.08 trillion on August 31, 2026. Total folios reached 28.35 crore by August 2026, while equity, hybrid and solution-oriented schemes accounted for about 21.62 crore folios, where the retail segment is especially important.",
"AUM is an asset stock, not a household-saving flow. It can rise because households contribute more, because existing holdings appreciate, because money moves between funds, or because of other valuation and structural effects. This is why the paper uses AUM to establish industry scale rather than as a direct measure of household investment.",
"The growth is nevertheless economically important. A larger mutual-fund industry creates a larger intermediation channel through which household savings can reach equity, debt and other securities. The household does not need to choose every security directly for its saving to affect the demand for market-linked assets.",
"The paper therefore uses mutual-fund AUM as evidence of the expansion of the intermediation layer and unique MF investors as evidence of distinct participation. Keeping the two together provides more information than either one alone."
],F6,'Figure 6. Mutual-fund industry AUM at selected August endpoints. Source: AMFI 2026.'))
sections.append(SEC('12. SIPs and Recurring Allocation',[
"SIPs matter because they turn market participation into a repeated flow. AMFI reports annual SIP contributions of ₹43,921 crore in FY2016–17, ₹1,00,084 crore in FY2019–20 and ₹3,49,589 crore in FY2025–26. The increase continued well beyond the initial COVID shock.",
"The pattern is informative but should not be overstated. A SIP contribution is money paid into a mutual-fund scheme. It is not necessarily an equity purchase by the fund on the same day, and it is not a unique-household measure. The fund decides its allocation within the scheme according to its mandate.",
"The persistence of the flow is nevertheless relevant to the demand-side argument. A market in which households repeatedly allocate through systematic mechanisms has a different demand structure from a market in which participation is dominated by one-time account openings.",
"The important question is therefore not ‘Are SIPs good?’ It is ‘What does the persistence of SIP contributions tell us about household allocation?’ The evidence supports the narrower statement that recurring mutual-fund allocation became a very large part of the household-facing financial system."
],F7,'Figure 7. Annual SIP contributions, FY2016–17 to FY2025–26. Source: AMFI; official month-wise SIP contribution series.'))
sections.append(SEC('13. The Current SIP Run-Rate',[
"AMFI reports ₹31,115 crore of SIP contributions in April 2026, ₹30,954 crore in May, ₹31,781 crore in June, ₹31,961 crore in July and ₹32,297 crore in August. Each month was above ₹30,000 crore. The April–August total was ₹1,58,108 crore.",
"The outstanding SIP-account count reached 1,075.32 lakh and contributing SIP accounts reached 1,001.79 lakh during April–August 2026. These are account measures. One investor can have more than one SIP, and the number of accounts can change through new registrations, discontinuations and completed tenures.",
"The August contribution is best treated as a current flow observation rather than a forecast. The paper does not multiply ₹32,297 crore by twelve and call the result an annual household investment estimate. Such an annualisation would impose an assumption about future months that is not contained in the source.",
"The useful conclusion is narrower: by August 2026, recurring mutual-fund allocation was operating at a very large monthly scale. That strengthens the argument that the post-COVID transformation was not only an account-opening event."
],F8,'Figure 8. SIP contributions remained above ₹30,000 crore in each month from April through August 2026. Source: AMFI 2026.'))
sections.append(SEC('14. Participation: A Larger Market Is Not a Universal Market',[
"SEBI’s Investor Survey 2025 estimates that 9.5% of Indian households were invested in securities-market products, corresponding to about 3.21 crore households out of 33.72 crore households. Urban participation was 15%, compared with 6% in rural households. Participation was 23% in the top nine metros and 17% in Maharashtra.",
"The survey also reports that 63% of households were aware of at least one securities-market product. This is the awareness rate, not the investment-participation rate. A household can know about shares, mutual funds or bonds without investing. The paper therefore keeps the 63% awareness measure separate from the 9.5% participation measure.",
"This participation gap changes the interpretation of the account and AUM boom. The market became much larger, but household coverage remained far from universal. The post-COVID transformation therefore looks more like deepening among a growing participant base than a complete transition of household savings into securities markets.",
"SEBI also reports associations between participation and income, education and urban residence. These are descriptive associations from a household survey. They should not be interpreted as causal estimates without a research design that addresses selection and confounding."
],os.path.join(FIG,'fig09_participation.png'),'Figure 9. Awareness and household participation across selected groups. Source: SEBI Investor Survey 2025.'))
sections.append(SEC('15. Household Ownership and the Scale of Financial Demand',[
"NSE’s India Ownership Tracker provides a direct measure of the household-linked ownership channel. In March 2026, direct individual ownership of NSE-listed companies was 9.1% of market capitalisation. When direct holdings and individual ownership through domestic mutual funds are combined, individuals held 18.7% of total NSE-listed market capitalisation.",
"The combined measure is useful because it captures both direct and intermediated household exposure. But it should not be described as direct household stock selection. The household may own a mutual-fund unit while the fund owns the underlying company shares.",
"The same report puts individuals’ total equity exposure at ₹76.5 lakh crore in March 2026. This is a stock of holdings. It is evidence that household-linked demand has become economically material, but it is not evidence of how much was purchased during March or during FY26.",
"The ownership evidence therefore adds a new layer to the paper. The story has moved from access and entry to a measurable ownership position. The remaining question is how much of the ownership stock reflects fresh household saving, valuation changes, or flows through intermediaries."
],F11,'Figure 10. Direct individual ownership and combined direct-plus-indirect individual ownership, alongside DMF and FPI ownership. Source: NSE India Ownership Tracker Q4 FY26.'))
sections.append(SEC('16. From Wealth to Flows: The Stock-Flow Problem',[
"NSE estimates cumulative household equity wealth creation since April 2020 at around ₹44 lakh crore, while individuals’ total equity exposure stood at ₹76.5 lakh crore in March 2026. These figures answer different questions. The first is cumulative change in estimated household equity wealth. The second is the value of holdings at a point in time.",
"Neither should be presented as cumulative household investment. Wealth can rise because securities appreciate. It can fall because markets decline. Fresh investment is another component. NSE’s methodology derives wealth changes from changes in the value of individual holdings adjusted for net fresh investment. This is why the paper keeps wealth and flow language separate.",
"Suppose a household owns shares worth ₹10 lakh and prices rise by 20%. The portfolio becomes ₹12 lakh without a new ₹2 lakh contribution. The household’s financial wealth increased, but the new investment was zero in that example. The reverse can also happen: a household contributes fresh money while prices fall.",
"The same accounting discipline applies to mutual-fund AUM. AUM is affected by contributions, withdrawals and market prices. SIP contribution is a flow. Household equity exposure is a stock. These measures are related, but they cannot be added together as if they were the same economic quantity."
],F12,'Figure 11. Individual equity exposure and cumulative household equity wealth creation are different measures. Source: NSE India Ownership Tracker Q4 FY26 and Market Pulse May 2026.'))
sections.append(SEC('17. From Household Demand to Corporate Financing',[
"A household buying an already-listed share in the secondary market transfers ownership from one holder to another. The transaction can improve liquidity and contribute to price discovery, but the payment normally goes to the seller rather than the issuing company. A primary-market subscription is different because the issuer receives financing for newly issued securities.",
"SEBI’s May 2026 study provides a more direct flow measure. Under the revised methodology, the household-sector measure includes NPISHs and savings flowing through securities-market primary channels reached ₹6,31,510 crore in FY2024–25. The components were ₹5,12,765 crore through mutual funds, ₹95,139 crore through equity, ₹22,400 crore through debt, ₹425 crore through REITs and ₹783 crore through InvITs. The source reports ₹59,452 crore of secondary-market flows and ₹6,90,963 crore in total. The paper reproduces the source totals rather than forcing the component rows to reconcile exactly; the reported primary components sum to ₹6,31,512 crore, while the source reports ₹6,31,510 crore as the primary total.",
"These figures improve the paper because they introduce a flow that is conceptually closer to saving and financing. But they still require careful interpretation. Mutual-fund primary flows are allocations to pooled vehicles rather than direct issuer financing by each household. Equity and debt primary-market flows are closer to issuance, but the aggregate table remains a household-savings estimate rather than a micro-level retail-allocation dataset.",
"The paper therefore makes a stronger statement than an account-only study can make: household saving through securities markets reached a substantial scale, and a large portion was routed through primary channels. The exact share originating from direct retail investors and the specific firms ultimately financed remain questions for issue-level analysis."
],F13,'Figure 12. Composition of the source-reported household-sector primary-market savings measure for FY2024–25, which includes NPISHs. Source: SEBI, Household Savings through Indian Securities Market (May 2026). The component rows reproduce the source; the source-reported primary total is ₹6,31,510 crore while the displayed component rows sum to ₹6,31,512 crore.'))
sections.append(SEC('18. Domestic Investors, Foreign Flows and Market Structure',[
"NSE reports that individuals’ combined direct and indirect ownership reached 18.7% in March 2026. Direct individual ownership was 9.1%, domestic mutual funds held 11.4%, and FPI ownership stood at 15.8%. These categories overlap in the sense that individual indirect ownership is partly represented through mutual funds, so they should not be added together as if they were mutually exclusive ownership buckets.",
"It is tempting to say that domestic investors are simply replacing foreign investors. That is too simple. Individuals, mutual funds, insurers, banks and other domestic institutions have different objectives and time horizons. Ownership shares are also stocks, while FPI and domestic flows are changes over a period.",
"The useful research question is whether a larger domestic base changes market sensitivity to foreign flows. A flow-return model can test whether the relationship between FPI flows and market returns changed after the expansion of domestic participation. But such a test requires consistent flow measures, return data, volatility controls and a carefully defined post-period.",
"The ownership evidence is therefore best used as a structural fact: the domestic household-linked presence in Indian equities became substantial. It is not proof that domestic investors mechanically offset every foreign sale or that ownership itself equals new capital formation."
],F14,'Figure 13. Selected ownership categories in NSE-listed companies, March 2026; categories overlap and are not additive. Source: NSE India Ownership Tracker Q4 FY26.'))
sections.append(SEC('19. Empirical Strategy: What Can Actually Be Estimated?',[
"The descriptive evidence establishes a major expansion, but a research paper should distinguish documented facts from estimates that still need to be made. The empirical strategy therefore has four possible models. None is presented as an estimated result in this manuscript unless the underlying dataset is assembled and validated.",
"First, an interrupted time-series model can test whether a participation series changed level or trend around the COVID period: <b>Y<sub>t</sub> = β<sub>0</sub> + β<sub>1</sub>Time<sub>t</sub> + β<sub>2</sub>Post<sub>t</sub> + β<sub>3</sub>PostTrend<sub>t</sub> + ε<sub>t</sub></b>. The purpose is to test whether the path changed, not to claim that COVID alone caused the change.",
"Second, new investor entry can be related to lagged market returns and volatility: <b>NewInvestors<sub>t</sub> = α + β<sub>1</sub>Return<sub>t−1</sub> + β<sub>2</sub>Return<sub>t−2</sub> + β<sub>3</sub>Volatility<sub>t−1</sub> + γX<sub>t</sub> + ε<sub>t</sub></b>. Third, market returns can be modelled against FPI and domestic flows with a post-period interaction. Fourth, a broker panel could relate profitability to clients, turnover and funding.",
"These specifications are useful because they convert the narrative into testable questions. They are not useful if the underlying data are inconsistent. The paper therefore refuses to show a regression table without validated observations, definitions and transformations."
],tab=T([['Model','Question','Required data','Status'],['Interrupted time series','Did level/trend change?','Consistent monthly/quarterly series','Specification only'],['Investor-entry model','Do earlier returns/volatility predict entry?','Monthly entry + market data','Specification only'],['Flow-return model','Did flow/return relationships change?','FPI/domestic flows + returns','Specification only'],['Broker panel','What explains broker profitability?','Comparable firm panel','Specification only']], [38*mm,52*mm,56*mm,38*mm]),tab_source='Author’s empirical design. These are specifications, not estimated results.'))
sections.append(SEC('20. Identification and Robustness',[
"The central identification problem is that the COVID period contained many shocks at once. Pandemic restrictions, market volatility, monetary conditions, digital adoption, broker competition and household behaviour changed over overlapping windows. An interrupted time-series break is therefore not the same as a causal estimate of COVID.",
"A credible design should test alternative intervention dates around February, March and April 2020. It should inspect pre-trends and use placebo breaks in 2017, 2018 and 2019. If the same method finds equally large breaks at arbitrary earlier dates, the COVID interpretation becomes weaker.",
"Alternative dependent variables are also important. If the break appears only in raw account counts but not in PAN-de-duplicated holders, that difference matters. If SIP contributions show persistence while direct accounts do not, the interpretation should reflect that. Aggregate monthly data can also exhibit serial correlation, so HAC/Newey-West standard errors may be appropriate.",
"The robustness plan should also separate CDSL and NSDL where possible, exclude extreme COVID-market months as a sensitivity check, use lags where economically sensible, and preserve direct versus indirect ownership distinctions. These tests do not guarantee causality. They show which parts of the story remain stable under reasonable alternative assumptions."
],F16,'Figure 14. Identification logic: a statistical break is evidence of a change in the series, not proof of a single cause. Source: author’s empirical design.'))
sections.append(SEC('21. Countervailing Interpretations',[
"Every major finding has an alternative interpretation. More accounts can mean more investors, but they can also mean multiple accounts. Higher SIP contributions can mean more recurring saving, but they do not reveal the number of unique households. Higher household wealth can reflect market appreciation. Higher domestic ownership can reflect valuation changes as well as net buying.",
"Digital intermediation has the same dual character. Lower onboarding and transaction costs can widen access. They can also make short-term trading easier. SEBI’s 2026 research programme separately examines individual trading behaviour and profitability in equity derivatives, which reinforces the analytical point that participation and trading behaviour are not the same outcome. The present paper does not use derivative outcomes as evidence for the demat or SIP series.",
"Selection is another issue. Households who enter securities markets may differ systematically from households that remain outside them. The survey’s urban-rural and geographic differences are consistent with this concern, although they do not by themselves explain the mechanisms.",
"Finally, entry does not equal persistence. Account opening tells us when someone enters the system. It does not tell us whether that person remains invested for five years, holds a diversified portfolio, or earns a positive realised return. A micro-level study would be needed to answer those questions."
],F15,'Figure 15. Counter-checks applied to the headline statistics. Source: author’s analytical framework.'))
sections.append(SEC('22. Discussion: What the Evidence Supports',[
"The evidence supports a coherent sequence. Distinct individual participation expanded substantially. Account counts expanded even faster, showing that multiple-account ownership matters. Digital brokerage became a dominant entry channel among the top ten CDSL DPs. Mutual-fund investors, AUM and SIP contributions expanded, providing a recurring allocation mechanism. Individual ownership became a material component of the NSE-listed market.",
"The evidence is strongest on scale and structure. It is weaker on causality. The post-COVID period is clearly different from the pre-COVID period in several measures, but the data do not isolate COVID from the other changes that happened at the same time. The appropriate interpretation is that COVID coincided with, and may have accelerated, a broader transition toward a larger household-facing financial market.",
"The capital-formation evidence is more nuanced than the account story. SEBI’s revised methodology shows that the household-sector measure, including NPISHs, reached ₹6,90,963 crore in FY2024–25, with ₹6,31,510 crore through primary channels. That provides a macro-level flow measure. Yet the flow includes mutual-fund saving and several instruments, so it cannot be equated with direct household funding of listed companies.",
"The phrase ‘from accounts to capital’ is therefore best understood as a research pathway. India moved from a smaller access base toward a much larger household demand base for financial assets. Whether that demand materially changed corporate financing, liquidity, market stability or household welfare remains a set of follow-up questions."
],tab=T([['Evidence layer','What the paper can say','What remains open'],['Participation','Distinct holders increased','Household-level persistence'],['Intermediation','Discount brokers became dominant among new accounts in top CDSL DPs','Causal effect of digital brokerage'],['Recurring allocation','SIP flows became very large and persistent','Unique households behind SIP accounts'],['Ownership','Individuals held 18.7% direct + indirect','Exact flow contribution to ownership'],['Capital formation','Household securities-market saving reached large scale','Issuer-level retail financing and firm outcomes']], [43*mm,65*mm,66*mm]),tab_source='SEBI, AMFI and NSE; interpretation is deliberately limited to what each evidence layer measures.'))
sections.append(SEC('23. Limitations and Reproducible Research Agenda',[
"The first limitation is aggregation. Most headline measures are national totals. They cannot identify individual holding periods, risk tolerance, portfolio concentration, realised returns or household-level causal effects. The second limitation is definition. An investor is not necessarily a household, a folio is not necessarily a person, and direct-plus-indirect ownership is not the same as direct ownership.",
"The third limitation is time alignment. SEBI’s PAN-de-duplicated investor series used here ends in March 2024. AMFI’s SIP data continue through August 2026. NSE ownership and household-equity exposure reach March 2026. SEBI’s household-savings study provides annual data through FY2024–25. These windows are reported separately rather than merged into a fictional common panel.",
"The fourth limitation is causal identification. A national time series cannot easily separate pandemic effects from technology adoption, market returns, policy changes and household behaviour. Stronger work could exploit regional variation, broker-level shocks, product-level changes or credible regulatory discontinuities.",
"The reproducibility plan is straightforward: archive every official source, preserve raw and transformed data separately, maintain a variable dictionary, record the exact observation period and unit of each variable, generate figures from code, and keep the final PDF tied to the same ledger used to create the figures. The package accompanying the manuscript contains the numeric ledger, workbook and individual figure files."
],F17,'Figure 16. Reproducible research workflow. Source: author’s research protocol.'))
sections.append(SEC('24. Conclusion',[
"India’s post-COVID retail securities expansion is larger than a story about demat accounts. The strongest evidence shows a sequence: distinct individual participation increased; account intensity increased even faster; digital brokerage became a dominant entry channel among leading CDSL DPs; mutual-fund assets and unique investors expanded; SIP contributions became a large recurring allocation mechanism; and individual ownership became a material component of the NSE-listed market.",
"The measurement lesson is just as important. Accounts are not people. SIP contributions are not household wealth. Mutual-fund AUM is not the same as household wealth. Household wealth is not the same thing as fresh investment. Ownership is not automatically new corporate financing. Keeping these distinctions intact makes the argument narrower, but also more credible.",
"The household-savings evidence provides an additional bridge. SEBI’s revised methodology estimates ₹6,90,963 crore of household-sector savings including NPISHs through the securities market in FY2024–25, including ₹6,31,510 crore through primary channels. This is evidence that household saving through securities markets can be measured as a flow. It is not evidence that every component directly financed companies or that the entire flow came from the newly entered retail investors documented in the demat series.",
"The central conclusion is therefore deliberately modest: India developed a substantially larger and more digitally mediated household demand base for financial assets after the COVID period. The next research step is to trace household-originated flows into specific primary issues, follow investor persistence and holding behaviour, and estimate the proposed models using a fully validated monthly or micro-level dataset. The interesting question is no longer simply whether the retail market grew. It is what that growth actually changed."
]))

# One substantive section per main-text page. This is intentional: each module is long enough to stand alone and avoids sparse half-sections.
for sec in sections:
    st += sec
    st.append(PageBreak())
# Remove last pagebreak before appendices? keep it: starts appendices cleanly.
st += [P('Appendix A. Core Data Ledger','h1'),T([['Measure','Period','Value','Source']]+ledger,[64*mm,34*mm,40*mm,42*mm],6.7),Spacer(1,6),P('The ledger is the control document for the headline statistics. It preserves unit, period and source so a current figure cannot be silently substituted for a historical figure.','small'),PageBreak()]
defs=[['Variable','Definition','Main caution'],['Unique individual demat holders','PAN-wise de-duplicated individual holders across depositories','Not households'],['Individual demat accounts','Accounts in SEBI’s individual category','Not unique people'],['Unique MF investors','Unique individual investors under the source methodology','Not folios'],['MF folios','Mutual-fund account records','Not unique people'],['SIP contribution','Money collected through SIPs during a period','Flow; not wealth'],['MF AUM','Assets managed by mutual-fund industry','Industry stock; not household-only wealth'],['Direct individual ownership','Direct individual share of NSE-listed market cap','Stock; not flow'],['Direct + indirect individual ownership','Direct holdings plus individual ownership through domestic mutual funds','Not direct holdings only'],['Household equity exposure','Value of individual-linked equity holdings estimated by NSE','Not cumulative investment'],['Household savings via securities market','Savings routed through primary and secondary securities-market channels under SEBI’s revised methodology','Broader than direct issuer financing']]
st += [P('Appendix B. Measurement Definitions','h1'),T(defs,[45*mm,76*mm,59*mm],6.7),Spacer(1,6),P('The paper uses “retail-demand economy” narrowly to mean household demand for financial assets and market-linked investment products. It does not mean consumer demand or aggregate consumption.','small'),PageBreak()]
eq=[['Specification','Equation'],['Growth','Growth_t = ((Y_t − Y_(t−1)) / Y_(t−1)) × 100'],['Interrupted time series','Y_t = β0 + β1 Time_t + β2 Post_t + β3 PostTrend_t + ε_t'],['Investor entry','NewInvestors_t = α + β1 Return_(t−1) + β2 Return_(t−2) + β3 Volatility_(t−1) + γX_t + ε_t'],['Flow-return','R_t = α + β1 FPI_t + β2 Domestic_t + β3 Post_t + β4(FPI_t × Post_t) + γX_t + ε_t'],['Broker panel','PAT_it = α_i + λ_t + β Clients_it + γ Turnover_it + δ Funding_it + ε_it'],['Stock-flow identity','H_t − H_(t−1) = I_t + V_t']]
st += [P('Appendix C. Empirical Specifications','h1'),T(eq,[45*mm,135*mm],7.0),Spacer(1,6),P('<b>Estimation rule:</b> These are specifications, not reported results. No coefficient, p-value, standard error or causal estimate is presented unless it has been estimated from a validated and reproducible dataset.','small'),PageBreak()]
refs=[
'Association of Mutual Funds in India (AMFI). 2026. <i>Indian Mutual Fund Industry Statistics</i>, August 2026.',
'Association of Mutual Funds in India (AMFI). 2026. <i>Month-wise SIP Contribution Statistics</i>, including FY2016–17 onward annual totals and April–August 2026 observations.',
'Chauhan, Kalpeshkumar Ambalal, and Jillet Sarah Sam. 2026. “Navigating Financialization in Rural India: Stock Market Participation Through Professional Communities.” <i>Journal of South Asian Development</i>. Published online August 31, 2026. DOI: 10.1177/09731741261478420.',
'National Stock Exchange of India (NSE). 2026. <i>India Ownership Tracker: Q4 FY26</i>, March 2026.',
'National Stock Exchange of India (NSE). 2026. <i>Market Pulse</i>, May 2026.',
'Securities and Exchange Board of India (SEBI). 2025. Kalyani H., Shyni Sunil, and Prabhas Kumar Rath. “Growth of Individual Investors in Indian Securities Market.” <i>SEBI Bulletin</i>, January 2025.',
'Securities and Exchange Board of India (SEBI). 2026. <i>Investor Survey 2025: Main Report</i>. Published January 20, 2026.',
'Securities and Exchange Board of India (SEBI). 2026. <i>Household Savings through Indian Securities Market</i>. May 20, 2026.',
'Securities and Exchange Board of India (SEBI). 2026. <i>Study: Trading Behaviour of Individual Traders in the Equity Derivatives Segment (FY25–FY26)</i>. August 20, 2026.',
'Securities and Exchange Board of India (SEBI). 2026. <i>Study: Profitability of Individual Traders in the Equity Derivatives Segment (FY25–FY26)</i>. August 20, 2026.',
'Sourirajan, Sunderarajan, and Natarajan, Subhashree. 2021. “Do retail mutual fund investments represent ‘dumb money’?” <i>IIMB Management Review</i> 33(1): 71–87. DOI: 10.1016/j.iimb.2021.03.004.',
'Central Depository Services (India) Limited (CDSL). Periodic depository statistics, as used in SEBI’s published investor-growth analysis.'
]
st += [P('References','h1')]+[P(f'[{i}] {r}') for i,r in enumerate(refs,1)]+[Spacer(1,6),P('<b>Source hierarchy.</b> Reported statistics are taken from official institutional publications wherever available. The manuscript does not replace an original institutional source with a secondary article when the original source is accessible.','small'),P('<b>Data integrity statement.</b> No unestimated regression coefficient, significance level or causal estimate is presented as an observed result. Numeric figures are generated from the included verified-data ledger and workbook.','small')]

def footer(c,d):
 c.saveState(); c.setFont('DS',7.1); c.setFillColor(HexColor('#666')); c.drawString(18*mm,8.5*mm,'Aditya Makan | From Accounts to Capital'); c.drawRightString(A4[0]-18*mm,8.5*mm,f'Page {d.page}'); c.restoreState()
D.build(st,onFirstPage=footer,onLaterPages=footer)

# ---------- QA ----------
r=PdfReader(PDF); texts=[p.extract_text() or '' for p in r.pages]; total=len(texts)
# Locate appendices by exact page heading, ignoring contents page
appendix_page=None
for i,t in enumerate(texts,1):
 if i >= 4 and re.search(r'(?m)^Appendix A\. Core Data Ledger\s*$',t): appendix_page=i; break
main_start=4; main_end=appendix_page-1 if appendix_page else None
main_pages=main_end-main_start+1 if main_end else None
# Figure captions and duplicates
caps=[]
for t in texts: caps += [int(x) for x in re.findall(r'Figure (\d+)\.',t)]
from collections import Counter
cnt=Counter(caps); dup={k:v for k,v in cnt.items() if v>1}
# Page densities
lens=[len(t.strip()) for t in texts[main_start-1:main_end]] if main_end else []
# headings in contents
content_text=texts[2]
# detect accidental repeated headings
heading_nums=re.findall(r'(?m)^([0-9]+)\. ',content_text)
qa={'total_pages':total,'main_pages':main_pages,'main_start':main_start,'main_end':main_end,'appendix_start':appendix_page,'figure_captions_in_order':caps,'duplicate_figure_numbers':dup,'min_main_chars':min(lens) if lens else None,'max_main_chars':max(lens) if lens else None,'word_count':len(' '.join(texts).split()),'pdf':PDF}
with open(os.path.join(ROOT,'QA_REPORT.json'),'w') as f: json.dump(qa,f,indent=2)
# copy script and readme
shutil.copy('/mnt/data/build_final_audited.py',os.path.join(ROOT,'build_final_audited.py'))
with open(os.path.join(ROOT,'README_REPRODUCIBILITY.txt'),'w') as f:
 f.write('Audited manuscript package. Core numeric inputs are in data/verified_data_ledger.csv and data/verified_data_tables.xlsx. Figures are separate files in figures/. Official primary sources are identified in the manuscript references and source notes. No unestimated regression results are presented.\n')
# zip
if os.path.exists(ZIP): os.remove(ZIP)
with zipfile.ZipFile(ZIP,'w',zipfile.ZIP_DEFLATED) as z:
 for base,dirs,files in os.walk(ROOT):
  for fn in files:
   p=os.path.join(base,fn); z.write(p,os.path.relpath(p,ROOT))
print(json.dumps(qa,indent=2))
