"""Build ReS AI opening deck from the checked-in workshop content.
Requires python-pptx, Pillow, qrcode. Run from the repository root.
"""
from pathlib import Path
import json, math, csv
from PIL import Image
import qrcode
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN

ROOT=Path(__file__).resolve().parent.parent
D=json.loads((ROOT/'presentation/content.json').read_text())
OUT=ROOT/'iros_res_ai_2026_opening.pptx'
A=ROOT/'presentation/assets'
WEB='https://embodied-nesy.github.io/'
CONF='https://2026.ieee-iros.org/'
WORKSHOPS=CONF+'program/workshops-tutorials/'
NAVY='102B3A'; INK='173343'; TEAL='007E86'; CYAN='69D7D8'; GOLD='FFC928'; GREY='536976'; PALE='EFF5F6'; WHITE='FFFFFF'
P=Presentation();P.slide_width=Inches(13.333333);P.slide_height=Inches(7.5)
P.core_properties.title='ReS AI 2026 — Opening & Workshop Proceedings'
P.core_properties.subject='Embodied Neuro-Symbolic AI for Reliable and Safe Robotics | IROS 2026'
P.core_properties.author='ReS AI 2026 Organizers'
P.core_properties.keywords='IROS 2026, ReS AI, neuro-symbolic robotics, workshop'

def rect(s,x,y,w,h,c,radius=False):
 a=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h));a.fill.solid();a.fill.fore_color.rgb=RGBColor.from_string(c);a.line.fill.background()
 if radius:a.adjustments[0]=0.08
 return a

def text(s,txt,x,y,w,h,size=24,color=INK,bold=False,align=None,link=None):
 a=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));f=a.text_frame;f.word_wrap=True
 f.margin_left=f.margin_right=0;f.margin_top=f.margin_bottom=0
 for i,line in enumerate(txt.split('\n')):
  p=f.paragraphs[0] if i==0 else f.add_paragraph();p.text=line;p.font.name='Arial';p.font.size=Pt(size);p.font.bold=bold;p.font.color.rgb=RGBColor.from_string(color);p.space_after=Pt(5)
  if align is not None:p.alignment=align
  if link:
   a.click_action.hyperlink.address=link
 return a

def line(s,x1,y1,x2,y2,c=TEAL,width=2):
 a=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2));a.line.color.rgb=RGBColor.from_string(c);a.line.width=Pt(width)
 return a

def circle(s,x,y,d,c):
 a=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x),Inches(y),Inches(d),Inches(d));a.fill.solid();a.fill.fore_color.rgb=RGBColor.from_string(c);a.line.fill.background();return a

def pic(s,path,x,y,w,h,crop=True):
 if crop:
  im=Image.open(path);iw,ih=im.size;ir=iw/ih;tr=w/h
  p=s.shapes.add_picture(str(path),Inches(x),Inches(y),width=Inches(w),height=Inches(h))
  if ir>tr:p.crop_left=p.crop_right=(1-tr/ir)/2
  else:p.crop_top=p.crop_bottom=(1-ir/tr)/2
  return p
 return s.shapes.add_picture(str(path),Inches(x),Inches(y),width=Inches(w),height=Inches(h))

def icon(s,kind,x,y,k=1,c=TEAL):
 # Original editable vector marks, built from PowerPoint shapes.
 if kind=='ground':
  nodes=[(.12,.12),(.76,.15),(.45,.48),(.12,.8),(.8,.8)]
  for i,j in [(0,2),(1,2),(2,3),(2,4),(0,1)]:line(s,x+nodes[i][0]*k,y+nodes[i][1]*k,x+nodes[j][0]*k,y+nodes[j][1]*k,c,2)
  for a,b in nodes:circle(s,x+(a-.065)*k,y+(b-.065)*k,.13*k,c)
 elif kind=='reason':
  for a,b in [(.05,.12),(.05,.7),(.68,.41)]:rect(s,x+a*k,y+b*k,.27*k,.22*k,c,True)
  line(s,x+.32*k,y+.23*k,x+.68*k,y+.52*k,c);line(s,x+.32*k,y+.81*k,x+.68*k,y+.52*k,c)
 elif kind=='adapt':
  a=s.shapes.add_shape(MSO_SHAPE.CIRCULAR_ARROW,Inches(x),Inches(y),Inches(k),Inches(k));a.fill.solid();a.fill.fore_color.rgb=RGBColor.from_string(c);a.line.fill.background()
 elif kind=='award':
  a=s.shapes.add_shape(MSO_SHAPE.STAR_5_POINT,Inches(x+.12*k),Inches(y),Inches(.76*k),Inches(.76*k));a.fill.solid();a.fill.fore_color.rgb=RGBColor.from_string(c);a.line.fill.background()
  line(s,x+.5*k,y+.76*k,x+.5*k,y+1*k,c,4);line(s,x+.25*k,y+1*k,x+.75*k,y+1*k,c,4)

def notes(s,extra='',source='index.html; https://embodied-nesy.github.io/'):
 s.notes_slide.notes_text_frame.text='Sources: '+source+'\nChecked 26 September 2026.\n'+extra

def base(title,kicker='WELCOME',dark=False):
 s=P.slides.add_slide(P.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=RGBColor.from_string(NAVY if dark else WHITE)
 rect(s,.55,.43,.42,.055,GOLD)
 text(s,kicker.upper(),1.1,.32,10,.3,12,CYAN if dark else TEAL,True)
 text(s,title,.6,.94,12.1,.95,34,WHITE if dark else INK,True)
 line(s,.6,6.99,12.72,6.99,'36505D' if dark else 'D8E4E8',.7)
 text(s,'ReS AI 2026  /  IROS • Pittsburgh  /  September 27',.6,7.12,10,.2,10,'A7BCC5' if dark else GREY)
 text(s,f'{len(P.slides):02d}',12.1,7.08,.6,.25,11,'A7BCC5' if dark else GREY,align=PP_ALIGN.RIGHT)
 notes(s);return s

# 01: Cover
s=P.slides.add_slide(P.slide_layouts[6]);rect(s,0,0,13.333,7.5,NAVY);pic(s,ROOT/'images/banner_full.png',8.55,0,4.783,7.5)
rect(s,.6,.6,.6,.065,GOLD);text(s,'IROS 2026 WORKSHOP',.6,.9,7,.4,16,CYAN,True)
text(s,'ReS AI',.57,1.47,7.7,1,68,WHITE,True)
text(s,'Embodied Neuro-Symbolic AI\nfor Reliable and Safe Robotics',.6,2.72,7.45,1.9,33,WHITE,True)
text(s,'September 27, 2026  •  Pittsburgh, PA, USA',.6,5.15,7.5,.4,19,WHITE)
text(s,'Room 411  •  Introductory remarks at 08:25 EDT',.6,5.68,7.5,.4,18,CYAN)
text(s,'David L. Lawrence Convention Center',.6,6.13,7.5,.4,16,'B9CFD8')
rect(s,10.51,.32,2.45,1.7,WHITE,True);pic(s,A/'iros2026-logo.png',10.63,.44,2.2,1.43,False)
notes(s,'Opening remarks are at 08:25 per the workshop website. The conference listing gives the workshop block as 08:30–12:30, Room 411. Logo from the official IROS homepage; banner from this repository. Slides 1–12 form the opening sequence; paper directory slides can be skipped during the five-minute introduction. Later slides are for session transitions and closing.',f'index.html; {WORKSHOPS}; {CONF}; https://www.ieee-ras.org/event/2026-ieee-rsj-international-conference-on-intelligent-robots-and-systems-iros-61738/')

# 02: Organizers
s=base('Meet the organizers','COMMUNITY')
for i,o in enumerate(D['organizers']):
 col=i%3;row=i//3;x=.6+col*4.13;y=2.02+row*2.35
 pic(s,ROOT/o['image'],x,y,1.12,1.35)
 text(s,o['name'],x+1.3,y+.05,2.62,.75,20,INK,True)
 text(s,o['affiliation'].replace('; ','\n'),x+1.3,y+.9,2.6,.98,15,GREY)
notes(s,'Thank the organizers; the following two slides acknowledge the complete reviewer roster.')
# Full reviewer roster, with display names preserved verbatim.
with (ROOT/'ReS AI 2026 Reviewers Status.csv').open(encoding='utf-8-sig') as f:
 reviewer_rows=list(csv.DictReader(f))
reviewers=sorted({r['name'].strip() for r in reviewer_rows if r['name'].strip()},key=str.casefold)
for page in range(math.ceil(len(reviewers)/22)):
 s=base(f'Program committee  /  {page+1} of {math.ceil(len(reviewers)/22)}','WITH THANKS TO OUR REVIEWERS')
 for i,name in enumerate(reviewers[page*22:(page+1)*22]):
  col=i//11;row=i%11;x=.75+col*6.25;y=2.03+row*.405
  text(s,name,x,y,5.7,.36,21,INK)
 notes(s,'All '+str(len(reviewers))+' distinct display names from the supplied reviewer-status CSV are included verbatim, alphabetized by display name. No filtering by assignments or completion; no affiliations or reviewer-performance designations are inferred. Reviewer IDs, assignment counts, and completion status are not displayed.','ReS AI 2026 Reviewers Status.csv, name column')
# 03
s=base('From capable robots to dependable robots','WHY WE ARE HERE')
text(s,'Learning expands what robots can do.',.65,2.0,11.9,.65,31,TEAL,True)
text(s,'Real-world deployment also demands reasoning about tasks, physical constraints, safety, and incomplete information.',.65,2.95,11.7,1.35,29)
for i,(a,b) in enumerate([('Reliable','Feasible behavior under changing conditions'),('Adaptable','Learning and recovery beyond familiar situations'),('Trustworthy','Interpretable, constraint-aware decisions')]):
 x=.65+4.12*i;rect(s,x,4.8,3.87,1.5,PALE,True);text(s,a,x+.2,5.02,3.45,.4,23,TEAL,True);text(s,b,x+.2,5.55,3.45,.65,18)
notes(s,'Paraphrase of the website About section. Audience includes robot learning, cognitive robotics, planning and control, embodied AI, and neuro-symbolic AI.')
# 04
s=base('Ground. Reason. Adapt.','THE EMBODIED NEURO-SYMBOLIC WORKFLOW')
for i,(kind,a,b) in enumerate([('ground','Ground observations','Objects, relations, and task-relevant constraints'),('reason','Reason over knowledge','Skills, tasks, environments, and safety requirements'),('adapt','Adapt through learning','Interpretable behavior that remains aware of constraints')]):
 x=.65+4.12*i;rect(s,x,2.05,3.87,3.75,PALE,True);icon(s,kind,x+.23,2.35,.9);text(s,a,x+.24,3.55,3.35,.78,25,INK,True);text(s,b,x+.24,4.5,3.35,1.1,21)
text(s,'Neural learning  +  structured representations  +  symbolic reasoning',.65,6.17,12,.5,23,TEAL,True)
# 05
s=base('Questions to carry into today’s discussions','RESEARCH QUESTIONS')
qs=['What should a robot represent explicitly—and what should it learn?','How can plans remain feasible, verifiable, and safe under uncertainty?','How can robots adapt without losing interpretability or constraint awareness?','What semantics—and what human input—does robot learning need?']
for i,q in enumerate(qs):
 y=2.0+i*1.08;circle(s,.7,y+.05,.43,TEAL);text(s,str(i+1),.7,y+.09,.43,.3,16,WHITE,True,PP_ALIGN.CENTER);text(s,q,1.42,y,11,.8,25)
notes(s,'Suggested framing questions synthesized from the About, Targeted Topics, and Debate sections. These are organizer discussion prompts, not claimed verbatim quotations.')
# 06
s=base('A meeting point for learning and reasoning','WORKSHOP TOPICS')
items=[('Architectures & representations','NeSy perception, planning, and control; ontologies, scene graphs, world models.'),('Planning & verification','Neural policies with symbolic planning; safety-aware and constraint-aware decisions.'),('Memory & grounding','Continual learning, experience reuse, and perceptual grounding for reasoning.'),('People & interpretability','Auditable decisions, human–robot interaction, and collaboration.')]
for i,(a,b) in enumerate(items):
 x=.65+(i%2)*6.2;y=2.02+(i//2)*2.22;rect(s,x,y,5.9,1.97,PALE,True);text(s,a,x+.22,y+.2,5.45,.5,24,TEAL,True);text(s,b,x+.22,y+.84,5.4,.95,21)
# 07 speakers
s=base('Invited speakers and panelists','RELIABLE EMBODIED AI')
jean=next(o for o in D['organizers'] if o['name']=='Jean Oh')
invited=D['speakers']+[dict(jean,time='11:20')]
for i,o in enumerate(invited):
 x=.65+i*2.065
 pic(s,ROOT/o['image'],x,2.05,1.82,2.05)
 text(s,o['time']+(' • PANEL' if o['name']=='Jean Oh' else ' • KEYNOTE'),x,4.32,1.86,.3,12,TEAL,True)
 text(s,o['name'],x,4.87,1.86,.78,19,INK,True)
 text(s,o['affiliation'],x,5.86,1.84,.87,14,GREY)
notes(s,'Five keynote speakers plus Jean Oh as an invited panelist at 11:20. Panel participation updated by the user; Jean’s affiliation and portrait come from the organizer section.','index.html; user panel correction')
# 08 counts
s=base('This morning, at a glance','PROGRAM',True)
for i,(n,label,detail) in enumerate([('5','Keynotes','Learning, symbols, concepts, evaluation, and recovery'),('18','Accepted papers','All accepted papers invited to present posters'),('4','Oral presentations','Two sessions at 09:30 and 11:00')]):
 x=.7+4.15*i;text(s,n,x,2.0,3.7,1.2,72,GOLD,True);text(s,label,x,3.43,3.6,.55,25,WHITE,True);text(s,detail,x,4.18,3.55,1.08,21,'C7D9E0')
line(s,.7,5.65,12.6,5.65,'36505D');text(s,'10:30  Posters + coffee     •     11:20  Debate     •     12:25  Award + closing',.7,6.04,12,.55,23,WHITE)
notes(s,'18 accepted papers = 4 Oral + 14 Poster entries in index.html, cross-checked against CSV decisions. Website submission guidelines invite all accepted papers for poster presentation. No reviewer scores or rejected submissions are included.','index.html; ReS AI 2026 Submission Status.csv (acceptance decisions only)')
# 09 schedule
s=base('The morning’s program','SEPTEMBER 27 • ALL TIMES EDT')
agenda=[('08:25','Introductory remarks','Organizers'),('08:30','Alessandra Sciutti','Keynote'),('08:50','Sebastian Scherer','Keynote'),('09:10','Jiayuan Mao','Keynote'),('09:30','Oral talks • Session 1','iFlax / Logic-VLA'),('09:50','Yezhou (YZ) Yang','Keynote'),('10:10','Joyce Chai','Keynote'),('10:30','Coffee + poster session','Meet the authors'),('11:00','Oral talks • Session 2','Conformal state estimation / ActionGround'),('11:20','Panel debate','Semantics, data, and human input'),('12:25','Best Paper Award + closing','Organizers')]
for i,(tm,a,b) in enumerate(agenda):
 col=0 if i<6 else 1;row=i if i<6 else i-6;x=.65+col*6.28;y=1.95+row*.78
 text(s,tm,x,y,.93,.4,23,TEAL,True);text(s,a,x+1.12,y,4.8,.39,21,INK,True);text(s,b,x+1.12,y+.4,4.8,.3,14,GREY)
notes(s,'Follow the workshop website’s exact start times and ordering. Conference listing: Room 411, 08:30–12:30. Workshop introductory remarks begin five minutes earlier, at 08:25. Do not infer individual contributed-talk lengths.',f'index.html; {WORKSHOPS}')
# 10 oral directory
s=base('Four contributed oral presentations','PAPER TRACK')
orals=[p for p in D['papers'] if p['type']=='Oral']
for i,p in enumerate(orals):
 y=2.02+i*1.12;text(s,'09:30' if i<2 else '11:00',.7,y,.9,.4,22,TEAL,True);text(s,p['title'],1.92,y,10.7,.9,23,INK,True,link=p['url'])
notes(s,'Times indicate session starts, not individual paper start times. Full paper titles link to OpenReview.')
# 11–12 poster directory
posters=[p for p in D['papers'] if p['type']=='Poster']
for page in range(2):
 s=base(f'Explore the poster track  /  {page+1} of 2','10:30 • COFFEE + POSTERS')
 for i,p in enumerate(posters[page*7:(page+1)*7]):
  y=1.92+i*.685;rect(s,.7,y+.08,.07,.33,TEAL);text(s,p['title'],.98,y,11.7,.64,20,INK,link=p['url'])
 notes(s,'These 14 papers are listed as Poster on the website. The four oral papers are also invited to poster presentation per the submission guidelines. Titles link to their public OpenReview pages. These directory slides may be skipped during opening remarks.')
# session transition helper

def keynote(o):
 s=base(o['name'],o['time']+' EDT • KEYNOTE')
 pic(s,ROOT/o['image'],.65,2.15,3.05,3.65)
 text(s,o['affiliation'],4.18,2.15,8.4,.8,22,TEAL,True)
 text(s,o['title'],4.18,3.26,8.3,2.7,33,INK,True)
 notes(s,'Session holding slide. Speaker abstract:\n'+o['abstract'],'index.html; IROS_ReS_AI_Abstracts.docx')

def oral_session(num,subset,tm):
 s=base(f'Oral talks • Session {num}',tm+' EDT • CONTRIBUTED RESEARCH',True)
 for i,p in enumerate(subset):
  y=2.2+i*2.0;text(s,f'0{i+1}',.75,y,1,.8,40,GOLD,True);text(s,p['title'],2.0,y,10.2,1.55,30,WHITE,True,link=p['url'])
 notes(s,'Session order and start time follow the workshop website. Individual talk durations and presenter names are not specified.')
keynote(D['speakers'][0]);keynote(D['speakers'][1]);keynote(D['speakers'][2]);oral_session(1,orals[:2],'09:30');keynote(D['speakers'][3]);keynote(D['speakers'][4])
# Coffee
qrcode.make(WEB).save(A/'workshop-qr.png')
s=base('Coffee, posters, and conversation','10:30 EDT',True)
text(s,'Meet the authors.\nExplore the ideas.',.7,2.1,8.8,1.7,43,WHITE,True)
text(s,'We resume at 11:00 with Oral Talks • Session 2.',.7,4.4,8.5,1.0,26,CYAN)
text(s,'Program, abstracts, and paper links',.7,5.72,8.5,.5,22,WHITE)
text(s,'embodied-nesy.github.io',.7,6.25,8.5,.4,21,CYAN,link=WEB)
pic(s,A/'workshop-qr.png',10,2.45,2.6,2.6,False)
oral_session(2,orals[2:],'11:00')
# Debate
s=base('Semantics, data, and human input','11:20 EDT • PANEL DEBATE',True)
text(s,'What semantics must the data capture for efficient robot learning, and how important is direct human input in generating and curating that data?',.7,1.97,11.9,2.48,32,WHITE,True)
panel=[D['speakers'][1],D['speakers'][2],D['speakers'][3],jean]
for i,o in enumerate(panel):
 x=.72+3.12*i;pic(s,ROOT/o['image'],x,4.78,1.05,1.2);text(s,o['name'],x+1.2,4.91,1.7,1.05,20,WHITE,True)
notes(s,'The debate question is reproduced verbatim from the workshop website. Panel roster corrected by the user: Sebastian Scherer, Jiayuan Mao, Yezhou (YZ) Yang, and Jean Oh. No moderator or audience submission platform is specified.','index.html; user panel correction')
# Award
s=base('Best Paper Award','12:25 EDT • AWARD + CLOSING',True)
icon(s,'award',.9,2.6,2.3,GOLD)
text(s,'Recognizing outstanding\nwork in embodied\nneuro-symbolic robotics',4,2.4,8.3,2.35,35,WHITE,True)
text(s,'Thank you to our authors and reviewers.',4,5.35,8.3,.8,25,CYAN)
notes(s,'Live announcement slide. No winning paper, award presenter, or finalized award-selection procedure appears in the repository. Announce the confirmed winner verbally or add it once selected. Do not infer the winner from review scores. Website states double-blind review and highest-rated papers receiving spotlight presentations; it does not document a separate award ranking process.')
# Thanks
s=base('Thank you for joining ReS AI 2026','CONTINUE THE CONVERSATION')
text(s,'Reliable. Adaptable. Trustworthy.',.7,2.03,11.9,.75,35,TEAL,True)
text(s,'Program • Talk abstracts • Accepted papers',.7,3.17,8.5,.8,26)
text(s,'embodied-nesy.github.io',.7,4.15,8.5,.6,28,TEAL,True,link=WEB)
text(s,'Remote participation: Join on Zoom',.7,5.24,8.4,.55,22,TEAL,link='https://cmu.zoom.us/s/96000840560')
text(s,'With thanks to our speakers, authors, reviewers,\norganizers, and everyone taking part.',.7,6.02,9.1,.68,21,GREY)
pic(s,A/'workshop-qr.png',10,3.12,2.6,2.6,False)
notes(s,'Workshop and Zoom hyperlinks are taken from index.html. QR code encodes the workshop website. Source links and editorial notes are recorded in slide notes throughout the deck.')
P.save(OUT)
print(f'Saved {OUT} ({len(P.slides)} slides)')
# Geometry sanity check: no object leaves the canvas.
for i,s in enumerate(P.slides,1):
 for sh in s.shapes:
  assert sh.left>=0 and sh.top>=0 and sh.left+sh.width<=P.slide_width+100 and sh.top+sh.height<=P.slide_height+100,(i,sh.name)
print('Geometry check passed')
