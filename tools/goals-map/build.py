import json, re
from pathlib import Path
import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent

wb = openpyxl.load_workbook(HERE / 'Learning_goals_11-14_classroom_v2.xlsx', data_only=True)
ws = wb['Learning goals']
rows = list(ws.iter_rows(values_only=True))[1:]

FORMAL = {
 'U1.K1':'Cell parts','U1.K2':'Specialised cells','U1.K3':'Diffusion & osmosis',
 'U1.K4':'Cell → tissue → organism','U1.A3':'Chloroplast hunt',
 'U1.A1':'Microscope & slides','U1.A2':'Potato osmosis','U1.C1':'Cell model',
 'U2.K1':'DNA & chromosomes','U2.K2':'Mutation: sickle & malaria','U2.K3':'Variation & selection',
 'U2.A1':'Punnett squares','U2.A2':'Selection on a new case','U2.A3':'Strawberry DNA','U2.C1':'Gene comic',
 'U3.K1':'Mosquito life cycle','U3.K2':'Flowers & pollination','U3.K3':'Human reproduction',
 'U3.A1':'Break the cycle','U3.A2':'Flower dissection','U3.C1':'Seed dispersal design',
 'U4.K1':'Body organisation','U4.K2':'Vaccines vs antibiotics','U4.K3':'Digestion & organs','U4.K4':'Microbes & health risks',
 'U4.A1':'Exercise physiology','U4.A2':'Food tests','U4.A3':'Pathogen journey','U4.C1':'Health campaign',
 'U5.K1':'Photosynthesis','U5.K2':'Respiration & fermentation','U5.K3':'Chloroplasts & mitochondria','U5.A1':'Leaf chromatography','U5.A2':'Yeast experiment','U5.C1':'Design an experiment',
 'U6.K1':'Food webs','U6.K2':'Tree of life','U6.K3':'Pollinators & toxins','U6.A1':'Identification keys','U6.A2':'Fieldwork survey','U6.C1':'Hochrindl field guide',
 'U7.K1':'Particle model','U7.K2':'Atoms & formulas','U7.K3':'Separating mixtures','U7.A1':'Particle explanations','U7.A2':'Separate a mixture','U7.A3':'Seawater salinity','U7.C1':'Water filter',
 'U8.K1':'Reactions & equations','U8.K2':'Acids, pH & energy','U8.A1':'Conservation of mass','U8.A2':'Cabbage indicator','U8.A3':'Rust investigation','U8.C1':'Hot & cold reactions',
 'U9.K1':'Snow & cholera','U9.K2':'Variables & causation','U9.A1':'Plan a fair test','U9.A2':'Data & graphs','U9.A3':'Claim vs evidence','U9.C1':'Full lab report',
}

KID = {
 'U1.K1':"What's inside a cell?",'U1.K2':'Cells with superpowers','U1.K3':'Why do raisins swell?',
 'U1.K4':'Is bark a different animal?','U1.A1':'Drive a microscope','U1.A2':'The potato experiment',
 'U1.A3':'The chloroplast hunt','U1.C1':'Build your own cell',
 'U2.K1':'The book written inside you','U2.K2':'The gene that fights malaria','U2.K3':'Why are we all different?',
 'U2.A1':'Predict the babies','U2.A2':'Evolution before your eyes','U2.A3':'Extract DNA in the kitchen','U2.C1':'Diary of a gene',
 'U3.K1':"A mosquito's four lives",'U3.K2':'What flowers are really for','U3.K3':'How humans begin',
 'U3.A1':"Break the mosquito's cycle",'U3.A2':'Take a flower apart','U3.C1':'Invent a flying seed',
 'U4.K1':'From cell to body','U4.K2':'Vaccine or antibiotic?','U4.K3':'The journey of your lunch','U4.K4':'Meet the microbes',
 'U4.A1':'Make your heart race','U4.A2':"What's really in your food?",'U4.A3':'Ride along with a germ','U4.C1':'Launch a health campaign',
 'U5.K1':'How plants eat sunlight','U5.K2':'Burning fuel without fire','U5.K3':'Solar panels & power plants',
 'U5.A1':'The hidden colours of leaves','U5.A2':'Make yeast burp','U5.C1':'Design your own experiment',
 'U6.K1':'Who eats whom?','U6.K2':'The tree of all life','U6.K3':'No bees, no breakfast',
 'U6.A1':'Name that creature','U6.A2':'Count a meadow','U6.C1':'Your Hochrindl field guide',
 'U7.K1':'What everything is made of','U7.K2':'The alphabet of matter','U7.K3':'Unmix the mixture',
 'U7.A1':'Why salt melts ice','U7.A2':'Rescue the salt','U7.A3':'How salty is your sea?','U7.C1':'Build a water filter',
 'U8.K1':'When stuff transforms','U8.K2':'Acid or alkali?','U8.A1':'Nothing ever disappears',
 'U8.A2':'The cabbage detector','U8.A3':'Make iron rust on purpose','U8.C1':'Hot packs & cold packs',
 'U9.K1':'The detective of Broad Street','U9.K2':'Does it really cause it?','U9.A1':'Plan a fair test',
 'U9.A2':'Turn numbers into pictures','U9.A3':"Don't get fooled",'U9.C1':'Your first real paper',
}

EMOJI = {
 'U1.K1':'🧫','U1.K2':'🩸','U1.K3':'🍇','U1.K4':'🌳','U1.A1':'🔬','U1.A2':'🥔','U1.A3':'🧅','U1.C1':'🏗️',
 'U2.K1':'🧬','U2.K2':'🌙','U2.K3':'🌈','U2.A1':'🎲','U2.A2':'🐛','U2.A3':'🍓','U2.C1':'💬',
 'U3.K1':'🦟','U3.K2':'🌼','U3.K3':'👶','U3.A1':'⛔','U3.A2':'🌷','U3.C1':'🪂',
 'U4.K1':'🫀','U4.K2':'💉','U4.K3':'🍽️','U4.K4':'🦠','U4.A1':'🏃','U4.A2':'🧪','U4.A3':'🚂','U4.C1':'📢',
 'U5.K1':'☀️','U5.K2':'🔥','U5.K3':'🔋','U5.A1':'🍃','U5.A2':'🍞','U5.C1':'💡',
 'U6.K1':'🕸️','U6.K2':'🌲','U6.K3':'🐝','U6.A1':'🔑','U6.A2':'🧭','U6.C1':'📔',
 'U7.K1':'⚛️','U7.K2':'🔤','U7.K3':'🔀','U7.A1':'❄️','U7.A2':'🧊','U7.A3':'🧂','U7.C1':'🚰',
 'U8.K1':'💥','U8.K2':'🍋','U8.A1':'⚖️','U8.A2':'🥬','U8.A3':'🔩','U8.C1':'🌡️',
 'U9.K1':'🕵️','U9.K2':'🎯','U9.A1':'📋','U9.A2':'📊','U9.A3':'🔍','U9.C1':'🏆',
}

TRAILS = [
 {'id':'malaria','name':'🦟 The Malaria Trail','desc':'Follow malaria from a gene to the fight against it',
  'steps':['U2.K2','U3.K1','U3.A1','U4.A3','U4.K2']},
 {'id':'water','name':'💧 The Water Trail','desc':'From particles to the sea — and one famous pump',
  'steps':['U7.K1','U1.K3','U1.A2','U7.C1','U9.K1','U7.A3']},
 {'id':'energy','name':'☀️ The Energy Trail','desc':'Sunlight to a racing heart',
  'steps':['U7.K2','U5.K1','U1.A3','U5.K3','U5.K2','U4.A1','U6.K1']},
 {'id':'dna','name':'🧬 The DNA Trail','desc':'From the nucleus to the tree of all life',
  'steps':['U1.K1','U2.K1','U2.A3','U2.K2','U2.K3','U6.K2']},
 {'id':'detective','name':'🕵️ The Detective Trail','desc':'Think like a scientist, start to finish',
  'steps':['U9.K1','U9.K2','U9.A1','U9.A2','U9.A3','U9.C1']},
]

UNITS = {
 1: {'name':'Cells & microscopy',        'deep':'#7733FF','soft':'#F1EAFF','ax':330,'ay':300},
 2: {'name':'Genes & evolution',         'deep':'#B05684','soft':'#FBEFF5','ax':640,'ay':195},
 3: {'name':'Reproduction & life cycles','deep':'#E8590C','soft':'#FFEFE3','ax':950,'ay':175},
 4: {'name':'Human body & health',       'deep':'#C2374B','soft':'#FDEDEF','ax':1265,'ay':330},
 5: {'name':'Energy for life',           'deep':'#0B8A75','soft':'#CBF7ED','ax':955,'ay':795},
 6: {'name':'Ecosystems & biodiversity', 'deep':'#5B9E22','soft':'#E2FDC7','ax':1270,'ay':665},
 7: {'name':'Matter & particles',        'deep':'#2563EB','soft':'#E7EFFD','ax':330,'ay':700},
 8: {'name':'Chemical reactions',        'deep':'#B8860B','soft':'#FFF6DC','ax':620,'ay':860},
 9: {'name':'Thinking like a scientist', 'deep':'#64748B','soft':'#EEF2F6','ax':800,'ay':490},
}

BT = 'https://meridian-one-jet.vercel.app/'
ES = 'https://meridian-one-jet.vercel.app/eatsun.html'
NODE_LINKS = {
 'U6.K2': [{'t':'Build a Tree — phylogenetics trainer','u':BT,'k':'trainer'}],
 'U1.A3': [{'t':'Eat the Sun — cells & photosynthesis trainer','u':ES,'k':'trainer'}],
 'U5.K1': [{'t':'Eat the Sun — cells & photosynthesis trainer','u':ES,'k':'trainer'}],
 'U5.K3': [{'t':'Eat the Sun — cells & photosynthesis trainer','u':ES,'k':'trainer'}],
}

nodes = []
for r in rows:
    code = str(r[0]).strip()
    if not re.match(r'^U\d+\.[KAC]\d+$', code): continue
    unit = int(code[1:code.index('.')])
    strand = code.split('.')[1][0]
    nodes.append({
        'id': code, 'unit': unit, 'strand': strand,
        'label': KID[code], 'formal': FORMAL[code], 'emoji': EMOJI[code],
        'text': str(r[4]).strip(),
        'evidence': str(r[5]).strip(),
        'links': NODE_LINKS.get(code, []),
    })

def E(s, t, type_, reason, links=None):
    return {'s': s, 't': t, 'type': type_, 'reason': reason, 'links': links or []}

edges = [
 # U1
 E('U1.K1','U1.K2','unit','Cell types build on knowing the parts.'),
 E('U1.K1','U1.A1','unit','Draw and label what you see down the microscope.'),
 E('U1.K3','U1.A2','unit','Osmosis predicted on paper, then tested in a potato.'),
 E('U1.K2','U1.C1','unit','The model must include one specialised cell.'),
 E('U1.K1','U1.A3','unit','Know the organelles, then hunt for chloroplasts.'),
 E('U1.K2','U1.K4','unit','Specialised cells team up into tissues and organs.'),
 E('U1.K1','U1.C1','unit','The model shows every organelle doing its job.'),
 # U2
 E('U2.K1','U2.K2','unit','A mutation is a change in the DNA you can now describe.'),
 E('U2.K2','U2.K3','unit','Harmful or useful depending on environment — that is what selection acts on.'),
 E('U2.K3','U2.A1','unit','Alleles become predictions in a Punnett square.'),
 E('U2.K3','U2.A2','unit','Selection logic applied to a case you were never taught.'),
 E('U2.K1','U2.A3','unit','Extract and actually see the DNA you learned about.'),
 E('U2.K2','U2.C1','unit','The comic must show mutation correctly.'),
 E('U2.K3','U2.C1','unit','The comic must show inheritance and selection correctly.'),
 # U3
 E('U3.K1','U3.A1','unit','Know the cycle to know where to break it.'),
 E('U3.K2','U3.A2','unit','Parts named on the diagram, then found in a real flower.'),
 E('U3.K2','U3.C1','unit','Dispersal is the last chapter of the flower story.'),
 # U4
 E('U4.K1','U4.A3','unit','The pathogen is traced through the systems you mapped.'),
 E('U4.K2','U4.A3','unit','On its way the pathogen meets the defences.'),
 E('U4.K4','U4.A3','unit','Know your microbe before you trace it.'),
 E('U4.K3','U4.A1','unit','Heart, lungs and circulation measured under load.'),
 E('U4.K3','U4.A2','unit','Nutrients named, then detected in real food.'),
 E('U4.K2','U4.C1','unit','A vaccination campaign needs the immunology.'),
 E('U4.K3','U4.C1','unit','A diet campaign needs the nutrition.'),
 E('U4.K4','U4.C1','unit','A smoking campaign needs the health effects.'),
 # U5
 E('U5.K1','U5.A1','unit','The pigments behind the photosynthesis equation, separated.'),
 E('U5.K2','U5.A2','unit','Fermentation from the comparison table, measured in yeast.'),
 E('U5.K1','U5.C1','unit','Design an experiment around photosynthesis.'),
 E('U5.K2','U5.C1','unit','…or around respiration.'),
 E('U5.K1','U5.K3','unit','Photosynthesis makes the sugar this goal follows.'),
 E('U5.K2','U5.K3','unit','Respiration burns it — in plants too.'),
 # U6
 E('U6.K1','U6.K3','unit','Toxins travel along the web you drew.'),
 E('U6.K2','U6.A1','unit','Classification in theory, identification keys in the field.'),
 E('U6.K1','U6.A2','unit','Webs and biodiversity need real field data.'),
 E('U6.K1','U6.C1','unit','Build the food web from your own observations.'),
 E('U6.K3','U6.C1','unit','Predict what one species loss does to the system.'),
 # U7
 E('U7.K1','U7.A1','unit','The particle model explains what you observed.'),
 E('U7.K3','U7.A2','unit','Choose the separation method, then run it.'),
 E('U7.K3','U7.C1','unit','Filtration scaled up into a working water filter.'),
 E('U7.K3','U7.A3','unit','Evaporation, now used to weigh the salt in the sea.'),
 # U8
 E('U8.K1','U8.A1','unit','The word equation predicts the sealed-bag masses.'),
 E('U8.K2','U8.A2','unit','pH theory tested with your own indicator.'),
 E('U8.K1','U8.A3','unit','Signs of a slow reaction, watched for a week.'),
 E('U8.K2','U8.C1','unit','Exothermic and endothermic, compared by experiment.'),
 # U9
 E('U9.K2','U9.A1','unit','Variables and controls become a written plan.'),
 E('U9.A1','U9.A2','unit','A planned test produces data worth recording.'),
 E('U9.A2','U9.C1','unit','Tables and graphs become a full report.'),
 E('U9.K2','U9.A3','unit','Correlation vs causation is the judging rule.'),
 E('U9.K1','U9.A3','unit',"Snow's 1854 logic applied to today's headlines."),
 # cross
 E('U7.K1','U1.K3','cross','The particle model is what makes diffusion and osmosis make sense.'),
 E('U7.K2','U5.K1','cross','H2O, CO2 and O2 — the formulas reappear in the photosynthesis equation.'),
 E('U8.K1','U5.K2','cross','Word equations are the grammar of the respiration equations.'),
 E('U7.K3','U5.A1','cross','Chromatography learned on ink, reused on leaf pigments.'),
 E('U1.K1','U2.K1','cross','DNA lives in the nucleus you labelled.'),
 E('U2.K2','U3.K1','cross','The sickle-cell story meets the mosquito that spreads malaria.'),
 E('U3.K1','U4.A3','cross','The life cycle delivers the pathogen into the body.'),
 E('U2.K3','U6.K2','cross','Variation and selection are what grow the tree of life.',
   [{'t':'Build a Tree — build the tree from differences','u':BT,'k':'trainer'}]),
 E('U3.K2','U6.K3','cross','Pollination is exactly why pollinators matter for food.'),
 E('U5.K2','U4.A1','cross','Respiration explains the racing heart and faster breathing.'),
 E('U4.K3','U5.K2','cross','Digested glucose is the fuel respiration burns.'),
 E('U8.K2','U4.K3','cross','Stomach acid: pH chemistry inside digestion.'),
 E('U7.C1','U4.A3','cross','Filtered is not safe to drink — the cholera lesson.'),
 E('U1.K2','U4.K1','cross','Specialised cells are what tissues and organs are made of.'),
 E('U3.K3','U2.K3','cross','Sexual reproduction is where variation comes from.'),
 E('U5.K1','U6.K1','cross','Photosynthesis feeds every food chain.'),
 E('U4.K4','U6.K2','cross','The microbe types each sit on their own branch of the tree.'),
 E('U9.A3','U4.C1','cross','Every claim in the campaign must survive the evidence test.'),
 E('U9.K1','U4.A3','cross','Cholera twice: as an epidemiological case and as a pathogen.'),
 E('U6.A2','U9.A2','cross','Field counts become tables, graphs and averages.'),
 E('U9.A1','U5.C1','cross','An experiment design is a fair test by another name.'),
 E('U1.A3','U5.K1','cross','Green parts have chloroplasts — and that is exactly where photosynthesis happens.',
   links=[{'t':'Eat the Sun — cells & photosynthesis trainer','u':ES,'k':'trainer'}]),
 E('U1.K4','U4.K1','cross','The same level ladder — cell, tissue, organ, organism — runs through the human body.'),
 E('U5.K3','U6.K1','cross','Squirrels cannot photosynthesise — that is why food chains exist.'),
 E('U9.K1','U6.K1','cross','The pine-cone hypotheses: lovable, testable — and checkable.'),
 E('U7.A3','U9.A2','cross','Salinity measurements become tables, averages and a world map.'),
]

# validation
ids = {n['id'] for n in nodes}
assert len(nodes) == 60, f'expected 60 nodes, got {len(nodes)}'
assert len(ids) == 60
pairs = set()
for e in edges:
    assert e['s'] in ids, 'unknown source ' + e['s']
    assert e['t'] in ids, 'unknown target ' + e['t']
    assert e['s'] != e['t']
    key = tuple(sorted([e['s'], e['t']]))
    assert key not in pairs, 'duplicate edge ' + str(key)
    pairs.add(key)
# every node connected?
deg = {i: 0 for i in ids}
for e in edges:
    deg[e['s']] += 1; deg[e['t']] += 1
isolated = [i for i, d in deg.items() if d == 0]
print('nodes:', len(nodes), 'edges:', len(edges),
      f"(unit {sum(1 for e in edges if e['type']=='unit')}, cross {sum(1 for e in edges if e['type']=='cross')})")
print('isolated nodes:', isolated or 'none')

assert set(KID) == ids, 'KID labels do not match goal codes: ' + str(set(KID) ^ ids)
assert set(EMOJI) == ids, 'EMOJI set does not match goal codes: ' + str(set(EMOJI) ^ ids)
assert set(FORMAL) == ids, 'FORMAL set does not match goal codes: ' + str(set(FORMAL) ^ ids)
for t in TRAILS:
    for s in t['steps']:
        assert s in ids, f"trail {t['id']}: unknown step {s}"
print('trails:', len(TRAILS))

data = {'units': UNITS, 'nodes': nodes, 'edges': edges, 'trails': TRAILS}
tpl = open(HERE / 'template.html', encoding='utf-8').read()
html = tpl.replace('__DATA__', json.dumps(data, ensure_ascii=False))
open(ROOT / 'goals-map-11-14.html', 'w', encoding='utf-8').write(html)
print('written', ROOT / 'goals-map-11-14.html', ',', len(html), 'bytes')
