import json, re
from pathlib import Path
import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
TEMPLATE = HERE.parent / 'goals-map' / 'template.html'   # shared with the 11-14 map

wb = openpyxl.load_workbook(HERE / 'Learning_goals_15-18_classroom.xlsx', data_only=True)
ws = wb['Learning goals']
rows = list(ws.iter_rows(values_only=True))[1:]

FORMAL = {
 'U1.K1':'Water & hydrogen bonds','U1.K2':'Biomolecules & protein structure','U1.K3':'Enzyme kinetics & inhibition',
 'U1.A1':'Catalase practical','U1.A2':'Food tests','U1.A3':'Enzyme case transfer','U1.C1':'Enzyme investigation design',
 'U2.K1':'Prokaryotes, eukaryotes & endosymbiosis','U2.K2':'Membranes & transport',
 'U2.A1':'SA:V & diffusion practical','U2.A2':'Potato osmolarity','U2.C1':'Transport model & cholera case',
 'U3.K1':'Transcription & translation','U3.K2':'Chromatin, repair & mutation','U3.K3':'Gene regulation & DNA tech',
 'U3.K4':'Somatic vs germline','U3.A1':'Codon problems & point mutations','U3.A2':'Gel electrophoresis case',
 'U3.A3':'Biotech case judgement','U3.C1':'Clinic rulebook',
 'U4.K1':'Somatic vs germ cells','U4.K2':'Mitosis, checkpoints & cancer','U4.K3':'Meiosis & inheritance patterns',
 'U4.A1':'Root-tip mitosis practical','U4.A2':'Genetics problems & chi-square','U4.C1':'Invent-an-organism problem',
 'U5.K1':'ATP & aerobic respiration','U5.K2':'Photosynthesis in depth','U5.K3':'Respiration vs fermentation',
 'U5.A1':'Leaf-disc photosynthesis','U5.A2':'Yeast rate practical','U5.C1':'Carbon-energy cycle model',
 'U6.K1':'Natural vs artificial selection','U6.K2':'Sickle-malaria co-evolution','U6.K3':'Hardy-Weinberg, speciation & cladograms',
 'U6.K4':'Ecosystems & climate impact','U6.A1':'Hardy-Weinberg calculation','U6.A2':'Build & read cladograms',
 'U6.A3':'Alpine fieldwork & Simpson index','U6.C1':'Breeding plan or policy brief',
 'U7.K1':'Microbes in space & biofilms','U7.K2':'Immunity & resistance','U7.K3':'Coordination & homeostasis',
 'U7.K4':'Gas exchange & circulation','U7.A1':'Biofilm case transfer','U7.A2':'Exercise physiology practical',
 'U7.A3':'Reaction-time t-test','U7.C1':'Space-mission briefing',
 'U8.K1':'Atoms & periodic trends','U8.K2':'Bonding & properties','U8.K3':'Shape, polarity & IMFs',
 'U8.A1':'Mole calculations','U8.A2':'Solubility practical','U8.A3':'Solution preparation','U8.C1':'3D molecular models',
 'U9.K1':'Reaction types & redox','U9.K2':'Energetics & rates','U9.K3':'Equilibrium & buffers',
 'U9.A1':'Rate practical','U9.A2':'Vinegar titration','U9.A3':'Buffer practical','U9.C1':'Consumer fair test',
 'U10.K1':'Bradford Hill guidelines','U10.K2':'Study designs','U10.A1':'Research design',
 'U10.A2':'Statistics toolkit','U10.A3':'Causality verdict','U10.C1':'Independent investigation (IA)',
}

KID = {
 'U1.K1':'Why ice floats','U1.K2':'The same four moves','U1.K3':'Locks, keys and saboteurs',
 'U1.A1':'Push catalase to the edge','U1.A2':'Interrogate a mystery food','U1.A3':'Milk, fever and poison',
 'U1.C1':'Plan the perfect experiment',
 'U2.K1':'Your cells have tenants','U2.K2':'Customs at the cell border',
 'U2.A1':'Why cells stay small','U2.A2':"Find the potato's secret number",'U2.C1':'When a gate jams open',
 'U3.K1':'Read, copy, build','U3.K2':'Typos, fixes and bookmarks','U3.K3':'Switches, scissors, photocopiers',
 'U3.K4':'Edits that outlive you','U3.A1':'One letter from disaster','U3.A2':'DNA on the witness stand',
 'U3.A3':'Judge the edit','U3.C1':'Run your own gene clinic',
 'U4.K1':'Two kinds of you','U4.K2':'The brakes that stop cancer','U4.K3':'The deck gets shuffled',
 'U4.A1':'Catch cells dividing','U4.A2':'Beat the odds (with a test)','U4.C1':'Breed your own monster',
 'U5.K1':'Three rooms, one fire','U5.K2':'Sunlight, split and stored','U5.K3':'The emergency generator',
 'U5.A1':'Make leaves float on light','U5.A2':'How hot does yeast like it?','U5.C1':'One cycle to run it all',
 'U6.K1':'Evolution under management','U6.K2':'An arms race in your blood','U6.K3':'The algebra of evolution',
 'U6.K4':"The planet's balance sheet",'U6.A1':'Count hidden alleles','U6.A2':'Draw the family tree of life',
 'U6.A3':'Audit an alpine meadow','U6.C1':'Brief the decision-makers',
 'U7.K1':'Salmonella goes to space','U7.K2':"Your body's standing army",'U7.K3':"The body's thermostat",
 'U7.K4':'Breathing at altitude','U7.A1':'Microbes on a Mars base','U7.A2':'Your heart vs the Alps',
 'U7.A3':'How fast are your nerves?','U7.C1':'Protect the crew',
 'U8.K1':"The periodic table's logic",'U8.K2':'Why salt snaps, metal bends','U8.K3':'The geometry of stickiness',
 'U8.A1':'Count atoms by weighing them','U8.A2':'Like dissolves like - prove it','U8.A3':'Mix a life-saving drink',
 'U8.C1':'Build life in 3D',
 'U9.K1':'Who stole the electrons?','U9.K2':'The hill reactions must climb','U9.K3':'Push it - it pushes back',
 'U9.A1':'Race the fizz','U9.A2':'How sour, exactly?','U9.A3':"Why your blood won't turn acid",
 'U9.C1':'Test what the ad claims',
 'U10.K1':'Nine clues to a cause','U10.K2':'Every study has a blind spot','U10.A1':'Ask nature the right way',
 'U10.A2':'Is the difference real?','U10.A3':'Headline on trial','U10.C1':'Your first real paper',
}

EMOJI = {
 'U1.K1':'🧊','U1.K2':'🧱','U1.K3':'🔐','U1.A1':'🫧','U1.A2':'🍱','U1.A3':'💊','U1.C1':'📐',
 'U2.K1':'🏠','U2.K2':'🛂','U2.A1':'📦','U2.A2':'🥔','U2.C1':'🧨',
 'U3.K1':'🏭','U3.K2':'🔖','U3.K3':'🖨️','U3.K4':'⏳','U3.A1':'✍️','U3.A2':'⚖️','U3.A3':'🧑‍⚖️','U3.C1':'🏥',
 'U4.K1':'🪞','U4.K2':'🛑','U4.K3':'🃏','U4.A1':'🔬','U4.A2':'🎲','U4.C1':'🐉',
 'U5.K1':'🔥','U5.K2':'☀️','U5.K3':'🔌','U5.A1':'🍃','U5.A2':'🫙','U5.C1':'♻️',
 'U6.K1':'🐄','U6.K2':'🩸','U6.K3':'🧮','U6.K4':'🌍','U6.A1':'🔢','U6.A2':'🌲','U6.A3':'⛰️','U6.C1':'🏛️',
 'U7.K1':'🚀','U7.K2':'🛡️','U7.K3':'🌡️','U7.K4':'🏔️','U7.A1':'👾','U7.A2':'🫀','U7.A3':'⏱️','U7.C1':'🧑‍🚀',
 'U8.K1':'⚛️','U8.K2':'🧂','U8.K3':'📏','U8.A1':'⚗️','U8.A2':'💧','U8.A3':'🥤','U8.C1':'🧩',
 'U9.K1':'🔋','U9.K2':'🎢','U9.K3':'🪃','U9.A1':'💨','U9.A2':'🍶','U9.A3':'🧪','U9.C1':'🛒',
 'U10.K1':'🕵️','U10.K2':'🗂️','U10.A1':'📝','U10.A2':'📉','U10.A3':'📰','U10.C1':'🏆',
}

TRAILS = [
 {'id':'molecule','name':'🧱 The Molecule Trail','desc':'Chemistry becomes biology: from atoms to the four families of life',
  'steps':['U8.K1','U8.K2','U8.K3','U1.K1','U1.K2','U8.C1']},
 {'id':'code','name':'🧬 The Code Trail','desc':'From a strand of DNA to a courtroom gel',
  'steps':['U3.K1','U3.A1','U3.K2','U3.K3','U3.A2']},
 {'id':'sickle','name':'🩸 The Sickle Trail','desc':'One letter, one parasite, one equation - the best story in biology',
  'steps':['U4.K3','U3.A1','U6.K2','U6.K3','U6.A1']},
 {'id':'atp','name':'⚡ The ATP Trail','desc':'From enzymes to your own pulse in the Alps',
  'steps':['U1.K3','U5.K1','U5.K3','U7.K4','U7.A2']},
 {'id':'space','name':'🚀 The Space Microbes Trail','desc':'Salmonella in orbit to a briefing for the crew',
  'steps':['U7.K1','U3.K3','U7.K2','U7.A1','U7.C1']},
 {'id':'causality','name':'🕵️ The Causality Trail','desc':'From Bradford Hill to your own investigation',
  'steps':['U10.K1','U10.K2','U10.A2','U10.A3','U10.C1']},
]

UNITS = {
 1:  {'name':'Biomolecules & enzymes',      'deep':'#7733FF','soft':'#F1EAFF','ax':330, 'ay':250},
 2:  {'name':'Cells & membranes',           'deep':'#B05684','soft':'#FBEFF5','ax':625, 'ay':165},
 3:  {'name':'Molecular genetics',          'deep':'#E8590C','soft':'#FFEFE3','ax':950, 'ay':160},
 4:  {'name':'Division & inheritance',      'deep':'#C2374B','soft':'#FDEDEF','ax':1265,'ay':295},
 5:  {'name':'Energy in cells',             'deep':'#0B8A75','soft':'#CBF7ED','ax':1290,'ay':625},
 6:  {'name':'Evolution & ecology',         'deep':'#5B9E22','soft':'#E2FDC7','ax':1000,'ay':830},
 7:  {'name':'Microbes & body systems',     'deep':'#0E7490','soft':'#E0F4FA','ax':655, 'ay':865},
 8:  {'name':'Atoms, bonding & the mole',   'deep':'#2563EB','soft':'#E7EFFD','ax':340, 'ay':770},
 9:  {'name':'Reactions & energy',          'deep':'#B8860B','soft':'#FFF6DC','ax':300, 'ay':505},
 10: {'name':'Thinking like a scientist',   'deep':'#64748B','soft':'#EEF2F6','ax':800, 'ay':485},
}

SITE = 'https://meridian-one-jet.vercel.app/'

def L(t, u, k):
    return {'t': t, 'u': u, 'k': k}

TRAINER_BT = L('Build a Tree - phylogenetics trainer', SITE, 'trainer')
TRAINER_MB = L('Make and Break - reactions & ATP trainer', SITE + 'makebreak.html', 'trainer')
TRAINER_CC = L('Copy the Code - DNA replication trainer', SITE + 'copycode.html', 'trainer')

NODE_LINKS = {
 'U1.K2': [TRAINER_MB],
 'U5.K1': [TRAINER_MB],
 'U3.K1': [TRAINER_CC],
 'U3.K2': [TRAINER_CC],
 'U6.A2': [TRAINER_BT],
 'U6.K3': [TRAINER_BT],
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
 # U1 Biomolecules & enzymes
 E('U1.K1','U1.K2','unit','Condensation releases a water molecule with every brick it lays.'),
 E('U1.K2','U1.K3','unit','Enzymes are proteins - their shape is their job.'),
 E('U1.K3','U1.A1','unit','Temperature and pH theory, tested on real catalase.'),
 E('U1.K2','U1.A2','unit','The molecules you described, detected in an unknown sample.'),
 E('U1.K3','U1.A3','unit','Same enzyme logic, case you have never seen.'),
 E('U1.K3','U1.C1','unit','Your investigation needs a mechanism worth asking about.'),
 E('U1.A1','U1.C1','unit','The catalase practical is the rehearsal for your own design.'),
 # U2 Cells & membranes
 E('U2.K1','U2.K2','unit','Every cell type you compared is wrapped in this membrane.'),
 E('U2.K2','U2.A1','unit','Diffusion has a speed limit - the cubes make it visible.'),
 E('U2.K2','U2.A2','unit','Osmosis theory becomes a number you can read off a graph.'),
 E('U2.K2','U2.C1','unit','Model one transport process - then break it.'),
 # U3 Molecular genetics
 E('U3.K1','U3.K2','unit','The text you can read is packed, copied and repaired.'),
 E('U3.K1','U3.A1','unit','Base pairing and codons, now with pencil and paper.'),
 E('U3.K2','U3.A1','unit','Predict what one typo does to the protein.'),
 E('U3.K3','U3.A2','unit','PCR and gels in theory, then a gel in your hands.'),
 E('U3.K3','U3.C1','unit','The clinic menu is built from these technologies.'),
 E('U3.K4','U3.A3','unit','Somatic or germline - the distinction that decides the case.'),
 E('U3.K4','U3.C1','unit','Every clinic rule stands on this distinction.'),
 # U4 Division & inheritance
 E('U4.K1','U4.K3','unit','Germ cells are where meiosis happens.'),
 E('U4.K1','U4.K2','unit','Somatic cells divide by mitosis - and must be policed.'),
 E('U4.K2','U4.A1','unit','The stages you memorised, found in a real root tip.'),
 E('U4.K3','U4.A2','unit','Inheritance patterns become solvable problems.'),
 E('U4.A2','U4.C1','unit','Solve enough problems and you can write one.'),
 # U5 Energy in cells
 E('U5.K1','U5.K3','unit','Take away the oxygen and see what the cell does.'),
 E('U5.K2','U5.A1','unit','Limiting factors, measured with floating leaf discs.'),
 E('U5.K3','U5.A2','unit','Fermentation theory, counted in bubbles of CO2.'),
 E('U5.K1','U5.C1','unit','Respiration is one half of the cycle you will build.'),
 E('U5.K2','U5.C1','unit','Photosynthesis is the other half.'),
 # U6 Evolution & ecology
 E('U6.K1','U6.K2','unit','The sharpest natural-selection case study there is.'),
 E('U6.K1','U6.C1','unit','A breeding plan is artificial selection with a budget.'),
 E('U6.K2','U6.A1','unit','The sickle allele, counted with the equation.'),
 E('U6.K3','U6.A1','unit','Hardy-Weinberg, loaded with real data.'),
 E('U6.K3','U6.A2','unit','Cladograms read in theory, then built from data.'),
 E('U6.K4','U6.A3','unit','Biodiversity claims need field numbers.'),
 E('U6.K4','U6.C1','unit','The policy brief defends what the diagrams describe.'),
 # U7 Microbes & body systems
 E('U7.K1','U7.A1','unit','Biofilm logic, moved to an environment nobody taught you.'),
 E('U7.K1','U7.C1','unit','The briefing starts with what space does to microbes.'),
 E('U7.K2','U7.C1','unit','Resistance and weakened immunity go in the briefing too.'),
 E('U7.K3','U7.A3','unit','Reaction time is the nervous system, measured.'),
 E('U7.K4','U7.A2','unit','Gas exchange and circulation, under load and at altitude.'),
 # U8 Atoms, bonding & the mole
 E('U8.K1','U8.K2','unit','Electron configurations decide who bonds with whom.'),
 E('U8.K2','U8.K3','unit','Bonds set the shape; shape sets the stickiness.'),
 E('U8.K1','U8.A1','unit','Atomic masses are what the mole counts.'),
 E('U8.K3','U8.A2','unit','Predict solubility from forces, then test it.'),
 E('U8.A1','U8.A3','unit','The calculation becomes a solution you can drink.'),
 E('U8.K3','U8.C1','unit','Shapes and bonds, now in your hands as models.'),
 # U9 Reactions & energy
 E('U9.K1','U9.A2','unit','Acid-base chemistry, measured drop by drop.'),
 E('U9.K2','U9.A1','unit','Collision theory predicts the fizz race.'),
 E('U9.K3','U9.A3','unit','Buffer theory against real pH meters.'),
 E('U9.K2','U9.C1','unit','Hand warmers are enthalpy you can sell.'),
 E('U9.K1','U9.C1','unit','Antacids are acid-base chemistry with marketing.'),
 # U10 Thinking like a scientist
 E('U10.K1','U10.A3','unit','The nine clues, applied to a headline you have never seen.'),
 E('U10.K2','U10.A3','unit','First ask: what design produced this claim?'),
 E('U10.A1','U10.C1','unit','The IA starts with a question asked properly.'),
 E('U10.A2','U10.C1','unit','The IA stands on honest statistics.'),
 # cross-unit bridges
 E('U8.K3','U1.K1','cross','Hydrogen bonding: explained by chemistry, exploited by water.'),
 E('U1.K2','U3.K1','cross','Proteins are what genes are for.'),
 E('U1.K3','U5.K1','cross','Respiration is a relay race run by enzymes.'),
 E('U2.K1','U5.K1','cross','Mitochondria: the tenants that pay rent in ATP.'),
 E('U2.C1','U8.A3','cross','The cure for cholera is a solution you can now mix yourself.'),
 E('U3.A1','U6.K2','cross','GAG to GTG: the typo behind the whole sickle story.'),
 E('U4.K3','U6.K3','cross','Meiosis makes the variation Hardy-Weinberg counts.'),
 E('U4.K1','U3.K4','cross','Somatic versus germline, met twice: in cells and in consequences.'),
 E('U5.K2','U6.K4','cross','Photosynthesis is the planet’s carbon intake.'),
 E('U5.K1','U7.K4','cross','Why you breathe: the electron transport chain wants that oxygen.'),
 E('U5.K3','U7.A2','cross','Uphill at Hochrindl, your muscles switch to the emergency generator.'),
 E('U6.K1','U7.K2','cross','Antibiotic resistance is natural selection on fast-forward.'),
 E('U8.K1','U9.K1','cross','Electron bookkeeping starts with the shell structure.'),
 E('U9.K2','U1.K3','cross','Enzymes are catalysts - the same energy hill, lowered.'),
 E('U9.A3','U7.K3','cross','The blood buffer is homeostasis written in pure chemistry.'),
 E('U4.A2','U10.A2','cross','Chi-square learned on crosses, reused on any data you collect.'),
 E('U7.A3','U10.A2','cross','The t-test twice: your reaction times are data too.'),
 E('U7.K2','U10.A3','cross','Health headlines are causality cases in the wild.'),
 E('U10.K2','U1.C1','cross','A good investigation plan is a study design in miniature.'),
 E('U10.A1','U5.A1','cross','Leaf discs behave only for a well-planned design.'),
 E('U10.A1','U9.C1','cross','The fair test is where design meets the shopping aisle.'),
 E('U7.K1','U3.K3','cross','Gene switches flipped by the environment - the operon logic, in orbit.'),
 E('U4.K2','U3.K2','cross','Checkpoints read the damage reports that repair enzymes write.'),
 E('U6.A3','U10.A2','cross','Simpson’s index is statistics with mud on its boots.'),
 E('U9.A2','U10.A2','cross','Uncertainty: where titration meets statistics.'),
 E('U8.A2','U2.K2','cross','Membranes exist because like dissolves like.'),
 E('U2.A1','U7.K4','cross','Alveoli: the surface-area lesson, built into your lungs.'),
 E('U1.A2','U4.A2','cross','Results against expectations - the chi-square habit starts early.'),
]

# validation
ids = {n['id'] for n in nodes}
assert len(nodes) == 68, f'expected 68 nodes, got {len(nodes)}'
assert len(ids) == 68
pairs = set()
for e in edges:
    assert e['s'] in ids, 'unknown source ' + e['s']
    assert e['t'] in ids, 'unknown target ' + e['t']
    assert e['s'] != e['t']
    key = tuple(sorted([e['s'], e['t']]))
    assert key not in pairs, 'duplicate edge ' + str(key)
    pairs.add(key)
deg = {i: 0 for i in ids}
for e in edges:
    deg[e['s']] += 1; deg[e['t']] += 1
isolated = [i for i, d in deg.items() if d == 0]
print('nodes:', len(nodes), 'edges:', len(edges),
      f"(unit {sum(1 for e in edges if e['type']=='unit')}, cross {sum(1 for e in edges if e['type']=='cross')})")
print('isolated nodes:', isolated or 'none')
assert not isolated, 'isolated: ' + str(isolated)

assert set(KID) == ids, 'KID mismatch: ' + str(set(KID) ^ ids)
assert set(EMOJI) == ids, 'EMOJI mismatch: ' + str(set(EMOJI) ^ ids)
assert set(FORMAL) == ids, 'FORMAL mismatch: ' + str(set(FORMAL) ^ ids)
assert len(set(EMOJI.values())) == len(EMOJI), 'duplicate emoji'
units_used = {n['unit'] for n in nodes}
assert units_used == set(UNITS), 'unit mismatch: ' + str(units_used ^ set(UNITS))
for t in TRAILS:
    for s in t['steps']:
        assert s in ids, f"trail {t['id']}: unknown step {s}"
print('trails:', len(TRAILS))

data = {'units': UNITS, 'nodes': nodes, 'edges': edges, 'trails': TRAILS}
tpl = open(TEMPLATE, encoding='utf-8').read()
html = tpl.replace('__DATA__', json.dumps(data, ensure_ascii=False)).replace('__AGES__', '15–18')
open(ROOT / 'goals-map-15-18.html', 'w', encoding='utf-8').write(html)
print('written', ROOT / 'goals-map-15-18.html', ',', len(html), 'bytes')
