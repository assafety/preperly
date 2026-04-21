import os
import io
import base64
import json
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import anthropic

# Point Flask to the public folder for static files
app = Flask(__name__, static_folder=os.path.join(os.path.dirname(__file__), '..', 'public'), static_url_path='')
CORS(app)


def hex2rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def get_colours(style):
    s = (style or '').lower()
    if 'bright' in s or 'colour' in s:
        return dict(dk='#1A1A2E', h1='#7C4DFF', h2='#E040FB', ac='#FF6D00', tx='#1A1A2E', lt='#FFF8E1', wh='#FFFFFF')
    if 'minimal' in s or 'white' in s or 'clean' in s:
        return dict(dk='#0D1B2A', h1='#1565C0', h2='#0D47A1', ac='#00ACC1', tx='#1A1A2E', lt='#F0F8FF', wh='#FFFFFF')
    if 'pastel' in s or 'friendly' in s:
        return dict(dk='#2D1B36', h1='#D81B60', h2='#AB47BC', ac='#26C6DA', tx='#37474F', lt='#FFF0F6', wh='#FFFFFF')
    if 'projector' in s or 'contrast' in s:
        return dict(dk='#000000', h1='#000000', h2='#222222', ac='#FFD600', tx='#000000', lt='#FFFFFF', wh='#FFFFFF')
    return dict(dk='#1A1D2E', h1='#1E2761', h2='#6C63FF', ac='#22C55E', tx='#1E2030', lt='#EEF0FF', wh='#FFFFFF')


def add_rect(slide, x, y, w, h, color):
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex2rgb(color)
    shape.line.fill.background()
    return shape


def add_text(slide, text, x, y, w, h, size=16, color='#FFFFFF', bold=False, italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = str(text or '')
    run.font.size = Pt(size)
    run.font.color.rgb = hex2rgb(color)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = 'Calibri'
    return tb


def add_bullets(slide, items, x, y, w, h, size=14, color='#1E2030'):
    if not items:
        return
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = p.add_run()
        run.text = '• ' + str(item)
        run.font.size = Pt(size)
        run.font.color.rgb = hex2rgb(color)
        run.font.name = 'Calibri'
    return tb


def build_pptx(plan, slides):
    C = get_colours(plan.get('slideStyle', ''))
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    blank = prs.slide_layouts[6]

    def new_slide(bg):
        s = prs.slides.add_slide(blank)
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = hex2rgb(bg)
        return s

    objs  = (plan.get('learningObjectives') or [])[:4]
    vocab = (plan.get('keyVocabulary') or [])[:4]
    hook  = slides.get('hook') or {}
    c1    = slides.get('content1') or {}
    c2    = slides.get('content2') or {}
    act   = slides.get('activity') or {}

    # Slide 1: Title
    s = new_slide(C['dk'])
    add_rect(s, 0, 0, 0.18, 5.625, C['h2'])
    add_text(s, (plan.get('subject') or '').upper(), 0.5, 0.7, 9, 0.5, size=11, color=C['h2'], bold=True)
    add_text(s, plan.get('lessonTitle') or '', 0.5, 1.25, 8, 1.9, size=30, color='#FFFFFF', bold=True)
    add_rect(s, 0.5, 3.28, 1.6, 0.07, C['h2'])
    add_text(s, f"{plan.get('yearGroup','')}  •  {plan.get('duration','')}", 0.5, 3.5, 9, 0.5, size=15, color='#8890B0')

    # Slide 2: Objectives
    s = new_slide(C['lt'])
    add_rect(s, 0, 0, 10, 1.1, C['h1'])
    add_text(s, 'LEARNING OBJECTIVES', 0.4, 0.25, 9, 0.6, size=14, color='#FFFFFF', bold=True)
    add_text(s, 'By the end of this lesson, you will be able to:', 0.4, 1.25, 9.2, 0.45, size=13, color=C['tx'], italic=True)
    for i, obj in enumerate(objs):
        y = 1.85 + i * 0.9
        add_rect(s, 0.4, y, 9.2, 0.75, C['wh'])
        add_rect(s, 0.4, y, 0.08, 0.75, C['h2'])
        add_text(s, str(i + 1), 0.55, y + 0.15, 0.4, 0.45, size=13, color=C['h2'], bold=True)
        add_text(s, obj, 1.1, y + 0.12, 8.2, 0.55, size=13, color=C['tx'])

    # Slide 3: Starter
    s = new_slide(C['dk'])
    add_text(s, '⚡  STARTER ACTIVITY', 0.5, 0.45, 9, 0.45, size=11, color=C['ac'], bold=True)
    add_text(s, hook.get('title') or 'Think About This...', 0.5, 0.95, 9, 1.1, size=24, color='#FFFFFF', bold=True)
    add_rect(s, 0.5, 2.1, 9, 0.07, C['h2'])
    add_bullets(s, hook.get('bullets') or [], 0.5, 2.25, 9, 3.0, size=15, color='#E8EAF2')

    # Slide 4: Content 1
    s = new_slide(C['wh'])
    add_rect(s, 0, 0, 10, 1.0, C['h1'])
    add_text(s, c1.get('title') or '', 0.4, 0.15, 9.2, 0.75, size=20, color='#FFFFFF', bold=True)
    add_bullets(s, c1.get('bullets') or [], 0.4, 1.15, 9.2, 4.2, size=14, color=C['tx'])

    # Slide 5: Content 2
    s = new_slide(C['lt'])
    add_rect(s, 0, 0, 10, 1.0, C['h2'])
    add_text(s, c2.get('title') or '', 0.4, 0.15, 9.2, 0.75, size=20, color='#FFFFFF', bold=True)
    add_bullets(s, c2.get('bullets') or [], 0.4, 1.15, 9.2, 4.2, size=14, color=C['tx'])

    # Slide 6: Vocabulary
    s = new_slide(C['lt'])
    add_rect(s, 0, 0, 10, 1.0, C['h2'])
    add_text(s, 'KEY VOCABULARY', 0.4, 0.15, 9, 0.75, size=20, color='#FFFFFF', bold=True)
    for i, v in enumerate(vocab):
        col = i % 2
        row = i // 2
        x = 0.4 if col == 0 else 5.2
        y = 1.2 + row * 1.7
        add_rect(s, x, y, 4.5, 1.45, C['wh'])
        add_rect(s, x, y, 0.1, 1.45, C['h2'])
        add_text(s, v.get('word') or '', x + 0.25, y + 0.1, 4.1, 0.45, size=14, color=C['h1'], bold=True)
        add_text(s, v.get('definition') or '', x + 0.25, y + 0.6, 4.1, 0.75, size=12, color='#7A80A0', italic=True)

    # Slide 7: Activity
    s = new_slide(C['dk'])
    add_text(s, '✏️  ACTIVITY TIME', 0.5, 0.45, 9, 0.45, size=11, color=C['ac'], bold=True)
    add_text(s, act.get('title') or 'Your Task', 0.5, 0.95, 9, 1.0, size=24, color='#FFFFFF', bold=True)
    if act.get('time'):
        add_rect(s, 0.5, 2.05, 2.5, 0.5, C['ac'])
        add_text(s, '⏱ ' + act['time'], 0.5, 2.1, 2.5, 0.45, size=13, color='#FFFFFF', bold=True)
    for i, step in enumerate((act.get('steps') or [])[:4]):
        add_text(s, f'{i+1}.  {step}', 0.5, 2.7 + i * 0.72, 9, 0.65, size=14, color='#E8EAF2')

    # Slide 8: Summary
    s = new_slide(C['wh'])
    add_rect(s, 0, 0, 10, 1.0, C['h1'])
    add_text(s, 'LESSON SUMMARY', 0.4, 0.15, 9, 0.75, size=20, color='#FFFFFF', bold=True)
    add_text(s, 'What have we learned today?', 0.4, 1.15, 9, 0.45, size=13, color='#7A80A0', italic=True)
    for i, obj in enumerate(objs):
        y = 1.75 + i * 0.9
        add_rect(s, 0.4, y, 9.2, 0.75, C['lt'])
        add_rect(s, 0.4, y, 0.07, 0.75, C['ac'])
        add_text(s, obj, 0.6, y + 0.12, 8.8, 0.55, size=13, color=C['tx'])

    # Slide 9: Exit Ticket
    s = new_slide(C['dk'])
    add_text(s, '🎫  EXIT TICKET', 0.5, 0.45, 9, 0.45, size=11, color=C['ac'], bold=True)
    add_text(s, 'Before You Go...', 0.5, 0.95, 9, 0.85, size=24, color='#FFFFFF', bold=True)
    add_rect(s, 0.5, 1.95, 9, 1.6, C['h2'])
    add_text(s, slides.get('exitQ') or "Summarise today's lesson in 3 key points.", 0.65, 2.05, 8.7, 1.4, size=16, color='#FFFFFF', italic=True)
    add_text(s, f"{plan.get('subject','')}  •  {plan.get('yearGroup','')}  •  {plan.get('lessonTitle','')}", 0.5, 5.1, 9, 0.35, size=9, color='#4A5070')

    buf = io.BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode('utf-8')


# ── Routes ────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/<path:path>')
def static_files(path):
    return send_from_directory(app.static_folder, path)

@app.route('/api/generate-plan', methods=['POST'])
def generate_plan():
    try:
        data = request.json
        api_key = data.get('apiKey', '')
        if not api_key:
            return jsonify({'error': 'No API key provided'}), 400

        subject      = data.get('subject', '')
        year_group   = data.get('yearGroup', '')
        topic        = data.get('topic', '')
        duration     = data.get('duration', '60 minutes')
        requirements = data.get('requirements', 'none')

        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model='claude-sonnet-4-6',
            max_tokens=4000,
            messages=[{
                'role': 'user',
                'content': (
                    'RESPOND WITH ONLY VALID JSON. NO BACKTICKS. NO EXPLANATION. START WITH { END WITH }.\n'
                    f'Create a UK school lesson plan: Subject:{subject}, Year:{year_group}, '
                    f'Topic:{topic}, Duration:{duration}, Notes:{requirements or "none"}\n'
                    'JSON:{"lessonTitle":"...","learningObjectives":["...","...","..."],'
                    '"keyVocabulary":[{"word":"...","definition":"..."},{"word":"...","definition":"..."},{"word":"...","definition":"..."}],'
                    '"resources":["...","..."],'
                    '"lessonPhases":[{"phase":"Starter","duration":"10 mins","teacherActivity":"...","studentActivity":"...","differentiation":"..."},'
                    '{"phase":"Direct Teaching","duration":"15 mins","teacherActivity":"...","studentActivity":"...","differentiation":"..."},'
                    '{"phase":"Main Activity","duration":"20 mins","teacherActivity":"...","studentActivity":"...","differentiation":"..."},'
                    '{"phase":"Plenary","duration":"5 mins","teacherActivity":"...","studentActivity":"...","differentiation":"..."}],'
                    '"assessmentStrategies":["...","..."],"homeworkSuggestion":"...","teacherNotes":"..."}'
                )
            }]
        )
        raw = message.content[0].text
        s = raw.index('{')
        e = raw.rindex('}')
        plan = json.loads(raw[s:e+1])
        plan.update({'subject': subject, 'yearGroup': year_group, 'duration': duration})
        return jsonify({'success': True, 'plan': plan})

    except Exception as ex:
        return jsonify({'success': False, 'error': str(ex)}), 500


@app.route('/api/build-pptx', methods=['POST'])
def build_pptx_route():
    try:
        data    = request.json
        api_key = data.get('apiKey', '')
        plan    = data.get('plan', {})
        if not api_key:
            return jsonify({'error': 'No API key provided'}), 400

        client = anthropic.Anthropic(api_key=api_key)
        objs   = (plan.get('learningObjectives') or [])[:3]
        message = client.messages.create(
            model='claude-sonnet-4-6',
            max_tokens=2000,
            messages=[{
                'role': 'user',
                'content': (
                    'RESPOND WITH ONLY VALID JSON. NO BACKTICKS. NO EXPLANATION. START WITH { END WITH }.\n'
                    f'Slide text for: "{plan.get("lessonTitle","")}" ({plan.get("subject","")}, {plan.get("yearGroup","")}).\n'
                    f'Objectives: {"; ".join(objs)}\n'
                    'JSON:{"hook":{"title":"...","bullets":["...","...","..."]},'
                    '"content1":{"title":"...","bullets":["...","...","..."]},'
                    '"content2":{"title":"...","bullets":["...","...","..."]},'
                    '"activity":{"title":"Activity: ...","time":"20 mins","steps":["...","...","..."]},'
                    '"exitQ":"..."}'
                )
            }]
        )
        raw    = message.content[0].text
        s      = raw.index('{')
        e      = raw.rindex('}')
        slides = json.loads(raw[s:e+1])

        pptx_b64 = build_pptx(plan, slides)
        return jsonify({'success': True, 'pptx': pptx_b64, 'slides': slides})

    except Exception as ex:
        return jsonify({'success': False, 'error': str(ex)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
