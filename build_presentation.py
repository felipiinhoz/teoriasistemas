import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Define Color Palette matching the source PDF
COLOR_BG_DARK = RGBColor(11, 19, 43)        # #0B132B - Deep navy for cover and conclusion
COLOR_BG_LIGHT = RGBColor(248, 250, 252)    # #F8FAFC - Clean off-white background
COLOR_CARD_LIGHT = RGBColor(255, 255, 255)  # #FFFFFF - Pure white card surface
COLOR_CARD_BORDER = RGBColor(226, 232, 240) # #E2E8F0 - Soft border
COLOR_CARD_DARK = RGBColor(15, 23, 42)      # #0F172A - Deep slate card
COLOR_PRIMARY_TEAL = RGBColor(13, 148, 136) # #0D9488 - Teal accent
COLOR_MINT_ACCENT = RGBColor(16, 185, 129)  # #10B981 - Mint green
COLOR_TEXT_MAIN = RGBColor(15, 23, 42)      # #0F172A - Dark text
COLOR_TEXT_MUTED = RGBColor(100, 116, 139)  # #64748B - Muted gray text
COLOR_TEXT_WHITE = RGBColor(255, 255, 255)  # #FFFFFF
COLOR_TEXT_LIGHT_GRAY = RGBColor(203, 213, 225) # #CBD5E1
COLOR_RED_BADGE = RGBColor(254, 226, 226)   # #FEE2E2
COLOR_RED_TEXT = RGBColor(185, 28, 28)      # #B91C1C
COLOR_GREEN_BADGE = RGBColor(204, 251, 241) # #CCFBF1
COLOR_GREEN_TEXT = RGBColor(15, 118, 110)   # #0F766E
COLOR_TEAL_BADGE = RGBColor(224, 242, 254)  # #E0F2FE
COLOR_TEAL_BADGE_TEXT = RGBColor(3, 105, 161)

FONT_HEADING = "Segoe UI"
FONT_BODY = "Segoe UI"

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    def set_slide_background(slide, color):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_header(slide, category, title):
        # Category Tag
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.7), Inches(0.4))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        tf_cat.margin_left = tf_cat.margin_right = tf_cat.margin_top = tf_cat.margin_bottom = 0
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.name = FONT_HEADING
        p_cat.font.size = Pt(11)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_PRIMARY_TEAL

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.95), Inches(11.7), Inches(0.75))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        tf_title.margin_left = tf_title.margin_right = tf_title.margin_top = tf_title.margin_bottom = 0
        p_title = tf_title.paragraphs[0]
        p_title.text = title
        p_title.font.name = FONT_HEADING
        p_title.font.size = Pt(28)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_TEXT_MAIN

    def set_speaker_notes(slide, notes_text):
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.text = notes_text.strip()

    # ==========================================
    # SLIDE 1: Capa (Dark Theme)
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1, COLOR_BG_DARK)

    # Decorative accent line
    line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(2.2), Inches(1.0), Inches(0.06))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_MINT_ACCENT
    line.line.fill.background()

    # Top Tag
    tag_box = s1.shapes.add_textbox(Inches(1.2), Inches(2.5), Inches(10.5), Inches(0.4))
    tf1 = tag_box.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "ADMINISTRAÇÃO GERAL  •  ESTRUTURA ACADÊMICA"
    p1.font.name = FONT_HEADING
    p1.font.size = Pt(13)
    p1.font.bold = True
    p1.font.color.rgb = COLOR_MINT_ACCENT

    # Main Title
    title_box = s1.shapes.add_textbox(Inches(1.2), Inches(3.0), Inches(10.5), Inches(1.5))
    tf2 = title_box.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "Teoria Sistêmica na\nAdministração"
    p2.font.name = FONT_HEADING
    p2.font.size = Pt(46)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_TEXT_WHITE

    # Subtitle
    sub_box = s1.shapes.add_textbox(Inches(1.2), Inches(4.7), Inches(9.5), Inches(1.2))
    tf3 = sub_box.text_frame
    tf3.word_wrap = True
    p3 = tf3.paragraphs[0]
    p3.text = "Da visão fragmentada ao pensamento integrado: compreendendo as organizações como sistemas vivos e abertos."
    p3.font.name = FONT_BODY
    p3.font.size = Pt(18)
    p3.font.color.rgb = COLOR_TEXT_LIGHT_GRAY

    # Bottom bar accent
    bot_line = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(6.1), Inches(1.5), Inches(0.04))
    bot_line.fill.solid()
    bot_line.fill.fore_color.rgb = COLOR_PRIMARY_TEAL
    bot_line.line.fill.background()

    set_speaker_notes(s1, """[ROTEIRO DE FALA - SLIDE 1: CAPA]
- Bom dia/boa tarde a todos. Hoje vamos explorar um dos marcos mais revolucionários da gestão moderna: a Teoria Sistêmica na Administração.
- Durante décadas, a administração tradicional tratou as empresas como máquinas e engrenagens isoladas. Cada departamento cuidava apenas do seu quadrado.
- A Teoria Sistêmica rompe com esse reducionismo: ela nos ensina a olhar para as empresas não como fábricas mecânicas fechadas, mas como organismos vivos, dinâmicos e abertos em constante troca com a sociedade, clientes e mercado.
- Ao longo desta apresentação, veremos os fundamentos, os principais teóricos, as propriedades operacionais e, principalmente, como essa visão se converte em vantagem competitiva real no processo de tomada de decisão.
Tempo estimado: 1 a 2 minutos.""")

    # ==========================================
    # SLIDE 2: Integrantes da Equipe
    # ==========================================
    s_team = prs.slides.add_slide(blank_layout)
    set_slide_background(s_team, COLOR_BG_LIGHT)
    add_header(s_team, "Administração Geral  •  Estrutura Acadêmica", "Integrantes da Equipe")

    team_members = [
        ("Cauê Oliveira Viana", "202310921", "CV"),
        ("Claudio Ávila Rosa Filho", "202310661", "CA"),
        ("Enzo Cardoso Ribeiro", "202211355", "EC"),
        ("Gustavo Moreira dos Santos", "202310663", "GS"),
        ("Iago Campos de Melo", "202311585", "IM"),
        ("João Felipe Santos Matos", "202310926", "JF"),
        ("João Gabryell Lopes Santana", "202310927", "JG"),
        ("Kaike Santos dos Santos", "202311309", "KS"),
        ("Tharcizio Rubens Santos Mota", "202211373", "TR")
    ]

    card_m_w = Inches(3.68)
    card_m_h = Inches(1.35)
    gap_m_x = Inches(0.35)
    gap_m_y = Inches(0.25)
    start_m_x = Inches(0.8)
    start_m_y = Inches(1.85)

    for idx, (m_name, m_mat, initials) in enumerate(team_members):
        row = idx // 3
        col = idx % 3
        mx = start_m_x + col * (card_m_w + gap_m_x)
        my = start_m_y + row * (card_m_h + gap_m_y)

        # Card container
        m_card = s_team.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx, my, card_m_w, card_m_h)
        m_card.fill.solid()
        m_card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        m_card.line.color.rgb = COLOR_CARD_BORDER
        m_card.line.width = Pt(1.5)

        # Initials circle avatar
        avatar = s_team.shapes.add_shape(MSO_SHAPE.OVAL, mx + Inches(0.25), my + Inches(0.3), Inches(0.75), Inches(0.75))
        avatar.fill.solid()
        avatar.fill.fore_color.rgb = COLOR_GREEN_BADGE
        avatar.line.fill.background()
        tf_av = avatar.text_frame
        p_av = tf_av.paragraphs[0]
        p_av.text = initials
        p_av.font.name = FONT_HEADING
        p_av.font.size = Pt(13)
        p_av.font.bold = True
        p_av.font.color.rgb = COLOR_GREEN_TEXT
        p_av.alignment = PP_ALIGN.CENTER

        # Member Name
        nb = s_team.shapes.add_textbox(mx + Inches(1.15), my + Inches(0.2), card_m_w - Inches(1.3), Inches(0.45))
        tf_nb = nb.text_frame
        tf_nb.word_wrap = True
        tf_nb.margin_left = tf_nb.margin_right = tf_nb.margin_top = tf_nb.margin_bottom = 0
        p_nb = tf_nb.paragraphs[0]
        p_nb.text = m_name
        p_nb.font.name = FONT_HEADING
        p_nb.font.size = Pt(14)
        p_nb.font.bold = True
        p_nb.font.color.rgb = COLOR_TEXT_MAIN

        # Member Registration (Matrícula)
        mb = s_team.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, mx + Inches(1.15), my + Inches(0.72), Inches(2.2), Inches(0.32))
        mb.fill.solid()
        mb.fill.fore_color.rgb = COLOR_TEAL_BADGE
        mb.line.fill.background()
        p_mb = mb.text_frame.paragraphs[0]
        p_mb.text = f"Matrícula: {m_mat}"
        p_mb.font.name = FONT_BODY
        p_mb.font.size = Pt(10)
        p_mb.font.bold = True
        p_mb.font.color.rgb = COLOR_TEAL_BADGE_TEXT
        p_mb.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s_team, """[ROTEIRO DE FALA - SLIDE 2: INTEGRANTES DA EQUIPE]
- Apresentação dos integrantes do grupo de trabalho:
  Cauê Oliveira Viana, Claudio Ávila Rosa Filho, Enzo Cardoso Ribeiro, Gustavo Moreira dos Santos, Iago Campos de Melo, João Felipe Santos Matos, João Gabryell Lopes Santana, Kaike Santos dos Santos e Tharcizio Rubens Santos Mota.
- Agradecemos ao professor e aos colegas pela presença e daremos início à exposição teórica.
Tempo estimado: 30 segundos a 1 minuto.""")

    # ==========================================
    # SLIDE 3: Objetivos de Aprendizagem
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2, COLOR_BG_LIGHT)
    add_header(s2, "Objetivos de Aprendizagem", "Por que Estudar a Teoria Sistêmica?")

    cards_s2 = [
        ("Visão Holística", "Compreensão total da organização sem isolamento de setores.", "Ver o todo integrado ao invés de fragmentos desconexos. Elimina a miopia departamental.", "👁️"),
        ("Interdependência", "Avaliação precisa dos impactos de decisões intersetoriais.", "Cada ação em vendas reverbera em produção, logística, RH e fluxo de caixa.", "🔗"),
        ("Adaptabilidade", "Resposta rápida e contínua às transformações do mercado.", "Capacidade dinâmica de sobreviver e prosperar em ambientes de alta incerteza.", "🔄")
    ]

    card_w = Inches(3.7)
    card_h = Inches(4.6)
    spacing = Inches(0.3)
    start_x = Inches(0.8)
    top_y = Inches(1.9)

    for i, (head, desc, detail, icon) in enumerate(cards_s2):
        cx = start_x + i * (card_w + spacing)
        # Card Background
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, top_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1.5)

        # Icon circle
        ic = s2.shapes.add_shape(MSO_SHAPE.OVAL, cx + Inches(0.4), top_y + Inches(0.4), Inches(0.7), Inches(0.7))
        ic.fill.solid()
        ic.fill.fore_color.rgb = COLOR_GREEN_BADGE
        ic.line.fill.background()
        tf_ic = ic.text_frame
        p_ic = tf_ic.paragraphs[0]
        p_ic.text = icon
        p_ic.font.size = Pt(16)
        p_ic.alignment = PP_ALIGN.CENTER

        # Title
        t_box = s2.shapes.add_textbox(cx + Inches(0.4), top_y + Inches(1.3), card_w - Inches(0.8), Inches(0.6))
        tf_t = t_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = head
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(20)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_MAIN

        # Description
        d_box = s2.shapes.add_textbox(cx + Inches(0.4), top_y + Inches(2.0), card_w - Inches(0.8), Inches(1.2))
        tf_d = d_box.text_frame
        tf_d.word_wrap = True
        p_d = tf_d.paragraphs[0]
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(14)
        p_d.font.bold = True
        p_d.font.color.rgb = COLOR_PRIMARY_TEAL

        # Detail text
        det_box = s2.shapes.add_textbox(cx + Inches(0.4), top_y + Inches(3.2), card_w - Inches(0.8), Inches(1.1))
        tf_det = det_box.text_frame
        tf_det.word_wrap = True
        p_det = tf_det.paragraphs[0]
        p_det.text = detail
        p_det.font.name = FONT_BODY
        p_det.font.size = Pt(13)
        p_det.font.color.rgb = COLOR_TEXT_MUTED

    set_speaker_notes(s2, """[ROTEIRO DE FALA - SLIDE 2: POR QUE ESTUDAR A TEORIA SISTÊMICA?]
- Três grandes competências justificam o domínio deste tema:
1. Visão Holística: 'Holos' vem do grego e significa 'todo'. Um gestor sistêmico não olha apenas para a sua mesa ou para a sua meta individual; ele enxerga a cadeia de ponta a ponta.
2. Interdependência: Em uma organização, nada acontece no vácuo. Se o marketing cria uma promoção agressiva sem alinhar com operações e estoque, a empresa gera frustração no cliente e gargalos logísticos. Decisões eficazes mapeiam esses elos.
3. Adaptabilidade: Vivemos na era da volatilidade. Empresas que se fecham em velhos manuais perecem; sistemas que escutam o ambiente e aprendem continuam evoluindo.
Tempo estimado: 1 minuto e meio.""")

    # ==========================================
    # SLIDE 3: Abordagem Sistêmica vs Tradicional
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3, COLOR_BG_LIGHT)
    add_header(s3, "Tópico 1  •  Relação Teórica", "Abordagem Sistêmica na Administração Geral")

    col_w = Inches(5.6)
    col_h = Inches(4.8)
    col1_x = Inches(0.8)
    col2_x = Inches(6.9)
    col_y = Inches(1.8)

    # Col 1: Tradicional
    c1 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col1_x, col_y, col_w, col_h)
    c1.fill.solid()
    c1.fill.fore_color.rgb = COLOR_CARD_LIGHT
    c1.line.color.rgb = COLOR_CARD_BORDER
    c1.line.width = Pt(1.5)

    badge1 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col1_x + Inches(0.5), col_y + Inches(0.4), Inches(2.5), Inches(0.4))
    badge1.fill.solid()
    badge1.fill.fore_color.rgb = COLOR_RED_BADGE
    badge1.line.fill.background()
    p_b1 = badge1.text_frame.paragraphs[0]
    p_b1.text = "Abordagem Tradicional"
    p_b1.font.bold = True
    p_b1.font.size = Pt(11)
    p_b1.font.color.rgb = COLOR_RED_TEXT
    p_b1.alignment = PP_ALIGN.CENTER

    t1_box = s3.shapes.add_textbox(col1_x + Inches(0.5), col_y + Inches(1.0), col_w - Inches(1.0), Inches(0.6))
    p_t1 = t1_box.text_frame.paragraphs[0]
    p_t1.text = "Visão Mecanicista"
    p_t1.font.name = FONT_HEADING
    p_t1.font.size = Pt(22)
    p_t1.font.bold = True
    p_t1.font.color.rgb = COLOR_TEXT_MAIN

    bullets1 = [
        ("✗", "Foco exclusivo em tarefas internas", "Atenção estrita aos postos de trabalho e eficiência de rotina isolada."),
        ("✗", "Departamentos isolados (Silos)", "Feudos corporativos onde as áreas não dialogam entre si."),
        ("✗", "Organização vista como sistema fechado", "Crença de que o ambiente externo é previsível ou ignorável.")
    ]
    for idx, (sym, main_txt, sub_txt) in enumerate(bullets1):
        item_box = s3.shapes.add_textbox(col1_x + Inches(0.5), col_y + Inches(1.8 + idx * 0.95), col_w - Inches(1.0), Inches(0.85))
        tf_i = item_box.text_frame
        tf_i.word_wrap = True
        tf_i.margin_left = tf_i.margin_right = tf_i.margin_top = tf_i.margin_bottom = 0
        p_main = tf_i.paragraphs[0]
        p_main.text = f"{sym}  {main_txt}"
        p_main.font.bold = True
        p_main.font.size = Pt(14)
        p_main.font.color.rgb = COLOR_RED_TEXT
        p_sub = tf_i.add_paragraph()
        p_sub.text = f"     {sub_txt}"
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = COLOR_TEXT_MUTED

    # Col 2: Sistêmica
    c2 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col2_x, col_y, col_w, col_h)
    c2.fill.solid()
    c2.fill.fore_color.rgb = COLOR_CARD_LIGHT
    c2.line.color.rgb = COLOR_PRIMARY_TEAL
    c2.line.width = Pt(2.5)

    badge2 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, col2_x + Inches(0.5), col_y + Inches(0.4), Inches(2.5), Inches(0.4))
    badge2.fill.solid()
    badge2.fill.fore_color.rgb = COLOR_GREEN_BADGE
    badge2.line.fill.background()
    p_b2 = badge2.text_frame.paragraphs[0]
    p_b2.text = "Abordagem Sistêmica"
    p_b2.font.bold = True
    p_b2.font.size = Pt(11)
    p_b2.font.color.rgb = COLOR_GREEN_TEXT
    p_b2.alignment = PP_ALIGN.CENTER

    t2_box = s3.shapes.add_textbox(col2_x + Inches(0.5), col_y + Inches(1.0), col_w - Inches(1.0), Inches(0.6))
    p_t2 = t2_box.text_frame.paragraphs[0]
    p_t2.text = "Visão Integrada"
    p_t2.font.name = FONT_HEADING
    p_t2.font.size = Pt(22)
    p_t2.font.bold = True
    p_t2.font.color.rgb = COLOR_TEXT_MAIN

    bullets2 = [
        ("✓", "Integração total dos subsistemas", "Processos transversais fluidos com cooperação intersetorial ativa."),
        ("✓", "Interação constante com o ambiente", "Monitoramento diário de concorrentes, clientes, leis e tecnologia."),
        ("✓", "Foco em objetivos globais", "O sucesso da empresa como um todo supera os ganhos isolados de uma área.")
    ]
    for idx, (sym, main_txt, sub_txt) in enumerate(bullets2):
        item_box = s3.shapes.add_textbox(col2_x + Inches(0.5), col_y + Inches(1.8 + idx * 0.95), col_w - Inches(1.0), Inches(0.85))
        tf_i = item_box.text_frame
        tf_i.word_wrap = True
        tf_i.margin_left = tf_i.margin_right = tf_i.margin_top = tf_i.margin_bottom = 0
        p_main = tf_i.paragraphs[0]
        p_main.text = f"{sym}  {main_txt}"
        p_main.font.bold = True
        p_main.font.size = Pt(14)
        p_main.font.color.rgb = COLOR_GREEN_TEXT
        p_sub = tf_i.add_paragraph()
        p_sub.text = f"     {sub_txt}"
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = COLOR_TEXT_MUTED

    set_speaker_notes(s3, """[ROTEIRO DE FALA - SLIDE 3: VISÃO MECANICISTA VS. VISÃO INTEGRADA]
- Façamos este contraste fundamental:
- À esquerda, temos a Administração Científica e Clássica (Taylor e Fayol). Para eles, a empresa era como uma máquina de relógio: peças intercambiáveis, regras rígidas e foco exclusivo para dentro da fábrica (sistema fechado).
- À direita, a Abordagem Sistêmica: percebeu-se que uma empresa não sobrevive isolada. Ela consome recursos do ambiente externo e precisa entregar valor para esse mesmo ambiente.
- Mais importante: na visão mecanicista, se cada setor bater sua meta, teoricamente tudo dá certo. Na visão sistêmica, descobrimos que metas locais conflitantes destroem o objetivo global. A integração é soberana!
Tempo estimado: 2 minutos.""")

    # ==========================================
    # SLIDE 4: Fluxo Fundamental de um Sistema
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4, COLOR_BG_LIGHT)
    add_header(s4, "Tópico 2  •  Aspectos Básicos", "O Fluxo Fundamental de um Sistema")

    steps = [
        ("Entradas (Inputs)", "Insumos, dados e matérias-primas", "Recursos humanos, financeiros, informação e tecnologia recebidos do ambiente.", False, "📥"),
        ("Processamento", "Transformação operacional interna", "Conversão produtiva, metodologias, linhas de montagem e tomada de decisão.", False, "⚙️"),
        ("Saídas (Outputs)", "Produtos, serviços e resultados", "Entregas tangíveis, valor gerado, rentabilidade e impactos no mercado.", False, "📤"),
        ("Feedback", "Retroalimentação e ajuste", "Avaliações, métricas, satisfação do cliente e correções de rumo em tempo real.", True, "🔄")
    ]

    sw = Inches(2.65)
    sh = Inches(4.5)
    gap = Inches(0.4)
    start_sx = Inches(0.8)
    sy = Inches(1.9)

    for i, (name, role, details, is_dark, icon) in enumerate(steps):
        sx = start_sx + i * (sw + gap)
        scard = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, sx, sy, sw, sh)
        scard.fill.solid()
        if is_dark:
            scard.fill.fore_color.rgb = COLOR_CARD_DARK
            scard.line.fill.background()
        else:
            scard.fill.fore_color.rgb = COLOR_CARD_LIGHT
            scard.line.color.rgb = COLOR_CARD_BORDER
            scard.line.width = Pt(1.5)

        # Icon / number
        ic_shape = s4.shapes.add_shape(MSO_SHAPE.OVAL, sx + Inches(0.35), sy + Inches(0.4), Inches(0.65), Inches(0.65))
        ic_shape.fill.solid()
        ic_shape.fill.fore_color.rgb = COLOR_PRIMARY_TEAL if is_dark else COLOR_GREEN_BADGE
        ic_shape.line.fill.background()
        p_ic = ic_shape.text_frame.paragraphs[0]
        p_ic.text = icon
        p_ic.font.size = Pt(15)
        p_ic.alignment = PP_ALIGN.CENTER

        # Title
        t_box = s4.shapes.add_textbox(sx + Inches(0.3), sy + Inches(1.3), sw - Inches(0.6), Inches(0.6))
        p_t = t_box.text_frame.paragraphs[0]
        p_t.text = name
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(17)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_WHITE if is_dark else COLOR_TEXT_MAIN

        # Role
        r_box = s4.shapes.add_textbox(sx + Inches(0.3), sy + Inches(1.95), sw - Inches(0.6), Inches(1.0))
        r_tf = r_box.text_frame
        r_tf.word_wrap = True
        p_r = r_tf.paragraphs[0]
        p_r.text = role
        p_r.font.name = FONT_BODY
        p_r.font.size = Pt(13)
        p_r.font.bold = True
        p_r.font.color.rgb = COLOR_MINT_ACCENT if is_dark else COLOR_PRIMARY_TEAL

        # Details
        det_box = s4.shapes.add_textbox(sx + Inches(0.3), sy + Inches(3.0), sw - Inches(0.6), Inches(1.2))
        det_tf = det_box.text_frame
        det_tf.word_wrap = True
        p_det = det_tf.paragraphs[0]
        p_det.text = details
        p_det.font.name = FONT_BODY
        p_det.font.size = Pt(12)
        p_det.font.color.rgb = COLOR_TEXT_LIGHT_GRAY if is_dark else COLOR_TEXT_MUTED

        # Arrow indicator between cards
        if i < 3:
            arrow = s4.shapes.add_textbox(sx + sw - Inches(0.05), sy + Inches(2.0), gap + Inches(0.1), Inches(0.5))
            p_arr = arrow.text_frame.paragraphs[0]
            p_arr.text = "➔"
            p_arr.font.size = Pt(20)
            p_arr.font.bold = True
            p_arr.font.color.rgb = COLOR_PRIMARY_TEAL
            p_arr.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s4, """[ROTEIRO DE FALA - SLIDE 4: FLUXO FUNDAMENTAL DE UM SISTEMA]
- Todo sistema dinâmico opera através deste circuito vital contínuo:
1. Entradas (Inputs): Insumos, matérias-primas, talentos, capitais e, cada vez mais, dados e informações.
2. Processamento (Throughput): A capacidade da organização de transformar esses recursos com inteligência, agilidade e eficiência.
3. Saídas (Outputs): O que devolvemos ao ecossistema: produtos de qualidade, serviços de valor, lucros e reputação.
4. Feedback (Retroalimentação - destacado à direita): Este é o coração da sobrevivência. O feedback nos diz se as saídas atenderam às expectativas. Sem feedback, um sistema fica cego e caminha para a falência.
Tempo estimado: 1 minuto e meio.""")

    # ==========================================
    # SLIDE 5: Hierarquização de Sistemas
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5, COLOR_BG_LIGHT)
    add_header(s5, "Tópico 3  •  Estrutura de Níveis", "Hierarquização de Sistemas")

    levels = [
        ("Supersistema (Suprasistema)", "Nível Macro",
         "Ambiente externo abrangente: mercado consumidor, concorrência, marcos regulatórios, avanços tecnológicos e contexto socioeconômico global.",
         COLOR_TEXT_MAIN, COLOR_CARD_LIGHT),
        ("Sistema", "Nível Mapeado",
         "A organização ou empresa considerada em sua totalidade focal: cultura corporativa, governança, identidade de marca e estratégia integrada.",
         COLOR_PRIMARY_TEAL, COLOR_CARD_LIGHT),
        ("Subsistemas", "Nível Micro",
         "Departamentos, divisões funcionais, setores operacionais e equipes multidisciplinares internas responsáveis por etapas específicas do trabalho.",
         COLOR_TEXT_MUTED, COLOR_CARD_LIGHT)
    ]

    bar_w = Inches(11.7)
    bar_h = Inches(1.4)
    start_by = Inches(1.9)
    gap_by = Inches(0.25)

    for idx, (title, badge, desc, accent_col, bg_col) in enumerate(levels):
        by = start_by + idx * (bar_h + gap_by)
        bar = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), by, bar_w, bar_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = bg_col
        bar.line.color.rgb = COLOR_CARD_BORDER
        bar.line.width = Pt(1.5)

        # Left vertical accent strip
        strip = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), by, Inches(0.12), bar_h)
        strip.fill.solid()
        strip.fill.fore_color.rgb = accent_col
        strip.line.fill.background()

        # Title
        t_box = s5.shapes.add_textbox(Inches(1.2), by + Inches(0.2), Inches(7.5), Inches(0.4))
        p_t = t_box.text_frame.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(20)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_MAIN

        # Description
        d_box = s5.shapes.add_textbox(Inches(1.2), by + Inches(0.65), Inches(8.5), Inches(0.65))
        d_tf = d_box.text_frame
        d_tf.word_wrap = True
        d_tf.margin_top = d_tf.margin_bottom = 0
        p_d = d_tf.paragraphs[0]
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(13)
        p_d.font.color.rgb = COLOR_TEXT_MUTED

        # Badge pill at right
        pill = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.5), by + Inches(0.45), Inches(1.7), Inches(0.45))
        pill.fill.solid()
        pill.fill.fore_color.rgb = COLOR_GREEN_BADGE if idx == 1 else COLOR_TEAL_BADGE
        pill.line.fill.background()
        p_pill = pill.text_frame.paragraphs[0]
        p_pill.text = badge
        p_pill.font.bold = True
        p_pill.font.size = Pt(12)
        p_pill.font.color.rgb = COLOR_GREEN_TEXT if idx == 1 else COLOR_TEAL_BADGE_TEXT
        p_pill.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s5, """[ROTEIRO DE FALA - SLIDE 5: HIERARQUIZAÇÃO DE SISTEMAS]
- O pensamento sistêmico nos ensina que a fronteira de um sistema depende do ponto de vista do observador (efeito matriosca ou bonecas russas):
- No Nível Macro (Supersistema): Temos o ecossistema econômico, as leis, a concorrência global e a cultura social. Nenhuma empresa controla o supersistema sozinha, mas todas são impactadas por ele.
- No Nível Mapeado (Sistema): É a empresa como um todo. É o foco da estratégia executiva.
- No Nível Micro (Subsistemas): São o marketing, finanças, compras, RH, TI. Para o diretor de TI, o departamento de TI é o seu sistema; mas para a presidência, ele é apenas um subsistema a serviço do todo.
- O segredo da governança sistêmica é garantir que o subsistema não tente se comportar como se fosse maior que o sistema focal.
Tempo estimado: 1 minuto e meio.""")

    # ==========================================
    # SLIDE 6: Princípios Práticos
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6, COLOR_BG_LIGHT)
    add_header(s6, "Tópico 4  •  Princípios Práticos", "Aplicação dos Conceitos Sistêmicos")

    concepts = [
        ("Sinergia", "O todo é maior que a soma das partes", "Trabalho conjunto cooperativo com resultado substancialmente superior à mera soma individual dos esforços (1 + 1 = 3).", COLOR_PRIMARY_TEAL, "⚡"),
        ("Homeostase", "Equilíbrio dinâmico e autorregulação", "Capacidade do sistema de resistir a perturbações externas e restaurar a estabilidade operacional através de autocorreção contínua.", COLOR_MINT_ACCENT, "⚖️"),
        ("Entropia Negativa", "Importação ativa de energia vital", "Tendência natural de qualquer sistema fechado é o caos e a morte (entropia). Sistemas abertos importam energia para reverter o desgaste.", COLOR_PRIMARY_TEAL, "🔋"),
        ("Equifinalidade", "Múltiplos caminhos, mesmo resultado", "Capacidade de alcançar o mesmo objetivo final a partir de diferentes condições iniciais e por variadas rotas estratégicas.", COLOR_MINT_ACCENT, "🎯")
    ]

    gw = Inches(5.6)
    gh = Inches(2.25)
    g_start_y = Inches(1.9)
    g_gap_y = Inches(0.3)
    g_gap_x = Inches(0.5)

    for idx, (title, tag, desc, color, icon) in enumerate(concepts):
        row = idx // 2
        col = idx % 2
        gx = Inches(0.8) + col * (gw + g_gap_x)
        gy = g_start_y + row * (gh + g_gap_y)

        gcard = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, gx, gy, gw, gh)
        gcard.fill.solid()
        gcard.fill.fore_color.rgb = COLOR_CARD_LIGHT
        gcard.line.color.rgb = COLOR_CARD_BORDER
        gcard.line.width = Pt(1.5)

        # Icon pill
        ic_b = s6.shapes.add_shape(MSO_SHAPE.OVAL, gx + Inches(0.35), gy + Inches(0.35), Inches(0.55), Inches(0.55))
        ic_b.fill.solid()
        ic_b.fill.fore_color.rgb = COLOR_GREEN_BADGE
        ic_b.line.fill.background()
        p_ic = ic_b.text_frame.paragraphs[0]
        p_ic.text = icon
        p_ic.font.size = Pt(14)
        p_ic.alignment = PP_ALIGN.CENTER

        # Title
        t_box = s6.shapes.add_textbox(gx + Inches(1.1), gy + Inches(0.3), gw - Inches(1.3), Inches(0.4))
        p_t = t_box.text_frame.paragraphs[0]
        p_t.text = title
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(18)
        p_t.font.bold = True
        p_t.font.color.rgb = color

        # Tag
        tag_box = s6.shapes.add_textbox(gx + Inches(1.1), gy + Inches(0.68), gw - Inches(1.3), Inches(0.35))
        p_tg = tag_box.text_frame.paragraphs[0]
        p_tg.text = tag
        p_tg.font.name = FONT_BODY
        p_tg.font.size = Pt(11)
        p_tg.font.bold = True
        p_tg.font.color.rgb = COLOR_TEXT_MUTED

        # Description
        d_box = s6.shapes.add_textbox(gx + Inches(0.35), gy + Inches(1.15), gw - Inches(0.7), Inches(0.95))
        d_tf = d_box.text_frame
        d_tf.word_wrap = True
        d_tf.margin_left = d_tf.margin_right = d_tf.margin_top = d_tf.margin_bottom = 0
        p_d = d_tf.paragraphs[0]
        p_d.text = desc
        p_d.font.name = FONT_BODY
        p_d.font.size = Pt(12)
        p_d.font.color.rgb = COLOR_TEXT_MAIN

    set_speaker_notes(s6, """[ROTEIRO DE FALA - SLIDE 6: CONCEITOS SISTÊMICOS FUNDAMENTAIS]
- Estes quatro termos são clássicos em provas, bancas e na literatura gerencial:
1. Sinergia: Quando dois times trabalham juntos de forma alinhada, o valor criado supera muito a soma dos dois trabalhando isolados.
2. Homeostase: Termo emprestado da biologia. Assim como o corpo humano regula a temperatura a 36,5°C faça frio ou calor, a empresa usa estoques de segurança, caixas de reserva e ajustes de preços para manter seu equilíbrio financeiro e produtivo.
3. Entropia Negativa (ou Negentropia): Pela física, tudo tende à desordem e ao colapso se for abandonado. As empresas combatem isso importando nova energia: contratando novos talentos, adotando novas tecnologias e oxigenando seus processos.
4. Equifinalidade: Não existe uma única 'receita de bolo' para lucrar ou vencer no mercado. Uma empresa pode crescer via inovação tecnológica de ponta, enquanto outra alcança o mesmo sucesso via excelência em atendimento ao cliente ou baixo custo.
Tempo estimado: 2 minutos.""")

    # ==========================================
    # SLIDE 7: Organizações como Sistemas Abertos
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7, COLOR_BG_LIGHT)
    add_header(s7, "Tópico 5  •  Propriedades Abertas", "Organizações como Sistemas Abertos")

    props = [
        ("Fronteiras Permeáveis", "Troca contínua de dados, clientes e recursos com o ecossistema externo.", "A empresa não tem muros impenetráveis: absorve demandas da sociedade e responde com inovação e valor."),
        ("Dinamismo Constante", "Reorganização flexível para sustentação do equilíbrio adaptativo.", "Frente a novas regulamentações ou concorrentes agressivos, o arranjo interno se adapta sem colapsar."),
        ("Ciclos de Eventos", "Reciclagem e reinvestimento contínuo para garantir longevidade.", "A receita gerada nas vendas é reinvestida em insumos, pesquisa e salários, retroalimentando o ciclo.")
    ]

    p_w = Inches(5.8)
    for idx, (head, subh, dtl) in enumerate(props):
        py = Inches(1.9) + idx * Inches(1.6)
        pcard = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), py, p_w, Inches(1.4))
        pcard.fill.solid()
        pcard.fill.fore_color.rgb = COLOR_CARD_LIGHT
        pcard.line.color.rgb = COLOR_CARD_BORDER
        pcard.line.width = Pt(1.5)

        # Dot badge
        dot = s7.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.1), py + Inches(0.25), Inches(0.4), Inches(0.4))
        dot.fill.solid()
        dot.fill.fore_color.rgb = COLOR_PRIMARY_TEAL
        dot.line.fill.background()

        t_box = s7.shapes.add_textbox(Inches(1.7), py + Inches(0.18), p_w - Inches(1.9), Inches(0.35))
        p_t = t_box.text_frame.paragraphs[0]
        p_t.text = head
        p_t.font.name = FONT_HEADING
        p_t.font.size = Pt(16)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_MAIN

        d_box = s7.shapes.add_textbox(Inches(1.7), py + Inches(0.55), p_w - Inches(2.0), Inches(0.75))
        d_tf = d_box.text_frame
        d_tf.word_wrap = True
        d_tf.margin_top = d_tf.margin_bottom = 0
        p_sub = d_tf.paragraphs[0]
        p_sub.text = subh
        p_sub.font.bold = True
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = COLOR_PRIMARY_TEAL
        p_dtl = d_tf.add_paragraph()
        p_dtl.text = dtl
        p_dtl.font.size = Pt(11)
        p_dtl.font.color.rgb = COLOR_TEXT_MUTED

    # Right side: Network Image
    if os.path.exists("network.png"):
        net_card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(1.9), Inches(5.5), Inches(4.8))
        net_card.fill.solid()
        net_card.fill.fore_color.rgb = COLOR_CARD_LIGHT
        net_card.line.color.rgb = COLOR_CARD_BORDER
        net_card.line.width = Pt(1.5)

        s7.shapes.add_picture("network.png", Inches(7.2), Inches(2.1), width=Inches(5.1))
        
        cap_box = s7.shapes.add_textbox(Inches(7.2), Inches(6.1), Inches(5.1), Inches(0.5))
        p_cap = cap_box.text_frame.paragraphs[0]
        p_cap.text = "Rede de Interconexões: Nós interdependentes e fluxo contínuo de estímulos."
        p_cap.font.size = Pt(10)
        p_cap.font.italic = True
        p_cap.font.color.rgb = COLOR_TEXT_MUTED
        p_cap.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s7, """[ROTEIRO DE FALA - SLIDE 7: ORGANIZAÇÕES COMO SISTEMAS ABERTOS]
- O que torna uma empresa um sistema aberto?
1. Fronteiras Permeáveis: Uma empresa não pode ser uma caixa-preta fechada. Ela compra do fornecedor, atende ao cliente, negocia com bancos e cumpre normas ambientais.
2. Dinamismo Constante: As condições nunca são estáticas. O mercado muda, os juros oscilam, novas IAs surgem. Manter o equilíbrio exige movimento contínuo (equilíbrio dinâmico).
3. Ciclos de Eventos: As trocas têm natureza cíclica. O produto sai -> o dinheiro entra -> novos insumos são comprados -> novos salários são pagos. Se o ciclo se rompe em qualquer nó dessa rede (como mostra a figura ao lado), o sistema inteiro entra em pane.
Tempo estimado: 1 minuto e meio.""")

    # ==========================================
    # SLIDE 8: Principais Teóricos
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8, COLOR_BG_LIGHT)
    add_header(s8, "Tópico 6  •  Fundamentação Teórica", "Principais Teóricos do Pensamento Sistêmico")

    theorists = [
        ("1", "Ludwig von Bertalanffy", "Criador da TGS", "Biólogo austríaco que concebeu a Teoria Geral dos Sistemas (década de 1950). Propôs que princípios universais regem sistemas em todos os ramos da ciência, da biologia à administração.", "Pioneirismo Conceitual"),
        ("2", "C. West Churchman", "Pioneiro da Gestão Sistêmica", "Filósofo e pesquisador operacional norte-americano. Foi quem primeiro aplicou com rigor o pensamento sistêmico no planejamento estratégico corporativo e na tomada de decisão gerencial.", "Aplicação Gerencial"),
        ("3", "Chester Barnard", "Sistemas Cooperativos", "Executivo e teórico que conceituou a organização como um 'sistema cooperativo aberto'. Demonstrou que a sobrevivência da firma depende do equilíbrio entre contribuições e incentivos das pessoas.", "Dimensão Humana & Aberta")
    ]

    tw = Inches(3.7)
    th = Inches(4.7)
    t_start_x = Inches(0.8)
    t_gap = Inches(0.3)
    t_y = Inches(1.9)

    for i, (num, name, role, bio, highlight) in enumerate(theorists):
        tx = t_start_x + i * (tw + t_gap)
        tcard = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx, t_y, tw, th)
        tcard.fill.solid()
        tcard.fill.fore_color.rgb = COLOR_CARD_LIGHT
        tcard.line.color.rgb = COLOR_CARD_BORDER
        tcard.line.width = Pt(1.5)

        # Number pill / badge
        nbadge = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx + Inches(0.4), t_y + Inches(0.4), Inches(0.65), Inches(0.65))
        nbadge.fill.solid()
        nbadge.fill.fore_color.rgb = COLOR_PRIMARY_TEAL if i == 1 else COLOR_CARD_DARK
        nbadge.line.fill.background()
        p_n = nbadge.text_frame.paragraphs[0]
        p_n.text = num
        p_n.font.bold = True
        p_n.font.size = Pt(16)
        p_n.font.color.rgb = COLOR_TEXT_WHITE
        p_n.alignment = PP_ALIGN.CENTER

        # Name
        n_box = s8.shapes.add_textbox(tx + Inches(0.4), t_y + Inches(1.25), tw - Inches(0.8), Inches(0.65))
        n_tf = n_box.text_frame
        n_tf.word_wrap = True
        p_nm = n_tf.paragraphs[0]
        p_nm.text = name
        p_nm.font.name = FONT_HEADING
        p_nm.font.size = Pt(18)
        p_nm.font.bold = True
        p_nm.font.color.rgb = COLOR_TEXT_MAIN

        # Role badge text
        r_box = s8.shapes.add_textbox(tx + Inches(0.4), t_y + Inches(1.9), tw - Inches(0.8), Inches(0.4))
        p_r = r_box.text_frame.paragraphs[0]
        p_r.text = role
        p_r.font.bold = True
        p_r.font.size = Pt(12)
        p_r.font.color.rgb = COLOR_PRIMARY_TEAL

        # Bio
        b_box = s8.shapes.add_textbox(tx + Inches(0.4), t_y + Inches(2.35), tw - Inches(0.8), Inches(1.6))
        b_tf = b_box.text_frame
        b_tf.word_wrap = True
        p_b = b_tf.paragraphs[0]
        p_b.text = bio
        p_b.font.name = FONT_BODY
        p_b.font.size = Pt(12)
        p_b.font.color.rgb = COLOR_TEXT_MUTED

        # Bottom highlight pill
        h_box = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx + Inches(0.4), t_y + Inches(4.05), tw - Inches(0.8), Inches(0.35))
        h_box.fill.solid()
        h_box.fill.fore_color.rgb = COLOR_GREEN_BADGE
        h_box.line.fill.background()
        p_h = h_box.text_frame.paragraphs[0]
        p_h.text = f"★ {highlight}"
        p_h.font.size = Pt(11)
        p_h.font.bold = True
        p_h.font.color.rgb = COLOR_GREEN_TEXT
        p_h.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s8, """[ROTEIRO DE FALA - SLIDE 8: PRINCIPAIS TEÓRICOS]
- É fundamental saber quem foram os pioneiros desta revolução:
1. Ludwig von Bertalanffy: Um biólogo que percebeu que a física clássica não conseguia explicar a vida porque tratava tudo como sistemas fechados. Em 1950, publicou as bases da TGS, provando que um organismo ou uma organização precisam de trocas ambientais para se manterem vivos.
2. C. West Churchman: Trouxe a TGS para o mundo dos negócios através da obra seminal 'The Systems Approach' (1968). Ensinou gestores a pensar em tradeoffs, objetivos globais e gestão de complexidade.
3. Chester Barnard: Já em 1938, na obra clássica 'As Funções do Executivo', antecipou que empresas são sistemas cooperativos sociais. Sem a disposição das pessoas em colaborar em prol de um propósito comum, a estrutura desmorona.
Tempo estimado: 2 minutos.""")

    # ==========================================
    # SLIDE 9: Sistemas Abertos vs Fechados
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9, COLOR_BG_LIGHT)
    add_header(s9, "Tópico 7  •  Classificação", "Sistemas Abertos vs. Sistemas Fechados")

    card_comp_w = Inches(5.6)
    card_comp_h = Inches(4.8)

    # Card Fechado
    cf = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), card_comp_w, card_comp_h)
    cf.fill.solid()
    cf.fill.fore_color.rgb = COLOR_CARD_LIGHT
    cf.line.color.rgb = COLOR_CARD_BORDER
    cf.line.width = Pt(1.5)

    cf_title = s9.shapes.add_textbox(Inches(1.3), Inches(2.3), card_comp_w - Inches(1.0), Inches(0.6))
    p_cft = cf_title.text_frame.paragraphs[0]
    p_cft.text = "Sistemas Fechados"
    p_cft.font.name = FONT_HEADING
    p_cft.font.size = Pt(24)
    p_cft.font.bold = True
    p_cft.font.color.rgb = COLOR_TEXT_MAIN

    f_points = [
        "Sem intercâmbio com o meio externo.",
        "Comportamento puramente determinístico, mecânico e previsível.",
        "Sujeito à entropia positiva máxima: desgaste, obsolescência e colapso inevitável.",
        "Não reage nem se adapta a mudanças do ambiente."
    ]
    cf_box = s9.shapes.add_textbox(Inches(1.3), Inches(3.0), card_comp_w - Inches(1.0), Inches(2.2))
    cf_tf = cf_box.text_frame
    cf_tf.word_wrap = True
    for p_idx, pt_str in enumerate(f_points):
        p = cf_tf.paragraphs[0] if p_idx == 0 else cf_tf.add_paragraph()
        p.text = f"•  {pt_str}"
        p.font.size = Pt(13)
        p.font.color.rgb = COLOR_TEXT_MUTED
        p.space_after = Pt(8)

    # Exemplo pill fechado
    ex_f = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.3), Inches(5.5), card_comp_w - Inches(1.0), Inches(0.65))
    ex_f.fill.solid()
    ex_f.fill.fore_color.rgb = RGBColor(241, 245, 249)
    ex_f.line.fill.background()
    p_exf = ex_f.text_frame.paragraphs[0]
    p_exf.text = "Exemplo Prático: Máquinas isoladas, circuitos fechados, motores industriais sem sensor externo."
    p_exf.font.size = Pt(11)
    p_exf.font.italic = True
    p_exf.font.color.rgb = COLOR_TEXT_MUTED

    # Card Aberto (Dark Hero Highlight)
    ca = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(1.9), card_comp_w, card_comp_h)
    ca.fill.solid()
    ca.fill.fore_color.rgb = COLOR_CARD_DARK
    ca.line.fill.background()

    ca_title = s9.shapes.add_textbox(Inches(7.4), Inches(2.3), card_comp_w - Inches(1.0), Inches(0.6))
    p_cat = ca_title.text_frame.paragraphs[0]
    p_cat.text = "Sistemas Abertos"
    p_cat.font.name = FONT_HEADING
    p_cat.font.size = Pt(24)
    p_cat.font.bold = True
    p_cat.font.color.rgb = COLOR_TEXT_WHITE

    a_points = [
        "Troca constante e vital de energia, insumos, dados e matérias.",
        "Alta capacidade adaptativa e flexibilidade frente a incertezas.",
        "Possuem homeostase dinâmica e entropia negativa (reorganização constante).",
        "Atingem equifinalidade: atingem metas por rotas diversas."
    ]
    ca_box = s9.shapes.add_textbox(Inches(7.4), Inches(3.0), card_comp_w - Inches(1.0), Inches(2.2))
    ca_tf = ca_box.text_frame
    ca_tf.word_wrap = True
    for p_idx, pt_str in enumerate(a_points):
        p = ca_tf.paragraphs[0] if p_idx == 0 else ca_tf.add_paragraph()
        p.text = f"✓  {pt_str}"
        p.font.size = Pt(13)
        p.font.color.rgb = COLOR_MINT_ACCENT if p_idx == 0 else COLOR_TEXT_LIGHT_GRAY
        p.space_after = Pt(8)

    # Exemplo pill aberto
    ex_a = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.4), Inches(5.5), card_comp_w - Inches(1.0), Inches(0.65))
    ex_a.fill.solid()
    ex_a.fill.fore_color.rgb = RGBColor(30, 41, 59)
    ex_a.line.fill.background()
    p_exa = ex_a.text_frame.paragraphs[0]
    p_exa.text = "Exemplo Prático: Empresas, ecossistemas digitais, startups, seres humanos e sociedades."
    p_exa.font.size = Pt(11)
    p_exa.font.italic = True
    p_exa.font.color.rgb = COLOR_MINT_ACCENT

    set_speaker_notes(s9, """[ROTEIRO DE FALA - SLIDE 9: SISTEMAS ABERTOS VS. FECHADOS]
- Esta é uma das distinções centrais da disciplina:
- No mundo físico inanimado, muitos sistemas são fechados: um relógio de corda funciona de forma determinística, sem trocar com o exterior. Se quebrar uma engrenagem, ele para.
- Organizações humanas NUNCA são sistemas fechados. Quando uma empresa age como se fosse um sistema fechado (ignorando reclamações de clientes, tendências tecnológicas ou concorrentes), ela é punida pela entropia: definha e quebra.
- As empresas mais bem-sucedidas do século XXI são sistemas abertos por excelência: capturam tendências no ar e ajustam sua proposta de valor em tempo real.
Tempo estimado: 1 minuto e meio.""")

    # ==========================================
    # SLIDE 10: Abordagem Sociotécnica
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10, COLOR_BG_LIGHT)
    add_header(s10, "Tópico 8  •  Abordagem Sociotécnica", "A Abordagem Sociotécnica")

    # Left: Factory Image
    if os.path.exists("factory.png"):
        s10.shapes.add_picture("factory.png", Inches(0.8), Inches(1.9), width=Inches(5.6))
        c_fac = s10.shapes.add_textbox(Inches(0.8), Inches(5.9), Inches(5.6), Inches(0.5))
        p_fac = c_fac.text_frame.paragraphs[0]
        p_fac.text = "Indústria 4.0 & Automação: Tecnologia avançada integrada ao fator humano."
        p_fac.font.size = Pt(10)
        p_fac.font.italic = True
        p_fac.font.color.rgb = COLOR_TEXT_MUTED
        p_fac.alignment = PP_ALIGN.CENTER

    # Right: 2 Pillar Cards
    socio_w = Inches(5.8)
    socio_x = Inches(6.8)

    # Card 1: Técnico
    c_tec = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, socio_x, Inches(1.9), socio_w, Inches(2.0))
    c_tec.fill.solid()
    c_tec.fill.fore_color.rgb = COLOR_CARD_LIGHT
    c_tec.line.color.rgb = COLOR_CARD_BORDER
    c_tec.line.width = Pt(1.5)

    t_box_tec = s10.shapes.add_textbox(socio_x + Inches(0.4), Inches(2.1), socio_w - Inches(0.8), Inches(0.4))
    p_tec = t_box_tec.text_frame.paragraphs[0]
    p_tec.text = "Subsistema Técnico"
    p_tec.font.name = FONT_HEADING
    p_tec.font.size = Pt(19)
    p_tec.font.bold = True
    p_tec.font.color.rgb = COLOR_PRIMARY_TEAL

    d_box_tec = s10.shapes.add_textbox(socio_x + Inches(0.4), Inches(2.6), socio_w - Inches(0.8), Inches(1.1))
    d_tf_tec = d_box_tec.text_frame
    d_tf_tec.word_wrap = True
    p_dt1 = d_tf_tec.paragraphs[0]
    p_dt1.text = "Tecnologia, equipamentos, infraestrutura física, softwares, linhas de produção e processos operacionais padronizados."
    p_dt1.font.size = Pt(13)
    p_dt1.font.color.rgb = COLOR_TEXT_MAIN

    # Card 2: Social
    c_soc = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, socio_x, Inches(4.1), socio_w, Inches(2.0))
    c_soc.fill.solid()
    c_soc.fill.fore_color.rgb = COLOR_CARD_LIGHT
    c_soc.line.color.rgb = COLOR_CARD_BORDER
    c_soc.line.width = Pt(1.5)

    t_box_soc = s10.shapes.add_textbox(socio_x + Inches(0.4), Inches(4.3), socio_w - Inches(0.8), Inches(0.4))
    p_soc = t_box_soc.text_frame.paragraphs[0]
    p_soc.text = "Subsistema Social"
    p_soc.font.name = FONT_HEADING
    p_soc.font.size = Pt(19)
    p_soc.font.bold = True
    p_soc.font.color.rgb = COLOR_MINT_ACCENT

    d_box_soc = s10.shapes.add_textbox(socio_x + Inches(0.4), Inches(4.8), socio_w - Inches(0.8), Inches(1.1))
    d_tf_soc = d_box_soc.text_frame
    d_tf_soc.word_wrap = True
    p_ds1 = d_tf_soc.paragraphs[0]
    p_ds1.text = "Pessoas, relações interpessoais, liderança, cultura organizacional, motivação humana, valores e competências."
    p_ds1.font.size = Pt(13)
    p_ds1.font.color.rgb = COLOR_TEXT_MAIN

    # Bottom Callout
    bot_callout = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, socio_x, Inches(6.25), socio_w, Inches(0.45))
    bot_callout.fill.solid()
    bot_callout.fill.fore_color.rgb = COLOR_GREEN_BADGE
    bot_callout.line.fill.background()
    p_cal = bot_callout.text_frame.paragraphs[0]
    p_cal.text = "Princípio da Otimização Conjunta: O sucesso exige equilibrar técnica e pessoas."
    p_cal.font.size = Pt(11)
    p_cal.font.bold = True
    p_cal.font.color.rgb = COLOR_GREEN_TEXT
    p_cal.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s10, """[ROTEIRO DE FALA - SLIDE 10: ABORDAGEM SOCIOTÉCNICA]
- Desenvolvida pelo Tavistock Institute de Londres (com Eric Trist e Fred Emery), a Abordagem Sociotécnica nasceu de um insight prático:
- Muitas empresas compram os computadores mais caros, instalam robôs modernos ou implementam novas IAs, mas a produtividade desaba. Por quê?
- Porque a organização é composta por DOIS subsistemas interligados:
  1. O Técnico: Máquinas, fluxos, algoritmos, métodos de trabalho.
  2. O Social: Pessoas, liderança, confiança, cultura e motivação.
- Não existe tecnologia que compense uma equipe desmotivada ou uma liderança tóxica. O segredo da alta performance é a 'otimização conjunta': a tecnologia deve capacitar e empoderar o ser humano, e as pessoas devem estar preparadas para extrair o melhor das máquinas.
Tempo estimado: 2 minutos.""")

    # ==========================================
    # SLIDE 11: Gestão Sistêmica no Processo Decisório
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11, COLOR_BG_LIGHT)
    add_header(s11, "Síntese Prática", "Gestão Sistêmica no Processo Decisório")

    dec_w = Inches(6.0)
    dec_points = [
        ("Mapeamento Prévio de Impactos", "Avaliação das consequências em cascata antes de implementar qualquer decisão estratégica (prevenção do efeito dominó)."),
        ("Eliminação de Silos Operacionais", "Derrubada de barreiras comunicacionais e cooperação horizontal fluida entre todos os departamentos."),
        ("Alinhamento Dinâmico com o Ambiente", "Sincronização ágil entre os objetivos internos da empresa e as demandas voláteis dos clientes e do mercado.")
    ]

    for idx, (head, text_str) in enumerate(dec_points):
        dy = Inches(1.9) + idx * Inches(1.55)
        dcard = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), dy, dec_w, Inches(1.35))
        dcard.fill.solid()
        dcard.fill.fore_color.rgb = COLOR_CARD_LIGHT
        dcard.line.color.rgb = COLOR_CARD_BORDER
        dcard.line.width = Pt(1.5)

        # Check circle
        ck = s11.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.15), dy + Inches(0.35), Inches(0.55), Inches(0.55))
        ck.fill.solid()
        ck.fill.fore_color.rgb = COLOR_MINT_ACCENT
        ck.line.fill.background()
        p_ck = ck.text_frame.paragraphs[0]
        p_ck.text = "✓"
        p_ck.font.bold = True
        p_ck.font.size = Pt(15)
        p_ck.font.color.rgb = COLOR_TEXT_WHITE
        p_ck.alignment = PP_ALIGN.CENTER

        dt_box = s11.shapes.add_textbox(Inches(1.9), dy + Inches(0.18), dec_w - Inches(2.1), Inches(0.35))
        p_dt = dt_box.text_frame.paragraphs[0]
        p_dt.text = head
        p_dt.font.name = FONT_HEADING
        p_dt.font.size = Pt(15)
        p_dt.font.bold = True
        p_dt.font.color.rgb = COLOR_TEXT_MAIN

        dd_box = s11.shapes.add_textbox(Inches(1.9), dy + Inches(0.55), dec_w - Inches(2.1), Inches(0.65))
        dd_tf = dd_box.text_frame
        dd_tf.word_wrap = True
        dd_tf.margin_top = dd_tf.margin_bottom = 0
        p_dd = dd_tf.paragraphs[0]
        p_dd.text = text_str
        p_dd.font.name = FONT_BODY
        p_dd.font.size = Pt(12)
        p_dd.font.color.rgb = COLOR_TEXT_MUTED

    # Right: Cubes Image
    if os.path.exists("cubes.jpg"):
        s11.shapes.add_picture("cubes.jpg", Inches(7.2), Inches(1.9), width=Inches(5.3))
        c_cub = s11.shapes.add_textbox(Inches(7.2), Inches(5.9), Inches(5.3), Inches(0.5))
        p_cub = c_cub.text_frame.paragraphs[0]
        p_cub.text = "Tomada de Decisão Sistêmica: Conectando lideranças e eliminando silos."
        p_cub.font.size = Pt(10)
        p_cub.font.italic = True
        p_cub.font.color.rgb = COLOR_TEXT_MUTED
        p_cub.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s11, """[ROTEIRO DE FALA - SLIDE 11: PROCESSO DECISÓRIO SISTÊMICO]
- Como transformar a teoria em prática diária de gestão?
- Três práticas indispensáveis:
1. Mapeamento prévio de impactos: Antes de cortar custos ou alterar um software, o líder sistêmico pergunta: 'Quem será afetado? O atendimento ao cliente vai sofrer? O pós-venda vai travar?'. Evita-se o clássico 'remédio que mata o paciente'.
2. Eliminação de silos: As reuniões não podem ser apenas de diretoria fechada. Criam-se equipes multidisciplinares e rituais de alinhamento entre áreas distintas.
3. Alinhamento dinâmico: Metas internas precisam conversar com a realidade externa. Não adianta bater a meta de produzir 1 milhão de itens se o mercado mudou e não quer mais aquele modelo.
Tempo estimado: 1 minuto e meio.""")

    # ==========================================
    # SLIDE 12: Conclusão / Frase de Impacto (Dark Theme)
    # ==========================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12, COLOR_BG_DARK)

    # Center Cap Icon / Badge
    cap_ic = s12.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.166), Inches(1.5), Inches(1.0), Inches(1.0))
    cap_ic.fill.solid()
    cap_ic.fill.fore_color.rgb = COLOR_PRIMARY_TEAL
    cap_ic.line.fill.background()
    p_cap = cap_ic.text_frame.paragraphs[0]
    p_cap.text = "🎓"
    p_cap.font.size = Pt(24)
    p_cap.alignment = PP_ALIGN.CENTER

    # Title
    t12 = s12.shapes.add_textbox(Inches(1.5), Inches(2.7), Inches(10.333), Inches(0.8))
    p_t12 = t12.text_frame.paragraphs[0]
    p_t12.text = "Visão Sistêmica como Vantagem Competitiva"
    p_t12.font.name = FONT_HEADING
    p_t12.font.size = Pt(36)
    p_t12.font.bold = True
    p_t12.font.color.rgb = COLOR_TEXT_WHITE
    p_t12.alignment = PP_ALIGN.CENTER

    # Quote Box
    q_box = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.2), Inches(3.7), Inches(8.933), Inches(1.5))
    q_box.fill.solid()
    q_box.fill.fore_color.rgb = RGBColor(20, 31, 60)
    q_box.line.color.rgb = COLOR_PRIMARY_TEAL
    q_box.line.width = Pt(1.5)

    q_tf = q_box.text_frame
    q_tf.word_wrap = True
    p_q = q_tf.paragraphs[0]
    p_q.text = "“A organização não é uma soma de partes isoladas,\nmas uma rede viva de interações contínuas.”"
    p_q.font.name = FONT_HEADING
    p_q.font.size = Pt(22)
    p_q.font.italic = True
    p_q.font.color.rgb = COLOR_MINT_ACCENT
    p_q.alignment = PP_ALIGN.CENTER

    # 3 Pillars summary
    pills_box = s12.shapes.add_textbox(Inches(1.5), Inches(5.6), Inches(10.333), Inches(0.8))
    p_pl = pills_box.text_frame.paragraphs[0]
    p_pl.text = "Pensamento Holístico   •   Integração Tecnológica & Humana   •   Adaptabilidade Contínua"
    p_pl.font.size = Pt(14)
    p_pl.font.bold = True
    p_pl.font.color.rgb = COLOR_TEXT_LIGHT_GRAY
    p_pl.alignment = PP_ALIGN.CENTER

    set_speaker_notes(s12, """[ROTEIRO DE FALA - SLIDE 12: CONCLUSÃO E SÍNTESE]
- Chegamos à conclusão da nossa apresentação com esta reflexão fundamental:
- 'A organização não é uma soma de partes isoladas, mas uma rede viva de interações contínuas.'
- Para quem administra, lidera ou empreende no século XXI, a Teoria Sistêmica não é apenas uma teoria de livro: é uma lente de navegação para a sobrevivência e para a alta performance.
- Líderes que pensam de forma sistêmica constroem organizações mais resilientes, empresas mais inovadoras e ambientes de trabalho mais cooperativos e humanos.
- Muito obrigado pela atenção de todos! Estou à disposição para dúvidas, comentários e debate.
Tempo estimado: 1 a 2 minutos.""")

    # ==========================================
    # SLIDE 13: Fontes e Referências
    # ==========================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13, COLOR_BG_LIGHT)
    add_header(s13, "Fontes & Leituras Recomendadas", "Fundamentação Bibliográfica e Fontes Visuais")

    # Col 1: Academic Sources
    ref_w = Inches(5.6)
    ref_h = Inches(4.8)
    
    rc1 = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), ref_w, ref_h)
    rc1.fill.solid()
    rc1.fill.fore_color.rgb = COLOR_CARD_LIGHT
    rc1.line.color.rgb = COLOR_CARD_BORDER
    rc1.line.width = Pt(1.5)

    t_r1 = s13.shapes.add_textbox(Inches(1.2), Inches(2.2), ref_w - Inches(0.8), Inches(0.4))
    p_tr1 = t_r1.text_frame.paragraphs[0]
    p_tr1.text = "📚 Obras Clássicas de Referência"
    p_tr1.font.name = FONT_HEADING
    p_tr1.font.size = Pt(18)
    p_tr1.font.bold = True
    p_tr1.font.color.rgb = COLOR_PRIMARY_TEAL

    academic_refs = [
        "BERTALANFFY, Ludwig von. Teoria Geral dos Sistemas: Fundamentos, Desenvolvimento e Aplicações. Petrópolis: Vozes, 1975.",
        "CHURCHMAN, C. West. The Systems Approach. New York: Dell Publishing, 1968.",
        "BARNARD, Chester. As Funções do Executivo. São Paulo: Atlas, 1979.",
        "KATZ, Daniel; KAHN, Robert L. Psicologia Social das Organizações. São Paulo: Atlas, 1978.",
        "CHIAVENATO, Idalberto. Introdução à Teoria Geral da Administração. 10. ed. Rio de Janeiro: GEN Atlas, 2020."
    ]
    
    rb_box = s13.shapes.add_textbox(Inches(1.2), Inches(2.8), ref_w - Inches(0.8), Inches(3.6))
    rb_tf = rb_box.text_frame
    rb_tf.word_wrap = True
    for idx, ref_text in enumerate(academic_refs):
        p = rb_tf.paragraphs[0] if idx == 0 else rb_tf.add_paragraph()
        p.text = f"•  {ref_text}"
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_TEXT_MAIN
        p.space_after = Pt(10)

    # Col 2: Digital & Visual Sources
    rc2 = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(1.9), ref_w, ref_h)
    rc2.fill.solid()
    rc2.fill.fore_color.rgb = COLOR_CARD_LIGHT
    rc2.line.color.rgb = COLOR_CARD_BORDER
    rc2.line.width = Pt(1.5)

    t_r2 = s13.shapes.add_textbox(Inches(7.3), Inches(2.2), ref_w - Inches(0.8), Inches(0.4))
    p_tr2 = t_r2.text_frame.paragraphs[0]
    p_tr2.text = "🌐 Estudos e Fontes Visuais"
    p_tr2.font.name = FONT_HEADING
    p_tr2.font.size = Pt(18)
    p_tr2.font.bold = True
    p_tr2.font.color.rgb = COLOR_PRIMARY_TEAL

    digital_refs = [
        ("MIT Sloan Management Review", "What is Systems Thinking in Business?", "https://executive.mit.edu/"),
        ("Frontiers in Psychology", "Systems Theory and Behavioral Interaction Networks", "https://www.frontiersin.org/"),
        ("NVIDIA Omniverse Manufacturing", "Digital Twins & Cyber-Physical Operational Systems", "https://nvidianews.nvidia.com/"),
        ("Tavistock Institute", "Socio-Technical Systems Design and Organizational Work", "Historical Research Papers")
    ]
    
    rd_box = s13.shapes.add_textbox(Inches(7.3), Inches(2.8), ref_w - Inches(0.8), Inches(3.6))
    rd_tf = rd_box.text_frame
    rd_tf.word_wrap = True
    for idx, (source_name, title_desc, url_str) in enumerate(digital_refs):
        p_name = rd_tf.paragraphs[0] if idx == 0 else rd_tf.add_paragraph()
        p_name.text = f"🔗 {source_name}"
        p_name.font.bold = True
        p_name.font.size = Pt(12)
        p_name.font.color.rgb = COLOR_TEXT_MAIN
        
        p_desc = rd_tf.add_paragraph()
        p_desc.text = f"   {title_desc}"
        p_desc.font.size = Pt(11)
        p_desc.font.color.rgb = COLOR_TEXT_MUTED

        p_u = rd_tf.add_paragraph()
        p_u.text = f"   Fonte: {url_str}"
        p_u.font.size = Pt(10)
        p_u.font.italic = True
        p_u.font.color.rgb = COLOR_PRIMARY_TEAL
        p_u.space_after = Pt(8)

    set_speaker_notes(s13, """[ROTEIRO DE FALA - SLIDE 13: FONTES E REFERÊNCIAS]
- Para aprofundamento, recomendamos as obras fundamentais de Bertalanffy, Churchman e Barnard, além das aplicações contemporâneas do MIT Sloan e estudos em sistemas sociotécnicos.
- Os links e fontes visuais das ilustrações estão devidamente creditados e catalogados para consulta acadêmica.""")

    output_filename = "apresentacao_teoria_sistemica.pptx"
    prs.save(output_filename)
    print(f"Presentation saved successfully as '{output_filename}' with {len(prs.slides)} slides!")

if __name__ == "__main__":
    create_deck()
