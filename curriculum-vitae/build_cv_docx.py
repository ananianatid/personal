from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


OUT = "cv-anatide-anani.docx"

INK = "111827"
BODY = "374151"
MUTED = "6B7280"
BLUE = "2563EB"
BLUE_DARK = "1E3A8A"
LINE = "D9E2EF"
SOFT = "F5F8FC"
WHITE = "FFFFFF"


def color(hex_value):
    return RGBColor.from_string(hex_value)


def set_font(run, size, bold=False, color_hex=BODY):
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color(color_hex)


def add_run(paragraph, text, size=9, bold=False, color_hex=BODY):
    run = paragraph.add_run(text)
    set_font(run, size=size, bold=bold, color_hex=color_hex)
    return run


def pformat(paragraph, before=0, after=3, line=1.12, align=None):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line
    if align is not None:
        paragraph.alignment = align


def paragraph_border_bottom(paragraph, color_hex=LINE, size="6"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "3")
    bottom.set(qn("w:color"), color_hex)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def no_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "nil")


def shade(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def section(doc, title):
    p = doc.add_paragraph()
    pformat(p, before=8, after=5, line=1.0)
    add_run(p, title.upper(), size=8.5, bold=True, color_hex=BLUE_DARK)
    paragraph_border_bottom(p, color_hex=LINE, size="7")


def meta_line(doc, text):
    p = doc.add_paragraph()
    pformat(p, after=2, line=1.08)
    add_run(p, text, size=8.4, color_hex=MUTED)


def entry(doc, title, date=None, org=None, desc=None, tags=None):
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(12.6)
    table.columns[1].width = Cm(4.2)
    for cell in table.rows[0].cells:
        no_borders(cell)
        set_cell_margins(cell, top=0, bottom=0, start=0, end=0)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    left, right = table.rows[0].cells
    p = left.paragraphs[0]
    pformat(p, after=0, line=1.0)
    add_run(p, title, size=9.3, bold=True, color_hex=INK)
    rp = right.paragraphs[0]
    pformat(rp, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    if date:
        add_run(rp, date, size=8.0, bold=True, color_hex=BLUE)
    if org:
        meta_line(doc, org)
    if desc:
        p = doc.add_paragraph()
        pformat(p, after=3, line=1.13)
        add_run(p, desc, size=8.4, color_hex=BODY)
    if tags:
        p = doc.add_paragraph()
        pformat(p, after=4, line=1.08)
        for i, tag in enumerate(tags):
            if i:
                add_run(p, "  ·  ", size=8.0, color_hex=MUTED)
            add_run(p, tag, size=8.0, bold=True, color_hex=BLUE_DARK)


def skill_line(doc, label, values):
    p = doc.add_paragraph()
    pformat(p, after=2, line=1.12)
    add_run(p, label + " : ", size=8.5, bold=True, color_hex=INK)
    add_run(p, ", ".join(values), size=8.4, color_hex=BODY)


def language_line(doc):
    p = doc.add_paragraph()
    pformat(p, after=2, line=1.1)
    parts = [("Français", "Avancé"), ("Anglais", "B2"), ("Ewe", "Natif")]
    for i, (lang, level) in enumerate(parts):
        if i:
            add_run(p, "   ·   ", size=8.4, color_hex=MUTED)
        add_run(p, lang, size=8.4, bold=True, color_hex=INK)
        add_run(p, f" {level}", size=8.3, color_hex=MUTED)


def bullet_line(doc, values):
    p = doc.add_paragraph()
    pformat(p, after=2, line=1.1)
    for i, value in enumerate(values):
        if i:
            add_run(p, "   ·   ", size=8.4, color_hex=MUTED)
        add_run(p, value, size=8.4, color_hex=BODY)


def configure(doc):
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.25)
    section.bottom_margin = Cm(1.15)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)
    section.header_distance = Cm(0.5)
    section.footer_distance = Cm(0.45)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    normal.font.size = Pt(8.5)
    normal.font.color.rgb = color(BODY)
    normal.paragraph_format.space_after = Pt(2)
    normal.paragraph_format.line_spacing = 1.1


def build():
    doc = Document()
    configure(doc)

    top = doc.add_table(rows=1, cols=2)
    top.autofit = False
    top.columns[0].width = Cm(9.9)
    top.columns[1].width = Cm(6.9)
    for cell in top.rows[0].cells:
        no_borders(cell)
        set_cell_margins(cell, top=0, bottom=80, start=0, end=0)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    left, right = top.rows[0].cells
    p = left.paragraphs[0]
    pformat(p, after=2, line=1.0)
    add_run(p, "Anatide ANANI", size=24, bold=True, color_hex=INK)
    p = left.add_paragraph()
    pformat(p, after=0, line=1.0)
    add_run(p, "Développeur Web Full-Stack", size=10.2, bold=True, color_hex=BLUE)
    add_run(p, "  ·  Disponible pour un stage", size=9.0, color_hex=MUTED)

    for text in ("Lomé, Togo", "ananianatid@gmail.com", "+228 70 07 68 29", "github.com/ananianatid"):
        p = right.add_paragraph() if right.paragraphs[0].text else right.paragraphs[0]
        pformat(p, after=1, line=1.0, align=WD_ALIGN_PARAGRAPH.RIGHT)
        add_run(p, text, size=8.3, color_hex=BODY)

    intro = doc.add_paragraph()
    pformat(intro, before=2, after=4, line=1.15)
    paragraph_border_bottom(intro, color_hex=BLUE, size="10")
    add_run(
        intro,
        "Passionné d'informatique depuis le lycée, j'ai démarré avec Arduino et l'IoT avant de plonger dans le développement web. Aujourd'hui en 3ème année de Licence en Génie Logiciel, je conçois des applications qui répondent à de vrais besoins, avec une approche orientée architecture propre et expérience utilisateur.",
        size=8.8,
        color_hex=BODY,
    )

    section(doc, "Compétences")
    skill_line(doc, "Backend", ["Laravel 10/12", "Statamic", "PHP 8+", "Node.js", "Express", "Sequelize", "Filament v3", "API REST", "Laravel Passport", "Sanctum"])
    skill_line(doc, "Frontend & mobile", ["JavaScript ES6+", "Expo", "HTML5", "CSS3", "Tailwind", "Bootstrap"])
    skill_line(doc, "Bases de données", ["MySQL", "SQLite", "PostgreSQL"])
    skill_line(doc, "Outils & DevOps", ["Git", "GitHub", "GitLab", "Vite", "Composer", "npm", "Supabase", "WordPress", "Docker"])
    skill_line(doc, "Fondements & scripting", ["Java", "Python", "C", "Bash"])

    section(doc, "Expérience")
    entry(
        doc,
        "Gestionnaire de site WordPress",
        "2024 - Présent",
        "UFW Verein · ONG internationale",
        "Administration et maintenance du site institutionnel en production. Mise à jour du contenu éditorial et des modules. Suivi des performances et résolution d'incidents.",
    )

    section(doc, "Projets")
    entry(
        doc,
        "logements_Mokpokpo",
        None,
        None,
        "Application complète de gestion des résidences universitaires. Panneau d'administration Filament v3, algorithme de scoring de priorité pour l'attribution des logements, gestion des contrats et paiements.",
        ["Laravel 12", "Filament v3", "PHP 8+", "MySQL", "Tailwind"],
    )
    entry(
        doc,
        "collegeDirectory",
        None,
        None,
        "Annuaire universitaire avec double panel admin/étudiant. SSO OAuth2 / OpenID Connect via Laravel Passport, fournisseur Socialite personnalisé et pattern shadow user.",
        ["Laravel", "Filament v3", "Passport", "OAuth2", "MySQL"],
    )
    entry(
        doc,
        "App Atelier de Couture",
        None,
        None,
        "Application mobile offline-first pour la gestion de commandes et mesures clients d'un atelier de couture. Modélisation dynamique des mesures.",
        ["React Native", "Expo", "SQLite", "Laravel", "MySQL"],
    )
    entry(
        doc,
        "UFW-verein",
        None,
        None,
        "Administration et maintenance du site institutionnel en production. Mise à jour du contenu éditorial et des modules. Suivi des performances et résolution d'incidents.",
        ["WordPress", "PHP"],
    )
    entry(
        doc,
        "OQSF Bénin",
        None,
        None,
        "Portail institutionnel de l'Observatoire de la Qualité des Services Financiers. Gestion des plaintes, ressources réglementaires et actualités financières.",
        ["Statamic", "Laravel", "PHP", "Tailwind"],
    )

    section(doc, "Formation")
    entry(
        doc,
        "Licence en Génie Logiciel - L3",
        "2023 - En cours",
        "Université · Lomé, Togo",
        "Formation en génie logiciel avec pratique des langages C, Python, Java, JavaScript, PHP. Projets académiques orientés conception et développement d'applications.",
    )
    entry(
        doc,
        "Certification OPEM B",
        "2022 - 2023",
        "Conception d'une maison à éclairage connecté Bluetooth",
        "Code Arduino, plan AutoCAD architecture, modélisation 3D Blender, schémas électroniques Fritzing et TinkerCAD.",
    )
    entry(
        doc,
        "Certification OPEM A",
        "2021 - 2022",
        "Conception d'un portail à ouverture automatique",
        "Conception et programmation d'un système de portail automatisé avec Arduino.",
    )
    entry(doc, "Baccalauréat Série D - Mention Bien", "2020 - 2023", "Lomé, Togo")

    section(doc, "Langues & loisirs")
    language_line(doc)
    bullet_line(doc, ["Modélisation 3D (Blender)", "Échecs (1350 Elo)", "Lecture"])

    footer = doc.sections[0].footer.paragraphs[0]
    pformat(footer, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_run(footer, "© 2025 Anatide ANANI · CV", size=7.0, color_hex=MUTED)

    doc.save(OUT)


if __name__ == "__main__":
    build()
