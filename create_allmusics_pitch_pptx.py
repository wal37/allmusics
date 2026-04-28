from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


OUT = Path("/Users/wale/Desktop/musically/musicalweb/ALLMUSICS_PITCH_DECK.pptx")

SLIDES = [
    {
        "bg": "0B0813",
        "accent": "FF5FA2",
        "title": "AllMusics",
        "body": "A music generation platform for artists, composers, film teams, and studios building songs, stems, moods, and sound ideas at production speed.",
        "label": "Cover",
        "bars": ["FF5FA2", "FFB55F", "8F7DFF"],
    },
    {
        "bg": "F7F2FF",
        "accent": "8F7DFF",
        "title": "The problem",
        "body": "Music creation tools are still split across ideation, generation, editing, versioning, and delivery. That slows artists, producers, and media teams that need fast output with usable control.",
        "label": "Problem",
        "cards": [
            ("Fragmented workflow", "Prompting, arranging, editing, and export often live in different tools."),
            ("Weak production control", "Creators get outputs but not enough structure for real creative iteration."),
            ("Slow commercial delivery", "Teams still need a system for versions, stems, and client-ready assets."),
        ],
    },
    {
        "bg": "120D22",
        "accent": "D4FF6D",
        "title": "Why now",
        "body": "Generative audio is moving from novelty toward production use. The next winner is not only a model, but the platform that wraps real workflow around it.",
        "label": "Why now",
        "cards": [
            ("Creators want speed", "Artists and producers need more throughput without killing momentum."),
            ("Media demand is rising", "Video, branded content, games, and social media need constant audio output."),
            ("Teams need control", "Studios and platforms need repeatability, rights, and export discipline."),
        ],
    },
    {
        "bg": "FFF8EE",
        "accent": "FF8E5E",
        "title": "Product",
        "body": "AllMusics is a prompt-to-sound platform that turns text, references, lyrics, and mood into songs, cues, stems, and editable audio directions from one operating layer.",
        "label": "Platform",
        "cards": [
            ("Song generation", "Generate complete tracks and musical drafts from prompts."),
            ("Stem control", "Separate output into layers that can be refined and reused."),
            ("Scene scoring", "Create cinematic and ambient pieces for film and visual work."),
            ("Project memory", "Save prompts, versions, and directions across creative sessions."),
        ],
    },
    {
        "bg": "0D0A18",
        "accent": "5AC8FF",
        "title": "Workflow",
        "body": "The operating rhythm is direction, generation, iteration, and delivery. AllMusics is built around how real audio work moves, not just how a model responds.",
        "label": "Workflow",
        "steps": [
            "Input style, mood, tempo, references, or lyrics.",
            "Generate candidate outputs and compare directions.",
            "Refine structure, intensity, length, or instrumentation.",
            "Export tracks, stems, and usable commercial assets.",
        ],
    },
    {
        "bg": "F7F2FF",
        "accent": "FF5FA2",
        "title": "Who it serves",
        "body": "AllMusics serves creators and teams that need audio output quickly, but still care about creative quality, asset organization, and delivery speed.",
        "label": "Customers",
        "cards": [
            ("Artists", "Prototype songs, hooks, and demos faster."),
            ("Producers", "Test styles, references, and arrangements."),
            ("Film teams", "Generate soundtrack drafts and scene moods."),
            ("Studios", "Manage version-heavy music workflows."),
        ],
    },
    {
        "bg": "140F26",
        "accent": "D4FF6D",
        "title": "Technology stack",
        "body": "The platform stack includes music generation models, prompt conditioning, vocal and arrangement tooling, version storage, export controls, and future API delivery for external products.",
        "label": "Technology",
        "cards": [
            ("Generation layer", "Audio models for music, vocals, and style variation."),
            ("Editing layer", "Extension, clipping, versioning, and stem operations."),
            ("Business layer", "Workspaces, licensing logic, export, and usage control."),
        ],
    },
    {
        "bg": "FFF8EE",
        "accent": "8F7DFF",
        "title": "Use cases",
        "body": "AllMusics can power direct creator workflows and become embedded inside broader software products that need music generation and audio ideation.",
        "label": "Use cases",
        "cards": [
            ("Creator tool", "Direct product for artists and producers."),
            ("Studio engine", "Internal workflow layer for music teams."),
            ("Media pipeline", "Faster soundtrack and branded content generation."),
            ("API platform", "Generation endpoints for partner products."),
        ],
    },
    {
        "bg": "0B0813",
        "accent": "FFB55F",
        "title": "Business model",
        "body": "The platform can monetize through creator subscriptions, team plans, commercial licensing, and API usage for software and media partners.",
        "label": "Business",
        "cards": [
            ("Individual creators", "Monthly usage-based plans."),
            ("Teams and studios", "Seats, workspaces, and shared projects."),
            ("Enterprise and API", "Embedded generation and commercial partnerships."),
        ],
    },
    {
        "bg": "F7F2FF",
        "accent": "5AC8FF",
        "title": "Differentiation",
        "body": "The defensible layer is workflow depth: generation plus control, versioning, stems, export discipline, and team-ready operations around creative output.",
        "label": "Moat",
        "cards": [
            ("Not just output", "AllMusics focuses on usable production workflow."),
            ("Not just individuals", "The system is built for teams and commercial work."),
            ("Not just one model", "The platform layer can evolve with the generation stack."),
        ],
    },
    {
        "bg": "120D22",
        "accent": "FF5FA2",
        "title": "Roadmap",
        "body": "Start with generation and creator workflows, then expand into deeper editing, team controls, and platform APIs for external software and media products.",
        "label": "Roadmap",
        "steps": [
            "Phase 1: Creator product and core song generation.",
            "Phase 2: Stem tools, scene scoring, and commercial exports.",
            "Phase 3: Team workspaces, licensing controls, and API access.",
        ],
    },
    {
        "bg": "FFF8EE",
        "accent": "D4FF6D",
        "title": "Vision",
        "body": "AllMusics can become the operating system for fast audio creation: a place where music, soundtrack, ideation, and delivery happen in one modern workflow.",
        "label": "Close",
        "bars": ["D4FF6D", "5AC8FF", "FF5FA2"],
    },
]


def esc(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def content_types(slide_count: int) -> str:
    overrides = "\n".join(
        f'  <Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, slide_count + 1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>
  <Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
{overrides}
</Types>
"""


def rels_root() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def presentation_xml(slide_count: int) -> str:
    slide_ids = "\n".join(
        f'    <p:sldId id="{256 + i}" r:id="rId{i + 1}"/>' for i in range(slide_count)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 saveSubsetFonts="1" autoCompressPictures="0">
  <p:sldMasterIdLst/>
  <p:sldIdLst>
{slide_ids}
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <a:defPPr/>
    <a:lvl1pPr marL="0" indent="0"/>
    <a:lvl2pPr marL="457200" indent="0"/>
    <a:lvl3pPr marL="914400" indent="0"/>
  </p:defaultTextStyle>
</p:presentation>
"""


def presentation_rels(slide_count: int) -> str:
    rels = "\n".join(
        f'  <Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i + 1}.xml"/>'
        for i in range(slide_count)
    )
    base = [
        rels,
        '  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
        '  <Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>',
        '  <Relationship Id="rId102" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>',
        '  <Relationship Id="rId103" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>',
    ]
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
""" + "\n".join(base) + "\n</Relationships>\n"


def simple_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""


def app_xml(slide_count: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office PowerPoint</Application>
  <Slides>{slide_count}</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>1</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Office Theme</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company>OpenAI</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
"""


def core_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>AllMusics Pitch Deck</dc:title>
  <dc:creator>OpenAI</dc:creator>
  <cp:lastModifiedBy>OpenAI</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-04-28T14:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-04-28T14:00:00Z</dcterms:modified>
</cp:coreProperties>
"""


def theme_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="0B0813"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="140F26"/></a:dk2>
      <a:lt2><a:srgbClr val="F7F2FF"/></a:lt2>
      <a:accent1><a:srgbClr val="FF5FA2"/></a:accent1>
      <a:accent2><a:srgbClr val="8F7DFF"/></a:accent2>
      <a:accent3><a:srgbClr val="5AC8FF"/></a:accent3>
      <a:accent4><a:srgbClr val="D4FF6D"/></a:accent4>
      <a:accent5><a:srgbClr val="FF8E5E"/></a:accent5>
      <a:accent6><a:srgbClr val="FFB55F"/></a:accent6>
      <a:hlink><a:srgbClr val="5AC8FF"/></a:hlink>
      <a:folHlink><a:srgbClr val="FF5FA2"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont>
        <a:latin typeface="Aptos Display"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Aptos"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Custom">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
"""


def view_props() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:normalViewPr/>
  <p:slideViewPr/>
  <p:notesTextViewPr/>
  <p:gridSpacing cx="72008" cy="72008"/>
</p:viewPr>
"""


def pres_props() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:showPr loop="0" useTimings="0"/>
</p:presentationPr>
"""


def table_styles() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5940675A-B579-460E-94D1-54222C63F5DA}"/>
"""


def shape_bg(shape_id: int, x: int, y: int, cx: int, cy: int, color: str) -> str:
    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="{shape_id}" name="Shape {shape_id}"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:style>
        <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
        <a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>
        <a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef>
        <a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
      </p:style>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>"""


def textbox(shape_id: int, x: int, y: int, cx: int, cy: int, text: str, size: int, color: str, bold: bool = False, font: str = "Aptos", align: str = "l") -> str:
    b = ' b="1"' if bold else ""
    return f"""
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="{shape_id}" name="TextBox {shape_id}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="{align}"/>
          <a:r>
            <a:rPr lang="en-US" sz="{size * 100}"{b}>
              <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
              <a:latin typeface="{font}"/>
            </a:rPr>
            <a:t>{esc(text)}</a:t>
          </a:r>
          <a:endParaRPr lang="en-US" sz="{size * 100}">
            <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
            <a:latin typeface="{font}"/>
          </a:endParaRPr>
        </a:p>
      </p:txBody>
    </p:sp>"""


def slide_xml(data: dict, idx: int) -> str:
    text_color = "F7F4FF" if data["bg"] in {"0B0813", "120D22", "140F26", "0D0A18"} else "111018"
    muted = "C8C0D9" if text_color == "F7F4FF" else "655E74"
    shapes = []
    sid = 2
    shapes.append(shape_bg(sid, 0, 0, 12192000, 6858000, data["bg"]))
    sid += 1
    shapes.append(shape_bg(sid, 730000, 620000, 10800000, 22000, data["accent"]))
    sid += 1
    shapes.append(textbox(sid, 840000, 740000, 1700000, 280000, data["label"].upper(), 11, data["accent"], True, "Aptos Display"))
    sid += 1
    shapes.append(textbox(sid, 840000, 1120000, 5500000, 720000, data["title"], 30, text_color, True, "Aptos Display"))
    sid += 1
    shapes.append(textbox(sid, 840000, 1960000, 5800000, 1360000, data["body"], 18, muted, False, "Aptos"))
    sid += 1

    if "bars" in data:
      x = 7550000
      widths = [3100000, 2400000, 3100000]
      heights = [820000, 720000, 2100000]
      tops = [1100000, 2260000, 3260000]
      for i, color in enumerate(data["bars"]):
          shapes.append(shape_bg(sid, x + (260000 if i == 1 else 0), tops[i], widths[i], heights[i], color))
          sid += 1

    if "cards" in data:
      cards = data["cards"]
      if len(cards) == 3:
          positions = [(840000, 3720000), (4240000, 3720000), (7640000, 3720000)]
          widths = [2800000, 2800000, 2800000]
      elif len(cards) == 4:
          positions = [(840000, 3560000), (3860000, 3560000), (6880000, 3560000), (9900000, 3560000)]
          widths = [2500000, 2500000, 2500000, 2100000]
      else:
          positions = [(7600000, 1600000), (7600000, 3040000), (7600000, 4480000)]
          widths = [3480000, 3480000, 3480000]
      for i, (title, body) in enumerate(cards):
          fill = "171127" if text_color == "F7F4FF" else "FFFFFF"
          line = data["accent"]
          x, y = positions[i]
          w = widths[i]
          shapes.append(shape_bg(sid, x, y, w, 1180000 if len(cards) == 4 else 1320000, fill))
          sid += 1
          shapes.append(shape_bg(sid, x, y, w, 18000, line))
          sid += 1
          shapes.append(textbox(sid, x + 180000, y + 150000, w - 260000, 260000, title, 15, text_color, True, "Aptos Display"))
          sid += 1
          shapes.append(textbox(sid, x + 180000, y + 460000, w - 260000, 580000, body, 11, muted, False, "Aptos"))
          sid += 1

    if "steps" in data:
        y = 3760000
        for i, step in enumerate(data["steps"], start=1):
            x = 980000 + (i - 1) * 2760000
            shapes.append(shape_bg(sid, x, y, 340000, 340000, data["accent"]))
            sid += 1
            shapes.append(textbox(sid, x + 70000, y + 65000, 220000, 180000, f"{i:02d}", 12, "111018", True, "Aptos Display", "ctr"))
            sid += 1
            shapes.append(textbox(sid, x, y + 500000, 2200000, 760000, step, 14, text_color, False, "Aptos"))
            sid += 1

    shapes.append(textbox(sid, 11280000, 250000, 480000, 220000, f"{idx:02d}", 10, muted if text_color == "F7F4FF" else text_color, True, "Aptos Display", "r"))
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
""" + "\n".join(shapes) + """
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"""


def build() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        OUT.unlink()
    with ZipFile(OUT, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types(len(SLIDES)))
        zf.writestr("_rels/.rels", rels_root())
        zf.writestr("docProps/app.xml", app_xml(len(SLIDES)))
        zf.writestr("docProps/core.xml", core_xml())
        zf.writestr("ppt/presentation.xml", presentation_xml(len(SLIDES)))
        zf.writestr("ppt/_rels/presentation.xml.rels", presentation_rels(len(SLIDES)))
        zf.writestr("ppt/theme/theme1.xml", theme_xml())
        zf.writestr("ppt/viewProps.xml", view_props())
        zf.writestr("ppt/presProps.xml", pres_props())
        zf.writestr("ppt/tableStyles.xml", table_styles())
        for idx, slide in enumerate(SLIDES, start=1):
            zf.writestr(f"ppt/slides/slide{idx}.xml", slide_xml(slide, idx))
            zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", simple_rels())


if __name__ == "__main__":
    build()
