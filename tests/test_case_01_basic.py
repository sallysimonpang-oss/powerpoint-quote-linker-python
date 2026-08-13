from hashlib import sha256
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from ppt_quote_linker.linker import link_presentation

CASE = Path(__file__).parent / "examples" / "case_01_basic"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"a": A_NS, "r": R_NS}


def _golden_templates() -> dict[str, str]:
    templates = {}
    with ZipFile(CASE / "expected.pptx") as package:
        for number in range(1, 6):
            root = etree.fromstring(package.read(f"ppt/diagrams/_rels/data{number}.xml.rels"))
            target = next(iter(root)).get("Target")
            templates[f"ppt/diagrams/data{number}.xml"] = target.rsplit("q=", 1)[0] + "q={query}"
    return templates


def _hyperlinks(path: Path) -> dict[str, list[tuple[str, str]]]:
    result = {}
    with ZipFile(path) as package:
        for name in package.namelist():
            if not name.startswith("ppt/diagrams/data") or not name.endswith(".xml"):
                continue
            rel_name = name.replace("/data", "/_rels/data") + ".rels"
            if rel_name not in package.namelist():
                continue
            rels = etree.fromstring(package.read(rel_name))
            targets = {item.get("Id"): item.get("Target") for item in rels}
            root = etree.fromstring(package.read(name))
            links = []
            for paragraph in root.xpath(".//a:p", namespaces=NS):
                current_id, current_text = None, []
                for run in paragraph.xpath("./a:r", namespaces=NS):
                    ids = run.xpath("./a:rPr/a:hlinkClick/@r:id", namespaces=NS)
                    rid = ids[0] if ids else None
                    text = "".join(run.xpath("./a:t/text()", namespaces=NS))
                    if current_id is not None and rid != current_id:
                        links.append(("".join(current_text), targets[current_id]))
                        current_text = []
                    current_id = rid
                    if rid:
                        current_text.append(text)
                if current_id:
                    links.append(("".join(current_text), targets[current_id]))
            if links:
                result[name] = links
    return result


def _content(path: Path) -> dict[str, tuple[str, str]]:
    result = {}
    with ZipFile(path) as package:
        for name in package.namelist():
            if name.startswith("ppt/diagrams/") and name.endswith(".xml"):
                root = etree.fromstring(package.read(name))
                text, bold = [], []
                for run in root.xpath(".//a:r", namespaces=NS):
                    value = "".join(run.xpath("./a:t/text()", namespaces=NS))
                    text.append(value)
                    rpr = run.find(f"{{{A_NS}}}rPr")
                    if rpr is not None and rpr.get("b") in {"1", "true"}:
                        bold.append(value)
                result[name] = ("".join(text), "".join(bold))
    return result


def test_case_01_matches_golden_hyperlink_semantics(tmp_path: Path) -> None:
    output = tmp_path / "output.pptx"
    assert link_presentation(CASE / "input.pptx", output, _golden_templates()) == 6
    actual = _hyperlinks(output)
    golden = _hyperlinks(CASE / "expected.pptx")

    # The manually edited deck is authoritative for link targets. PowerPoint's
    # selection UI changed three opening quote glyphs, omitted one closing quote,
    # and captured one trailing space, so exact ranges follow the written rule.
    assert {name: [url for _, url in links] for name, links in actual.items()} == {
        name: [url for _, url in links] for name, links in golden.items()
    }
    assert [text for links in actual.values() for text, _ in links] == [
        "“The gospel is the announcement or proclamation of Jesus as the long-awaited Messiah of Israel’s hope who through his life, death, burial, resurrection, and ascension conquers sin and death—personal, systemic—in order to unleash the redemption of God—that is, the kingdom of God, for the transformation of humans and systems.”(28)",
        "“The good news is summarized well”",
        "“the person and work of Christ as prophet, priest , and king, a catholic ecclesiology, and the future consummation of the kingdom.”(63)",
        "“By God’s grace from start to finish, those who trust in Jesus Christ are saved from the guilt and penalty of their former sins, renewed in God’s image, and empowered by the Holy Spirit to love and serve God and neighbor in true holiness all the remaining days of their lives.”(108)",
        "“My understanding of the gospel is the good news of Jesus Christ to nonbelievers who are in darkness. I begin with my personal experience of the salvation of Christ, which will help me explain the gospel in a more tangible way”",
        "“A Liberation Gospel perspective testifies to God’s present voice and actions in the lives of those on the margins. Born in poverty yet free, Jesus died the death of an enslaved person. A liberation perspective of the gospel marks Jesus’s social status as one among the oppressed and reads with other ‘non-persons.’”",
    ]


def test_processing_preserves_text_and_bold_formatting(tmp_path: Path) -> None:
    output = tmp_path / "output.pptx"
    link_presentation(CASE / "input.pptx", output, _golden_templates())
    assert _content(output) == _content(CASE / "input.pptx")


def test_reprocessing_is_byte_for_byte_idempotent(tmp_path: Path) -> None:
    first, second = tmp_path / "first.pptx", tmp_path / "second.pptx"
    templates = _golden_templates()
    assert link_presentation(CASE / "input.pptx", first, templates) == 6
    assert link_presentation(first, second, templates) == 0
    assert sha256(first.read_bytes()).digest() == sha256(second.read_bytes()).digest()
