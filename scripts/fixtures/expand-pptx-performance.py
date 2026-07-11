#!/usr/bin/env python3

"""把真实业务 PPTX 扩展到指定页数，用于不提交大文件的性能验收。"""

import argparse
from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--slides", type=int, default=79)
    args = parser.parse_args()
    if args.slides <= 0:
        raise ValueError("页数必须大于零。")

    with ZipFile(args.source) as archive:
        entries = {entry.filename: archive.read(entry.filename) for entry in archive.infolist()}
    presentation = etree.fromstring(entries["ppt/presentation.xml"])
    relationships = etree.fromstring(entries["ppt/_rels/presentation.xml.rels"])
    content_types = etree.fromstring(entries["[Content_Types].xml"])
    slide_list = presentation.find(f"{{{P}}}sldIdLst")
    if slide_list is None or not len(slide_list):
        raise ValueError("源文件没有页面。")
    originals = list(slide_list)
    relation_by_id = {item.get("Id"): item for item in relationships}
    slide_paths = []
    for item in originals:
        relation = relation_by_id[item.get(f"{{{R}}}id")]
        slide_paths.append("ppt/" + relation.get("Target").lstrip("/"))
    max_slide_number = max(int(Path(path).stem.removeprefix("slide")) for path in slide_paths)
    max_slide_id = max(int(item.get("id")) for item in originals)
    relationship_numbers = [
        int(item.get("Id")[3:]) for item in relationships
        if item.get("Id", "").startswith("rId") and item.get("Id")[3:].isdigit()
    ]
    next_relationship = max(relationship_numbers, default=0) + 1

    while len(slide_list) < args.slides:
        source_path = slide_paths[(len(slide_list) - len(originals)) % len(slide_paths)]
        max_slide_number += 1
        target_path = f"ppt/slides/slide{max_slide_number}.xml"
        entries[target_path] = entries[source_path]
        source_rels = f"ppt/slides/_rels/{Path(source_path).name}.rels"
        target_rels = f"ppt/slides/_rels/{Path(target_path).name}.rels"
        if source_rels in entries:
            entries[target_rels] = entries[source_rels]
        relationship_id = f"rId{next_relationship}"
        next_relationship += 1
        relationship = etree.SubElement(relationships, f"{{{PR}}}Relationship")
        relationship.set("Id", relationship_id)
        relationship.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
        relationship.set("Target", f"slides/{Path(target_path).name}")
        max_slide_id += 1
        slide_id = deepcopy(originals[(len(slide_list) - len(originals)) % len(originals)])
        slide_id.set("id", str(max_slide_id))
        slide_id.set(f"{{{R}}}id", relationship_id)
        slide_list.append(slide_id)
        override = etree.SubElement(content_types, f"{{{CT}}}Override")
        override.set("PartName", f"/{target_path}")
        override.set("ContentType", "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")

    entries["ppt/presentation.xml"] = etree.tostring(presentation, xml_declaration=True, encoding="UTF-8", standalone=True)
    entries["ppt/_rels/presentation.xml.rels"] = etree.tostring(relationships, xml_declaration=True, encoding="UTF-8", standalone=True)
    entries["[Content_Types].xml"] = etree.tostring(content_types, xml_declaration=True, encoding="UTF-8", standalone=True)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(args.output, "w", ZIP_DEFLATED) as output:
        for name, data in entries.items():
            output.writestr(name, data)
    print(args.output)


if __name__ == "__main__":
    main()
