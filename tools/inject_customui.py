#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
inject_customui.py  -  chen Ribbon (customUI) vao file .xlam / .xlsm da build.

Excel chi nap Ribbon khi trong goi OOXML co part customUI/customUI.xml
va mot Relationship tro toi part do trong _rels/.rels.

Neu customUI.xml tham chieu image="icoX" (icon rieng),
can nhung them cac file PNG trong customUI/images/ + customUI.xml.rels.

Cach dung:
    python inject_customui.py duong_dan.xlam customUI\\customUI.xml

Ghi de truc tiep len file .xlam (tao ban .tmp roi thay the).
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
import sys
import zipfile

CT_PART = "[Content_Types].xml"
RELS_PART = "_rels/.rels"
UI_PART = "customUI/customUI.xml"
UI_RELS_PART = "customUI/customUI.xml.rels"
REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
CT_TYPE = "application/xml"
IMG_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
# customUI.xml declares <image id="..." r:id="rIcN" /> -> rels Id must be "rIcN"
IMG_ID_RE = re.compile(r'<image\s+id="([^"]+)"\s+r:id="([^"]+)"')


def _read_all(path: str) -> tuple[list[str], dict[str, bytes]]:
    with zipfile.ZipFile(path, "r") as z:
        names = z.namelist()
        data = {n: z.read(n) for n in names}
    return names, data


def _ensure_rel(rels_xml: str) -> tuple[str, str]:
    if REL_TYPE in rels_xml:
        rid = re.search(r'Id="([^"]+)"[^>]*' + re.escape(REL_TYPE), rels_xml)
        return rels_xml, (rid.group(1) if rid else "(da co)")
    ids = set(re.findall(r'Id="([^"]+)"', rels_xml))
    rid = "rIdNXCustomUI"
    i = 1
    while rid in ids:
        rid = "rIdNXCustomUI%d" % i
        i += 1
    rel = '<Relationship Id="%s" Type="%s" Target="%s"/>' % (rid, REL_TYPE, UI_PART)
    if "</Relationships>" not in rels_xml:
        raise SystemExit("LOI: _rels/.rels khong co the </Relationships>")
    return rels_xml.replace("</Relationships>", rel + "</Relationships>", 1), rid


def _ensure_ct(ct_xml: str, images: dict[str, bytes]) -> tuple[str, bool]:
    added = False
    # ensure Default for png (for embedded ribbon images)
    if images and not re.search(r'<Default\s+Extension="png"', ct_xml, re.I):
        d = '<Default Extension="png" ContentType="image/png"/>'
        if "</Types>" not in ct_xml:
            raise SystemExit("LOI: [Content_Types].xml khong co the </Types>")
        ct_xml = ct_xml.replace("</Types>", d + "</Types>", 1)
        added = True
    if re.search(r'PartName="/customUI/customUI\.xml"', ct_xml, re.I):
        return ct_xml, added
    if re.search(r'<Default\s+Extension="xml"', ct_xml, re.I):
        # Da co Default cho duoi .xml -> khong can Override
        return ct_xml, added
    ov = '<Override PartName="/%s" ContentType="%s"/>' % (UI_PART, CT_TYPE)
    if "</Types>" not in ct_xml:
        raise SystemExit("LOI: [Content_Types].xml khong co the </Types>")
    return ct_xml.replace("</Types>", ov + "</Types>", 1), True


def inject(xlam: str, ui_file: str) -> int:
    if not os.path.isfile(xlam):
        raise SystemExit("LOI: khong thay file %s" % xlam)
    if not os.path.isfile(ui_file):
        raise SystemExit("LOI: khong thay file %s" % ui_file)

    with open(ui_file, "rb") as fh:
        ui_bytes = fh.read()
    ui_text = ui_bytes.decode("utf-8-sig", "replace")
    if "customUI" not in ui_text:
        raise SystemExit("LOI: %s khong phai customUI.xml" % ui_file)

    # optional: collect PNG images next to customUI.xml in images/ subfolder
    img_dir = os.path.join(os.path.dirname(os.path.abspath(ui_file)), "images")
    images: dict[str, bytes] = {}
    if os.path.isdir(img_dir):
        for p in sorted(glob.glob(os.path.join(img_dir, "*.png"))):
            key = "customUI/images/" + os.path.basename(p)
            with open(p, "rb") as fh:
                images[key] = fh.read()
    has_images = len(images) > 0

    names, data = _read_all(xlam)
    if CT_PART not in data:
        raise SystemExit("LOI: %s khong phai goi OOXML" % xlam)
    if RELS_PART not in data:
        raise SystemExit("LOI: thieu %s" % RELS_PART)

    ct_xml = data[CT_PART].decode("utf-8-sig", "replace")
    rels_xml = data[RELS_PART].decode("utf-8-sig", "replace")

    ct_xml, ct_added = _ensure_ct(ct_xml, images)
    rels_xml, rid = _ensure_rel(rels_xml)

    # build customUI rels for embedded images: Id = r:id declared in <image ... r:id="rIcN" />
    ui_rels_xml = None
    if has_images:
        pairs = [(m.group(1), m.group(2)) for m in IMG_ID_RE.finditer(ui_text)]
        by_base = {os.path.splitext(os.path.basename(k))[0]: k for k in images.keys()}
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
        seen_rids: set[str] = set()
        for img_id, rid_img in pairs:
            part = by_base.get(img_id)
            if part is None:
                raise SystemExit("LOI: customUI.xml khai image id=%s nhung thieu file images/%s.png" % (img_id, img_id))
            target = "images/" + os.path.basename(part)
            parts.append('<Relationship Id="%s" Type="%s" Target="%s"/>' % (rid_img, IMG_REL_TYPE, target))
            seen_rids.add(rid_img)
        # images not declared in <image> (fallback: use basename as Id)
        for part in sorted(images.keys()):
            base = os.path.splitext(os.path.basename(part))[0]
            if base not in [p[0] for p in pairs]:
                # skip unreferenced images; only referenced ones go to rels
                continue
        parts.append('</Relationships>')
        ui_rels_xml = ("\n".join(parts) + "\n").encode("utf-8")

    tmp = xlam + ".tmp"
    if os.path.exists(tmp):
        os.remove(tmp)

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(CT_PART, ct_xml.encode("utf-8"))
        for n in names:
            if n in (CT_PART, RELS_PART, UI_PART, UI_RELS_PART):
                continue
            if n.startswith("customUI/images/"):
                continue
            if n.endswith("/"):
                continue
            z.writestr(n, data[n])
        z.writestr(RELS_PART, rels_xml.encode("utf-8"))
        z.writestr(UI_PART, ui_bytes)
        if has_images:
            for part, blob in sorted(images.items()):
                z.writestr(part, blob)
            z.writestr(UI_RELS_PART, ui_rels_xml)

    # ---- kiem chung lai goi vua ghi
    with zipfile.ZipFile(tmp, "r") as z:
        bad = z.testzip()
        if bad:
            raise SystemExit("LOI: zip hong o %s" % bad)
        got = z.namelist()
        if UI_PART not in got:
            raise SystemExit("LOI: khong ghi duoc %s" % UI_PART)
        if REL_TYPE not in z.read(RELS_PART).decode("utf-8", "replace"):
            raise SystemExit("LOI: thieu relationship customUI")
        if has_images:
            if UI_RELS_PART not in got:
                raise SystemExit("LOI: thieu %s" % UI_RELS_PART)
            rels_out = z.read(UI_RELS_PART).decode("utf-8", "replace")
            if "rIc1" not in rels_out:
                raise SystemExit("LOI: thieu relationship image trong %s" % UI_RELS_PART)

    shutil.move(tmp, xlam)

    print("OK  customUI  -> %s" % os.path.basename(xlam))
    print("    part       : %s (%d bytes)" % (UI_PART, len(ui_bytes)))
    print("    relationship: %s" % rid)
    print("    content-type: %s" % ("them Override" if ct_added else "dung Default xml"))
    if has_images:
        print("    images     : %d file -> customUI/images/ + %s" % (len(images), UI_RELS_PART))
    return 0


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Chen Ribbon customUI vao .xlam/.xlsm")
    ap.add_argument("xlam", help="file .xlam/.xlsm can chen")
    ap.add_argument("customui", help="file customUI.xml")
    a = ap.parse_args(argv)
    return inject(a.xlam, a.customui)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
