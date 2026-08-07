"""Estimate text overflow for every text box in a pptx (CJK-aware)."""
import sys, zipfile, re
from defusedxml import minidom

EMU = 914400.0
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def char_w(ch):
    o = ord(ch)
    if 0x3000 <= o <= 0x9FFF or 0xFF00 <= o <= 0xFF60 or 0xFFE0 <= o <= 0xFFE6:
        return 1.0          # fullwidth
    if ch in "0123456789":
        return 0.56
    if ch.isupper():
        return 0.66
    if ch == " ":
        return 0.28
    return 0.52             # lowercase / punctuation


def txt(node):
    return "".join(n.firstChild.nodeValue for n in node.getElementsByTagNameNS(A, "t")
                   if n.firstChild)


def analyze(path):
    z = zipfile.ZipFile(path)
    slides = sorted((n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                    key=lambda n: int(re.findall(r"\d+", n)[0]))
    problems = []
    for idx, name in enumerate(slides, 1):
        doc = minidom.parseString(z.read(name))
        for sp in doc.getElementsByTagNameNS("http://schemas.openxmlformats.org/presentationml/2006/main", "sp"):
            body = sp.getElementsByTagNameNS(A, "bodyPr")
            xfrm = sp.getElementsByTagNameNS(A, "off")
            ext = sp.getElementsByTagNameNS(A, "ext")
            if not body or not xfrm or not ext:
                continue
            w = int(ext[0].getAttribute("cx")) / EMU
            h = int(ext[0].getAttribute("cy")) / EMU
            bp = body[0]
            def ins(attr, dflt):
                v = bp.getAttribute(attr)
                return int(v) / EMU if v else dflt
            lIns, rIns = ins("lIns", 0.1), ins("rIns", 0.1)
            tIns, bIns = ins("tIns", 0.05), ins("bIns", 0.05)
            avail_w = w - lIns - rIns
            avail_h = h - tIns - bIns
            if avail_w <= 0.1:
                continue
            total = 0.0
            any_text = False
            for p in sp.getElementsByTagNameNS(A, "p"):
                t = txt(p)
                szs = [int(r.getAttribute("sz")) / 100.0
                       for r in p.getElementsByTagNameNS(A, "rPr") if r.getAttribute("sz")]
                sz = max(szs) if szs else 18.0
                lnspc = p.getElementsByTagNameNS(A, "lnSpc")
                mult = 1.0
                if lnspc:
                    pct = lnspc[0].getElementsByTagNameNS(A, "spcPct")
                    if pct:
                        mult = int(pct[0].getAttribute("val")) / 100000.0
                spc_after = 0.0
                sa = p.getElementsByTagNameNS(A, "spcAft")
                if sa:
                    pts = sa[0].getElementsByTagNameNS(A, "spcPts")
                    if pts:
                        spc_after = int(pts[0].getAttribute("val")) / 100.0 / 72.0
                bullet = bool(p.getElementsByTagNameNS(A, "buChar") or p.getElementsByTagNameNS(A, "buAutoNum"))
                if not t:
                    total += sz / 72.0 * mult
                    continue
                any_text = True
                em = sz / 72.0
                indent = 0.22 if bullet else 0.0
                line_cap = max(1.0, (avail_w - indent) / em)
                lines = 0
                for seg in t.split("\n"):
                    cur = 0.0
                    n = 1
                    for ch in seg:
                        cw = char_w(ch)
                        if cur + cw > line_cap:
                            n += 1
                            cur = cw
                        else:
                            cur += cw
                    lines += n
                total += lines * em * mult * 1.2 + spc_after
            if any_text and total > avail_h * 1.0:
                problems.append((idx, round(total, 2), round(avail_h, 2),
                                 round(total / avail_h, 2), txt(sp)[:58].replace("\n", "/")))
    return problems


if __name__ == "__main__":
    probs = analyze(sys.argv[1])
    for p in sorted(probs, key=lambda x: -x[3]):
        print(f"slide {p[0]:>3}  need {p[1]}\" / have {p[2]}\"  ratio {p[3]}  | {p[4]}")
    print(f"\n{len(probs)} potential overflow(s)")
