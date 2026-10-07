"""Vectorize the FB-JELCZ logo bitmap and export it as DXF, SVG and an extruded STEP solid.

Usage: python3 make_logo.py [height_mm] [thickness_mm]
"""
import sys
from pathlib import Path

import cv2
import numpy as np
import ezdxf
import cadquery as cq

HERE = Path(__file__).parent
SRC = HERE / "fb-jelcz-source.png"
HEIGHT_MM = float(sys.argv[1]) if len(sys.argv) > 1 else 30.0   # height of the lettering
THICK_MM = float(sys.argv[2]) if len(sys.argv) > 2 else 3.0     # extrusion depth
UPSCALE = 12

COLORS = {  # BGR reference colours sampled from the source image
    "blue": (np.array([138, 76, 26]), (0x1A, 0x4C, 0x8A)),
    "red": (np.array([53, 56, 225]), (0xE1, 0x38, 0x35)),
}


def colour_mask(img, bgr):
    dist = np.linalg.norm(img.astype(float) - bgr, axis=2)
    white = np.linalg.norm(img.astype(float) - 255, axis=2)
    # pixel belongs to this colour if it is closer to it than to white (anti-aliasing included)
    weight = np.clip(white / (white + dist + 1e-9), 0, 1)
    other = min((np.linalg.norm(img.astype(float) - c, axis=2) for n, (c, _) in COLORS.items()
                 if not np.array_equal(c, bgr)))
    weight[other < dist] = 0
    return (weight * 255).astype(np.uint8)


def trace(mask):
    big = cv2.resize(mask, None, fx=UPSCALE, fy=UPSCALE, interpolation=cv2.INTER_CUBIC)
    big = cv2.GaussianBlur(big, (0, 0), UPSCALE * 0.35)
    _, bw = cv2.threshold(big, 127, 255, cv2.THRESH_BINARY)
    contours, hier = cv2.findContours(bw, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_NONE)
    shapes = []  # (outer, [holes])
    for i, c in enumerate(contours):
        if hier[0][i][3] != -1 or cv2.contourArea(c) < (UPSCALE * 2) ** 2:
            continue
        holes = []
        j = hier[0][i][2]
        while j != -1:
            if cv2.contourArea(contours[j]) > (UPSCALE * 1.5) ** 2:
                holes.append(contours[j])
            j = hier[0][j][0]
        shapes.append((c, holes))
    return shapes


def simplify(c):
    c = cv2.approxPolyDP(c, UPSCALE * 0.25, True)
    return c.reshape(-1, 2).astype(float)


def main():
    img = cv2.imread(str(SRC))
    layers = {name: trace(colour_mask(img, bgr)) for name, (bgr, _) in COLORS.items()}

    # common scale: lettering bounding box height -> HEIGHT_MM, origin bottom-left, Y up
    pts = np.vstack([simplify(o) for s in layers.values() for o, _ in s])
    xmin, ymin = pts.min(0)
    xmax, ymax = pts.max(0)
    k = HEIGHT_MM / (ymax - ymin)

    def to_mm(p):
        return [((x - xmin) * k, (ymax - y) * k) for x, y in p]

    polys = {name: [(to_mm(simplify(o)), [to_mm(simplify(h)) for h in hs]) for o, hs in shapes]
             for name, shapes in layers.items()}

    # DXF: one layer per colour, closed polylines (import into an Inventor sketch)
    doc = ezdxf.new("R2010", setup=True)
    doc.units = ezdxf.units.MM
    msp = doc.modelspace()
    for name, (_, rgb) in COLORS.items():
        layer = doc.layers.add(name.upper())
        layer.rgb = rgb
        for outer, holes in polys[name]:
            for loop in [outer, *holes]:
                msp.add_lwpolyline(loop, close=True, dxfattribs={"layer": name.upper()})
    doc.saveas(HERE / "fb-jelcz.dxf")

    # SVG preview
    W, H = (xmax - xmin) * k, HEIGHT_MM
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="-2 -2 {W+4:.2f} {H+4:.2f}" '
           f'width="{(W+4)*10:.0f}" height="{(H+4)*10:.0f}">']
    for name, (_, rgb) in COLORS.items():
        d = ""
        for outer, holes in polys[name]:
            for loop in [outer, *holes]:
                d += "M" + "L".join(f"{x:.3f},{H-y:.3f}" for x, y in loop) + "Z"
        out.append(f'<path fill="#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}" fill-rule="evenodd" d="{d}"/>')
    out.append("</svg>")
    (HERE / "fb-jelcz.svg").write_text("\n".join(out))

    # STEP: extruded solids, one compound per colour
    assy = cq.Assembly(name="FB-JELCZ")
    for name, (_, rgb) in COLORS.items():
        solids = []
        for outer, holes in polys[name]:
            wp = cq.Workplane("XY").polyline(outer).close()
            for h in holes:
                wp = wp.polyline(h).close()
            solids.append(wp.extrude(THICK_MM).val())
        body = cq.Compound.makeCompound(solids)
        assy.add(body, name=name, color=cq.Color(*(c / 255 for c in rgb)))
    assy.save(str(HERE / "fb-jelcz.step"))
    print(f"logo size: {W:.1f} x {H:.1f} x {THICK_MM} mm")


if __name__ == "__main__":
    main()
