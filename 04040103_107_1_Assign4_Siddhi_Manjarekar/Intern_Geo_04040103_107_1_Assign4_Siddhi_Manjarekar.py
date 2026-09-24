import math
import random
from pathlib import Path
import xlsxwriter

VERTEX_TRIPLES = [
    ('X', 'Y', 'Z'), ('P', 'Q', 'R'), ('A', 'B', 'C'), ('L', 'M', 'N'),
    ('D', 'E', 'F'), ('M', 'N', 'O'), ('R', 'S', 'T'), ('U', 'V', 'W'),
    ('E', 'F', 'G'), ('S', 'T', 'U'),
]


def is_valid_triangle(a, b, c):
    return (a + b > c) and (b + c > a) and (a + c > b)


def make_triangle_sides(used):
    for _ in range(1000):
        a = random.randint(4, 12)
        b = random.randint(4, 12)
        c = random.randint(4, 12)
        if not is_valid_triangle(a, b, c):
            continue
        key = tuple(sorted([a, b, c]))
        if key in used:
            continue
        return a, b, c, key
    return None


def third_vertex(a, b, c):
    # v0 = (0, 0), v1 = (a, 0); v0v2 = c, v1v2 = b
    x = (c ** 2 - b ** 2 + a ** 2) / (2 * a)
    y = math.sqrt(max(c ** 2 - x ** 2, 0))
    return x, y


def rotate_pt(x, y, angle):
    ca, sa = math.cos(angle), math.sin(angle)
    return x * ca - y * sa, x * sa + y * ca


def make_rotated_triangle(a, x2, y2):
    """Rotate (and sometimes mirror) the triangle (0,0), (a,0), (x2,y2)
    by a random angle so the construction isn't always drawn with a
    horizontal base. Returns the three transformed points."""
    angle = random.uniform(0, 2 * math.pi)
    mirror = random.random() < 0.5

    raw_pts = [(0.0, 0.0), (a, 0.0), (x2, y2)]
    if mirror:
        raw_pts = [(-px, py) for px, py in raw_pts]

    return [rotate_pt(px, py, angle) for px, py in raw_pts]


def correct_option_text(v0, v1, v2):
    en = (f"Point of intersection of the arcs is the required point ${v2}$. "
          f"Join this point with ${v0}$ and ${v1}$, to get the required "
          f"$\\triangle {v0}{v1}{v2}$.<br>")
    mr = (f"कंसांचा छेदबिंदू हा हवा असलेला बिंदू ${v2}$ आहे. हा बिंदू ${v0}$ "
          f"आणि ${v1}$ यांच्याशी जोडून आवश्यक $\\triangle {v0}{v1}{v2}$ "
          f"मिळेल.<br>")
    return en, mr


def wrong_option_pool(v0, v1, v2):
    return [
        (f"Draw the perpendicular bisector of segment ${v0}{v1}$ to locate "
         f"the point ${v2}$.<br>",
         f"बिंदू ${v2}$ शोधण्यासाठी रेषाखंड ${v0}{v1}$ चा लंबदुभाजक काढा.<br>"),
        (f"Draw the bisector of $\\angle {v1}{v0}{v2}$ at point ${v0}$ to "
         f"locate the point ${v2}$.<br>",
         f"बिंदू ${v2}$ शोधण्यासाठी बिंदू ${v0}$ येथे $\\angle {v1}{v0}{v2}$ "
         f"चा दुभाजक काढा.<br>"),
        (f"Join the point of intersection with ${v0}$ only, to get the "
         f"required $\\triangle {v0}{v1}{v2}$.<br>",
         f"आवश्यक $\\triangle {v0}{v1}{v2}$ मिळवण्यासाठी छेदबिंदू फक्त "
         f"${v0}$ शी जोडा.<br>"),
        (f"Draw one more arc of any radius from point ${v2}$ to verify the "
         f"triangle.<br>",
         f"त्रिकोणाची अचूकता ठरवण्यासाठी बिंदू ${v2}$ पासून कोणत्याही "
         f"त्रिज्येचा आणखी एक कंस काढा.<br>"),
    ]


def build_question(v0, v1, v2, a, b, c):
    eng_stem = (
        f"We are to draw a $\\triangle {v0}{v1}{v2}$ whose sides have "
        f"lengths as $l({v0}{v1}) = {a}$ cm, $l({v1}{v2}) = {b}$ cm, "
        f"$l({v0}{v2}) = {c}$ cm.<br>\n"
        f"Intersecting arcs of required value are drawn from points ${v0}$ "
        f"and ${v1}$, on same side of ${v0}{v1}$, as shown in the figure "
        f"below.<br>\n"
        f"What will be the next step for this construction?<br>\n"
    )
    mr_stem = (
        f"आपल्याला एक $\\triangle {v0}{v1}{v2}$ असा काढायचा आहे, की ज्याच्या "
        f"भुजांची मापे, $l({v0}{v1}) = {a}$ सेमी, $l({v1}{v2}) = {b}$ सेमी, "
        f"$l({v0}{v2}) = {c}$ सेमी अशी असतील.<br>\n"
        f"${v0}$ आणि ${v1}$ पासून भुजा ${v0}{v1}$ च्या एकाच बाजूला असणारे आणि "
        f"एकमेकांना छेदणारे दोन कंस, खालील आकृतीत दाखवल्याप्रमाणे काढले.<br>\n"
        f"या रचनेचा पुढील टप्पा कोणता असेल?<br>\n"
    )
    # Next-step statements are already listed as the four answer options
    # (in the separate answer columns), so they are NOT repeated here.
    return eng_stem + "#" + mr_stem


def build_solution(correct_letter, v0, v1, v2):
    restated = (
        f"Point of intersection of the arcs is the required point ${v2}$. "
        f"Join line segment ${v0}{v2}$ and ${v1}{v2}$ and get the required "
        f"$\\triangle {v0}{v1}{v2}$."
    )
    restated_mr = (
        f"कंसांचा छेदबिंदू हा हवा असलेला बिंदू ${v2}$ आहे. रेषाखंड "
        f"${v0}{v2}$ आणि ${v1}{v2}$ जोडा म्हणजे आवश्यक "
        f"$\\triangle {v0}{v1}{v2}$ मिळेल."
    )

    en = (
        f"Ans: ${correct_letter}$ - {restated}<br>"
        f"Point of intersection of the arcs which are already drawn is the "
        f"required third vertex of this triangle.<br>"
        f"Label this point as point ${v2}$.<br>"
        f"Join line segment ${v0}{v2}$ and ${v1}{v2}$.<br>"
        f"$\\triangle {v0}{v1}{v2}$ is the required triangle.<br>"
        f"Such step is mentioned only in option ${correct_letter}$.<br>"
        f"Steps mentioned in other options do not give the required "
        f"triangle.<br>"
        f"Therefore, option ${correct_letter}$, <b>{restated}</b> is the "
        f"answer.<br>"
    )
    mr = (
        f"#उत्तर: ${correct_letter}$ : {restated_mr}<br>"
        f"एकमेकांना छेदणाऱ्या कंसांचा छेदबिंदू हा आपल्याला हवा असलेला "
        f"तिसरा शिरोबिंदू आहे.<br>"
        f"या बिंदूला ${v2}$ असे नाव देऊन रेषाखंड ${v0}{v2}$ आणि ${v1}{v2}$ "
        f"जोडा.<br>"
        f"$\\triangle {v0}{v1}{v2}$ हा आपल्याला हवा असलेला त्रिकोण आहे.<br>"
        f"हा टप्पा फक्त पर्याय ${correct_letter}$ मध्ये दिला आहे.<br>"
        f"इतर पर्यायांत दिलेले टप्पे वापरून आवश्यक त्रिकोण मिळत नाही."
        f"<br>"
        f"म्हणून, पर्याय ${correct_letter}$ - <b>{restated_mr}</b> हे "
        f"उत्तर आहे.<br>"
    )
    return en, mr


SVG_W, SVG_H, SVG_MARGIN = 340, 260, 38


def _svg_transform(points):
    """Generic bounding-box transform for a set of math-space points ->
    svg-space, so it works no matter how the triangle is rotated."""
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)

    span_x = max(max_x - min_x, 1e-6)
    span_y = max(max_y - min_y, 1e-6)

    avail_w = SVG_W - 2 * SVG_MARGIN
    avail_h = SVG_H - 2 * SVG_MARGIN

    scale = min(avail_w / span_x, avail_h / span_y)
    scale = max(min(scale, 18), 5)

    pad_x = (avail_w - span_x * scale) / 2
    pad_y = (avail_h - span_y * scale) / 2

    def to_svg(px, py):
        sx = SVG_MARGIN + pad_x + (px - min_x) * scale
        # flip y: math-space "up" should render toward the top of the svg
        sy = SVG_MARGIN + pad_y + (max_y - py) * scale
        return sx, sy

    return to_svg, scale


def _arc_polyline_points(cx, cy, r, theta1_deg, theta2_deg, to_svg, steps=20):
    pts = []
    step_size = (theta2_deg - theta1_deg) / steps
    for i in range(steps + 1):
        t = math.radians(theta1_deg + step_size * i)
        mx, my = cx + r * math.cos(t), cy + r * math.sin(t)
        pts.append(to_svg(mx, my))
    return " ".join(f"{px:.1f},{py:.1f}" for px, py in pts)


def _label_offset(vx, vy, cx, cy, dist=18):
    """Push a vertex label away from the triangle's centroid (in svg
    space), so labels stay clear of the figure at any rotation."""
    dx, dy = vx - cx, vy - cy
    norm = math.hypot(dx, dy) or 1.0
    return dx / norm * dist, dy / norm * dist


def get_arc_geometry(P0, P1, P2, b, c):
    """
    Compute intersection angles and robust outward directions for both arcs
    so labels diverge into open space away from the other arc and vertex P2.
    """
    ang0 = math.atan2(P2[1] - P0[1], P2[0] - P0[0])
    ang1 = math.atan2(P2[1] - P1[1], P2[0] - P1[0])

    delta = 0.05
    q_pos = (P0[0] + c * math.cos(ang0 + delta), P0[1] + c * math.sin(ang0 + delta))
    q_neg = (P0[0] + c * math.cos(ang0 - delta), P0[1] + c * math.sin(ang0 - delta))
    dir0 = 1.0 if math.hypot(q_pos[0] - P1[0], q_pos[1] - P1[1]) > math.hypot(q_neg[0] - P1[0], q_neg[1] - P1[1]) else -1.0

    r_pos = (P1[0] + b * math.cos(ang1 + delta), P1[1] + b * math.sin(ang1 + delta))
    r_neg = (P1[0] + b * math.cos(ang1 - delta), P1[1] + b * math.sin(ang1 - delta))
    dir1 = 1.0 if math.hypot(r_pos[0] - P0[0], r_pos[1] - P0[1]) > math.hypot(r_neg[0] - P0[0], r_neg[1] - P0[1]) else -1.0

    return ang0, dir0, ang1, dir1


def get_construction_bbox_pts(P0, P1, P2, b, c, ang0, dir0, ang1, dir1):
    """Include extended arc sweeps and radius label offsets in bounding box."""
    pts = [P0, P1, P2]
    for factor in [-0.25, 0.6]:
        pts.append((P0[0] + (c + 1.2) * math.cos(ang0 + dir0 * factor), P0[1] + (c + 1.2) * math.sin(ang0 + dir0 * factor)))
        pts.append((P1[0] + (b + 1.2) * math.cos(ang1 + dir1 * factor), P1[1] + (b + 1.2) * math.sin(ang1 + dir1 * factor)))
    return pts


def _arc_and_label_svg(P0, P1, P2, b, c, to_svg, scale, ang0, dir0, ang1, dir1):
    """
    Draw construction arcs and outward, non-overlapping radius labels.
    """
    sep_deg0 = max(22.0, min(45.0, math.degrees(38.0 / (c * scale))))
    sep_deg1 = max(22.0, min(45.0, math.degrees(38.0 / (b * scale))))

    def calc_lbls(s0, s1):
        lbl_ang0 = ang0 + dir0 * math.radians(s0)
        lbl_ang1 = ang1 + dir1 * math.radians(s1)
        p_arc0 = to_svg(P0[0] + c * math.cos(lbl_ang0), P0[1] + c * math.sin(lbl_ang0))
        p_arc1 = to_svg(P1[0] + b * math.cos(lbl_ang1), P1[1] + b * math.sin(lbl_ang1))
        sp0 = to_svg(*P0)
        sp1 = to_svg(*P1)
        u0 = (p_arc0[0] - sp0[0], p_arc0[1] - sp0[1])
        len0 = math.hypot(*u0) or 1.0
        lbl0 = (p_arc0[0] + 16 * u0[0] / len0, p_arc0[1] + 16 * u0[1] / len0)

        u1 = (p_arc1[0] - sp1[0], p_arc1[1] - sp1[1])
        len1 = math.hypot(*u1) or 1.0
        lbl1 = (p_arc1[0] + 16 * u1[0] / len1, p_arc1[1] + 16 * u1[1] / len1)
        return lbl0, lbl1, s0, s1

    lbl0, lbl1, s0, s1 = calc_lbls(sep_deg0, sep_deg1)
    if math.hypot(lbl0[0] - lbl1[0], lbl0[1] - lbl1[1]) < 36:
        s0 += 12.0
        s1 += 12.0
        lbl0, lbl1, s0, s1 = calc_lbls(s0, s1)

    arc0_start = math.degrees(ang0) - dir0 * 10.0
    arc0_end = math.degrees(ang0) + dir0 * (s0 + 10.0)
    arc0_pts = _arc_polyline_points(P0[0], P0[1], c, arc0_start, arc0_end, to_svg)

    arc1_start = math.degrees(ang1) - dir1 * 10.0
    arc1_end = math.degrees(ang1) + dir1 * (s1 + 10.0)
    arc1_pts = _arc_polyline_points(P1[0], P1[1], b, arc1_start, arc1_end, to_svg)

    def clamp(x, y):
        return (
            max(18, min(x, SVG_W - 18)),
            max(18, min(y, SVG_H - 12))
        )

    lbl0 = clamp(*lbl0)
    lbl1 = clamp(*lbl1)

    return (
        f"<polyline points='{arc0_pts}' "
        f"fill='none' "
        f"stroke='black' "
        f"stroke-width='1.2'/>"

        f"<polyline points='{arc1_pts}' "
        f"fill='none' "
        f"stroke='black' "
        f"stroke-width='1.2'/>"

        f"<text "
        f"x='{lbl0[0]:.1f}' "
        f"y='{lbl0[1]:.1f}' "
        f"text-anchor='middle' "
        f"dominant-baseline='middle' "
        f"font-size='11' "
        f"fill='black' "
        f"paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' "
        f"font-weight='bold'>"
        f"{c} cm"
        f"</text>"

        f"<text "
        f"x='{lbl1[0]:.1f}' "
        f"y='{lbl1[1]:.1f}' "
        f"text-anchor='middle' "
        f"dominant-baseline='middle' "
        f"font-size='11' "
        f"fill='black' "
        f"paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' "
        f"font-weight='bold'>"
        f"{b} cm"
        f"</text>"
    )


def build_question_svg(v0, v1, v2, b, c, P0, P1, P2):
    ang0, dir0, ang1, dir1 = get_arc_geometry(P0, P1, P2, b, c)
    bbox_pts = get_construction_bbox_pts(P0, P1, P2, b, c, ang0, dir0, ang1, dir1)
    to_svg, scale = _svg_transform(bbox_pts)

    p0 = to_svg(*P0)
    p1 = to_svg(*P1)
    pi = to_svg(*P2)

    arcs_svg = _arc_and_label_svg(P0, P1, P2, b, c, to_svg, scale, ang0, dir0, ang1, dir1)

    cx = (p0[0] + p1[0] + pi[0]) / 3
    cy = (p0[1] + p1[1] + pi[1]) / 3
    l0x, l0y = _label_offset(p0[0], p0[1], cx, cy, dist=18)
    l1x, l1y = _label_offset(p1[0], p1[1], cx, cy, dist=18)

    def clamp(x, low, high):
        return max(low, min(x, high))

    tx0 = clamp(p0[0] + l0x, 14, SVG_W - 14)
    ty0 = clamp(p0[1] + l0y, 14, SVG_H - 14)

    tx1 = clamp(p1[0] + l1x, 14, SVG_W - 14)
    ty1 = clamp(p1[1] + l1y, 14, SVG_H - 14)

    return (
        f"<svg xmlns='http://www.w3.org/2000/svg' "
        f"width='{SVG_W}' height='{SVG_H}' "
        f"viewBox='0 0 {SVG_W} {SVG_H}' "
        f"font-family='Arial' font-size='15'>"
        f"<line x1='{p0[0]:.1f}' y1='{p0[1]:.1f}' "
        f"x2='{p1[0]:.1f}' y2='{p1[1]:.1f}' "
        f"stroke='black' stroke-width='1.2'/>"
        f"{arcs_svg}"
        f"<circle cx='{pi[0]:.1f}' cy='{pi[1]:.1f}' r='1.8' fill='black'/>"
        f"<circle cx='{p0[0]:.1f}' cy='{p0[1]:.1f}' r='1.8' fill='black'/>"
        f"<circle cx='{p1[0]:.1f}' cy='{p1[1]:.1f}' r='1.8' fill='black'/>"
        f"<text x='{tx0:.1f}' y='{ty0:.1f}' fill='black' text-anchor='middle' dominant-baseline='middle' paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' font-weight='bold'>{v0}</text>"
        f"<text x='{tx1:.1f}' y='{ty1:.1f}' fill='black' text-anchor='middle' dominant-baseline='middle' paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' font-weight='bold'>{v1}</text>"
        f"</svg>"
    )


def build_solution_svg(v0, v1, v2, b, c, P0, P1, P2):
    ang0, dir0, ang1, dir1 = get_arc_geometry(P0, P1, P2, b, c)
    bbox_pts = get_construction_bbox_pts(P0, P1, P2, b, c, ang0, dir0, ang1, dir1)
    to_svg, scale = _svg_transform(bbox_pts)

    p0 = to_svg(*P0)
    p1 = to_svg(*P1)
    p2 = to_svg(*P2)

    arcs_svg = _arc_and_label_svg(P0, P1, P2, b, c, to_svg, scale, ang0, dir0, ang1, dir1)

    cx = (p0[0] + p1[0] + p2[0]) / 3
    cy = (p0[1] + p1[1] + p2[1]) / 3
    l0x, l0y = _label_offset(p0[0], p0[1], cx, cy)
    l1x, l1y = _label_offset(p1[0], p1[1], cx, cy)
    l2x, l2y = _label_offset(p2[0], p2[1], cx, cy)

    def clamp(x, low, high):
        return max(low, min(x, high))

    tx0 = clamp(p0[0] + l0x, 14, SVG_W - 14)
    ty0 = clamp(p0[1] + l0y, 14, SVG_H - 14)

    tx1 = clamp(p1[0] + l1x, 14, SVG_W - 14)
    ty1 = clamp(p1[1] + l1y, 14, SVG_H - 14)

    tx2 = clamp(p2[0] + l2x, 14, SVG_W - 14)
    ty2 = clamp(p2[1] + l2y, 14, SVG_H - 14)

    return (
        f"<svg xmlns='http://www.w3.org/2000/svg' "
        f"width='{SVG_W}' height='{SVG_H}' "
        f"viewBox='0 0 {SVG_W} {SVG_H}' "
        f"font-family='Arial' font-size='15'>"

        f"<line x1='{p0[0]:.1f}' y1='{p0[1]:.1f}' "
        f"x2='{p1[0]:.1f}' y2='{p1[1]:.1f}' "
        f"stroke='black' stroke-width='1.2'/>"

        f"<line x1='{p0[0]:.1f}' y1='{p0[1]:.1f}' "
        f"x2='{p2[0]:.1f}' y2='{p2[1]:.1f}' "
        f"stroke='black' stroke-width='1.2'/>"

        f"<line x1='{p1[0]:.1f}' y1='{p1[1]:.1f}' "
        f"x2='{p2[0]:.1f}' y2='{p2[1]:.1f}' "
        f"stroke='black' stroke-width='1.2'/>"

        f"{arcs_svg}"

        f"<circle cx='{p0[0]:.1f}' cy='{p0[1]:.1f}' r='1.8' fill='black'/>"
        f"<circle cx='{p1[0]:.1f}' cy='{p1[1]:.1f}' r='1.8' fill='black'/>"
        f"<circle cx='{p2[0]:.1f}' cy='{p2[1]:.1f}' r='1.8' fill='black'/>"

        f"<text x='{tx0:.1f}' y='{ty0:.1f}' fill='black' text-anchor='middle' dominant-baseline='middle' paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' font-weight='bold'>{v0}</text>"
        f"<text x='{tx1:.1f}' y='{ty1:.1f}' fill='black' text-anchor='middle' dominant-baseline='middle' paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' font-weight='bold'>{v1}</text>"
        f"<text x='{tx2:.1f}' y='{ty2:.1f}' fill='black' text-anchor='middle' dominant-baseline='middle' paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' font-weight='bold'>{v2}</text>"

        f"</svg>"
    )


def create_workbook():
    target_dir = Path(__file__).parent
    filename = target_dir / "Intern_Geo_04040103_107_1_Assign4_Siddhi_Manjarekar.xlsx"

    headers = [
        "Sr. No", "Question Type", "Answer Type", "Topic Number",
        "Question (Text Only)", "Correct Answer 1", "Correct Answer 2",
        "Correct Answer 3", "Correct Answer 4", "Wrong Answer 1",
        "Wrong Answer 2", "Wrong Answer 3", "Time in seconds",
        "Difficulty Level", "Question (Image/ Audio/ Video)",
        "Contributor's Registered mailId", "Solution (Text Only)",
        "Solution (Image/ Audio/ Video)", "Variation Number"
    ]

    used_sides = set()
    used_combos = set()
    rows_written = 0
    TOTAL_ROWS = 50
    rows = []

    while rows_written < TOTAL_ROWS:
        v0, v1, v2 = random.choice(VERTEX_TRIPLES)
        sides = make_triangle_sides(used_sides)
        if sides is None:
            used_sides.clear()
            continue
        a, b, c, key = sides

        combo = (v0, v1, v2, key)
        if combo in used_combos:
            continue

        wrongs = random.sample(wrong_option_pool(v0, v1, v2), 3)
        correct = correct_option_text(v0, v1, v2)

        all_opts = [correct] + wrongs
        random.shuffle(all_opts)
        letters = ['A', 'B', 'C', 'D']
        options = [(letters[i], all_opts[i][0], all_opts[i][1]) for i in range(4)]

        correct_letter = next(lbl for lbl, en, mr in options if en == correct[0])
        wrong_entries = [(lbl, en, mr) for lbl, en, mr in options if en != correct[0]]

        q_text = build_question(v0, v1, v2, a, b, c)
        sol_en, sol_mr = build_solution(correct_letter, v0, v1, v2)

        used_sides.add(key)
        used_combos.add(combo)

        row_num = rows_written + 2

        x2, y2 = third_vertex(a, b, c)
        P0, P1, P2 = make_rotated_triangle(a, x2, y2)

        q_svg = build_question_svg(v0, v1, v2, b, c, P0, P1, P2)
        s_svg = build_solution_svg(v0, v1, v2, b, c, P0, P1, P2)

        q_text_full = q_text + "\n" + q_svg + "<br>"
        solution_full = (
            sol_en +
            "\n" + s_svg + "<br>" +
            "\n" + sol_mr +
            "\n" + s_svg + "<br>"
        )
        correct_full = f"${correct_letter}$. {correct[0]}#${correct_letter}$. {correct[1]}"
        wrong_full = [f"${lbl}$. {en}#${lbl}$. {mr}" for lbl, en, mr in wrong_entries]

        row_data = [
            rows_written + 1, "Text", 1, "04040103", q_text_full,
            correct_full, None, None, None,
            wrong_full[0], wrong_full[1], wrong_full[2],
            120, 1, None, "siddhimanjarekar274@gmail.com",
            solution_full, None, 107
        ]
        rows.append(row_data)
        rows_written += 1

    marker_row = ["****"] + [None] * 2 + ["****"] + [None] * (len(headers) - 4)
    rows.append(marker_row)

    wb = xlsxwriter.Workbook(str(filename))

    # 1. Sheet Index 0 (First Tab): Instruction Sheet (completely empty)
    ws_inst = wb.add_worksheet("Instruction")

    # 2. Sheet Index 1 (Second Tab): Questions Sheet (main dataset)
    ws_q = wb.add_worksheet("Questions")

    # Formats matching OpenXML specifications
    fmt_header = wb.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_size': 11,
        'align': 'left',
        'valign': 'top',
        'text_wrap': False
    })

    fmt_general = wb.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'align': 'left',
        'valign': 'top',
        'text_wrap': False
    })

    fmt_wrap = wb.add_format({
        'font_name': 'Calibri',
        'font_size': 11,
        'align': 'left',
        'valign': 'top',
        'text_wrap': True
    })

    # Text format (built-in format 49 / '@') with applyNumberFormat=1
    fmt_topic = wb.add_format({
        'num_format': 49,
        'font_name': 'Calibri',
        'font_size': 11,
        'align': 'left',
        'valign': 'top',
        'text_wrap': False
    })

    # Set row 0 format (Header Row: Bold, Wrapped, Left/Top aligned)
    ws_q.set_row(0, None, fmt_header)

    # Write Header row
    for c_idx, col_name in enumerate(headers):
        ws_q.write(0, c_idx, col_name, fmt_header)

    # Freeze Header Row (Row 1 frozen, so data scrolls under header at A2)
    ws_q.freeze_panes(1, 0)

    # Standard column widths
    widths = {
        'A': 3,  'B': 3,  'C': 3,  'D': 5,  'E': 60, 'F': 30,
        'G': 2,  'H': 2,  'I': 2,  'J': 30, 'K': 30, 'L': 30,
        'M': 5,  'N': 3,  'O': 3,  'P': 5,  'Q': 75, 'R': 3,  'S': 10
    }
    for col_letter, width in widths.items():
        col_idx = ord(col_letter) - ord('A')
        if col_letter == 'D':
            ws_q.set_column(col_idx, col_idx, width, fmt_topic)
        else:
            ws_q.set_column(col_idx, col_idx, width)

    wrap_cols = {4, 5, 9, 10, 11, 16}

    # Write data rows
    for r_idx, r_data in enumerate(rows, start=1):
        for c_idx, val in enumerate(r_data):
            if val is None:
                continue
            if c_idx == 3 and val != "****":
                ws_q.write_string(r_idx, c_idx, str(val), fmt_topic)
            elif isinstance(val, (int, float)):
                ws_q.write_number(r_idx, c_idx, val, fmt_general)
            elif c_idx in wrap_cols:
                ws_q.write_string(r_idx, c_idx, str(val), fmt_wrap)
            else:
                ws_q.write_string(r_idx, c_idx, str(val), fmt_general)

    # Activate Questions sheet so activeTab=1
    ws_q.activate()
    wb.close()
    print(f"Saved: {filename}  ({rows_written} rows)")


if __name__ == "__main__":
    create_workbook()