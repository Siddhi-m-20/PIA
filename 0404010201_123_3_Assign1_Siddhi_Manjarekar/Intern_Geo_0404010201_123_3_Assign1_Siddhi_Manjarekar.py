import random
import math
from pathlib import Path
import xlsxwriter

CONFIG = {
    "contributor_email": "siddhimanjarekar274@gmail.com",
    "topic_number": "0404010201",
    "variation_number": 123,
    "assign_no": 1,
    "your_name": "Siddhi_Manjarekar",
    "total_questions": 50,
    "time_in_seconds": 150,
    "difficulty": 3,
}

COLUMNS = [
    "Sr. No", "Question Type", "Answer Type", "Topic Number",
    "Question (Text Only)", "Correct Answer 1", "Correct Answer 2",
    "Correct Answer 3", "Correct Answer 4", "Wrong Answer 1",
    "Wrong Answer 2", "Wrong Answer 3", "Time in seconds",
    "Difficulty Level", "Question (Image/ Audio/ Video)",
    "Contributor's Registered mailId", "Solution (Text Only)",
    "Solution (Image/ Audio/ Video)", "Variation Number"
]

def get_arc_path(vx, vy, adj1x, adj1y, adj2x, adj2y, r):
    d1 = math.hypot(adj1x - vx, adj1y - vy)
    px = vx + (adj1x - vx) * r / d1
    py = vy + (adj1y - vy) * r / d1
    d2 = math.hypot(adj2x - vx, adj2y - vy)
    qx = vx + (adj2x - vx) * r / d2
    qy = vy + (adj2y - vy) * r / d2
    cross = (px - vx) * (qy - vy) - (py - vy) * (qx - vx)
    sweep = 1 if cross > 0 else 0
    return (f"<path d='M {px:.1f},{py:.1f} A {r},{r} 0 0,{sweep} {qx:.1f},{qy:.1f}'"
            f" fill='none' stroke='black' stroke-width='1.2'/>")


def unit_vec(ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    m = math.hypot(dx, dy)
    if m == 0:
        return 0.0, 0.0
    return dx / m, dy / m


def bisector_unit(vx, vy, a1x, a1y, a2x, a2y):
    u1x, u1y = unit_vec(vx, vy, a1x, a1y)
    u2x, u2y = unit_vec(vx, vy, a2x, a2y)
    bx, by = u1x + u2x, u1y + u2y
    m = math.hypot(bx, by)
    if m == 0:
        return -u1y, u1x
    return bx / m, by / m

def generate_svg(given_angle, orientation, forced_labels):
    canvas = 340
    cx, cy = canvas / 2, canvas / 2
    theta = math.radians(given_angle)

    if given_angle >= 140:
        L = 98
    elif given_angle >= 130:
        L = 105
    elif given_angle >= 115:
        L = 112
    elif given_angle >= 100:
        L = 118
    elif given_angle >= 90:
        L = 124
    elif given_angle <= 25:
        L = 136
    elif given_angle <= 40:
        L = 132
    elif given_angle <= 55:
        L = 130
    elif given_angle <= 70:
        L = 128
    else:
        L = 126

    H = L * math.cos(theta / 2)
    W = 2 * L * math.sin(theta / 2)

    min_H = 50
    if H < min_H:
        scale = min_H / H
        H = min_H
        W *= scale
        L = int(L * scale)
        if W > 270:
            w_scale = 270 / W
            W = 270
            H *= w_scale
            L = int(L * w_scale)

    bx1, by1 = cx,         cy - H * 0.58
    bx2, by2 = cx - W / 2, cy + H * 0.42
    bx3, by3 = cx + W / 2, cy + H * 0.42

    def rotate(x, y, deg):
        rad = math.radians(deg)
        dx, dy = x - cx, y - cy
        return (cx + dx * math.cos(rad) - dy * math.sin(rad),
                cy + dx * math.sin(rad) + dy * math.cos(rad))

    x1, y1 = rotate(bx1, by1, orientation)
    x2, y2 = rotate(bx2, by2, orientation)
    x3, y3 = rotate(bx3, by3, orientation)

    gcx = (x1 + x2 + x3) / 3
    gcy = (y1 + y2 + y3) / 3

    def outward(px, py, dist=24):
        dx, dy = px - gcx, py - gcy
        m = math.hypot(dx, dy)
        if m == 0:
            return px, py
        return px + dx / m * dist, py + dy / m * dist

    labels = forced_labels

    def clamp_pos(px, py, margin=14):
        return (max(margin, min(canvas - margin, px)),
                max(margin, min(canvas - margin, py)))

    lx1, ly1 = clamp_pos(*outward(x1, y1, 25))
    lx2, ly2 = clamp_pos(*outward(x2, y2, 25))
    lx3, ly3 = clamp_pos(*outward(x3, y3, 25))

    if given_angle >= 130:
        arc_r = 13.0
    elif given_angle >= 110:
        arc_r = 15.0
    elif given_angle >= 90:
        arc_r = 17.0
    elif given_angle <= 35:
        arc_r = 24.0
    else:
        arc_r = 20.0

    bix, biy = bisector_unit(x1, y1, x2, y2, x3, y3)
    apex_to_base = math.hypot((x2 + x3) / 2 - x1, (y2 + y3) / 2 - y1)

    # For obtuse angles, center the label halfway between arc and base line to guarantee
    # equal, collision-free clearance on both sides.
    if given_angle >= 95:
        label_push = (arc_r + apex_to_base) / 2.0
    else:
        target = arc_r + 14.0
        label_push = min(target, apex_to_base - 18.0)

    dvx = x1 + bix * label_push
    dvy = y1 + biy * label_push

    side_len_base = math.hypot(x1 - x2, y1 - y2)

    if given_angle >= 120:
        base_arc_r = 15.0
        arc_gap = 4.0
        q_push = base_arc_r + arc_gap + 15.0
    elif given_angle >= 90:
        base_arc_r = 17.0
        arc_gap = 4.5
        q_push = base_arc_r + arc_gap + 14.0
    else:
        base_arc_r = 20.0
        arc_gap = 5.0
        q_push = base_arc_r + arc_gap + 13.0

    q_push = min(q_push, side_len_base * 0.42)

    b2x, b2y = bisector_unit(x2, y2, x1, y1, x3, y3)
    b3x, b3y = bisector_unit(x3, y3, x1, y1, x2, y2)

    qx2 = x2 + b2x * q_push
    qy2 = y2 + b2y * q_push
    qx3 = x3 + b3x * q_push
    qy3 = y3 + b3y * q_push

    svg  = (f"<svg viewBox='0 0 {canvas} {canvas}' xmlns='http://www.w3.org/2000/svg' "
            f"width='{canvas}' height='{canvas}' "
            f"style='display:block; margin-left: auto; margin-right: auto; margin-top: 5px; margin-bottom: 5px;'>")

    svg += (f"<polygon points='{x1:.1f},{y1:.1f} {x2:.1f},{y2:.1f} {x3:.1f},{y3:.1f}'"
            f" fill='none' stroke='black' stroke-width='1.2'/>")

    deg_fontsize = 13 if given_angle >= 110 else 14
    svg += (f"<text x='{dvx:.1f}' y='{dvy:.1f}' "
            f"font-family='Arial' font-size='{deg_fontsize}' "
            f"dominant-baseline='central' text-anchor='middle' "
            f"paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' "
            f"fill='black'>{given_angle}°</text>")

    q_fontsize = 12 if given_angle >= 120 else 13
    svg += (f"<text x='{qx2:.1f}' y='{qy2:.1f}' "
            f"font-family='Arial' font-size='{q_fontsize}' "
            f"dominant-baseline='central' text-anchor='middle' "
            f"paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' "
            f"fill='black'>?</text>")
    svg += (f"<text x='{qx3:.1f}' y='{qy3:.1f}' "
            f"font-family='Arial' font-size='{q_fontsize}' "
            f"dominant-baseline='central' text-anchor='middle' "
            f"paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' "
            f"fill='black'>?</text>")

    for lx, ly, lab in [(lx1, ly1, labels[0]), (lx2, ly2, labels[1]), (lx3, ly3, labels[2])]:
        svg += (f"<text x='{lx:.1f}' y='{ly:.1f}' "
                f"font-family='Arial' font-size='19' "
                f"dominant-baseline='central' text-anchor='middle' "
                f"paint-order='stroke' stroke='white' stroke-width='3.5px' stroke-linejoin='round' "
                f"fill='black'>{lab}</text>")

    for vx_, vy_ in [(x1, y1), (x2, y2), (x3, y3)]:
        svg += f"<circle cx='{vx_:.1f}' cy='{vy_:.1f}' r='1.8' fill='black'/>"

    if given_angle == 90:
        u2x, u2y = unit_vec(x1, y1, x2, y2)
        u3x, u3y = unit_vec(x1, y1, x3, y3)
        sq = 14
        px1_, py1_ = x1 + u2x * sq, y1 + u2y * sq
        px2_, py2_ = x1 + u3x * sq, y1 + u3y * sq
        pmx, pmy   = x1 + (u2x + u3x) * sq, y1 + (u2y + u3y) * sq
        svg += (f"<polyline points='{px1_:.1f},{py1_:.1f} "
                f"{pmx:.1f},{pmy:.1f} {px2_:.1f},{py2_:.1f}'"
                f" fill='none' stroke='black' stroke-width='1.2'/>")
    else:
        svg += get_arc_path(x1, y1, x2, y2, x3, y3, arc_r)

    for vx, vy, ax1, ay1, ax2, ay2 in [
        (x2, y2, x1, y1, x3, y3),
        (x3, y3, x1, y1, x2, y2),
    ]:
        svg += get_arc_path(vx, vy, ax1, ay1, ax2, ay2, base_arc_r)
        svg += get_arc_path(vx, vy, ax1, ay1, ax2, ay2, base_arc_r + arc_gap)

    svg += "</svg>"
    return svg.replace("\n", " ").replace("    ", "").strip()


def generate_distractors(given_angle, equal_angles):
    distractors = set()
    distractors.add(180 - given_angle)
    distractors.add(given_angle // 2)
    distractors.add(given_angle)
    distractors.add((360 - given_angle) // 2)
    distractors.add(equal_angles + 10)

    candidate = abs(equal_angles - 10)
    if candidate >= 10:
        distractors.add(candidate)

    if given_angle < 90:
        v = (90 - given_angle) // 2
        if v >= 10:
            distractors.add(v)
    else:
        distractors.add(45)

    distractors.discard(equal_angles)
    distractors.discard(0)
    distractors = {d for d in distractors if 10 <= d <= 170}

    while len(distractors) < 6:
        val = random.randint(15, 85)
        if val != equal_angles:
            distractors.add(val)

    return list(distractors)[:6]


def build_excel():
    acute_pool  = list(range(36, 90, 2))
    obtuse_pool = list(range(100, 142, 2))

    random.shuffle(acute_pool)
    acutes  = acute_pool[:24]
    obtuses = obtuse_pool

    all_angles = acutes + [90] * 5 + obtuses
    random.shuffle(all_angles)

    orientation_pool = []
    step = 360 / 50
    for i in range(50):
        base = int(i * step)
        jitter = random.randint(-8, 8)
        orientation_pool.append((base + jitter) % 360)
    random.shuffle(orientation_pool)

    all_label_triplets = [
        list("ABC"), list("DEF"), list("GHI"), list("JKL"),
        list("MNO"), list("PQR"), list("STU"), list("VWX"),
        list("XYZ"), list("EFG"), list("HIJ"), list("KLM"),
        list("NOP"), list("QRS"), list("TUV"), list("UVW"),
        list("RST"), list("LMN"),
    ]
    shuffled_labels = all_label_triplets[:]
    random.shuffle(shuffled_labels)
    label_pool = []
    while len(label_pool) < 50:
        label_pool.extend(shuffled_labels)
    label_pool = label_pool[:50]

    used_signatures = set()
    rows = []

    for i in range(50):
        given_angle  = all_angles[i]
        equal_angles = (180 - given_angle) // 2

        orientation = orientation_pool[i]
        labels      = label_pool[i]

        for attempt in range(150):
            svg = generate_svg(given_angle, orientation, labels)

            q_text = (
                f"In a certain triangle, as shown in the figure below, one angle measures ${given_angle}^{{\\circ}}$ "
                f"and the remaining two angles are equal in measure. "
                f"What is the measure of each remaining angle?<br>#\n"
                f"खाली आकृती मध्ये दाखविल्या नुसार, एका त्रिकोणाचा एक कोन ${given_angle}^{{\\circ}}$ मापाचा आहे आणि "
                f"त्याचे उरलेले दोन कोन समान मापाचे आहेत. "
                f"तर त्या उरलेल्या कोनांचे माप किती असेल?<br>\n"
                f"{svg}<br>"
            )

            all_distractors      = generate_distractors(given_angle, equal_angles)
            selected_distractors = random.sample(all_distractors, 3)

            # Build option string without labels — just the value
            correct_ans_str = f"${equal_angles}^{{\\circ}}$<br>"
            wrong_ans_strs  = [f"${d}^{{\\circ}}$<br>" for d in selected_distractors]

            signature = (
                given_angle,
                orientation,
                tuple(labels),
                tuple(sorted(selected_distractors)),
            )

            if signature not in used_signatures:
                used_signatures.add(signature)
                break

            orientation = (orientation + random.randint(12, 35)) % 360

        rem      = 180 - given_angle

        # Solution text matching the required format (no "Option X" labels)
        solution = (
            f"Ans: ${equal_angles}^{{\\circ}}$<br>\n"
            f"The sum of measures of all angles of a triangle is always $180^{{\\circ}}$.<br>\n"
            f"Measure of given angle = ${given_angle}^{{\\circ}}$.<br>\n"
            f"$\\therefore$ Sum of remaining angles = $180^{{\\circ}} - {given_angle}^{{\\circ}} = {rem}^{{\\circ}}$.<br>\n"
            f"Since, remaining two angles are equal in measure, we divide this remaining quantity equally in two parts:<br>\n"
            f"Measure of each remaining angle = $\\dfrac{{{rem}^{{\\circ}}}}{{2}} = {equal_angles}^{{\\circ}}$.<br>\n"
            f"Hence, ${equal_angles}^{{\\circ}}$ is the answer.<br>#\n"
            f"उत्तर: ${equal_angles}^{{\\circ}}$<br>\n"
            f"त्रिकोणाच्या सर्व कोनांच्या मापांची बेरीज नेहेमी $180^{{\\circ}}$ असते.<br>\n"
            f"दिलेला कोन = ${given_angle}^{{\\circ}}$ या मापाचा आहे.<br>\n"
            f"$\\therefore$ उरलेल्या कोनांच्या मापांची बेरीज = $180^{{\\circ}} - {given_angle}^{{\\circ}} = {rem}^{{\\circ}}$ असेल.<br>\n"
            f"उरलेले दोन कोन समान मापाचे असल्याने, आपण या उरलेल्या किमतीचे दोन समान भाग करू:<br>\n"
            f"$\\therefore$ प्रत्येक उरलेल्या कोनाचे माप = $\\dfrac{{{rem}^{{\\circ}}}}{{2}} = {equal_angles}^{{\\circ}}$ असेल.<br>\n"
            f"म्हणून,  ${equal_angles}^{{\\circ}}$ हे उत्तर.<br>"
        )

        row_data = [
            i + 1,
            "Text",
            1,
            str(CONFIG["topic_number"]),
            q_text,
            correct_ans_str,
            None,
            None,
            None,
            wrong_ans_strs[0],
            wrong_ans_strs[1],
            wrong_ans_strs[2],
            int(CONFIG["time_in_seconds"]),
            int(CONFIG["difficulty"]),
            None,
            str(CONFIG["contributor_email"]),
            solution,
            None,
            int(CONFIG["variation_number"]),
        ]
        rows.append(row_data)

    rows.append(["****"] + [None] * 2 + ["****"] + [None] * (len(COLUMNS) - 4))

    out_dir = Path(__file__).parent
    out_name = "Intern_Geo_0404010201_123_3_Assign1_Siddhi_Manjarekar.xlsx"
    filepath = out_dir / out_name

    wb = xlsxwriter.Workbook(str(filepath))

    # 1. Sheet Index 0 (First Tab): Instruction Sheet (completely empty)
    ws_inst = wb.add_worksheet("Instruction")

    # 2. Sheet Index 1 (Second Tab): Questions Sheet (main dataset)
    ws_q = wb.add_worksheet("Questions")

    # Formats matching OpenXML specifications
    fmt_header = wb.add_format({
        'align': 'left',
        'valign': 'top',
        'bold': True,
        'text_wrap': False
    })

    fmt_wrap = wb.add_format({
        'align': 'left',
        'valign': 'top',
        'text_wrap': True
    })

    fmt_nowrap = wb.add_format({
        'align': 'left',
        'valign': 'top',
        'text_wrap': False
    })

    # Text format (built-in format 49 / '@') with applyNumberFormat=1
    fmt_topic = wb.add_format({
        'num_format': 49,
        'align': 'left',
        'valign': 'top',
        'text_wrap': False
    })

    # Write Header row
    for c_idx, col_name in enumerate(COLUMNS):
        ws_q.write(0, c_idx, col_name, fmt_header)

    ws_q.freeze_panes(1, 0)

    col_widths = {
        "A": 3, "B": 3, "C": 3, "D": 5, "E": 60, "F": 15,
        "G": 2, "H": 2, "I": 2, "J": 15, "K": 15, "L": 15,
        "M": 5, "N": 3, "O": 3, "P": 5, "Q": 75, "R": 3, "S": 10,
    }
    for col, width in col_widths.items():
        if col == "D":
            ws_q.set_column(f"{col}:{col}", width, fmt_topic)
        else:
            ws_q.set_column(f"{col}:{col}", width)

    # Write data rows
    for r_idx, r_data in enumerate(rows, start=1):
        for c_idx, val in enumerate(r_data):
            if val is None:
                continue
            fmt = fmt_wrap if c_idx in [4, 5, 9, 10, 11, 16] else fmt_nowrap
            if c_idx == 3 and val != "****":
                ws_q.write_string(r_idx, c_idx, str(val), fmt_topic)
            elif isinstance(val, (int, float)):
                ws_q.write_number(r_idx, c_idx, val, fmt)
            else:
                ws_q.write_string(r_idx, c_idx, str(val), fmt)

    # Activate Questions sheet so activeTab=1
    ws_q.activate()
    wb.close()
    print(f"Saved: {filepath} ({len(rows) - 1} rows, {len(used_signatures)} unique signatures)")
    return str(filepath)


if __name__ == "__main__":
    build_excel()