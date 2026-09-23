import os
import random
from pathlib import Path
import xlsxwriter

CONFIG = {
    'contributor_email': 'siddhimanjarekar274@gmail.com',
    'topic_number': '04040103',
    'variation_number': 102,
    'assign_no': 3,
    'your_name': 'Siddhi_Manjarekar',
    'total_questions': 50,
    'time_in_seconds': 120,
    'difficulty': 1,
    'answer_type': 1,
}

COLUMNS = [
    'Sr. No', 'Question Type', 'Answer Type', 'Topic Number',
    'Question (Text Only)', 'Correct Answer 1', 'Correct Answer 2',
    'Correct Answer 3', 'Correct Answer 4', 'Wrong Answer 1',
    'Wrong Answer 2', 'Wrong Answer 3', 'Time in seconds',
    'Difficulty Level', 'Question (Image/ Audio/ Video)',
    "Contributor's Registered mailId", 'Solution (Text Only)',
    'Solution (Image/ Audio/ Video)', 'Variation Number'
]


def is_valid_triangle(a, b, c):
    return (a + b > c) and (b + c > a) and (a + c > b)


def make_valid_triplet(used):
    for _ in range(1000):
        a = round(random.uniform(4.0, 15.0), 1)
        b = round(random.uniform(4.0, 15.0), 1)
        min_c = round(abs(a - b) + 0.5, 1)
        max_c = round(a + b - 0.5, 1)
        if min_c > max_c:
            continue
        c = round(random.uniform(min_c, max_c), 1)
        t = tuple(sorted([a, b, c]))
        if is_valid_triangle(*t) and t not in used:
            return t
    return None


def make_invalid_triplet(used):
    for _ in range(1000):
        a = round(random.uniform(4.0, 15.0), 1)
        b = round(random.uniform(4.0, 15.0), 1)
        choice = random.randint(0, 2)
        if choice == 0:
            c = round(a + b + random.uniform(0.1, 4.0), 1)
        elif choice == 1:
            c = round(random.uniform(1.0, 5.0), 1)
            a = round(b + c + random.uniform(0.1, 4.0), 1)
        else:
            c = round(random.uniform(1.0, 5.0), 1)
            b = round(a + c + random.uniform(0.1, 4.0), 1)
        t = tuple(sorted([a, b, c]))
        if not is_valid_triangle(*t) and t not in used:
            return t
    return None


def fmt(t):
    return f"({t[0]}, {t[1]}, {t[2]})"


def build_solution(options, invalid, correct_letter):
    a, b, c = sorted(invalid)  # a <= b <= c
    sum_bc = round(b + c, 1)
    sum_ac = round(a + c, 1)
    sum_ab = round(a + b, 1)

    en = (
        f"Ans: ${correct_letter}$.<br>\n"
        f"For three given lengths to form a triangle, the sum of the lengths of any two sides must be greater than the length of the third side.<br>\n"
        f"We will have to verify this for all given triplets.<br>\n"
        f"After checking this way, the only triplet which does not obey this rule is, ${fmt(invalid)}$.<br>\n"
        f"For the triplet ${fmt(invalid)}$, we get,<br>\n"
        f"${b} + {c} = {sum_bc} > {a}$, the remaining side,<br>\n"
        f"${a} + {c} = {sum_ac} > {b}$, the remaining side, but<br>\n"
        f"${a} + {b} = {sum_ab}$.<br>\n"
        f"This sum is not greater than ${c}$, the length of remaining side.<br>\n"
        f"Since, given condition is not satisfied, this triplet can not form a triangle.<br>\n"
        f"For all other triplets, all three inequalities are satisfied, so they can form triangles.<br>\n"
        f"Therefore, ${correct_letter}$ is the answer.<br>"
    )

    mr = (
        f"#उत्तर: ${correct_letter}$.<br>\n"
        f"त्रिकोण तयार होण्यासाठी, त्रिकुटामधील कोणत्याही दोन बाजूंची बेरीज ही तिसऱ्या (उरलेल्या) बाजूपेक्षा मोठी असली पाहिजे.<br>\n"
        f"आपल्याला दिलेल्या सर्व त्रिकुटांची या अटीच्या पूर्ततेसाठी पडताळणी करावी लागेल.<br>\n"
        f"या पद्धतीने तपासल्यावर असे दिसते की, या नियमाचे पालन न करणारे एकमेव त्रिकुट ${fmt(invalid)}$ आहे.<br>\n"
        f"${fmt(invalid)}$ या त्रिकुटासाठी आपल्याला,<br>\n"
        f"${b} + {c} = {sum_bc} > {a}$ जी उरलेली बाजू आहे.<br>\n"
        f"${a} + {c} = {sum_ac} > {b}$, जी उरलेली बाजू आहे, पण<br>\n"
        f"${a} + {b} = {sum_ab}$ आणि<br>\n"
        f" ही बेरीज उरलेली बाजू ${c}$ पेक्षा मोठी नाही, असे मिळते.<br>\n"
        f"म्हणजेच, हे त्रिकुट दिलेली अट पूर्ण करीत नाही, त्यामुळे फक्त  याच त्रिकुटाने त्रिकोण तयार होऊ शकत नाही.<br>\n"
        f"इतर सर्व त्रिकुटांसाठी दिलेली अट पूर्ण होत असल्यामुळे ते त्रिकोण तयार करू शकतात.<br>\n"
        f"म्हणून, ${correct_letter}$ हे उत्तर.<br>"
    )

    return en + "\n" + mr


def generate_dataset_rows():
    used_triplets = set()
    used_option_sets = set()
    rows = []

    for i in range(1, CONFIG['total_questions'] + 1):
        while True:
            invalid = make_invalid_triplet(used_triplets)
            if invalid is None:
                continue

            temp_used = set(used_triplets) | {invalid}
            valids = []
            ok = True
            for _ in range(3):
                v = make_valid_triplet(temp_used)
                if v is None:
                    ok = False
                    break
                valids.append(v)
                temp_used.add(v)
            if not ok:
                continue

            option_set = frozenset([invalid] + valids)
            if option_set in used_option_sets:
                continue

            used_triplets.add(invalid)
            for v in valids:
                used_triplets.add(v)
            used_option_sets.add(option_set)
            break

        all_triplets = [invalid] + valids
        random.shuffle(all_triplets)
        letters = ['A', 'B', 'C', 'D']
        options = list(zip(letters, all_triplets))

        correct_letter = next(lbl for lbl, t in options if t == invalid)
        wrong_letters = [lbl for lbl, t in options if t != invalid]

        eng_opts = "".join(f"${lbl}$. {fmt(t)}<br>" for lbl, t in options)

        q_text = (
            f"Below given are some triplets representing the side lengths. "
            f"Which of the triplets cannot form a triangle?<br>\n"
            f"{eng_opts}#\n"
            f"खाली रेषाखंडांच्या लांबीसह काही त्रिकूटे दिली आहेत. "
            f"यातील कोणते त्रिकुट त्रिकोण तयार करू शकत नाही?<br>\n"
            f"{eng_opts}"
        )

        solution = build_solution(options, invalid, correct_letter)

        row_data = [
            i,
            'Text',
            int(CONFIG['answer_type']),
            str(CONFIG['topic_number']),
            q_text,
            f"${correct_letter}$.<br>",
            None,
            None,
            None,
            f"${wrong_letters[0]}$.<br>",
            f"${wrong_letters[1]}$.<br>",
            f"${wrong_letters[2]}$.<br>",
            int(CONFIG['time_in_seconds']),
            int(CONFIG['difficulty']),
            None,
            str(CONFIG['contributor_email']),
            solution,
            None,
            int(CONFIG['variation_number']),
        ]
        rows.append(row_data)

    # Marker row at row 52
    rows.append(['****'] + [None] * 2 + ['****'] + [None] * (len(COLUMNS) - 4))
    return rows


def build_excel():
    rows = generate_dataset_rows()
    parent_dir = Path(__file__).parent
    filename = parent_dir / 'Intern_Geo_04040103_102_1_Assign3_Siddhi_Manjarekar.xlsx'

    wb = xlsxwriter.Workbook(str(filename))

    # Sheet 0 (Tab 0): Instruction Sheet (completely empty)
    ws_inst = wb.add_worksheet('Instruction')

    # Sheet 1 (Tab 1): Questions Sheet (main dataset)
    ws_q = wb.add_worksheet('Questions')

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
    fmt_topic = wb.add_format({
        'num_format': 49,
        'align': 'left',
        'valign': 'top',
        'text_wrap': False
    })

    # Header row
    for c_idx, col_name in enumerate(COLUMNS):
        ws_q.write(0, c_idx, col_name, fmt_header)

    ws_q.freeze_panes(1, 0)

    col_widths = {
        'A': 3, 'B': 3, 'C': 3, 'D': 5, 'E': 60, 'F': 15,
        'G': 2, 'H': 2, 'I': 2, 'J': 15, 'K': 15, 'L': 15,
        'M': 5, 'N': 3, 'O': 3, 'P': 5, 'Q': 75, 'R': 3, 'S': 10,
    }
    for col, width in col_widths.items():
        if col == 'D':
            ws_q.set_column(f'{col}:{col}', width, fmt_topic)
        else:
            ws_q.set_column(f'{col}:{col}', width)

    # Write data rows
    for r_idx, r_data in enumerate(rows, start=1):
        for c_idx, val in enumerate(r_data):
            if val is None:
                continue
            fmt = fmt_wrap if c_idx in [4, 5, 9, 10, 11, 16] else fmt_nowrap
            if c_idx == 3 and val != '****':
                ws_q.write_string(r_idx, c_idx, str(val), fmt_topic)
            elif isinstance(val, (int, float)):
                ws_q.write_number(r_idx, c_idx, val, fmt)
            else:
                ws_q.write_string(r_idx, c_idx, str(val), fmt)

    # Activate Questions tab (activeTab = 1)
    ws_q.activate()
    wb.close()
    print(f'Successfully generated portal-ready workbook: {filename}')
    return str(filename)


if __name__ == '__main__':
    build_excel()