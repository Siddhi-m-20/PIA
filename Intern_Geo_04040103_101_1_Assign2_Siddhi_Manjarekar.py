import os
import random
from pathlib import Path
import xlsxwriter

CONFIG = {
    'contributor_email': 'siddhimanjarekar274@gmail.com',
    'topic_number': '04040103',
    'variation_number': 101,
    'assign_no': 2,
    'your_name': 'Siddhi_Manjarekar',
    'total_questions': 50,
    'time_in_seconds': 60,
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

TRIANGLE_NAMES = [
    ('A', 'B', 'C'),
    ('P', 'Q', 'R'),
    ('L', 'M', 'N'),
    ('X', 'Y', 'Z'),
    ('D', 'E', 'F'),
    ('S', 'T', 'U'),
    ('G', 'H', 'I'),
    ('J', 'K', 'L'),
    ('M', 'N', 'O'),
    ('R', 'S', 'T'),
]


def fmt_option(val):
    return f'${val}$ cm.<br>#\n${val}$ सेमी<br>'


def generate_wrong_options(max_c, count=6):
    '''
    Generate count wrong options that are >= max_c (sum of two sides).
    These are invalid because the third side must be STRICTLY less than the sum of the other two sides.
    Ensures all values are distinct and rounded to 1 decimal place.
    '''
    wrong_vals = set()
    # Exact sum (= max_c) is invalid because third side must be strictly less
    wrong_vals.add(round(max_c, 1))

    attempts = 0
    while len(wrong_vals) < count and attempts < 1000:
        attempts += 1
        offset = round(random.uniform(0.2, 3.5), 1)
        candidate = round(max_c + offset, 1)
        if candidate not in wrong_vals:
            wrong_vals.add(candidate)

    extra = 0.2
    while len(wrong_vals) < count:
        candidate = round(max_c + extra, 1)
        if candidate not in wrong_vals:
            wrong_vals.add(candidate)
        extra = round(extra + 0.2, 1)

    return sorted(list(wrong_vals))[:count]


def generate_dataset_rows():
    rows = []
    used_combinations = set()

    for i in range(1, CONFIG['total_questions'] + 1):
        while True:
            tri = random.choice(TRIANGLE_NAMES)
            V1, V2, V3 = tri

            a = round(random.uniform(4.0, 15.0), 1)
            b = round(random.uniform(4.0, 15.0), 1)

            min_c = round(abs(a - b), 1)
            max_c = round(a + b, 1)

            # Ensure sufficient gap between min_c and max_c
            if max_c - min_c < 2.5:
                continue

            # Pick correct_c strictly between min_c and max_c
            lower_bound = max(min_c + 0.5, max_c - 3.2)
            upper_bound = max_c - 0.6
            if lower_bound >= upper_bound:
                continue

            correct_c = round(random.uniform(lower_bound, upper_bound), 1)
            if not (min_c < correct_c < max_c):
                continue

            key = (V1, V2, V3, a, b, correct_c)
            if key in used_combinations:
                continue

            used_combinations.add(key)
            break

        all_wrong_6 = generate_wrong_options(max_c, count=6)
        candidate_wrongs = [v for v in all_wrong_6 if v != correct_c]
        displayed_wrong_3 = random.sample(candidate_wrongs, 3)

        correct_ans_str = fmt_option(correct_c)
        wrong_ans_strs = [fmt_option(v) for v in displayed_wrong_3]

        q_text = (
            f'If two sides are given as ${V1}{V2} = {a}$ cm and ${V1}{V3} = {b}$ cm, '
            f'then, for $\\triangle {V1}{V2}{V3}$ to be feasible, the side ${V2}{V3}$ can have length as _______ cm.<br>#\n'
            f'जर दोन बाजू ${V1}{V2} = {a}$ सेमी आणि ${V1}{V3} = {b}$ सेमी दिल्या असतील, '
            f'तर $\\triangle {V1}{V2}{V3}$ तयार होण्यासाठी बाजू ${V2}{V3}$ ची लांबी _______ सेमी असू शकते.<br>'
        )

        solution = (
            f'Ans: ${correct_c}$ cm.<br>\n'
            f'For a triangle to be formed, the sum of any two sides must be greater than the third side.<br>\n'
            f'Here, ${V1}{V2} = {a}$ cm and ${V1}{V3} = {b}$ cm.<br>\n'
            f'Sum of the two given sides $= {a} + {b} = {max_c}$ cm.<br>\n'
            f'So, the third side ${V2}{V3}$ must be strictly less than ${max_c}$ cm.<br>\n'
            f'In the given options, the only side whose length is $<{max_c}$ cm is the side with length ${correct_c}$ cm.<br>\n'
            f'All other options are greater than or equal to ${max_c}$ cm, so they cannot form a triangle.<br>\n'
            f'Hence, the side ${V2}{V3}$ can have the length as ${correct_c}$ cm.<br>\n'
            f'Therefore, ${correct_c}$ cm is the answer.<br>\n'
            f'#उत्तर: ${correct_c}$ सेमी<br>\n'
            f'त्रिकोण तयार होण्यासाठी कोणत्याही दोन बाजूंची बेरीज तिसऱ्या बाजूपेक्षा मोठी असली पाहिजे.<br>\n'
            f'येथे, ${V1}{V2} = {a}$ सेमी आणि ${V1}{V3} = {b}$ सेमी.<br>\n'
            f'दोन बाजूंची बेरीज $= {a} + {b} = {max_c}$ सेमी.<br>\n'
            f'म्हणून बाजू ${V2}{V3}$ ची लांबी ${max_c}$ सेमी पेक्षा कमी असली पाहिजे.<br>\n'
            f'दिलेल्या पर्यायांपैकी, फक्त ${correct_c}$ सेमी ही लांबी ही अट पूर्ण करते.<br>\n'
            f'इतर सर्व पर्याय ${max_c}$ सेमी पेक्षा मोठे किंवा समान असल्यामुळे त्या लांबीने त्रिकोण तयार होऊ शकत नाही.<br>\n'
            f'म्हणून, त्रिकोण तयार होण्यासाठी, बाजू ${V2}{V3}$ ची लांबी ${correct_c}$ सेमी असू शकते.<br>\n'
            f'म्हणून, ${correct_c}$ सेमी हे उत्तर.<br>'
        )

        row_data = [
            i,
            'Text',
            int(CONFIG['answer_type']),
            str(CONFIG['topic_number']),
            q_text,
            correct_ans_str,
            None,
            None,
            None,
            wrong_ans_strs[0],
            wrong_ans_strs[1],
            wrong_ans_strs[2],
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
    filename = parent_dir / 'Intern_Geo_04040103_101_1_Assign2_Siddhi_Manjarekar.xlsx'

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