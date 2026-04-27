import random
import unicodedata
import openpyxl
import warnings
warnings.filterwarnings('ignore')

def normalize_marathi_text(text):
    return unicodedata.normalize('NFC', text)

def get_standard_name(standard):
    number_words_eng = {
        1: "1^{st}", 2: "2^{nd}", 3: "3^{rd}", 4: "4^{th}", 5: "5^{th}",
        6: "6^{th}", 7: "7^{th}", 8: "8^{th}", 9: "9^{th}", 10: "10^{th}",
        11: "11^{th}", 12: "12^{th}"
    }
    number_words_mar = {
        1: "पहिलीच्या", 2: "दुसरीच्या", 3: "तिसरीच्या", 4: "चौथीच्या", 5: "पाचवीच्या",
        6: "सहावीच्या", 7: "सातवीच्या", 8: "आठवीच्या", 9: "नववीच्या", 10: "दहावीच्या",
        11: "अकरावीच्या", 12: "बारावीच्या"
    }
    return number_words_eng.get(standard, "Sixth"), number_words_mar.get(standard, "सहावीच्या")

def create_pictograph_table(quantities, teachers_eng, teachers_mar, scale=3):
    face_image = '<img src=\'https://portal.coepvlab.ac.in/VirtualMathsLab/media/090201_108_emoji1.png\' style=\'width:24px;height:24px;vertical-align:middle;\'>'
    table_html = '<table border=\'1\' style=\'border-collapse:collapse;width:100%;text-align:left;\'>'
    table_html += '<tr style=\'background-color:lightgray;\'>'
    table_html += '<th style=\'padding:8px;font-size:14px;\'>Teacher Name / शिक्षकांचे नाव</th>'
    table_html += '<th style=\'padding:8px;font-size:14px;\'>Number of Students / विद्यार्थ्यांची संख्या</th>'
    table_html += '</tr>'
    for teacher_eng, teacher_mar, qty in zip(teachers_eng, teachers_mar, quantities):
        table_html += '<tr>'
        table_html += f'<td style=\'padding:8px;font-size:14px;\'>{teacher_eng} / {teacher_mar}</td>'
        table_html += f'<td style=\'padding:8px;\'>{face_image * qty}</td>'
        table_html += '</tr>'
    table_html += '</table>'
    return table_html

def create_chart_svg(quantities, teachers_eng, teachers_mar, scale=3):
    actual_quantities = [q * scale for q in quantities]
    y_max = max(24, int(max(actual_quantities) + 3))
    y_max = ((y_max + 2) // 3) * 3

    chart_height = 280
    bar_heights = [(q / y_max) * chart_height for q in actual_quantities]

    origin_x = 100
    grid_width = 470
    grid_end_x = origin_x + grid_width
    start_x = origin_x + 35

    top_margin = 45
    svg_height = chart_height + 120

    svg = f'<svg width=\'750\' height=\'{svg_height}\' xmlns=\'http://www.w3.org/2000/svg\'>'
    svg += '<rect width=\'100%\' height=\'100%\' fill=\'white\'/>'

    for i in range(0, y_max + 1, 3):
        y_pos = chart_height - (i / y_max) * chart_height + top_margin
        if i % 6 == 0:
            svg += f'<line x1=\'{origin_x}\' y1=\'{y_pos}\' x2=\'{grid_end_x}\' y2=\'{y_pos}\' stroke=\'gray\' stroke-width=\'1\' opacity=\'0.8\'/>'
            svg += f'<text x=\'{origin_x - 10}\' y=\'{y_pos + 4}\' text-anchor=\'end\' font-size=\'13\' font-family=\'Arial\'>{i}</text>'
        else:
            svg += f'<line x1=\'{origin_x}\' y1=\'{y_pos}\' x2=\'{grid_end_x}\' y2=\'{y_pos}\' stroke=\'gray\' stroke-width=\'0.5\' opacity=\'0.6\'/>'

    svg += f'<line x1=\'{origin_x}\' y1=\'{top_margin - 10}\' x2=\'{origin_x}\' y2=\'{chart_height + top_margin}\' stroke=\'black\' stroke-width=\'1\'/>'
    svg += f'<polygon points=\'{origin_x},{top_margin - 10} {origin_x - 5},{top_margin} {origin_x + 5},{top_margin}\' fill=\'black\'/>'
    svg += f'<text x=\'{origin_x}\' y=\'{top_margin - 25}\' text-anchor=\'middle\' font-family=\'Arial\' font-size=\'12\'>Y-Axis / Y-अक्ष</text>'

    arrow_tip_x = grid_end_x + 5
    svg += f'<line x1=\'{origin_x}\' y1=\'{chart_height + top_margin}\' x2=\'{arrow_tip_x}\' y2=\'{chart_height + top_margin}\' stroke=\'black\' stroke-width=\'1\'/>'
    svg += f'<polygon points=\'{arrow_tip_x},{chart_height + top_margin} {arrow_tip_x - 10},{chart_height + top_margin - 5} {arrow_tip_x - 10},{chart_height + top_margin + 5}\' fill=\'black\'/>'

    mid_y = (chart_height / 2) + top_margin
    svg += f'<text transform=\'rotate(-90, 40, {mid_y})\' x=\'40\' y=\'{mid_y}\' text-anchor=\'middle\' font-family=\'Arial\' font-size=\'14\'>'
    svg += f'<tspan x=\'40\' dy=\'-0.7em\'>Number of Students</tspan>'
    svg += f'<tspan x=\'40\' dy=\'1.4em\'>विद्यार्थ्यांची संख्या</tspan>'
    svg += f'</text>'

    bar_width = 60
    bar_spacing = 90
    bar_colors = ["steelblue", "coral", "darkgray", "gold", "mediumseagreen"]

    for i, (height, color, qty) in enumerate(zip(bar_heights, bar_colors, actual_quantities)):
        x_pos = start_x + i * bar_spacing
        bar_y = chart_height + top_margin - height
        svg += f'<rect x=\'{x_pos}\' y=\'{bar_y}\' width=\'{bar_width}\' height=\'{height}\' fill=\'{color}\' stroke=\'black\' stroke-width=\'0.7\'/>'

    label_y_eng = chart_height + top_margin + 20
    for i, teacher_eng in enumerate(teachers_eng):
        x_center = start_x + i * bar_spacing + (bar_width / 2)
        svg += f'<text x=\'{x_center}\' y=\'{label_y_eng}\' text-anchor=\'middle\' font-size=\'12\' font-family=\'Arial\' font-weight=\'bold\'>{teacher_eng}</text>'

    label_y_mar = chart_height + top_margin + 38
    for i, teacher_mar in enumerate(teachers_mar):
        x_center = start_x + i * bar_spacing + (bar_width / 2)
        svg += f'<text x=\'{x_center}\' y=\'{label_y_mar}\' text-anchor=\'middle\' font-size=\'12\' font-family=\'Arial\' font-weight=\'bold\'>{normalize_marathi_text(teacher_mar)}</text>'

    axis_label_x = grid_end_x + 15
    axis_label_y_eng = chart_height + top_margin + 5
    axis_label_y_mar = chart_height + top_margin + 20
    svg += f'<text x=\'{axis_label_x}\' y=\'{axis_label_y_eng}\' text-anchor=\'start\' font-size=\'13\' font-family=\'Arial\'>X-Axis</text>'
    svg += f'<text x=\'{axis_label_x}\' y=\'{axis_label_y_mar}\' text-anchor=\'start\' font-size=\'13\' font-family=\'Arial\'>X-अक्ष</text>'

    title_x_center = origin_x + (grid_width / 2)
    title_y = chart_height + top_margin + 65
    svg += f'<text x=\'{title_x_center}\' y=\'{title_y}\' text-anchor=\'middle\' font-size=\'14\' font-family=\'Arial\' font-weight=\'bold\'>Teacher Name / शिक्षकांचे नाव</text>'

    svg += '</svg>'
    return svg

def generate_wrong_answers(quantities, teachers_eng, teachers_mar, scale=3, all_teachers_eng=None, all_teachers_mar=None):
    wrong_answers_set = set()
    wrong_options = []
    attempts = 0
    max_attempts = 100
    if all_teachers_eng is None:
        all_teachers_eng = teachers_eng
    if all_teachers_mar is None:
        all_teachers_mar = teachers_mar
    alt_teachers_eng = [t for t in all_teachers_eng if t not in teachers_eng]
    alt_teachers_mar = [t for t in all_teachers_mar if t not in teachers_mar]
    alt_scales = [2, 4, 5, 6]
    while len(wrong_options) < 3 and attempts < max_attempts:
        attempts += 1
        temp_quantities = quantities[:]
        temp_teachers_eng = teachers_eng[:]
        temp_teachers_mar = teachers_mar[:]
        temp_scale = scale
        method = random.randint(0, 6)
        if method == 0:
            idx = random.randrange(len(temp_quantities))
            change = random.choice([-2, -1, 1, 2])
            temp_quantities[idx] = max(2, min(8, temp_quantities[idx] + change))
        elif method == 1:
            random.shuffle(temp_quantities)
        elif method == 2:
            idx1, idx2 = random.sample(range(len(temp_quantities)), 2)
            temp_quantities[idx1], temp_quantities[idx2] = temp_quantities[idx2], temp_quantities[idx1]
        elif method == 3:
            indices = random.sample(range(len(temp_quantities)), 2)
            for idx in indices:
                change = random.choice([-1, 1])
                temp_quantities[idx] = max(2, min(8, temp_quantities[idx] + change))
        elif method == 4:
            if alt_teachers_eng and alt_teachers_mar:
                num_changes = random.randint(1, min(2, len(alt_teachers_eng)))
                indices_to_change = random.sample(range(5), num_changes)
                available_eng = alt_teachers_eng[:]
                available_mar = alt_teachers_mar[:]
                for idx in indices_to_change:
                    if available_eng and available_mar:
                        new_eng = random.choice(available_eng)
                        eng_idx = all_teachers_eng.index(new_eng)
                        new_mar = all_teachers_mar[eng_idx]
                        temp_teachers_eng[idx] = new_eng
                        temp_teachers_mar[idx] = new_mar
                        available_eng.remove(new_eng)
                        available_mar.remove(new_mar)
        elif method == 5:
            temp_scale = random.choice(alt_scales)
        else:
            idx = random.randrange(len(temp_quantities))
            change = random.choice([-1, 1])
            temp_quantities[idx] = max(2, min(8, temp_quantities[idx] + change))
            if alt_teachers_eng and alt_teachers_mar:
                change_idx = random.randint(0, 4)
                available = [t for t in alt_teachers_eng if t not in temp_teachers_eng]
                if available:
                    new_eng = random.choice(available)
                    eng_idx = all_teachers_eng.index(new_eng)
                    new_mar = all_teachers_mar[eng_idx]
                    temp_teachers_eng[change_idx] = new_eng
                    temp_teachers_mar[change_idx] = new_mar
        wrong_key = (tuple(temp_quantities), tuple(temp_teachers_eng), temp_scale)
        correct_key = (tuple(quantities), tuple(teachers_eng), scale)
        if wrong_key != correct_key and wrong_key not in wrong_answers_set:
            wrong_answers_set.add(wrong_key)
            wrong_options.append({
                'quantities': temp_quantities[:],
                'teachers_eng': temp_teachers_eng[:],
                'teachers_mar': temp_teachers_mar[:],
                'scale': temp_scale
            })
    return wrong_options if len(wrong_options) == 3 else None

def create_solution_text(quantities, teachers_eng, teachers_mar, correct_option_letter, scale=3):
    q0_val = quantities[0] * scale
    q1_val = quantities[1] * scale
    q2_val = quantities[2] * scale
    q3_val = quantities[3] * scale
    q4_val = quantities[4] * scale
    face_symbol = '<img src=\'https://portal.coepvlab.ac.in/VirtualMathsLab/media/090201_108_emoji1.png\' style=\'width:20px;height:20px;vertical-align:middle;\'>'
    eng_sol = (
        f"Ans : ${correct_option_letter}$.<br>"
        f"The pictograph provided to us gives following information.<br>"
        f"$1.$ Number of teachers the students like are $5$.<br>"
        f"$2.$ Names of the teachers liked by students are {teachers_eng[0]}, {teachers_eng[1]}, {teachers_eng[2]}, {teachers_eng[3]} and {teachers_eng[4]}.<br>"
        f"$3$. Scale used is $1$ face symbol {face_symbol} $= {scale}$ students.<br>"
        f"$4.$Teacher liked by number of students is as follows :  <br>"
        f"Number of students liking a teacher can be calculated as,<br>"
        f"$=$ number of faces $\\times {scale}$.<br>"
        f"{teachers_eng[0]} is liked by ${quantities[0]}$ faces $\\times {scale} = {q0_val}$ students,<br>"
        f"{teachers_eng[1]} is liked by ${quantities[1]}$ faces $\\times {scale} = {q1_val}$ students,<br>"
        f"{teachers_eng[2]} is liked by ${quantities[2]}$ faces $\\times {scale} = {q2_val}$ students,<br>"
        f"{teachers_eng[3]} is liked by ${quantities[3]}$ faces $\\times {scale} = {q3_val}$ students,<br>"
        f"{teachers_eng[4]} is liked by ${quantities[4]}$ faces $\\times {scale} = {q4_val}$ students,<br>"
        f"From this information, we know that, required bar graph will have $5$ bars corresponding to these $5$ teachers.<br>"
        f"To find correct bar graph, corresponding to given pictograph, we need to check following conditions.<br>"
        f"$i)$ Check for $5$ bars corresponding to $5$ teachers.<br>"
        f"$ii)$ Check for the names of teachers as given.<br>"
        f"$iii)$ Check for scale such that the number of students will match with the corresponding teacher.<br>"
        f"$iv)$ Check for bar width and the space between the bars. Both must be uniform.<br>"
        f"By comparing all bargraphs with conditions mentioned in $(i)$ to $(iv)$, we can easily conclude that, the bar graph shown in option ${correct_option_letter}$ is the required bar graph.<br>"
        f"In, remaining bar graphs at least one of three conditions mentioned in $(i)$ to $(iv)$ is not satisfied.<br>"
        f"Hence, option ${correct_option_letter}$ is the answer.<br>#"
    )
    mar_sol = (
        f"उत्तर : ${correct_option_letter}$.<br>"
        f"आपल्याला दिलेल्या चित्रालेखावरून पुढील माहिती मिळते.<br>"
        f"$1.$ विद्यार्थ्यांना आवडणाऱ्या शिक्षकांची संख्या $5$ आहे.<br>"
        f"$2.$ विद्यार्थ्यांना आवडणारे शिक्षक {teachers_mar[0]}, {teachers_mar[1]}, {teachers_mar[2]}, {teachers_mar[3]} आणि {teachers_mar[4]} असे आहेत.<br>"
        f"$3.$ चित्रालेखात वापरलेले प्रमाण $1$ चेहेऱ्याचे चिन्ह {face_symbol} $= {scale}$ विद्यार्थी.<br>"
        f"$4.$ वेगवेगळे शिक्षक आवडत असणाऱ्या विद्यार्थ्यांची संख्या खालील प्रमाणे आहे :  <br>"
        f"प्रत्येक शिक्षक आवडणाऱ्या विद्यार्थ्यांची संख्या<br>"
        f"$=$ चेहेऱ्याच्या चिन्हांची संख्या $\\times {scale}$.<br>"
        f"शिक्षक {teachers_mar[0]} : चेहेऱ्याची ${quantities[0]}$ चिन्हे $\\times {scale} = {q0_val}$ विद्यार्थ्यांना आवडतात.<br>"
        f"शिक्षक {teachers_mar[1]} : चेहेऱ्याची ${quantities[1]}$ चिन्हे $\\times {scale} = {q1_val}$ विद्यार्थ्यांना आवडतात.<br>"
        f"शिक्षक {teachers_mar[2]} : चेहेऱ्याची ${quantities[2]}$ चिन्हे $\\times {scale} = {q2_val}$ विद्यार्थ्यांना आवडतात.<br>"
        f"शिक्षक {teachers_mar[3]} : चेहेऱ्याची ${quantities[3]}$ चिन्हे $\\times {scale} = {q3_val}$ विद्यार्थ्यांना आवडतात.<br>"
        f"शिक्षक {teachers_mar[4]} : चेहेऱ्याची ${quantities[4]}$ चिन्हे $\\times {scale} = {q4_val}$ विद्यार्थ्यांना आवडतात.<br>"
        f"या माहिती नुसार असे समजते की, आपल्याला हव्या स्तंभालेखात $5$ शिक्षकांसाठी $5$ स्तंभ असायला हवेत.<br>"
        f"दिलेल्या चित्रालेखाला अनुसरून, हव्या असलेल्या स्तंभालेखा साठी पुढील बाबी तपासाव्या लागतील.<br>"
        f"$i)$ $5$ शिक्षकांसाठी $5$ स्तंभ आहेत ना.... <br>"
        f"$ii)$शिक्षकांची नावे हवी आहेत तशीच आहेत ना..... <br>"
        f"$iii)$ योग्य प्रमाण वापरला गेलेला असा स्तंभालेख, ज्यात शिक्षकांनुसार विद्यार्थ्यांची संख्या जुळते.<br>"
        f"$iv)$ प्रत्येक स्तंभाची रुंदी आणि लगतच्या दोन स्तंभां मधील अंतर हे दोन्ही सामान आहे.<br>"
        f"अट क्रमांक  $(i)$ ते  $(iv)$ नुसार दिलेल्या सर्व स्तंभालेखांची तपासणी केली असता, आपल्या असे लक्षात येते की, फक्त पर्याय ${correct_option_letter}$ हाच असा पर्याय आहे, ज्यात आवश्यक सर्व अटी पूर्ण होतात. <br>"
        f"उरलेल्या इतर सर्व स्तंभालेखात अट  $(i)$ ते  $(iv)$ या पैकी किमान एक अट तरी पूर्ण होत नाही.<br>"
        f"म्हणून, पर्याय ${correct_option_letter}$ हे उत्तर.<br>"
    )
    return eng_sol + mar_sol


def create_option_chart_html(quantities, teachers_eng, teachers_mar, scale=3):
    chart_svg = create_chart_svg(quantities, teachers_eng, teachers_mar, scale)
    if not chart_svg:
        return ""

    eng_title = "Favorite Teachers"
    mar_title = normalize_marathi_text("आवडते शिक्षक")
    eng_scale = f"Scale: 1 unit = {scale} students"
    mar_scale = normalize_marathi_text(f"प्रमाण: 1 एकक = {scale} विद्यार्थी")

    complete_chart = (
        f'<div style=\'text-align:center;border:1px solid lightgray;padding:10px;max-width:720px;\'>'
        f'<span style=\'font-size:22px;\'>{eng_title} / {mar_title}</span><br>'
        f'<span style=\'font-size:15px;\'>({eng_scale} / {mar_scale})</span>'
        f'<div style=\'display:flex;align-items:center;justify-content:center;margin-top:10px;\'>'
        f'{chart_svg}'
        f'</div>'
        f'</div>'
    )
    return complete_chart

def create_pictograph_html(quantities, teachers_eng, teachers_mar, scale=3):
    pictograph_table = create_pictograph_table(quantities, teachers_eng, teachers_mar, scale)
    face_image = '<img src=\'https://portal.coepvlab.ac.in/VirtualMathsLab/media/090201_108_emoji1.png\' style=\'width:20px;height:20px;vertical-align:middle;\'>'
    eng_title = "Favorite Teachers / आवडते शिक्षक"
    eng_scale = f"Scale: 1 {face_image} = {scale} students"
    mar_scale = normalize_marathi_text(f"प्रमाण: 1 {face_image} = {scale} विद्यार्थी")
    complete_pictograph = (
        f'<div style=\'text-align:center;border:1px solid lightgray;padding:10px;max-width:650px;\'>'
        f'<span style=\'font-size:22px;\'>{eng_title}</span><br>'
        f'<span style=\'font-size:15px;\'>({eng_scale} / {mar_scale})</span>'
        f'<div style=\'margin-top:10px;\'>'
        f'{pictograph_table}'
        f'</div>'
        f'</div>'
    )
    return complete_pictograph

def main():
    filename = "Intern_09030201_108_3_Assign8_Siddhi_Manjarekar.xlsx"
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)
    workbook.create_sheet("Instruction")
    sheet = workbook.create_sheet("Questions")
    header = ["Sr. No", "Question Type", "Answer Type", "Topic Number", "Question (Text Only)",
              "Correct Answer 1", "Correct Answer 2", "Correct Answer 3", "Correct Answer 4",
              "Wrong Answer 1", "Wrong Answer 2", "Wrong Answer 3", "Time in seconds",
              "Difficulty Level", "Question (Image/ Audio/ Video)", "Contributor's Registered mailId",
              "Solution (Text Only)", "Solution (Image/ Audio/ Video)", "Variation Number"]
    sheet.append(header)
    try:
        num_questions = int(input("How many questions do you want to enter? "))
        if num_questions > 200:
            print("Warning: Requesting more than 200 questions may result in duplicates or take a long time.")
            proceed = input("Do you want to continue? (yes/no): ").strip().lower()
            if proceed != 'yes':
                return
    except ValueError:
        print("Invalid input. Please enter a number.")
        return
    all_teachers_eng = [
        "Shri Desai", "Shri Patil", "Shri Kulkarni", "Shri Joshi", "Shri Patel",
        "Shri Kamble", "Shri Jadhav", "Shri Pawar", "Shri Shinde", "Shri More",
        "Sau. Deshpande", "Sau. Joshi", "Sau. Kulkarni", "Sau. Patil", "Sau. Desai",
        "Shri Bhosale", "Shri Sawant", "Shri Chavan", "Shri Kale", "Shri Nikam",
        "Sau. Pawar", "Sau. Rane", "Sau. Gaikwad", "Ku. Bhosale", "Ku. Sawant",
        "Ku. Shinde", "Ku. Jadhav", "Shri Mane", "Shri Gaikwad", "Sau. Salunkhe"
    ]
    all_teachers_mar = [
        "श्री. देसाई", "श्री. पाटील", "श्री. कुलकर्णी", "श्री. जोशी", "श्री. पटेल",
        "श्री. कांबळे", "श्री. जाधव", "श्री. पवार", "श्री. शिंदे", "श्री. मोरे",
        "सौ. देशपांडे", "सौ. जोशी", "सौ. कुलकर्णी", "सौ. पाटील", "सौ. देसाई",
        "श्री. भोसले", "श्री. सावंत", "श्री. चव्हाण", "श्री. काळे", "श्री. निकम",
        "सौ. पवार", "सौ. राणे", "सौ. गायकवाड", "कु. भोसले", "कु. सावंत",
        "कु. शिंदे", "कु. जाधव", "श्री. माने", "श्री. गायकवाड", "सौ. साळुंखे"
    ]
    option_letters = ['A', 'B', 'C', 'D']
    scale = 3
    generated_questions = set()
    i = 1
    rows_to_add = []
    max_retries = num_questions * 10
    retry_count = 0
    while i <= num_questions and retry_count < max_retries:
        retry_count += 1
        quantities = random.sample(range(2, 8), 5)
        chosen_indices = random.sample(range(len(all_teachers_eng)), 5)
        teachers_eng5 = [all_teachers_eng[idx] for idx in chosen_indices]
        teachers_mar5 = [all_teachers_mar[idx] for idx in chosen_indices]
        question_key = (tuple(sorted(quantities)), tuple(sorted(teachers_eng5)))
        if question_key in generated_questions:
            continue
        wrong_answers_list = generate_wrong_answers(
            quantities,
            teachers_eng5,
            teachers_mar5,
            scale,
            all_teachers_eng,
            all_teachers_mar
        )
        if not wrong_answers_list:
            continue
        generated_questions.add(question_key)
        pictograph_html = create_pictograph_html(quantities, teachers_eng5, teachers_mar5, scale)
        standard = random.choice([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
        std_eng, std_mar = get_standard_name(standard)
        eng_question = (
            f"Information is collected from class ${std_eng}$ students, about the teacher they like most. For this information a pictograph is drawn as shown.<br>"
            f"For this pictograph an equivalent bar graph is to be drawn.<br>"
            f"Which of the following Bar graphs is the correct representation for this?<br>"
        )
        mar_question = normalize_marathi_text(
            f"सोबत {std_mar} वर्गातील विद्यार्थ्यांकडून त्यांना सगळ्यात जास्त आवडणाऱ्या शिक्षकां संदर्भात माहिती जमा केली. या माहिती साठी बनविलेला चित्रालेख सोबत दिला आहे.<br>"
            f"या चित्रालेखाला अनुसरून स्तंभालेख तयार केला असता पुढील पैकी कोणता स्तंभालेख बरोबर आहे?<br>"
        )
        question_text = f'<p>{eng_question}#{mar_question}{pictograph_html}</p>'
        correct_chart_html = create_option_chart_html(quantities, teachers_eng5, teachers_mar5, scale)
        if not correct_chart_html:
            continue
        wrong_charts_html = []
        for wrong_option in wrong_answers_list:
            wrong_chart_html = create_option_chart_html(
                wrong_option['quantities'], 
                wrong_option['teachers_eng'], 
                wrong_option['teachers_mar'], 
                wrong_option['scale']
            )
            if not wrong_chart_html:
                break
            wrong_charts_html.append(wrong_chart_html)
        if len(wrong_charts_html) != 3:
            continue
        all_charts = [correct_chart_html] + wrong_charts_html
        random.shuffle(all_charts)
        correct_option_letter = option_letters[all_charts.index(correct_chart_html)]
        formatted_options = [f"${letter}$.<br>{chart}"
                             for letter, chart in zip(option_letters, all_charts)]
        correct_answer_formatted = formatted_options[all_charts.index(correct_chart_html)]
        wrong_answers_formatted = [opt for opt in formatted_options if opt != correct_answer_formatted]
        solution = create_solution_text(quantities, teachers_eng5, teachers_mar5, correct_option_letter, scale)
        row_data = [i, "Text", 1, "09030201", question_text, correct_answer_formatted,
                    "", "", "", wrong_answers_formatted[0], wrong_answers_formatted[1], 
                    wrong_answers_formatted[2], 180, 3, "", "siddhimanjarekar274@gmail.com",
                    solution, "", 108]
        rows_to_add.append(row_data)
        i += 1
    if i <= num_questions:
        print(f"\nWarning: Could only generate {i-1} unique questions out of {num_questions} requested.")
    for row in rows_to_add:
        sheet.append(row)
    sheet.append(["****"])
    workbook.save(filename)
    print(f"\nSuccessfully generated {len(rows_to_add)} questions in '{filename}'!")

if __name__ == "__main__":
    main()