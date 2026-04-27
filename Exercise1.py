import random
import pandas as pd
import string

def get_random_labels():
    letters = list(string.ascii_uppercase)
    for c in ['O', 'I']:
        if c in letters:
            letters.remove(c)
    return random.sample(letters, 8)

def generate_svg(labels, transversal_slope):
    A, B, C, D, M, N, P, Q = labels
    
    if transversal_slope == 1:
        mx, my = 150, 100
        nx, ny = 250, 200
        px, py = 100, 50
        qx, qy = 300, 250
    else:
        mx, my = 250, 100
        nx, ny = 150, 200
        px, py = 300, 50
        qx, qy = 100, 250

    svg = f"""<svg width="400" height="300" xmlns="http://www.w3.org/2000/svg">
    <style>
        .line {{ stroke: black; stroke-width: 2; }}
        .text {{ font-family: Arial, sans-serif; font-size: 16px; fill: black; font-weight: bold; }}
        .point {{ fill: red; }}
    </style>
    <line x1="50" y1="100" x2="350" y2="100" class="line" />
    <line x1="50" y1="200" x2="350" y2="200" class="line" />
    <line x1="{px}" y1="{py}" x2="{qx}" y2="{qy}" class="line" />
    
    <circle cx="80" cy="100" r="4" class="point"/>
    <text x="75" y="90" class="text">{A}</text>
    <circle cx="320" cy="100" r="4" class="point"/>
    <text x="315" y="90" class="text">{B}</text>
    <circle cx="{mx}" cy="{my}" r="4" class="point"/>
    <text x="{mx - 15}" y="{my - 10}" class="text">{M}</text>

    <circle cx="80" cy="200" r="4" class="point"/>
    <text x="75" y="190" class="text">{C}</text>
    <circle cx="320" cy="200" r="4" class="point"/>
    <text x="315" y="190" class="text">{D}</text>
    <circle cx="{nx}" cy="{ny}" r="4" class="point"/>
    <text x="{nx + 10}" y="{ny + 20}" class="text">{N}</text>

    <circle cx="{px}" cy="{py}" r="4" class="point"/>
    <text x="{px - 20}" y="{py + 5}" class="text">{P}</text>
    <circle cx="{qx}" cy="{qy}" r="4" class="point"/>
    <text x="{qx + 10}" y="{qy + 5}" class="text">{Q}</text>
</svg>"""
    return svg

def generate_question_bank(n):
    generated_configs = set()
    data = []
    
    while len(data) < n:
        labels = tuple(get_random_labels())
        slope = random.choice([1, -1])
        
        # Configuration uniqueness check
        config_hash = (tuple(sorted(labels)), slope)
        if config_hash in generated_configs:
            continue
        generated_configs.add(config_hash)
        
        A, B, C, D, M, N, P, Q = labels
        
        # Determine logical Left and Right based on X coordinates
        # M is intersection on upper line (AB), N is on lower line (CD)
        left_upper, right_upper = A, B
        left_lower, right_lower = C, D
        
        # Define internal alternate pairs
        # Pair 1: Left Upper Internal & Right Lower Internal
        pair1 = (f"{left_upper}{M}{N}", f"{right_lower}{N}{M}")
        # Pair 2: Right Upper Internal & Left Lower Internal
        pair2 = (f"{right_upper}{M}{N}", f"{left_lower}{N}{M}")
        
        # Randomly choose which pair is given in the question
        if random.choice([True, False]):
            given_pair = pair1
            target_pair = pair2
        else:
            given_pair = pair2
            target_pair = pair1

        # Format Given Pairs and Answer
        given_en = f"∠{given_pair[0]} and ∠{given_pair[1]}"
        given_mr = f"∠{given_pair[0]} आणि ∠{given_pair[1]}"
        
        ans_en = f"∠{target_pair[0]} and ∠{target_pair[1]}"
        ans_mr = f"∠{target_pair[0]} आणि ∠{target_pair[1]}"
        
        # Distractor 1: Same side interior angles (incorrect relationship)
        d1_en = f"∠{given_pair[0]} and ∠{target_pair[1]}"
        d1_mr = f"∠{given_pair[0]} आणि ∠{target_pair[1]}"
        
        # Distractor 2: Corresponding angles
        corresp_angle = f"{P}{M}{left_upper}" if given_pair[0] == f"{left_upper}{M}{N}" else f"{P}{M}{right_upper}"
        d2_en = f"∠{corresp_angle} and ∠{given_pair[1]}"
        d2_mr = f"∠{corresp_angle} आणि ∠{given_pair[1]}"
        
        # Distractor 3: Linear pair / Adjacent Interior
        d3_en = f"∠{target_pair[0]} and ∠{given_pair[0]}"
        d3_mr = f"∠{target_pair[0]} आणि ∠{given_pair[0]}"
        
        # Shuffle Distractors safely
        options_en = [d1_en, d2_en, d3_en]
        options_mr = [d1_mr, d2_mr, d3_mr]
        distractors = list(zip(options_en, options_mr))
        random.shuffle(distractors)
        
        q_en = f"For the figure shown below, one pair of internal alternate angles is {given_en}.\nWhich is the other pair of internal alternate angles?"
        q_mr = f"खाली दाखविलेल्या आकृतीमध्ये, {given_mr} ही आंतर व्युत्क्रम कोनांची एक जोडी आहे.\nआंतर व्युत्क्रम कोनांची दुसरी जोडी कोणती असेल?"
        
        exp_en = f"Given: Two parallel lines cut by a transversal. Internal alternate angles lie between the parallel lines and on opposite sides of the transversal. Since the given pair is {given_en}, the remaining corresponding pair on the opposite sides is {ans_en}."
        exp_mr = f"दिलेले: एका छेदिकेने दोन समांतर रेषांना छेदले आहे. आंतर व्युत्क्रम कोन समांतर रेषांच्या मध्ये आणि छेदिकेच्या विरुद्ध बाजूस असतात. दिलेली जोडी {given_mr} असल्यामुळे, विरुद्ध बाजूंवरील दुसरी संबंधित जोडी {ans_mr} ही आहे."
        
        # SVG
        svg_code = generate_svg(labels, slope)
        
        data.append({
            "SR_NO": len(data) + 1,
            "QUESTION_EN": q_en,
            "QUESTION_MR": q_mr,
            "CORRECT_ANSWER_EN": ans_en,
            "CORRECT_ANSWER_MR": ans_mr,
            "OPTION_1_EN": distractors[0][0],
            "OPTION_1_MR": distractors[0][1],
            "OPTION_2_EN": distractors[1][0],
            "OPTION_2_MR": distractors[1][1],
            "OPTION_3_EN": distractors[2][0],
            "OPTION_3_MR": distractors[2][1],
            "SVG_CODE": svg_code,
            "EXPLANATION_EN": exp_en,
            "EXPLANATION_MR": exp_mr
        })
        
    return data

def export_to_excel(data, filename="Geometry_Internal_Alternate_Angles.xlsx"):
    """Converts the dictionaries to a Pandas DataFrame and exports as Excel."""
    df = pd.DataFrame(data)
    
    # Organize columns logically
    columns_order = [
        "SR_NO", 
        "QUESTION_EN", "QUESTION_MR", 
        "CORRECT_ANSWER_EN", "CORRECT_ANSWER_MR",
        "OPTION_1_EN", "OPTION_1_MR",
        "OPTION_2_EN", "OPTION_2_MR",
        "OPTION_3_EN", "OPTION_3_MR",
        "SVG_CODE",
        "EXPLANATION_EN", "EXPLANATION_MR"
    ]
    df = df[columns_order]
    
    df.to_excel(filename, index=False, engine='openpyxl')
    print(f"Successfully generated {len(data)} questions and saved to '{filename}'.")

if __name__ == "__main__":
    try:
        n = int(input("Enter the number of unique geometry questions to generate: "))
        if n <= 0:
            print("Please enter a positive integer.")
        else:
            q_data = generate_question_bank(n)
            export_to_excel(q_data)
    except ValueError:
        print("Invalid input. Please enter a valid integer.")