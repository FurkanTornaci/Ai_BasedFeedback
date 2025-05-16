import json
import time
import openai
from itertools import product

# === Set your OpenAI API key ===
openai.api_key = "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"

# === Question text and correct answers (Q1–Q9) ===
question_texts = {
    "Q1": "The term 'power-driven vessel' refers to any vessel __________.",
    "Q2": "A vessel 'restricted in her ability to maneuver' is one which __________.",
    "Q3": "Which statement is TRUE, according to the Rules?",
    "Q4": "When two power-driven vessels are meeting head-on and there is a risk of collision, which action is required to be taken?",
    "Q5": "Which describes a head-on situation?",
    "Q6": "Your vessel is underway in reduced visibility. You hear the fog signal of another vessel about 30° on your starboard bow. If danger of collision exists, which action(s) are you required to take?",
    "Q7": "You have sighted three red lights in a vertical line on another vessel dead ahead at night. Which vessel would display these lights?",
    "Q8": "By radar alone, you detect a vessel ahead on a collision course, about 3 miles distant. Which action should you take?",
    "Q9": "A vessel may exhibit lights other than those prescribed by the Rules as long as the additional lights meet which requirement(s)?"
}

correct_answers = {
    "Q1": 'with propelling machinery in use',
    "Q2": 'from the nature of her work is unable to maneuver as required by the rules',
    "Q3": 'A vessel engaged in fishing while underway shall, so far as possible, keep out of the way of a vessel restricted in her ability to maneuver.',
    "Q4": 'Both vessels alter course to starboard',
    "Q5": 'Seeing both sidelights of a vessel directly ahead',
    "Q6": 'Reduce your speed to the minimum at which it can be kept on course',
    "Q7": 'A vessel constrained by her draft',
    "Q8": 'Turn the vessel to starboard',
    "Q9": 'do not impair the visibility or distinctive character of the prescribed lights'
}

wrong_answer = "Unclear or incorrect choice"

# === Build the user prompt ===
def build_prompt(case_data):
    bullet_layout = ""
    for qkey, response in case_data.items():
        bullet_layout += (
            f"- Question: {response['text']}\n"
            f"  Correct Answer: {correct_answers[qkey]}\n"
            f"  Response: {response['feedback']}\n\n"
        )

    prompt = (
        "You are a maritime training officer reviewing a trainee's assessment performance. "
        "Based on their answers, write clear, supportive feedback in 3–4 short paragraphs (maximum 500 words). "
        "Do not include question numbers or scores.\n\n"
        "Below are the trainee's responses:\n\n" + bullet_layout + "\n\n" +
        "Analyze the types of questions they got wrong and identify common themes. Write your feedback as a summary: "
        "recognize strengths first, then discuss areas needing improvement, and finally offer motivating advice.\n\n"
        "Mention specific COLREG rules they should revisit, but do not repeat the original questions or answers. "
        "Keep the tone practical, encouraging, and suitable for a seafarer preparing for bridge duties."
    )
    return prompt

# === Generate all possible combinations for Q1–Q9 ===
combinations = list(product([0, 1], repeat=9))
results = {}

# === Main processing loop ===
for idx, combo in enumerate(combinations, start=1):
    case_id = f"case_{idx:03}"
    binary = ''.join(str(bit) for bit in combo)
    responses = {}

    print(f"Processing {case_id} ({idx}/{len(combinations)})...")  # Progress log

    for i, is_correct in enumerate(combo, start=1):
        q_key = f"Q{i}"
        responses[q_key] = {
            "text": question_texts[q_key],
            "answer": correct_answers[q_key] if is_correct else wrong_answer,
            "feedback": "Correct" if is_correct else "Incorrect"
        }

    prompt = build_prompt(responses)

    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo-0125",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=600
        )
        feedback = response.choices[0].message.content.strip()
        time.sleep(1.2)  # Respect API rate limits
    except Exception as e:
        feedback = f"⚠️ Error: {str(e)}"

    results[case_id] = {
        "binary": binary,
        "responses": responses,
        "feedback": feedback
    }

# === Save the results to a JSON file ===
with open("approachB_feedback_Q1_Q9.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("✅ All 512 feedback cases saved to 'approachB_feedback_Q1_Q9.json'")
