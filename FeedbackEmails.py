import win32com.client as win32
import re

# === Step 1: Real question texts ===
question_texts = {
    'Q1': 'The term "power-driven vessel" refers to any vessel __________.',
    'Q2': 'A vessel "restricted in her ability to maneuver" is one which __________.',
    'Q3': 'Which statement is TRUE, according to the Rules?',
    'Q4': 'When two power-driven vessels are meeting head-on and there is a risk of collision, which action is required to be taken?',
    'Q5': 'Which describes a head-on situation?',
    'Q6': 'Your vessel is underway in reduced visibility. You hear the fog signal of another vessel about 30° on your starboard bow. If danger of collision exists, which action(s) are you required to take?',
    'Q7': 'You have sighted three red lights in a vertical line on another vessel dead ahead at night. Which vessel would display these lights?',
    'Q8': 'By radar alone, you detect a vessel ahead on a collision course, about 3 miles distant. Your radar plot shows this to be a meeting situation. Which action should you take?',
    'Q9': 'A vessel may exhibit lights other than those prescribed by the Rules as long as the additional lights meet which requirement(s)?',
    'Q10': 'Which signal should be used to indicate that your vessel is in distress?'
}

# === Step 2: Define correct answers ===
correct_answers = {
    'Q1': 'with propelling machinery onboard whether in use or not',
    'Q2': 'from the nature of her work is unable to maneuver as required by the rules',
    'Q3': 'A vessel engaged in fishing while underway shall, so far as possible, keep out of the way of a vessel restricted in her ability to maneuver.',
    'Q4': 'Both vessels alter course to starboard',
    'Q5': 'Seeing both sidelights of a vessel directly ahead',
    'Q6': 'Reduce your speed to the minimum at which it can be kept on course',
    'Q7': 'A vessel constrained by her draft',
    'Q8': 'Turn the vessel to starboard',
    'Q9': 'do not impair the visibility or distinctive character of the prescribed lights',
    'Q10': 'Continuously sounding the fog whistle'
}

# === Step 3: Simulated user responses (replace these with actual response data) ===
user_answers = {
    'Q1': 'with propelling machinery onboard whether in use or not',
    'Q2': 'due to adverse weather conditions is unable to maneuver as required by the rules',
    'Q3': 'A vessel engaged in fishing while underway shall, so far as possible, keep out of the way of a vessel restricted in her ability to maneuver.',
    'Q4': 'Both vessels alter course to starboard',
    'Q5': 'Seeing both sidelights of a vessel directly ahead',
    'Q6': 'Reduce your speed to the minimum at which it can be kept on course',
    'Q7': 'A vessel dredging',
    'Q8': 'Maintain course and speed and sound at least five short blasts of the whistle',
    'Q9': 'All of the above',
    'Q10': 'Continuously sounding the fog whistle'
}

# === Step 4: All options per question ===
all_options = {
    'Q1': [
        'with propelling machinery onboard whether in use or not',
        'traveling at a speed greater than that of the current',
        'with propelling machinery in use',
        'making way against the current'
    ],
    'Q2': [
        'from the nature of her work is unable to maneuver as required by the rules',
        'due to adverse weather conditions is unable to maneuver as required by the rules',
        'through some exceptional circumstance is unable to maneuver as required by the rules',
        'has lost steering and is unable to maneuver'
    ],
    'Q3': [
        'A vessel engaged in fishing while underway shall, so far as possible, keep out of the way of a vessel restricted in her ability to maneuver.',
        'A vessel constrained by her draft shall keep out of the way of a vessel engaged in fishing.',
        'A vessel not under command shall avoid impeding the safe passage of a vessel constrained by her draft.',
        'A vessel not under command shall keep out of the way of a vessel restricted in her ability to maneuver.'
    ],
    'Q4': [
        'Sound at least five short and rapid blasts',
        'Back down',
        'Both vessels alter course to starboard',
        'Both vessels shall stop their engines'
    ],
    'Q5': [
        'Seeing two forward white towing lights in a vertical line on a towing vessel directly ahead',
        'Seeing one red light of a vessel directly ahead',
        'Seeing both sidelights of a vessel directly ahead',
        'Seeing both sidelights of a vessel directly off your starboard beam'
    ],
    'Q6': [
        'Reduce your speed to the minimum at which it can be kept on course',
        'Slow your engines and let the other vessel pass ahead of you',
        'Alter course to port and pass the other vessel on its port side',
        'Alter course to starboard to pass around the other vessel\'s stern'
    ],
    'Q7': [
        'A vessel aground',
        'A vessel dredging',
        'A vessel moored over a wreck',
        'A vessel constrained by her draft'
    ],
    'Q8': [
        'Turn the vessel to starboard',
        'Maintain course and speed and do not sound any whistle signals',
        'Maintain course and speed and sound at least five short blasts of the whistle',
        'Turn the vessel to port'
    ],
    'Q9': [
        'do not impair the visibility or distinctive character of the prescribed lights',
        'have a lesser range of visibility than the prescribed lights',
        'are not the same color as either sidelight',
        'All of the above'
    ],
    'Q10': [
        'Sounding four or more short rapid blasts on the whistle',
        'Displaying a large red flag',
        'Displaying three black balls in a vertical line',
        'Continuously sounding the fog whistle'
    ]
}

# === Step 5: Construct replacement fields ===
fields = {}
for qid, options in all_options.items():
    correct = correct_answers[qid]
    selected = user_answers[qid]
    is_correct = selected == correct

    fields[f'{qid}_text'] = question_texts[qid]
    fields[f'{qid}_feedback_icon'] = '<p style="color: green;">✔ Correct</p>' if is_correct else '<p style="color: red;">✖ Incorrect</p>'
    fields[f'{qid}_feedback_message'] = (
        '<p style="margin-top: 16px; font-size: 16px; color: #28a745;"><strong>🎉 Congratulations! You answered correctly.</strong></p>'
        if is_correct else
        f'<div style="margin-top: 20px; padding: 12px 16px; background-color: #f7f7f7; border-left: 4px solid #28a745; border-radius: 6px;"><p style="margin: 0; font-size: 16px;"><strong style="color: #333;">Correct Answer:</strong> <span style="color: #28a745; font-weight: bold;">{correct}</span></p></div>'
    )

    for idx, opt in enumerate(options):
        label_key = f'{qid}_opt{idx+1}'
        style_key = f'{qid}_opt{idx+1}_style'

        if opt == selected and is_correct:
            style = '#d4edda'
        elif opt == selected and not is_correct:
            style = '#f8d7da'
        elif opt == correct and not is_correct:
            style = '#d4edda'
        else:
            style = 'transparent'

        fields[label_key] = opt
        fields[style_key] = style

# === Step 6: Load improved HTML template with explanation ===
with open('untitled8.html', 'r', encoding='utf-8') as f:
    html_body = f.read()

# Replace placeholders
for key, value in fields.items():
    html_body = html_body.replace('${e://Field/' + key + '}', value)

# Clean up unused placeholders
html_body = re.sub(r'\${e://Field/[^}]+}', '', html_body)

# === Step 7: Send email via Outlook ===
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'ipekgolbol@gmail.com'
mail.Subject = 'Your Personalized COLREGs Feedback Report'
mail.HTMLBody = html_body
mail.Send()

print("Email sent successfully.")
