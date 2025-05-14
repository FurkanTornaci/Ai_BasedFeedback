import pandas as pd
import win32com.client as win32
import re

# === Step 1: Load CSV and clean columns ===
file_path = "Clean_with_Feedback_Cleaned.csv"
df = pd.read_csv(file_path, encoding='windows-1252')


# Clean column names: strip whitespace and convert to lowercase
df.rename(columns=lambda x: x.strip().lower(), inplace=True)

# Remove rows without email
df.dropna(subset=['email'], inplace=True)

# === Step 2: Define correct answers ===
correct_answers = {
    'q1': 'with propelling machinery onboard whether in use or not',
    'q2': 'from the nature of her work is unable to maneuver as required by the rules',
    'q3': 'A vessel engaged in fishing while underway shall, so far as possible, keep out of the way of a vessel restricted in her ability to maneuver.',
    'q4': 'Both vessels alter course to starboard',
    'q5': 'Seeing both sidelights of a vessel directly ahead',
    'q6': 'Reduce your speed to the minimum at which it can be kept on course',
    'q7': 'A vessel constrained by her draft',
    'q8': 'Turn the vessel to starboard',
    'q9': 'do not impair the visibility or distinctive character of the prescribed lights',
    'q10': 'Continuously sounding the fog whistle'
}

# === Step 3: Define full question texts ===
question_texts = {
    'q1': 'The term "power-driven vessel" refers to any vessel __________.',
    'q2': 'A vessel "restricted in her ability to maneuver" is one which __________.',
    'q3': 'Which statement is TRUE, according to the Rules?',
    'q4': 'When two power-driven vessels are meeting head-on and there is a risk of collision, which action is required to be taken?',
    'q5': 'Which describes a head-on situation?',
    'q6': 'Your vessel is underway in reduced visibility. You hear the fog signal of another vessel about 30° on your starboard bow. If danger of collision exists, which action(s) are you required to take?',
    'q7': 'You have sighted three red lights in a vertical line on another vessel dead ahead at night. Which vessel would display these lights?',
    'q8': 'By radar alone, you detect a vessel ahead on a collision course, about 3 miles distant. Your radar plot shows this to be a meeting situation. Which action should you take?',
    'q9': 'A vessel may exhibit lights other than those prescribed by the Rules as long as the additional lights meet which requirement(s)?',
    'q10': 'Which signal should be used to indicate that your vessel is in distress?'
}

# === Step 4: All options per question ===
all_options = {
    'q1': [
        'with propelling machinery onboard whether in use or not',
        'traveling at a speed greater than that of the current',
        'with propelling machinery in use',
        'making way against the current'
    ],
    'q2': [
        'from the nature of her work is unable to maneuver as required by the rules',
        'due to adverse weather conditions is unable to maneuver as required by the rules',
        'through some exceptional circumstance is unable to maneuver as required by the rules',
        'has lost steering and is unable to maneuver'
    ],
    'q3': [
        'A vessel engaged in fishing while underway shall, so far as possible, keep out of the way of a vessel restricted in her ability to maneuver.',
        'A vessel constrained by her draft shall keep out of the way of a vessel engaged in fishing.',
        'A vessel not under command shall avoid impeding the safe passage of a vessel constrained by her draft.',
        'A vessel not under command shall keep out of the way of a vessel restricted in her ability to maneuver.'
    ],
    'q4': [
        'Sound at least five short and rapid blasts',
        'Back down',
        'Both vessels alter course to starboard',
        'Both vessels shall stop their engines'
    ],
    'q5': [
        'Seeing two forward white towing lights in a vertical line on a towing vessel directly ahead',
        'Seeing one red light of a vessel directly ahead',
        'Seeing both sidelights of a vessel directly ahead',
        'Seeing both sidelights of a vessel directly off your starboard beam'
    ],
    'q6': [
        'Reduce your speed to the minimum at which it can be kept on course',
        'Slow your engines and let the other vessel pass ahead of you',
        'Alter course to port and pass the other vessel on its port side',
        'Alter course to starboard to pass around the other vessel\'s stern'
    ],
    'q7': [
        'A vessel aground',
        'A vessel dredging',
        'A vessel moored over a wreck',
        'A vessel constrained by her draft'
    ],
    'q8': [
        'Turn the vessel to starboard',
        'Maintain course and speed and do not sound any whistle signals',
        'Maintain course and speed and sound at least five short blasts of the whistle',
        'Turn the vessel to port'
    ],
    'q9': [
        'do not impair the visibility or distinctive character of the prescribed lights',
        'have a lesser range of visibility than the prescribed lights',
        'are not the same color as either sidelight',
        'All of the above'
    ],
    'q10': [
        'Sounding four or more short rapid blasts on the whistle',
        'Displaying a large red flag',
        'Displaying three black balls in a vertical line',
        'Continuously sounding the fog whistle'
    ]
}


traditional_explanations = {
    'q1': 'A power-driven vessel is defined as any vessel being propelled by machinery that is actively engaged.',
    'q2': 'A vessel is considered “restricted in her ability to maneuver” when her operations limit her ability to move according to standard navigation rules.',
    'q3': 'In hierarchy of vessel priorities, a vessel restricted in her ability to maneuver has precedence over a fishing vessel.',
    'q4': 'In a head-on situation, both power-driven vessels must alter course to starboard to avoid collision.',
    'q5': 'Seeing both sidelights of a vessel directly ahead indicates a head-on approach.',
    'q6': 'Rule 19 requires vessels to reduce speed to the minimum at which it can be kept on course during restricted visibility.',
    'q7': 'Three red lights in a vertical line indicate a vessel engaged in underwater operations such as dredging.',
    'q8': 'When radar shows a meeting situation, standard COLREGs apply—turn to starboard to avoid collision.',
    'q9': 'Additional lights are allowed if they do not impair the visibility or character of the required navigation lights.',
    'q10': 'A continuous fog whistle is an internationally recognized sound signal for distress.'
}

# === Step 5: Load HTML feedback template ===
with open('untitled8.html', 'r', encoding='utf-8') as f:
    template_html = f.read()

# === Step 6: Initialize Outlook ===
outlook = win32.Dispatch('outlook.application')

# === Step 7: Loop through each participant ===
for index, row in df.iterrows():
    email = row['email']
    if pd.isna(email):
        continue

    # Retrieve participantId and feedbackType
    participant_id = str(row.values[0])
    feedback_type = str(row.get('feedbacktype', '999'))
    feedback_link = f"https://stratheng.eu.qualtrics.com/jfe/form/SV_3kocAh1wEfLTXxQ?participantId={participant_id}&feedbackType={feedback_type}"

    user_answers = {qid: row[qid] for qid in correct_answers}
    feedback_raw = str(row['feedbackgenerated']).strip().replace('\n', '<br><br>')
    overall_feedback = f"<p style='font-size: 15px; color: #333;'>{feedback_raw}</p>"

    fields = {
        'overall_feedback_AI': overall_feedback,
        'feedback_link': feedback_link
    }

    for qid, options in all_options.items():
        qid_uc = qid.upper()
        correct = correct_answers[qid]
        selected = user_answers[qid]
        is_correct = selected == correct

        fields[f'{qid_uc}_text'] = question_texts[qid]
        fields[f'{qid_uc}_feedback_icon'] = '<p style="color: green;">✔ Correct</p>' if is_correct else '<p style="color: red;">✖ Incorrect</p>'
        fields[f'{qid_uc}_feedback_message'] = (
            '<p style="margin-top: 16px; font-size: 16px; color: #28a745;"><strong>🎉 Congratulations! You answered correctly.</strong></p>'
            if is_correct else
            f'<div style="margin-top: 20px; padding: 12px 16px; background-color: #f7f7f7; border-left: 4px solid #28a745; border-radius: 6px;"><p style="margin: 0; font-size: 16px;"><strong style="color: #333;">Correct Answer:</strong> <span style="color: #28a745; font-weight: bold;">{correct}</span></p></div>'
        )

        for idx, opt in enumerate(options):
            label_key = f'{qid_uc}_opt{idx+1}'
            style_key = f'{qid_uc}_opt{idx+1}_style'

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

    # Fill the template with the feedback
    html_body = template_html
    for key, value in fields.items():
        html_body = html_body.replace('${e://Field/' + key + '}', value)
    html_body = html_body.replace('${feedback_link}', feedback_link)
    html_body = re.sub(r'\${e://Field/[^}]+}', '', html_body)

    # === Send email ===
    mail = outlook.CreateItem(0)
    mail.To = email
    
    if feedback_type.lower() == 'human':
        subject = "Feedback Approach A – COLREGs Assessment"
    else:
        subject = "Feedback Approach B – COLREGs Assessment"
    
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Send()

    print(f"Email sent to {email} | Link: {feedback_link}")

print("All emails sent.")
