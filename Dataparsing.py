import pdfplumber
import re
import pandas as pd

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text()
    except Exception as e:
        print(f"Error extracting text from PDF: {e}")
    return text

def parse_data(text):
    # Define patterns for extraction
    date_pattern = r'\b(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+\w+\s+\d{2},\s+\d{4}\b'
    case_number_pattern = r'\bCV-\d{6}-\d{2}/[A-Z]{2}\b'
    calendar_number_pattern = r'Calendar Number -([^/]+)'
    calendar_sequence_number_pattern = r'Calendar Sequence.*?-(.*)'
    plaintiff_attorney_pattern = r'Plaintiff Attorney: ([A-Za-z &.,]*)'
    defendant_attorney_pattern = r'Defendant Attorney:([^\n]*)'
    case_description_pattern = r'CV-\d{6}-\d{2}/\w{2} - (.*)'
    court_part_pattern = r'(?<=Part - ).+'
    judge_pattern = r'Judge -([^\n]*)'
    on_for_trial_pattern = r'On For - ([^/]+)'
    appearance_time_pattern = r'(?<=Appearance Time )\d{2}:\d{2}'
    court_name_pattern = r'.+(?= / Part)'

    # Initialize data storage
    data = {
        'Date': [],
        'Case Number': [],
        'Calendar Number': [],
        'Calendar Sequence Number': [],
        'Plaintiff Attorney': [],
        'Defendant Attorney': [],
        'Case Description': [],
        'Court Part': [],
        'Judge': [],
        'On For - Trial': [],
        'Appearance Time': [],
        'Court Name': []
    }

    # Split text by dates and process each section
    sections = re.split(f'({date_pattern})', text)[1:]  # Include date as split key, skip empty sections

    current_date = None
    for i in range(0, len(sections), 2):
        date = sections[i].strip()
        section_text = sections[i + 1].strip()

        if re.match(date_pattern, date):
            current_date = date
        
        if current_date:
            # Extract data for the current section
            case_numbers = re.findall(case_number_pattern, section_text)
            calendar_numbers = re.findall(calendar_number_pattern, section_text)
            calendar_sequence_numbers = re.findall(calendar_sequence_number_pattern, section_text)
            plaintiff_attorneys = re.findall(plaintiff_attorney_pattern, section_text)
            defendant_attorneys = re.findall(defendant_attorney_pattern, section_text)
            case_descriptions = re.findall(case_description_pattern, section_text)
            court_parts = re.findall(court_part_pattern, section_text)
            judges = re.findall(judge_pattern, section_text)
            on_for_trials = re.findall(on_for_trial_pattern, section_text)
            appearance_times = re.findall(appearance_time_pattern, section_text)
            court_names = re.findall(court_name_pattern, section_text)

            # Use the maximum length of all lists to determine how many entries to create
            max_length = max(
                len(case_numbers),
                len(calendar_numbers),
                len(calendar_sequence_numbers),
                len(plaintiff_attorneys),
                len(defendant_attorneys),
                len(case_descriptions),
                len(court_parts),
                len(judges),
                len(on_for_trials),
                len(appearance_times),
                len(court_names)
            )

            # Extend each list to match max_length and maintain their association
            case_numbers += [''] * (max_length - len(case_numbers))
            calendar_numbers += [''] * (max_length - len(calendar_numbers))
            calendar_sequence_numbers += [''] * (max_length - len(calendar_sequence_numbers))
            plaintiff_attorneys += [''] * (max_length - len(plaintiff_attorneys))
            defendant_attorneys += [''] * (max_length - len(defendant_attorneys))
            case_descriptions += [''] * (max_length - len(case_descriptions))
            court_parts += [''] * (max_length - len(court_parts))
            judges += [''] * (max_length - len(judges))
            on_for_trials += [''] * (max_length - len(on_for_trials))
            appearance_times += [''] * (max_length - len(appearance_times))
            court_names += [''] * (max_length - len(court_names))

            # Extend data dictionary with lists
            data['Date'].extend([current_date] * max_length)
            data['Case Number'].extend(case_numbers)
            data['Calendar Number'].extend(calendar_numbers)
            data['Calendar Sequence Number'].extend(calendar_sequence_numbers)
            data['Plaintiff Attorney'].extend(plaintiff_attorneys)
            data['Defendant Attorney'].extend(defendant_attorneys)
            data['Case Description'].extend(case_descriptions)
            data['Court Part'].extend(court_parts)
            data['Judge'].extend(judges)
            data['On For - Trial'].extend(on_for_trials)
            data['Appearance Time'].extend(appearance_times)
            data['Court Name'].extend(court_names)

    return data

def save_to_excel(data, output_path):
    try:
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False)
    except ValueError as ve:
        print(f"ValueError saving data to Excel: {ve}")
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

def main(pdf_path, output_path):
    text = extract_text_from_pdf(pdf_path)
    if text:
        parsed_data = parse_data(text)
        save_to_excel(parsed_data, output_path)
    else:
        print("No text extracted from the PDF.")

# Specify the path to your PDF file and the desired output Excel file
pdf_path = 'Sanders-E-court-Date-Attorney-Aug-Dec-2024.pdf'
output_path = 'Sanders-E-court-Date-Attorney-Aug-Dec-2024(new).xlsx'

main(pdf_path, output_path)
