import argparse
from docx import Document
import re

MONTHS = ('Zeroember', 'January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December')


def parse(filename):
    doc = Document(filename)

    paragraphs = doc.paragraphs

    # the first paragraph is the title, "Year in Pixels - [year]"
    year_match = re.match('Year in Pixels - ([0-9]{4,})', paragraphs[0].text.strip())
    assert year_match is not None

    year = int(year_match[1])

    days = list()
    for paragraph in paragraphs[1:]:
        assert len(paragraph.runs) % 3 == 0

        for i in range(0, len(paragraph.runs), 3):

            runs = paragraph.runs[i:i+3]

            date = runs[0].text
            date_match = re.match('([0-9]{1,2}) ([A-Za-z]+)', date.strip())
            assert date_match is not None
            day = int(date_match[1])
            month = MONTHS.index(date_match[2])

            emotions = runs[1].text.split(',')
            emotions = [emotion.strip() for emotion in emotions if emotion.strip() != '']

            body = runs[2].text.strip()

            color = '#' + str(runs[0].font.color.rgb)

            day = dict(month=month, day=day, emotions=emotions, body=body, color=color)
            days.append(day)

    return days

def export_csv(days):
    import csv
    import io

    output = io.StringIO()

    writer = csv.DictWriter(output, fieldnames=('day', 'month', 'emotions', 'body', 'color'))
    for day in days:
        writer.writerow(day)

    output.seek(0)
    return output.read()

def export_json(days):
    import json

    return json.dumps(days, indent='\t')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('file')
    parser.add_argument('-o', '--format', default='json', choices=['csv', 'json'])
    args = parser.parse_args()

    days = parse(args.file)
    if args.format == 'csv':
        print(export_csv(days))
    elif args.format == 'json':
        print(export_json(days))
