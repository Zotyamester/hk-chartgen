#!/usr/bin/python3

import sys

from openpyxl import load_workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList


def get_answer_distribution_for_question(dataset):
    question = dataset[0]
    answers = dataset[1:]

    distribution = {}

    for answer in answers:
        if answer == None:
            continue

        if answer not in distribution:
            distribution[answer] = 0
        distribution[answer] += 1

    return (question, distribution)


def make_charts(inputfile, outputfile):
    workbook = load_workbook(filename=inputfile)

    datasheet = workbook.active

    columns = [column for column in datasheet.iter_cols(values_only=True)]

    distributions = [get_answer_distribution_for_question(
        dataset) for dataset in columns]

    distsheet = workbook.create_sheet("distributions")
    for question, distribution in distributions:
        distsheet.append((question,))
        for answer, count in distribution.items():
            distsheet.append((answer, count))

    sheet_idx = 0
    row_idx = 1
    for question, distribution in distributions:
        chartsheet = workbook.create_chartsheet('diagram%d' % sheet_idx)

        title_row = row_idx  # solely for debug purposes
        entry_count = len(distribution.items())
        first_data_row = row_idx + 1
        last_data_row = row_idx + entry_count

        chart = PieChart()
        labels = Reference(distsheet, min_col=1,
                           min_row=first_data_row, max_row=last_data_row)
        data = Reference(distsheet, min_col=2,
                         min_row=first_data_row, max_row=last_data_row)
        chart.add_data(data)
        chart.set_categories(labels)
        chart.title = question
        chartsheet.add_chart(chart)

        chart.dataLabels = DataLabelList()
        chart.dataLabels.dLblPos = 'outEnd'
        chart.dataLabels.showPercent = True

        sheet_idx += 1
        row_idx += 1 + entry_count

    workbook.save(outputfile)


if __name__ == '__main__':
    inputfile = 'responses.xlsx'
    if len(sys.argv) > 1:
        inputfile = sys.argv[1]
    outputfile = 'analytics.xlsx'
    if len(sys.argv) > 2:
        outputfile = sys.argv[2]
        
    make_charts(inputfile, outputfile)
