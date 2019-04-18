from __future__ import print_function
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML, CSS

__author__ = 'Richard Craggs'


# Mostly taken from this example - https://pbpython.com/pdf-reports.html

env = Environment(loader=FileSystemLoader('.'))
template = env.get_template("template/student_report_template.html")

# Load the data frames from the spreadsheet
df_attainment = pd.read_excel("data/example_sg_course.xlsx", sheet_name="Attainment")
df_achievements = pd.read_excel("data/example_sg_course.xlsx", sheet_name="Achievements")

# Get the arrays we need
bundles = df_achievements['Bundle'].unique()
students = df_attainment['Student'].unique()
attainments = df_achievements['Assessment']

df_attainment_with_bundles = df_attainment.merge(df_achievements, on='Assessment')

# For each student, create a report
for student in students:

    student_df = df_attainment_with_bundles[df_attainment_with_bundles['Student'] == student]

    # create a mapping for this student from bundle to the attainments in that bundle
    bundle_df = {elem: pd.DataFrame for elem in bundles}
    for key in bundle_df.keys():
        bundle_df[key] = student_df[:][student_df.Bundle == key]

    print("Generating report for " + student)
    html_out = template.render(student_name=student, student_data=student_df, bundle_names=bundles, bundles=bundle_df)
    HTML(string=html_out).write_pdf(student + "_report.pdf", stylesheets=[CSS('template/bare.min.css')])