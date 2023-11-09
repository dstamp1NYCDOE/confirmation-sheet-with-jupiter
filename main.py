import pandas as pd
import glob
import sys

from reportlab.platypus import Paragraph, Spacer, PageBreak, Table, TableStyle, SimpleDocTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape

def main(data):
    week_of = data['week_of']

    rdsc_files = glob.glob(f"data/{week_of}/*RDSC*.xlsx")
    

    df_lst = []
    for rdsc_file in rdsc_files:
        df = pd.read_excel(rdsc_file, skiprows=3)
        df = df.rename(columns={'Student ID':'StudentID','Attd. Date':'Date'})

        dff = df[['StudentID','Student Name','Teacher','Date']]
        dfff = df[['StudentID','Student Name','Teacher.1','Date']]
        dfff = dfff.rename(columns={'Teacher.1':'Teacher'})
        
        df_lst.append(dff)
        df_lst.append(dfff)
    
    df = pd.concat(df_lst)
    df = df.dropna()
    df['Date'] = pd.to_datetime( df['Date'], format="%m/%d/%y")
    


    ## process attendance
    attendance_df = pd.read_csv(f"data/{week_of}/attendance.csv")
    attendance_df['Date'] = pd.to_datetime( attendance_df['Date'])
    attendance_df['Pd'] = attendance_df['Period'].apply(lambda x: x[1:])

    rdsc_students = df['StudentID'].unique()
    attendance_df = attendance_df[attendance_df['StudentID'].isin(rdsc_students)]
    
    temp_lst = []
    for (student, date), attendance_df in attendance_df.groupby(['StudentID','Date']):
 
        attendance_dict = attendance_df[['Pd','Attendance']].set_index('Pd').T.to_dict()
        attendance_dict = { x:y['Attendance'] for (x, y) in attendance_dict.items()}

        attendance_dict['StudentID'] = student 
        attendance_dict['Date'] = date 

        temp_lst.append(attendance_dict)

    parsed_attd_df = pd.DataFrame(temp_lst).fillna('')
    parsed_cols = ['StudentID', 'Date','1','2','3','4','5','6','7','8','9']
    parsed_attd_df = parsed_attd_df[parsed_cols]

    parsed_attd_df = df[['StudentID','Date','Student Name']].merge(
        parsed_attd_df, 
        on=['StudentID','Date'],
        how='left'
    )

    ## build letters 

    styles = getSampleStyleSheet()

    styles.add(ParagraphStyle(name='Normal_RIGHT',
                              parent=styles['Normal'],
                              alignment=TA_RIGHT,
                              ))

    styles.add(ParagraphStyle(name='BodyJustify',parent=styles['BodyText'],alignment=TA_JUSTIFY,))
    letter_head = [
    Paragraph('High School of Fashion Industries',styles['Normal']),
    Paragraph('225 W 24th St',styles['Normal']),
    Paragraph('New York, NY 10011',styles['Normal']),
    Paragraph('Principal, Daryl Blank',styles['Normal']),
    ]

    closing = [
    Spacer(width=0, height=0.25*inch),
    Paragraph('Warmly,',styles['Normal_RIGHT']),
    Paragraph('Derek Stampone',styles['Normal_RIGHT']),
    Paragraph('Assistant Principal',styles['Normal_RIGHT']),
    ]

    directions_txt = """Confirmation of Attendance scan sheets are used to change attendance of students marked absent on the daily attendance roster (Blue Sheets) but marked present for 1 or 2 periods on the SPAT sheets (White Sheets). This may happen because (1) a student arrived after period 3, (2) a student was not in class period 3, or (3) a student was marked present by mistake in 1 or 2 periods. Subject class teachers should confirm the attendance of each student on the sheet by entering absent, late or present. Teachers should sign each sheet."""
    directions_paragraph = Paragraph(directions_txt,styles['BodyText'])

    intro_txt = """Below you will find a list of students per date showing their Jupiter attendance across all periods. What was entered in Jupiter may be different than what was bubbled on the white sheet. This information may be useful for you as you are confirming student attendance."""
    intro_paragraph = Paragraph(intro_txt,styles['BodyText'])

    flowables = []
    for teacher, students_df in df.groupby('Teacher'):
        flowables.extend(letter_head)
        paragraph = Paragraph(f"Dear {teacher},",styles['BodyText'])
        flowables.append(paragraph)

        flowables.append(directions_paragraph)

        flowables.append(Spacer(width=0, height=0.25*inch))

        flowables.append(intro_paragraph)
        flowables.append(Spacer(width=0, height=0.25*inch))

        for date, students_dff in students_df.groupby('Date'):            
            students_lst = students_dff['StudentID']
            student_attd_df = parsed_attd_df[(parsed_attd_df['Date'] == date) & (parsed_attd_df['StudentID'].isin(students_lst))]
            student_attd_df = student_attd_df.drop_duplicates()

            paragraph = Paragraph(f"{date.strftime("%d-%b-%Y")}",styles['BodyText'])
            flowables.append(paragraph)    
            attd_grid_flowable = return_attd_grid_as_table(student_attd_df, ['Student Name','1','2','3','4','5','6','7','8','9'])
            flowables.append(attd_grid_flowable)
            flowables.append(Spacer(width=0, height=0.25*inch))

        flowables.extend(closing)

        flowables.append(PageBreak())
    
    filename = f'output//{week_of}_Confirmation_Cover_Sheets.pdf'
    my_doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        topMargin=0.50*inch,
        leftMargin=1.25*inch,
        rightMargin=1.25*inch,
        bottomMargin=0.25*inch
        )
    my_doc.build(flowables)


    ## spreadsheet of expected confirmation sheets per teacher 

    confirmation_sheets_per_teacher_df = pd.pivot_table(
        df,
        columns='Date',
        index='Teacher',
        values='StudentID',
        aggfunc='count'
    ).fillna(0)
    for col in confirmation_sheets_per_teacher_df.columns:
        confirmation_sheets_per_teacher_df[col] = confirmation_sheets_per_teacher_df[col] > 0
    
    confirmation_sheets_per_teacher_df['TotalSheets'] = confirmation_sheets_per_teacher_df.sum(axis=1)
    
    filename = f'output//{week_of}_Confirmation_Sheets_Tracker.xlsx'
    writer = pd.ExcelWriter(filename)
    confirmation_sheets_per_teacher_df.to_excel(writer, sheet_name=week_of)

    writer.close()

    return True

def return_attd_grid_as_table(df, cols):
    table_data = df[cols].values.tolist()
    table_data.insert(0, cols)
    attendance_col_widths = [2.5*inch] + 9*[0.25*inch]
    t = Table(table_data, colWidths=attendance_col_widths,
              repeatRows=1, rowHeights=None)
    t.setStyle(TableStyle([
        ('ALIGN', (0, 0), (100, 100), 'CENTER'),
        ('VALIGN', (0, 0), (100, 100), 'MIDDLE'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('LEFTPADDING', (0, 0), (100, 100), 1),
        ('RIGHTPADDING', (0, 0), (100, 100), 1),
        ('BOTTOMPADDING', (0, 0), (100, 100), 1),
        ('TOPPADDING', (0, 0), (100, 100), 1),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), (0xD0D0FF, None)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
    ]))
    return t



if __name__ == "__main__":
    try:
        week_of = sys.argv[1]
    except:
        week_of = input("What is the week of you are running the report? YYYY_MM_DD ")
    
    data = {
        'week_of':f"Week_of_{week_of}"
    }
    main(data)
