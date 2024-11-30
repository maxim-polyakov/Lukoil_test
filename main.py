from spire.doc import *
from spire.doc.common import *
import pandas as pd
import json

def read_json():
    # Open and read the JSON file
    with open('config.json', encoding='utf-8') as file:
        data = json.load(file)
    return data
def csv(json):
    csv = pd.read_csv(json['csvFilePath'])
    return csv
def document(df, json):
#
#
    # Create a Document object
    n = len(df.index)
    doc = Document()

    # Add a section
    section = doc.AddSection()
    # Create a table

    titleParagraph = section.AddParagraph()
    titleParagraph.AppendText(str(json['documentTitle']))
    titleParagraph.ApplyStyle(BuiltinStyle.Heading1)
    titleParagraph.Format.HorizontalAlignment = HorizontalAlignment.Center

    table = section.AddTable(True)

    # Set the row number and column number of table

    table.ResetCells(n + 2, 4)

    # Set the width of table

    table.PreferredWidth = PreferredWidth(WidthType.Percentage, int(100))


    # Set row height

    for i in range(0, table.Rows.Count):

        table.Rows[i].Height = 20.0

    # Horizontally merge cells

    table.ApplyHorizontalMerge(0, 0, 3)

    # Get a cell


    cell = table.Rows[1].Cells[0]
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    paragraph = cell.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.AppendText('№')


    cell = table.Rows[1].Cells[1]
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    paragraph = cell.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.AppendText(df.columns[0])


    cell = table.Rows[1].Cells[2]
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    paragraph = cell.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.AppendText(df.columns[1])


    cell = table.Rows[1].Cells[3]
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    paragraph = cell.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.AppendText(df.columns[2])


    cell = table.Rows[0].Cells[0]
    cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
    paragraph = cell.AddParagraph()
    paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
    paragraph.AppendText('Протокол')

    for i in range(2, n + 2):
        cell = table.Rows[i].Cells[0]
        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
        paragraph = cell.AddParagraph()
        paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
        paragraph.AppendText(str(i - 1))

        cell = table.Rows[i].Cells[1]
        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
        paragraph = cell.AddParagraph()
        paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
        paragraph.AppendText(str(df['Фамилия'][i - 2]))

        cell = table.Rows[i].Cells[2]
        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
        paragraph = cell.AddParagraph()
        paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
        paragraph.AppendText(str(df['Имя'][i - 2]))

        cell = table.Rows[i].Cells[3]
        cell.CellFormat.VerticalAlignment = VerticalAlignment.Middle
        paragraph = cell.AddParagraph()
        paragraph.Format.HorizontalAlignment = HorizontalAlignment.Center
        paragraph.AppendText(str(df['Оценка'][i - 2]))

    bodyParagraph_1 = section.AddParagraph()
    bodyParagraph_1.AppendText(' ')
    # Split the cell into 3 smaller cells


    bodyParagraph_2 = section.AddParagraph()
    bodyParagraph_2.AppendText('Обучение провел ' + str(json['employee']['position']) +
                               ' ' + str(json['employee']['lastName']) +
                               ' ' + str(json['employee']['firstName'][0]) + '.' +
                               ' ' + str(json['employee']['middleName'][0]) + '.' +
                               ' ' + '______________(подпись)')


    doc.SaveToFile("output/Протокол обучения пользователей №1.docx", FileFormat.Docx2013)

    doc.Close()


if __name__ == "__main__":
    json = read_json()
    df = csv(json)
    document(df=df, json=json)





