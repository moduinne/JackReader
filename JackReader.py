from docx import Document
#returns all the paragraphs in the document by order
def obtainText(docFileName):
    document = Document(docFileName)
    finalText = []
    for line in document.paragraphs:
        finalText.append(line.text)
    return '\n'.join(finalText)

#returns all the tables in the document by order
def obtainTables(docFileName):
    document = Document(docFileName)
    finalTables = []
    for table in document.tables:
        finalTables.append(table)
    return finalTables

#returns a give tables rowList as a dictionaries
def returnTableRowList(table):
    data =  []
    keys = None
    for i, row in enumerate(table.rows):
        text = (cell.text for cell in row.cells)
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys,text))
        data.append(row_data)
    return data

def returnNumColDictionary(table):
    heads = []
    td = {}
    col = 0
    for i in table:
        heads.append(i)
    while col < len(table) :
        td[col] = heads[col]
        col += 1
    return td

for i in obtainTables("Mock.docx"):
    print(returnTableRowList(i))
    #print (i)

