import comtypes.client
from win32com import client

file_path = r"C:/Users/Administrator/Desktop/E/"
out_path = r"C:/Users/Administrator/Desktop/B/"

# def convert_word_to_pdf(inputFile, outputFile):
#     ''' the following lines that are commented out are items that others shared with me they used when 
#     running loops to stop some exceptions and errors, but I have not had to use them yet (knock on wood) '''
#     word = comtypes.client.CreateObject('Word.Application')
#     doc = word.Documents.Open(inputFile)
#     doc.SaveAs(outputFile, FileFormat = 17)
#     doc.close()
#     word.visible = False
#     word.Quit()

# def convert_word_to_pdf(inputFile, outputFile):
#     word = client.DispatchEx("Word.Application")
#     worddoc = word.Documents.Open(inputFile ,ReadOnly = 1)
#     worddoc.SaveAs(outputFile, FileFormat = 17)
#     worddoc.Close()

# def main():
#     for name in os.listdir(file_path):
#         print(name)
#         file = file_path + name
#         outfile = out_path + name[:-5] +'.pdf'
#         if file.split(".")[-1] == 'docx':
#             convert_word_to_pdf(file,outfile)
#         print("^"*30) 