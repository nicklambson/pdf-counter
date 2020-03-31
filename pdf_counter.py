# This script takes a folder path like C:\Users\User\Desktop\test and searches
# for only .pdf files in that folder and all subfolders, generating a report that displays
# file path, page count of each pdf, author and pdf generator, and optionally, total word count. 
# Requires the PyPDF2 package and pathlib installed to run


from pathlib import Path, PurePath
import PyPDF2
import csv


page_count = 0
file_count = 0
word_count = 0

dir = input('Paste folder path as text: ') 
folder = Path(dir)
print('')

csv_report = 'PDF_report.csv'
filepath = folder / csv_report


count_words = input('Count words in each pdf? y/n (experimental) ')
print('') # this word counter doesn't work all the time, and wont work for languages like Chinese...

if count_words == 'y':
    count_words = True
    header = ['File', 'Pages', 'Producer', 'Creator', 'Author', 'Word Count']
else:
    count_words = False
    header = ['File', 'Pages', 'Producer', 'Creator', 'Author']
    


with open(filepath, 'w', encoding='utf-16', newline='') as csv_report:
    writer = csv.writer(csv_report, delimiter='\t') 
    writer.writerow(header)
    
    for file in folder.glob('**/*.pdf'):
        csv_row = []
        
        pdf = file.absolute()
        print('File:', pdf)
        file_count += 1
        csv_row.append(pdf)
              
        with open(pdf,'rb') as FileObj:
            pdfReader = PyPDF2.PdfFileReader(FileObj)
            pages = pdfReader.numPages
            print('pages:', pages)
            page_count += pages
            csv_row.append(pages)
            stringObj = " "

            info = pdfReader.getDocumentInfo()
            producer = info.producer
            csv_row.append(producer)
            creator = info.creator
            csv_row.append(creator)
            author = info.author
            csv_row.append(author)
            
            if count_words == True:
                print("counting words...")
                for page in range(pdfReader.numPages): 
                    stringObj += pdfReader.getPage(page).extractText()
                    text = stringObj.split(' ')

                    for word in text:
                        if word == (''):
                            text.remove(word)

                if text != ['']:            
                    csv_row.append(len(text))
                    word_count += len(text)
                    print('words:',len(text))
                    
                else:
                    print('words: 0')
                    csv_row.append('0')
            else:
                pass
      
                
        print('')
        #print(split)
        writer.writerow(csv_row)
        
print('-=-=-=-=-=-=-=-=-=-')
print('Results:')
print('Total PDFs:', file_count)
print('Total pages:', page_count)
print('Total words:', word_count)
print('Report generated in root folder!')
print(filepath)
print('-=-=-=-=-=-=-=-=-=-')

csv_report.close()


'''
filepath = folder / name
with filepath.open('w', encoding="utf-8") as f:
    f.write('Page count: %s \n' % count)
    f.write('Total PDFs: %s' % file_count)
f.close()
'''


