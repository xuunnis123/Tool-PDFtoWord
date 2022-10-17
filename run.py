from pdf2docx import Converter
import PySimpleGUI as sg


def pdf2word(file_path):
    file_name = file_path.split('.')[0]
    doc_file = f'{file_name}.docx'
    p2w = Converter(file_path)
    p2w.convert(doc_file, start=0, end=None)
    p2w.close()
    return doc_file

def main():
    # theme setting
    sg.theme('LightBlue5')
    # window setting
    layout = [
        [sg.Text('pdfToword', font=('微軟雅黑', 12)),
         sg.Text('', key='filename', size=(50, 1), font=('微軟雅黑', 10), text_color='blue')],
        [sg.Output(size=(80, 10), font=('微軟雅黑', 10))],
        [sg.FilesBrowse('Select Files', key='file', target='filename'), sg.Button('Start to convert..'), sg.Button('exit')]]
    # windows
    window = sg.Window("Ezra Lin PDF to word", layout, font=("微軟雅黑", 15), default_element_size=(50, 1))

    # 多文件
    while True:

        event, values = window.read()
        print(event, values)

        if event == "Start to convert..":
            # one
            if values['file'] and values['file'].split('.')[1] == 'pdf':
                filename = pdf2word(values['file'])
                print('file total:1')
                print('\n' + 'Successful!' + '\n')
                print('File location:', filename)
            # multi
            elif values['file'] and values['file'].split(';')[0].split('.')[1] == 'pdf':
                print('files total:{}'.format(len(values['file'].split(';'))))
                for f in values['file'].split(';'):
                    filename = pdf2word(f)
                    print('\n' + 'Successful!' + '\n')
                    print('Files location:', filename)
            else:
                print('Please confirm the format of files is PDF!')
        if event in (None, 'exit'):
            break

    window.close()

def __init__(self):
    main()