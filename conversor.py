import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QMessageBox
from openpyxl import Workbook
from ofxparse import OfxParser

class OFXToExcelConverter(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 400, 200)
        self.setWindowTitle('OFX to Excel Converter')

        self.openButton = QPushButton('Abrir arquivo OFX', self)
        self.openButton.setGeometry(50, 50, 300, 50)
        self.openButton.clicked.connect(self.openOFXFile)

    def openOFXFile(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, 'Abrir Arquivo OFX', '', 'OFX Files (*.ofx);;All Files (*)', options=options)

        if filePath:
            try:
                ofx = OfxParser.parse(open(filePath, 'rb'))
                workbook = Workbook()
                worksheet = workbook.active
                worksheet.title = 'Transações' 

                headers = ['Data', 'Descrição', 'Tipo', 'Valor']
                worksheet.append(headers)

                for transaction in ofx.account.statement.transactions:
                    row = [transaction.date.strftime('%Y-%m-%d'), transaction.memo, transaction.type, transaction.amount]
                    worksheet.append(row)

                # Definir a largura das colunas para 25 unidades
                for column in worksheet.columns:
                    max_length = 0
                    column = [cell for cell in column]
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

                excelFilePath, _ = QFileDialog.getSaveFileName(self, 'Salvar Arquivo Excel', '', 'Excel Files (*.xlsx);;All Files (*)', options=options)

                if excelFilePath:
                    workbook.save(excelFilePath)
                    QMessageBox.information(self, 'Conversão Concluída', 'O arquivo Excel foi criado com sucesso.')

            except Exception as e:
                QMessageBox.critical(self, 'Erro', f'Ocorreu um erro ao converter o arquivo OFX: {str(e)}')

def main():
    app = QApplication(sys.argv)
    window = OFXToExcelConverter()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
