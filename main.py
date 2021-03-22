import gspread
import logging
import csv


class Spreadsheet:

    def __init__(self):
        logging.basicConfig(filename='sheets.log', filemode='w', format='%(asctime)s - %(levelname)s - %(message)s',
                            level=logging.INFO)

        # alterar o link da planilha conforme suas credenciais
        self.link_doc = 'https://docs.google.com/spreadsheets/d/1R35f_ayWwfywTSN14PifndqG2YjMOida6zgnnNO0XxI/edit#gid=0'

        try:
            self.gc = gspread.service_account()
            # self.gc = gspread.service_account(filename=r'C:\Users\Alana\Desktop\pythonProject\testsheets-308323'
            #                                          r'-085f8c0d2f56.json')
        except Exception:
            print("ERROR: Spreadsheet access blocked.")
            print("INFO:Verify credentials")
            logging.error("Spreadsheet access blocked.", exc_info=True)
            logging.info("Verify credentials")

    def correction(self):
        self.sh = self.gc.open_by_url(self.link_doc)
        if self.sh:
            print("Access completed in spreadsheet.", self.sh)
            logging.info("Access completed in spreadsheet.")
            logging.info(self.sh)

        self.worksheet = self.sh.sheet1

        self.cell_list_A = self.worksheet.range('A2:A125')
        self.cell_list_B = self.worksheet.range('B2:B125')
        self.cell_list_C = self.worksheet.range('C2:C125')
        self.cell_list_D = self.worksheet.range('D2:D125')

        print("  ")
        print("Verificando Coluna A... ")
        logging.info("Verificando Coluna A... ")
        # verificação de células vazias na coluna A e altereação da célula vaiza para NULL
        for i, val in enumerate(self.cell_list_A):
            if self.cell_list_A[i].value == "":
                print("Coluna A Linha " + str(i + 2) + " Célula está vazia, verificar data, alterado para NULL.")
                logging.warning("Coluna A Linha " + str(i + 2) + " Celula esta vazia, verificar data, alterado para "
                                                                 "NULL.")
                self.cell_list_A[i].value = "NULL"

        # verificação de células vazias na coluna B e alteração da célula vazia para NULL
        # correção da palavra Lava-roupas.
        print("  ")
        print("Verificando Coluna B... ")
        logging.info("Verificando Coluna B... ")
        for i, val in enumerate(self.cell_list_B):
            if self.cell_list_B[i].value == "":
                print("Coluna B Linha " + str(i + 2) + " Célula está vazia, verificar item.")
                logging.warning("Coluna B Linha " + str(i + 2) + " Celula esta vazia, verificar item, alterado para "
                                                                 "NULL.")
                self.cell_list_B[i].value = "NULL"

            elif self.cell_list_B[i].value == "Lava roupas":
                self.cell_list_B[i].value = "Lava-roupas"
                print("Coluna B Linha " + str(i + 2) + " Correção da palavra Lava roupas.")
                logging.info("Coluna B Linha " + str(i + 2) + " Correcao da palavra Lava roupas.")

        # verificação de células vazias na coluna C e alteração da célula vazia para NULL
        # correção do cifrão distorcido no início dos valores
        print("  ")
        print("Verificando Coluna C... ")
        logging.info("Verificando Coluna C... ")
        for i, val in enumerate(self.cell_list_C):
            if self.cell_list_C[i].value == "":
                print("Coluna C Linha " + str(i + 2) + " Célula está vazia, verificar valor.")
                logging.warning("Coluna C Linha " + str(i + 2) + " Celula esta vazia, verificar valor, alterado para "
                                                                 "NULL.")
                self.cell_list_C[i].value = "NULL"

            elif self.cell_list_C[i].value.startswith("$"):
                self.cell_list_C[i].value = self.cell_list_C[i].value
                print("Coluna C Linha " + str(i + 2) + " Correção do $ efetuada ao inicio do valor.")
                logging.info("Coluna C Linha " + str(i + 2) + " Correcao do $ efetuada ao inicio do valor.")

        print("  ")
        print("Verificando Coluna D... ")
        logging.info("Verificando Coluna D... ")

        # verificação de células vazias na coluna D e alteração da célula vazia para NULL
        # correção da símbolo de percentual no valor do imposto.
        for i, val in enumerate(self.cell_list_D):
            if self.cell_list_D[i].value == "":
                print("Coluna D Linha " + str(i + 2) + " Célula está vazia, verificar imposto.")
                logging.warning("Coluna D Linha " + str(i + 2) + " Celula esta vazia, verificar imposto, alterado para "
                                                                 "NULL.")
                self.cell_list_D[i].value = "NULL"

            elif not self.cell_list_D[i].value.endswith("%"):
                self.cell_list_D[i].value = self.cell_list_D[i].value + "%"
                print("Coluna D Linha " + str(i + 2) + " Correção do % efetuada ao final do imposto.")
                logging.info("Coluna D Linha " + str(i + 2) + " Correcao do % efetuada ao final do imposto.")

        # atualiza celulas em cada coluna
        self.worksheet.update_cells(self.cell_list_A)
        self.worksheet.update_cells(self.cell_list_B)
        self.worksheet.update_cells(self.cell_list_C, value_input_option="USER_ENTERED")
        self.worksheet.update_cells(self.cell_list_D)

        self.worksheet_list = self.sh.worksheets()
        return self.worksheet_list

    def save_csv(self):
        self.correction()
        for i, worksheet in enumerate(self.worksheet_list):
            filename_sheets = "Corrigido - Prova Engenharia de Dados - Planilha de compras.csv"
            with open(filename_sheets, 'w', newline='') as file:
                writer = csv.writer(file)
                writer.writerows(worksheet.get_all_values())

        print("Correction completed.")
        print("Access the file: " + filename_sheets)
        logging.info(">> Correction completed.")
        logging.info("Access the file: " + filename_sheets)


if __name__ == '__main__':
    run = Spreadsheet()
    run.save_csv()
