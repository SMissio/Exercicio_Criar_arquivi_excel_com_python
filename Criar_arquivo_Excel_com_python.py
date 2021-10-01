#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import xlsxwriter as opcoesDOXL

import os

nomeArquivo = 'C:\\Users\\User1\\Desktop\\RPA1\\xlsx\\Exemplo3.xlsx'
myplan = opcoesDOXL.Workbook(nomeArquivo)

sheetdata = myplan.add_worksheet("Dados")


sheetdata.write("A1", "Numero1")
sheetdata.write("B1", "Numero2")
sheetdata.write("C1", "Fórmula")

sheetdata.write("A2", 11)
sheetdata.write("A3", 18)
sheetdata.write("A4", 10)
sheetdata.write("A5", 5)
sheetdata.write("A8", "Sylvia")

sheetdata.write("B2", 22)
sheetdata.write("B3", 15)
sheetdata.write("B4", 10)
sheetdata.write("B5", 5)
sheetdata.write("B8", "Suzi")

sheetdata.write_formula("C2", "=A2+B2")
sheetdata.write_formula("C3", "=A3-B3")
sheetdata.write_formula("C4", "=A4*B4")
sheetdata.write_formula("C5", "=A5/B5")
sheetdata.write_formula("C8", '=CONCATENATE(A8," ",B8)')
                        
sheetdata.set_column('A:C',15)
                        
myplan.close()

os.startfile(nomeArquivo)

