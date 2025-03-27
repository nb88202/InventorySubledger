import docx
import pandas as pd
import re
import xlwings as xw
import numpy as np
#  set the file path to the text file of the report.  run this.  and then run_a_macro to set the pages. 




path=r"C:\Users\jl\Documents\invsub\result.xlsx"
excel_workbook= xw.Book(path).sheets['Sheet1']

with open(r"C:\Users\jl\Documents\invsub\report.txt", "r") as f:
    
        contents = f.readlines() #list containing all strings

        newarray =[]
        array2 = []
        array3 =[]
        array4 = []
        array5=[]
        array6=[]
        array7=[]
        array8 = []
        

        word1='  **** Warehouse Id'
        word2='   ** Location Id '
        word3 ='*'
        word4 ='INVENTORY           STANDARD          WEIGHTED '
        word5='Inventory Subledger                      Berkel & Company, Contractors '
        word6='Sorted by Balance Sheet'
        word7 = 'AVERAGE          UNITS        VALUE'
        word8 = 'Inventory Account #'
        word9 = ' ============='
        word10 = 'continued'
        word11 = 'Retain                                       Non-Confidential'
        word12 = 'xxx'
        word13 = ' Total **** '
        word14 ='WHID:'
        word15 = 'LOC:'

        for line in contents:
                if word1 in line:
                        newarray.append(line)
                
    
        for line in contents:
                if word1 in line:
                        array2.append(line)
                if word2 in line:
                        array2.append(line)
                if word3 not in line:
                        array2.append(line)
        
        for line in array2:
                if word4 in line:
                        array2.remove(line)
                if word5 in line:
                        array2.remove(line)
                if word6 in line:
                        array2.remove(line)

        for line in array2:
                if word7 in line:
                        array2.remove(line)   
                if word8 in line:
                        array2.remove(line)   

        for line in array2:
                 if word9 in line:
                         array2.remove(line)
                 if word10 in line:
                         array2.remove(line)
                 if word11 in line:
                         array2.remove(line)
        
        for line in array2:
                array3.append(line.replace(" \n","xxx"))
              

        for line in array3:
                if word10 in line:
                        array3.remove(line)
                if word12 in line:
                        array3.remove(line)

        for sub in array3:
               array4.append(re.sub('\n','',sub))
            
        for line in array4:
                if word10 in line:
                        array4.remove(line)

        for line in array4:
                newline = line.replace(",","")
                array5.append(newline)


        for line in array5:
           if word13 in line:
                   array5.remove(line)

        for line in array5:
                if word1 in line:
                    array6.append("WHID:" + line[19:24])
                elif word2 in line:
                    array6.append("LOC:" +line[19:23])
                else:
                    array6.append(line[0:11]+","+line[12:47]+","+ line[48:60] + ","+ line[61:80]+","+line[81:98]+","+line[99:110]+","+line[111:200])
        

        for line in array6:
               newString =  re.sub("\s+", "", line.strip()) 
               array7.append(newString)

    
        for line in array7:
                if word14 in line[0:5]:
                     array8.append("YYY,,,,,,,,"+ line)
                elif word15 in line[0:4]:
                     array8.append("ZZZ,,,,,,,"+line)
                else:
                     array8.append(line)
      

        df = pd.DataFrame([sub.split(",") for sub in array8]) 

        df = df[df[0].str.contains("xxx")==False]
        df[7] = df[7].str[4:8]
        df = df.replace([None], [np.nan], regex=True)
        df[7]=df[7].fillna(method='ffill')
        df[8]=df[8].str[5:10]
        df[8]=df[8].fillna(method='ffill')
        
        df= df.drop(df[df[0] == 'ZZZ'].index)
        df= df.drop(df[df[0] == 'YYY'].index)
      
        mneg = df[5].str.endswith("-")
        df[5] = df[5].str.rstrip("-")
        df[5] = pd.to_numeric(df[5])
        df.loc[mneg, 5] *= -1
        
        df[6] = df[6].replace([np.nan], '0000000', regex=True)
        mneg2 = df[6].str.endswith("-")
        df[6] = df[6].str.rstrip("-")
        df[6] = pd.to_numeric(df[6])
        df.loc[mneg2, 6] *= -1



        df.rename(columns={0:"Product ID", 1:"Description", 2:"Account",3:"Std_Cost", 4:"WA_Cost", 5:"Units",6:"Value", 7:"Location_ID", 8:"Warehouse_ID"}, inplace=True)

        excel_workbook.range("A1").options(index=False, header=True).value=df
                             

        


