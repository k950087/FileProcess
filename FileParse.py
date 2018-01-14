import pandas as pd
import os

#################################
readPath=r'C:\ta01'  #讀檔路徑
outPath=r'C:\ta01'   #輸出路徑
#################################
folderPath=os.listdir(readPath)                 #文件夾下有哪些檔案
fileCount=''.join(folderPath).count('.xlsx')    #合併字串後計算字元

for i in range(fileCount):
    i+=1
    filePath = readPath+'\CK'+str(i)+'.xlsx' #0+1
    readFile = pd.ExcelFile(filePath)
    file15 = readFile.parse(0)
    file16 = readFile.parse(1)
    file17 = readFile.parse(2)
    file15.to_csv(outPath+'\CK'+str(i)+'_'+readFile.sheet_names[0]+'.csv')
    file16.to_csv(outPath+'\CK'+str(i)+'_'+readFile.sheet_names[1]+'.csv')
    file17.to_csv(outPath+'\CK'+str(i)+'_'+readFile.sheet_names[2]+'.csv')




