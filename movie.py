import xlrd
import pandas as pd
import numpy as np
import glob
file_location = "/Users/apple/Desktop/movie theater/movie.xlsx"
workbook = xlrd.open_workbook(file_location)
worksheet = workbook.sheet_by_index(0)
all_data = pd.DataFrame()
for f in glob.glob('movie*.xlsx'):
    df = pd.read_excel(f)
    all_data = all_data.append(df, ignore_index = True)
print(all_data)
sheet2 = pd.read_excel('movie.xlsx',sheet_name = 'ticket types')
print(sheet2)
sheet3 = pd.read_excel('movie.xlsx',sheet_name = 'theaters')
print(sheet3)
sheet4 = pd.read_excel('movie.xlsx',cheet_name = 'movies')
print(sheet4)

total_rows = worksheet.nrows
total_cloumn = worksheet.ncols

def q2():
    while True:
        for row_cursor in range(1,total_rows):
            if worksheet.cell(row_cursor,1).value=='The Seaborn Identity':
                if worksheet.cell(row_cursor,2).value=='adult' and worksheet.cell(row_cursor,3).value==3:
                    d = worksheet.cell(row_cursor,0).value
                    return(d)
print('The second largest theater to sale the movie the seaborn identity is',q2())

def q3():
    while True:
        for row_cursor in range(1, total_rows):
            if worksheet.cell(row_cursor,2).value=='senior' and worksheet.cell(row_cursor,3).value==2:
                d = worksheet.cell(row_cursor,0).value
                if d=="The Frame":
                    return('Portland')
                    
                elif d=="The Empirical House":
                    return('New York')
                elif d=="Richie's Famous Minimax Theater":
                    return('Los Angeles')
                elif d=="Sumdance Cinemas":
                    return('Los Angeles')

print("The Third largest city to sale only with senior is",q3())

def q4():
    dict = {"The Sumif All Fears":124,"The Seaborn Identity":119,"The Matrices":136,"There's Something About Merging":119,"Mamma Median!":108,"Harry Porter":158,"Kung Fu pandas":92,"While You Were Sorting":103,"10 Things I Hate About Unix":97}
    while True:
        for row_cursor in range(1,total_rows):
            if worksheet.cell(row_cursor,2).value=="senior" and worksheet.cell(row_cursor,3).value==4:
                d = worksheet.cell(row_cursor,1).value
            if worksheet.cell(row_cursor,2).value=="adult" and worksheet.cell(row_cursor,3).value==2:
                d1 = worksheet.cell(row_cursor,1).value
        if dict[d]>dict[d1]:
            return(d)
        else:
            return(d1)
print("The longest movie to have third maximum sales is",q4())
