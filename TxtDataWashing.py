'''
process txt data(with blank spaces, /n and '' each line)

3 elements in the front  of each row is shared, then each 5 elements is a piece of info

split() -> convert item(string) to item(list)
strip -> keep item(string) still string
'''

import csv
import numpy as np
import pandas as pd
updated_rows = []
with open(r"C:\Users\User\PycharmProjects\pythonProject2\turnstile-110528.txt", "r") as f:
    all_data = f.readlines()
    #print(all_data)
    #print(type(all_data))  # list, each line is a string

    for index,item in enumerate(all_data): #item is the row
        item = all_data[index].split()#[sentence 1] remove /n and balankspace #string
        #item = all_data[index].strip('') #[sentence 2] string still have /n and balankspace
        #item = item.split(",") #[sentence 3] no half ' and ] at the end of each line
                                # dosen't remove blank space and /n

        #print(item[0])
        #print(type(item[0]))#string
        #print(item)
        #print(type(item))  #list
        list2 = item[0].split(",") #list 2 is each row in list form,each element is string
        #print(list2)
        #print(type(list2)) #list
        #now I need to stack all list2 to array

        updated_row = item[0:3]
        for i in range(0, len(item[3:len(item)])):
            updated_row.append(item[i + 3])
            if (i + 1) % 5 == 0:
                updated_rows.append(updated_row)
                updated_row = item[0:3]
                print(updated_row)
    '''
    for name in filenames:
        with open('updated_' + name, 'wb') as f:
            writer = csv.writer(f)
            writer.writerows(updated_rows)

    
    array1.append(list2)
    print(array1)
    print(type(array1)) #array1 is list include many list[[list2],[],...,[],[]]

    
    
    
    #data_txt = np.loadtxt('turnstile-110528.txt')
    #data_txtDF = pd.DataFrame(data_txt)
    #data_txtDF.to_csv('turnstile-110528.csv', index=False)
    '''
