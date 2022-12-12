import xlwings as xw
import pandas as pd
import random

SET_NUM = set([1,2,3,4,5,6,7,8,9])

CH = [True,False]

def check_num(num,j):
    for i in range(1,10):
        #return True
        if xw.Range((i,j+15)).value == num:
            return False
        else:
            return True


def put_num(i,j,avl):
    if i>6:
        return(0)

    print(i,j)
    src_row = xw.Range((i,1),(i,9)).value
    dst_row = xw.Range((i,1+15),(i,9+15)).value
    avl_l =  list(SET_NUM - set(dst_row)-set(src_row))
    #if j==1:
    #    avl_l =  list(SET_NUM - set(src_row))
    #else:
    #    avl_l = avl
    random.shuffle(avl_l)
    print(avl_l)
    if src_row[j-1]:
        xw.Range((i,j+15)).value = src_row[j-1]
    else:
        if avl_l:
            num = avl_l[0]
            if check_num(num,j) == False:
                print("check NG:", i,j,num)
                xw.Range((i,j+15)).value = ''
                put_num(i,j-1,avl_l)
            else:
                #car_num = avl_l.pop(0)
                print("put:",avl_l[0])
                xw.Range((i,j+15)).value = avl_l[0] #car_num

    if j>=9:
        put_num(i+1,1,avl_l)
    else:
        put_num(i,j+1,avl_l)


def sudoku():
    wb.sheets['数独'].activate()
    put_num(1,1,[])


def df_test():
    sheet = wb.sheets[1]
    global df
    df = sheet.range('A1:D5').options(pd.DataFrame, header=1,index=False).value
    df['bbbb'] = df['bbbb'] *10 

    sheet = wb.sheets[2]
    sheet.range('A1:D5').options(pd.DataFrame, header=1,index=False).value = df


def simple_test():
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"

def main():
    global wb
    wb = xw.Book.caller()

    #simple_test()
    #df_test()
    sudoku()


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("test.xlsx").set_mock_caller()
    main()
