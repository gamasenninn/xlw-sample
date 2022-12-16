import xlwings as xw
import pandas as pd
import random

def next_idx(board):
    for y in range(9):
        for x in range(9):
            if board[y][x] == 0:
                return y, x
    return -1, -1

def check_val(board, y, x, v):
    if v not in board[y] and v not in [row[x] for row in board]:
        sx, sy = (x//3) * 3, (y//3) * 3
        for bky in board[sy:sy+3]:
            if v in bky[sx:sx+3]:
                return False
        return True
    return False

def solve_sd(board, y, x):
    y, x = next_idx(board)
    if -1 in [y,x]: 
        return True     
    for val in [1, 2, 3, 4, 5, 6, 7, 8, 9]:
        if check_val(board, y, x, val):
            board[y][x] = val
            if solve_sd(board, y, x):
                return True
            board[y][x] = 0
    return False

def _solve_sd(board, y, x):
    y, x = next_idx(board)
    if -1 in [y,x]: 
        return True     
    for val in [1, 2, 3, 4, 5, 6, 7, 8, 9]:
        if check_val(board, y, x, val):
            board[y][x] = val
            if solve_sd(board, y, x):
                return True
            board[y][x] = 0
            print("cut back:",y,x)
    return False




def count_x(x_grid):
    xc = 0
    for x in x_grid:
        if x in [1]:
            xc += 1
    return xc

def grid_next_idx(grid):
    y_list = list(range(4))
    random.shuffle(y_list)
    for y in y_list:
        if count_x(grid[y]) >= 20:
            continue
        x_list = list(range(30))
        random.shuffle(x_list)
        for x in x_list:
            if grid[y][x] == 0:
                return y, x
    return -1, -1

def check_shift_val(grid, y, x):
    apply_count = 0
    for i in range(4):       
        if grid[i][x] == 1:
            apply_count += 1
            #print(y,x,apply_count)
    if apply_count > 3 :
        print("koko:",y,x,apply_count)
        return False
    return True
def count_shift(grid):
    dc_l =[]
    for x in range(31):
        dc = 0
        for y in range(4):
            if grid[y][x] == 1:
                dc += 1
        dc_l.append(dc)
    return dc_l

def solve_shift(grid,y,x):
    y,x = grid_next_idx(grid)
    if -1 in [y,x]: 
        return True     
    if check_shift_val(grid,y,y):
        grid[y][x] = 1
        xw.Range((y+2,x+4)).value = 1
        if solve_shift(grid,y,x):
            return True
        grid[y][x] = 0
        xw.Range((y+2,x+4)).value = 0
        print("Track back:",y,x)
    return False

def swp_next_idx(grid):
    dc_l = count_shift(grid)
    print("swap_l:",dc_l)
    max_index = dc_l.index(max(dc_l))
    print("max_index:",max_index)
    min_index = dc_l.index(min(dc_l))
    print("min_index:",min_index)


    return -1,-1

def swap_shift(grid,y,x):
    y,x = swp_next_idx(grid)
    if -1 in [y,x]: 
        return True     
    grid[y][x] = 1
    if swap_shift(grid,y,x):
        return True
    grid[y][x] = 0
    return False

#----------simple model --------------
y_range = 5
x_range = 10 
count_max = 4
def model_next_idx(grid,oy,ox):
    y_list = list(range(y_range))
    #random.shuffle(y_list)
    for y in y_list:
        if count_x(grid[y]) >= 5:
            continue
        x_list = list(range(x_range))
        random.shuffle(x_list)
        for x in x_list:
            if x != ox:
                if grid[y][x] == 0:
                    return y, x
    return -1, -1

def check_model_val(grid, y, x):
    #return True
    #print("check:",y,x)
    apply_count = 0
    for i in range(y_range):       
        if grid[i][x] == 1:
            apply_count += 1
            #print(y,x,apply_count)
    #print("check:",y,x,apply_count)
    if apply_count > 2:
        print("koko:",y,x,apply_count)
        return False
    return True

def solve_model(grid,y,x):
    y,x = model_next_idx(grid,y,x)
    #print("new y,x:",y,x)
    if -1 in [y,x]: 
        print("ループ終了")
        return True     
    for i in range(11):
        if check_model_val(grid,y,x):
            grid[y][x] = 1
            xw.Range((y+1,x+16)).value = 1
            #print("pre recursive:",y,x,grid)
            if solve_model(grid,y,x):
            #if solve_model(grid,y,x):
                #print("after recursive:",y,x,grid)
                return True
            grid[y][x] = 0
            xw.Range((y+1,x+16)).value = 0
            print("Track back:",y,x,grid)
                #solve_model(grid,y,x)
    #solve_model(grid,y,x)
    #print("--------")
    return False


#-------------------------------------
def main():
    wb = xw.Book.caller()

    sh_model = wb.sheets["model"]


    model_grid = [[0 if v is None else int(v) for v in row] for row in sh_model.range('A1:J5').value]
    sh_model.range('P1').value = model_grid
    solve_model(model_grid,0,0)

    for row in model_grid:
        print(row)



    return #テストのためここから先はすすませない

    sh_src = wb.sheets["conf"]

    #global df
    #df = sh_src.range('A5').expand('table').options(pd.DataFrame, header=1,index=False).value
    #print(df)

    #print(df.loc[1,'CODs'])

    #df['CODs'] = df['CODs'].apply(lambda x: [int(y) for y in x.split(',')])
    #df['RODs'] = df['RODs'].apply(lambda x: [int(y) for y in x.split(',')])

    #print(df)
    #print(df.loc[1,'CODs'])

    sh_shit_in = wb.sheets["shift_in"]
    global df_person
    df_person = sh_shit_in.range('A1:D1').expand('table').options(pd.DataFrame, header=1,index=False).value
    print(df_person)
   
    wb.sheets['shift_in'].activate()
    global date_grid
    date_grid = [[0 if v is None else int(v) for v in row] for row in xw.Range('F2:AJ5').value]
    print("solve_shift")
    for row in date_grid:
        print(row)

    wb.sheets['shift_out'].activate()
 
    solve_shift(date_grid,0,0)
    for row in date_grid:
        print(row)

    #wb.sheets['shift_out'].activate()
    #xw.Range('D2').value = date_grid

    print(count_shift(date_grid))
    #swap_shift(date_grid,0,0)
    #xw.Range('D10').value = date_grid

if __name__ == "__main__":
    xw.Book("shift.xlsx").set_mock_caller()
    main()
