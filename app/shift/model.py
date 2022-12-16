import xlwings as xw
import pandas as pd
import random
#----------simple model --------------
y_range = 5
x_range = 10 
count_max = 5
max_day = 2

def next_idx(grid,y,x):
    y_list = list(range(y_range))
    #random.shuffle(y_list)
    for idx,y in enumerate(y_list):
        if grid[y].count(1) >= count_max:
            continue
        x_list = list(range(x_range))
        random.shuffle(x_list)
        for x in x_list:
                if grid[y][x] == 0:
                    return y, x      
    return -1, -1

def check_model_val(grid, y, x):
    apply_count = 0
    for i in range(y_range):       
        if grid[i][x] == 1:
            apply_count += 1
    if apply_count > max_day:
        print("koko:",y,x,apply_count)
        return False   
    return True

def count_grid(grid):
    dc_l =[]
    for x in range(x_range):
        dc = 0
        for y in range(y_range):
            if grid[y][x] == 1:
                dc += 1
        dc_l.append(dc)
    return dc_l

def find_max_idx(d_l):
    return d_l.index(max(d_l))

def find_min_idx(d_l):
    min_v = 99999
    for idx,v  in enumerate(d_l) :
        if (v < min_v) and not(v <=0)  :
            min_v = v
            min_idx = idx
    return min_idx
    
def find_max_idx_y_x(grid):
    max_idx = find_max_idx(count_grid(grid))
    for y in range(y_range):
        if grid[y][max_idx] == 1:
            return y,max_idx

def find_min_idx_y_x(grid):
    min_idx = find_min_idx(count_grid(grid))
    for y in range(y_range):
        if grid[y][min_idx] == 0:
            return y,min_idx




def solve_shift(grid,y,x,check_fun=check_model_val,next_fun=next_idx):
    y,x = next_fun(grid,y,x)
    if -1 in [y,x]: 
        return True     
    for i in range(x_range):
        if check_fun(grid,y,x):
            grid[y][x] = 1
            xw.Range((y+1,x+16)).value = 1
            if solve_shift(grid,y,x):
                return True
            grid[y][x] = 0
            xw.Range((y+1,x+16)).value = 0
            print("Track back:",y,x,grid)
    return False

#----- swap model logic -----


#-------------------------------------
def main():
    wb = xw.Book.caller()

    sh_model = wb.sheets["model"]

    global model_grid
    model_grid = [[0 if v is None else int(v) for v in row] for row in sh_model.range('A1:J5').value]
    sh_model.range('P1').value = model_grid

    solve_shift(model_grid,0,0)
    for row in model_grid:
        print(row)

if __name__ == "__main__":
    xw.Book("model.xlsx").set_mock_caller()
    main()
