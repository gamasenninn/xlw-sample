import xlwings as xw
import pandas as pd
import random
import time
#----------simple model --------------
y_range = 5
x_range = 20 
count_max = 12
max_day = 4   #これ以上いてはいけない
min_day = 3   #これ以上いること
out_x= 1
out_y= 11

def next_idx(grid,y,x):
    y_list = list(range(y_range))
    random.shuffle(y_list)
    for idx,y in enumerate(y_list):
        if grid[y].count(1) >= count_max:
            continue
        x_list = list(range(x_range))
        random.shuffle(x_list)
        for x in x_list:
                if grid[y][x] == 0:
                    return y, x      
    return -1, -1

def find_block(row):
    r_cnt = 0
    ov = 1
    s_idx = e_idx = -1
    for i,v in enumerate(row):
        if v == 1: 
            if v == ov:
                if r_cnt == 0:
                    s_idx = i
                r_cnt += 1
                ov = v
                e_idx = i
        else:
            ov = 0
            r_cnt = 0
    return r_cnt,s_idx,e_idx

def find_continuous_sequence(arr):
  continuous_sequence_list = []
  current_sequence = []
  start_index = 0
  end_index = 0

  for i, num in enumerate(arr):
    if not current_sequence or num == current_sequence[-1]:
      current_sequence.append(num)
      end_index = i
    else:
      continuous_sequence_list.append((current_sequence[0], end_index-start_index+1, start_index, end_index))
      current_sequence = [num]
      start_index = i
      end_index = i

  if current_sequence:
    continuous_sequence_list.append((current_sequence[0], end_index-start_index+1, start_index, end_index))

  random.shuffle(continuous_sequence_list)
  return continuous_sequence_list


def check_model_val(grid, y, x):
    apply_count = 0
    is_max = False
    is_continuos = False
    for i in range(y_range):       
        if grid[i][x] == 1:
            apply_count += 1
    if apply_count >= max_day:
        #print("too many:",y,x,apply_count)
        is_max = True
    #5連投防止
    arr_block =  find_continuous_sequence(grid[y])
    for v, cnt, si, ei  in arr_block:
        if v == 1:
            if cnt >3:
                #print("連投?",y,cnt,si,ei)
                if  si-1 <= x and x <= ei+1:
                    print("連投",y,cnt,si,ei)
                    is_continuos = True
                    break

    if (is_max or is_continuos):
        print("false:",is_max,is_continuos,y,x)
        return False
    return True

def check_all_true(grid, y, x):
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
        if (v < min_v) and not(v <=-1)  :
            min_v = v
            min_idx = idx
    return min_idx
    
def find_max_idx_y_x(grid):
    max_idx = find_max_idx(count_grid(grid))
    for y in range(y_range):
        if grid[y][max_idx] >= 0:
            return y,max_idx

def find_min_idx_y_x(grid):
    min_idx = find_min_idx(count_grid(grid))
    for y in range(y_range):
        if grid[y][min_idx] == 0:
            return y,min_idx

def find_swap_idx(grid):
    dc_l = count_grid(grid)
    print(dc_l)
    max_idx = find_max_idx(dc_l)
    min_idx = find_min_idx(dc_l)
    print(max_idx,min_idx)
    if dc_l[min_idx] <= min_day :
        for y in range(y_range):
            if (grid[y][max_idx] == 1) and (grid[y][min_idx] == 0):
                return y,max_idx,min_idx
    return -1,-1,-1


def swap_model(grid):
    for i in range(x_range):
        y,max_x,min_x = find_swap_idx(grid)
        if y == -1 and max_x == -1 and min_x == -1:
            return True
        sw = grid[y][max_x]
        grid[y][max_x] = grid[y][min_x]
        grid[y][min_x] = sw

        xw.Range((y+out_y,max_x+out_x)).value = grid[y][max_x]
        xw.Range((y+out_y,min_x+out_x)).value = grid[y][min_x]

    #swap_model(grid)



def solve_shift(grid,y,x,check_fun=check_model_val,next_fun=next_idx):
    y,x = next_fun(grid,y,x)
    if -1 in [y,x]: 
        return True     
    for i in range(2):
        if check_fun(grid,y,x):
            grid[y][x] = 1
            xw.Range((y+out_y,x+out_x)).value = 1
            if solve_shift(grid,y,x,check_fun=check_fun):
                return True
            grid[y][x] = 0
            xw.Range((y+out_y,x+out_x)).value = 0
            #print("Track back:",y,x,grid)
            solve_shift(grid,y,x,check_fun=check_fun)
    solve_shift(grid,y,x,check_fun=check_fun)
    #return False

#----- swap model logic -----


#-------------------------------------
def main():
    wb = xw.Book.caller()

    sh_model = wb.sheets["model"]

    global model_grid
    model_grid = [[0 if v is None else int(v) for v in row] for row in sh_model.range('A1:T5').value]
    sh_model.range('A11').value = model_grid

#    solve_shift(model_grid,0,0,check_fun=check_all_true)
#    for row in model_grid:
#        print(row)
    
#    time.sleep(2)

    solve_shift(model_grid,0,0)
    for row in model_grid:
        print(row)
    sh_model.range('A11').value = model_grid

    time.sleep(2)
    print("--swap--")
    swap_model(model_grid)



if __name__ == "__main__":
    xw.Book("model.xlsx").set_mock_caller()
    main()
