import xlwings as xw
import pandas as pd

input_grid = []

def find_next_idx(grid):
    for y in range(9):
        for x in range(9):
            if grid[y][x] == 0:
                return y, x
    return -1, -1

def is_valid(grid, y, x, value):
    if value not in grid[y]:
        if value not in [i[x] for i in grid]:            
            blk_x, blk_y = (x//3) * 3, (y//3) * 3
            blk_grid = [i[blk_x:blk_x + 3] for i in grid[blk_y:blk_y + 3]]
            if value not in sum(blk_grid, []):
                return True
    return False

def solve_sudoku(grid, y=0, x=0):
    y, x = find_next_idx(grid)
    if y == -1 or x == -1:
        return True
    for value in range(1, 10):
        if is_valid(grid, y, x, value):
            grid[y][x] = value
            #xw.Range((y+1, x+16)).value = value
            if solve_sudoku(grid, y, x):
                return True
            grid[y][x] = 0
            #xw.Range((y+1, x+16)).value = ''
    return False

def main():
    global wb
    wb = xw.Book.caller()
    wb.sheets['数独'].activate()

    input_grid = [[0 if v is None else int(v) for v in row] for row in xw.Range('A1:I9').value]
    solve_sudoku(input_grid)
    xw.Range('P1').value = input_grid


if __name__ == "__main__":
    xw.Book("sudoku.xlsx").set_mock_caller()
    main()


