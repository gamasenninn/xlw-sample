import xlwings as xw

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

def main():
    wb = xw.Book.caller()
    wb.sheets['数独'].activate()
    i_board = [[0 if v is None else int(v) for v in row] for row in xw.Range('A1:I9').value]
    solve_sd(i_board,0,0)
    xw.Range('P1').value = i_board

if __name__ == "__main__":
    xw.Book("sudoku.xlsx").set_mock_caller()
    main()
