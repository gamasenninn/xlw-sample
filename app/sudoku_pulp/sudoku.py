import xlwings as xw
import pulp as p
import itertools

def get_val(i, j, x):
    for v in range(1, 10):
        if p.value(x[i][j][v]) == 1:
            return v
    return None

def main():
    wb = xw.Book.caller()
    wb.sheets['数独'].activate()
    board = xw.Range('A1:I9').value
    #-------ここから数独を線形最適化で解く-------
    prob = p.LpProblem('Sudoku')
    x = p.LpVariable.dicts('x', (range(0, 9),range(0, 9),range(1, 10)), cat='Binary')
    for i,j in itertools.product(range(0,9),range(0,9)):  #かならず数字が埋まるための制約
        prob += p.lpSum([x[i][j][v] for v in range(1, 10)]) == 1
    for i,v in itertools.product(range(0,9),range(1,10)): #行方向の制約
        prob += p.lpSum([x[i][j][v] for j in range(0, 9)]) == 1
    for j,v in itertools.product(range(0,9),range(1,10)): #列方向の制約
        prob += p.lpSum([x[i][j][v] for i in range(0, 9)]) == 1
    for k,v in itertools.product(range(0,9),range(1,10)): #ブロックごとの制約
        prob += p.lpSum([x[k // 3 * 3 + l // 3][k % 3 * 3 + l % 3][v]  for l in range(0, 9)]) == 1
    for i, row in enumerate(board): #初期状態の設定（あらかじめ数字がうまったボードを作成）
        for j,v in enumerate(row):
            if v != None:
                prob += x[i][j][int(v)] == 1
    prob.solve()
    #-------解析終了-------
    xw.Range('P1').value = [[ get_val(i,j,x)for j in range(0,9)] for i in range(0,9)]

if __name__ == "__main__":
    xw.Book("sudoku.xlsx").set_mock_caller()
    main()
