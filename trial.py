import openpyxl
lost = openpyxl.load_workbook("lost.xlsx")
sheet = lost['lost']
students = [
    [1, "Alice", 20, 85, 90, 88],
    [2, "Bob", 21, 78, 83, 85],
    [3, "Charlie", 19, 92, 88, 91],
    [4, "David", 22, 70, 75, 78],
    [5, "Eva", 20, 95, 100, 94],
    [6, "Frank", 23, 88, 84, 87],
    [7, "Grace", 21, 90, 92, 90],
    [8, "Hannah", 19, 76, 80, 75],
    [9, "Ian", 22, 82, 78, 80],
    [10, "Jack", 20, 91, 89, 93]
]

alphabets = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

rw = 1

for row in students:
    col = 0
    for column in row: 
        cell = alphabets[col]+str(rw)
        print(f"{column} {cell}", end="  ")
        sheet[cell] = column
        col+=1
    rw += 1
    print("")

lost.save(filename="lost.xlsx")