n = int(input('input the total item numbers for array:'))
arr = [float(input('input the numbers:')) for _ in range(n)]

for i in range(n):
    for j in range(len(arr) - 1):
         if arr[j] > arr[j + 1]:
            arr[j], arr[j + 1] = arr[j + 1], arr[j]

print('The result of the Bubble sorting:', arr)