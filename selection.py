lst = [64, 34, 25, 5, 22, 11, 90, 12]

n = len(lst)
for i in range(n-1):
  min_index = i
  for j in range(i+1, n):
     if lst[j] < lst[min_index]:
       min_index = j
  min_value = lst.pop(min_index)
  lst.insert(i, min_value)

print('The result:', lst)
