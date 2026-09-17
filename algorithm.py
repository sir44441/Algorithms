def sorting_method(lst):
    for i in range(1, len(lst)):
        key = lst[i]
        j = i-1
        while key < lst[j] and j >= 0:
            lst[j+1] = lst[j]
            j-=1
        lst[j+1] = key
    return lst

numbers = [2,8,4,10,3,5,6,1]
sorting = sorting_method(numbers)
print(sorting)