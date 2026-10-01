n = int(input('input the total item numbers for array:'))
arr = [float(input('input the numbers:')) for _ in range(n)]
def heapify(arr, i):
    left = 2 * i + 1
    right = 2 * i + 2
    largest = i

    if left < len(arr) and arr[left] > arr[largest]:
        largest = left
    if right < len(arr) and arr[right] > arr[largest]:
        largest = right
    if largest != i:
        arr[i], arr[largest] = arr[largest], arr[i]


def heap_sort(arr):
    n = len(arr)
    for i in range(n // 2 - 1, -1, -1):
        heapify(arr, i)


    for i in range(n - 1, 0, -1):
        arr[0], arr[i] = arr[i], arr[0]
        heapify(arr, i)

heap_sort(arr)
print("Sorted array:", arr)

