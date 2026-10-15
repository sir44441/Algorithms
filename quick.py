def swap(arr, i, j):
    arr[i], arr[j] = arr[j], arr[i]


def median_of_three(arr, low, high):

    mid = (low + high) // 2

    if arr[low] > arr[mid]:
        swap(arr, low, mid)
    if arr[low] > arr[high]:
        swap(arr, low, high)
    if arr[mid] > arr[high]:
        swap(arr, mid, high)

    swap(arr, mid, high)
    return arr[high]


def partition(arr, low, high):
    pivot = median_of_three(arr, low, high)

    i = low - 1
    for j in range(low, high):
        if arr[j] < pivot:
            i += 1
            swap(arr, i, j)

    swap(arr, i + 1, high)
    return i + 1


def quickSort(arr, low, high):
    if low < high:
        pi = partition(arr, low, high)

        quickSort(arr, low, pi - 1)
        quickSort(arr, pi + 1, high)


if __name__ == "__main__":
    test_lists = [
        [54, 26, 93, 17, 77, 31, 44, 55, 20],
        [1, 2, 3, 4, 5, 6],
        [4, 2, -3, 12, 4, 1, -5, 6, 0, 12]
    ]

    for idx, lst in enumerate(test_lists, 1):
        print(f"--- Test {idx} ---")
        print("Original:", lst)

        arr_copy = lst.copy()

        quickSort(arr_copy, 0, len(arr_copy) - 1)

        print("Sorted:  ", arr_copy)
