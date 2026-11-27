# Algorithms

A collection of Python implementations of sorting, string-searching, and matrix multiplication algorithms, with supplementary programming exercises.

## Implementations

| Category | Algorithm | File |
| --- | --- | --- |
| Sorting | Bubble sort | [bubble_sort.py](bubble_sort.py) |
| Sorting | Heap sort | [heap_sorting.py](heap_sorting.py) |
| Sorting | Insertion sort | [insertion_sort.py](insertion_sort.py) |
| Sorting | Merge sort | [merge_sort.py](merge_sort.py) |
| Sorting | Quicksort | [quick_sort.py](quick_sort.py) |
| Sorting | Selection sort | [selection_sort.py](selection_sort.py) |
| String searching | Boyer–Moore | [boyer_moore.py](boyer_moore.py) |
| String searching | Knuth–Morris–Pratt (KMP) | [kmp.py](kmp.py) |
| Matrix multiplication | Standard multiplication | [matrix_multiply.py](matrix_multiply.py) |
| Matrix multiplication | Strassen's method | [strassen.py](strassen.py) |

## Requirements

Python 3.9 or later. No third-party packages are required.

## Running the programs

Clone the repository and open its directory:

```bash
git clone https://github.com/sir44441/Algorithms.git
cd Algorithms
```

Run a script using Python, for example:

```bash
python bubble_sort.py
```

On systems where Python is available as `python3`, use `python3` instead.
Some sorting scripts prompt for input; others run predefined examples.

### String searching

Run either search algorithm:

```bash
python boyer_moore.py
```

or:

```bash
python kmp.py
```

Enter the text on the first line and the pattern on the second:

```text
ABABABABA
ABA
```

Expected output:

```text
Matches: 0 2 4 6
```

Positions are zero-based, and overlapping matches are included. If no match is found, the program prints `Matches: none`.

### Matrix multiplication

Run either implementation:

```bash
python matrix_multiply.py
```

or:

```bash
python strassen.py
```

The first input line contains `r k c`: matrix A has `r` rows and `k` columns, and matrix B has `k` rows and `c` columns.
Enter the rows of A followed by the rows of B, using space-separated integers without brackets.

Example input:

```text
2 3 2
1 2 3
4 5 6
7 8
9 10
11 12
```

Expected output:

```text
58 64
139 154
```

Both implementations support compatible rectangular matrices. Strassen's implementation pads matrices internally to a power-of-two square size and trims the result to the required dimensions.

## Exercises

The [exercises](exercises/) directory contains supplementary programming exercises and problem-solving challenges.

| Exercise | File |
| --- | --- |
| Two eggs and 100 floors | [2_eggs_100_floors.py](exercises/2_eggs_100_floors.py) |

Run the exercise with:

```bash
python exercises/2_eggs_100_floors.py
```

## Scope

These implementations are intended for learning and experimentation.
