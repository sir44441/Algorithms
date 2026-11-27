def dimensions(a, b):
    if not a or not b or not a[0] or not b[0]:
        raise ValueError("Matrices must be nonempty.")
    r, k, c = len(a), len(a[0]), len(b[0])
    if any(len(row) != k for row in a) or any(len(row) != c for row in b):
        raise ValueError("Matrices must have equal-length rows.")
    if len(b) != k:
        raise ValueError("A columns must equal B rows.")
    return r, k, c

def matrix_multiply(a, b):
    r, k, c = dimensions(a, b)
    result = [[0] * c for _ in range(r)]
    for i in range(r):
        for j in range(c):
            for t in range(k):
                result[i][j] += a[i][t] * b[t][j]
    return result

def read_matrix(rows, cols):
    matrix = []
    for _ in range(rows):
        row = list(map(int, input().split()))
        if len(row) != cols:
            raise ValueError(f"Each row must contain {cols} integers.")
        matrix.append(row)
    return matrix

if __name__ == "__main__":
    r, k, c = map(int, input().split())
    if min(r, k, c) <= 0:
        raise ValueError("Dimensions must be positive.")
    a = read_matrix(r, k)
    b = read_matrix(k, c)
    for row in matrix_multiply(a, b):
        print(*row)
