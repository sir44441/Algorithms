def dimensions(a, b):
    if not a or not b or not a[0] or not b[0]:
        raise ValueError("Matrices must be nonempty.")
    r, k, c = len(a), len(a[0]), len(b[0])
    if any(len(row) != k for row in a) or any(len(row) != c for row in b):
        raise ValueError("Matrices must have equal-length rows.")
    if len(b) != k:
        raise ValueError("A columns must equal B rows.")
    return r, k, c

def add(a, b, sign=1):
    return [[x + sign * y for x, y in zip(ar, br)]
            for ar, br in zip(a, b)]


def split(a):
    h = len(a) // 2
    return ([row[:h] for row in a[:h]], [row[h:] for row in a[:h]],
            [row[:h] for row in a[h:]], [row[h:] for row in a[h:]])


def recurse(a, b):
    if len(a) == 1:
        return [[a[0][0] * b[0][0]]]
    a11, a12, a21, a22 = split(a)
    b11, b12, b21, b22 = split(b)
    m1 = recurse(add(a11, a22), add(b11, b22))
    m2 = recurse(add(a21, a22), b11)
    m3 = recurse(a11, add(b12, b22, -1))
    m4 = recurse(a22, add(b21, b11, -1))
    m5 = recurse(add(a11, a12), b22)
    m6 = recurse(add(a21, a11, -1), add(b11, b12))
    m7 = recurse(add(a12, a22, -1), add(b21, b22))
    c11 = add(add(add(m1, m4), m5, -1), m7)
    c12 = add(m3, m5)
    c21 = add(m2, m4)
    c22 = add(add(add(m1, m2, -1), m3), m6)
    return [left + right for left, right in zip(c11, c12)] + [
        left + right for left, right in zip(c21, c22)]


def strassen(a, b):
    r, k, c = dimensions(a, b)
    size = 1
    while size < max(r, k, c):
        size *= 2
    padded_a = [[0] * size for _ in range(size)]
    padded_b = [[0] * size for _ in range(size)]
    for i in range(r):
        padded_a[i][:k] = a[i][:]
    for i in range(k):
        padded_b[i][:c] = b[i][:]
    result = recurse(padded_a, padded_b)
    return [row[:c] for row in result[:r]]

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
    for row in strassen(a, b):
        print(*row)
