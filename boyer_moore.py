def boyer_moore(text, pattern):
    n, m = len(text), len(pattern)
    if m == 0:
        return list(range(n + 1))
    last = {char: i for i, char in enumerate(pattern)}
    shift = [0] * (m + 1)
    border = [0] * (m + 1)
    i, j = m, m + 1
    border[i] = j
    while i > 0:
        while j <= m and pattern[i - 1] != pattern[j - 1]:
            if shift[j] == 0:
                shift[j] = j - i
            j = border[j]
        i -= 1
        j -= 1
        border[i] = j
    j = border[0]
    for i in range(m + 1):
        if shift[i] == 0:
            shift[i] = j
        if i == j:
            j = border[j]

    matches = []
    start = 0
    while start <= n - m:
        j = m - 1
        while j >= 0 and pattern[j] == text[start + j]:
            j -= 1
        if j < 0:
            matches.append(start)
            start += shift[0]
        else:
            start += max(shift[j + 1], j - last.get(text[start + j], -1))
    return matches

if __name__ == "__main__":
    text = input()
    pattern = input()
    matches = boyer_moore(text, pattern)
    print("Matches:", " ".join(map(str, matches)) if matches else "none")
