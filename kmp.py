def kmp(text, pattern):
    if not pattern:
        return list(range(len(text) + 1))
    lps = [0] * len(pattern)
    j = 0
    for i in range(1, len(pattern)):
        while j > 0 and pattern[i] != pattern[j]:
            j = lps[j - 1]
        if pattern[i] == pattern[j]:
            j += 1
        lps[i] = j

    matches = []
    j = 0
    for i, char in enumerate(text):
        while j > 0 and char != pattern[j]:
            j = lps[j - 1]
        if char == pattern[j]:
            j += 1
        if j == len(pattern):
            matches.append(i - j + 1)
            j = lps[j - 1]
    return matches

if __name__ == "__main__":
    text = input()
    pattern = input()
    matches = kmp(text, pattern)
    print("Matches:", " ".join(map(str, matches)) if matches else "none")
