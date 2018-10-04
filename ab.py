import string

def abn(s):
    assert all(c in string.ascii_letters for c in s)
    s = s.upper()
    return sum((ord(c) - ord("A") + 1) * (26 ** i)
               for i, c in enumerate(s[::-1]))

assert abn("") == 0
assert abn("A") == 1
assert abn("B") == 2
assert abn("Z") == 26
assert abn("AA") == 27
assert abn("AB") == 28
assert abn("BA") == 53
assert abn("ZZ") == 702
assert abn("AAA") == 703
assert abn("AAB") == 704
assert abn("AMJ") == 1024
assert abn("BXX") == 2000
assert abn("HELLO") == 3752127
assert abn("WORLD") == 10786572
assert abn("FXSHRXW") == 2147483647
assert abn("") == 0
assert abn("a") == 1
assert abn("b") == 2
assert abn("z") == 26
assert abn("aa") == 27
assert abn("ab") == 28
assert abn("ba") == 53
assert abn("zz") == 702
assert abn("aaa") == 703
assert abn("aab") == 704
assert abn("amj") == 1024
assert abn("bxx") == 2000
assert abn("hello") == 3752127
assert abn("world") == 10786572
assert abn("fxshrxw") == 2147483647

def nab(n):
    assert n >= 0
    letters = []
    while n != 0:
        letters.append(string.ascii_uppercase[(n % 26) - 1])
        n = ((n - 1) // 26)
    return "".join(letters)[::-1]

assert nab(0) == ""
assert nab(1) == "A"
assert nab(2) == "B"
assert nab(26) == "Z"
assert nab(27) == "AA"
assert nab(28) == "AB"
assert nab(53) == "BA"
assert nab(702) == "ZZ"
assert nab(703) == "AAA"
assert nab(704) == "AAB"
assert nab(1024) == "AMJ"
assert nab(2000) == "BXX"
assert nab(3752127) == "HELLO"
assert nab(10786572) == "WORLD"
assert nab(2147483647) == "FXSHRXW"
