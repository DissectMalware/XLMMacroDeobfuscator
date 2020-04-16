def column_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def column_number(s):
    number = 0
    power = 1
    for character in s:
        character = character.upper()
        digit = (ord(character) - ord('A')) *power
        number = number + digit
        power = power * 26

    return number + 1