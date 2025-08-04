import math
def roundup(number):
    number = str(number)
    ones_digit = number[-1]
    ones_digit = float(ones_digit)
    number = float(number)

    if ones_digit < 5:
        rounded = float(math.ceil((number + 5) / 10) * 10)
    else:
        rounded = float(math.ceil(number / 10) * 10)

    return rounded
