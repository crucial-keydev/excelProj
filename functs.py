import math

def get_amount():
    number = input('Amount: ')
    while True:
        try:
            number = float(number)
        except:
            print('Invalid input')
        
        if type(number) == float:
            break
        else:
            number = input('Amount: ')

    return number

def get_perc():
    number = input('Percentage: ')
    while True:
        try:
            number = int(number)
        except:
            print('Invalid input')
        
        if type(number) == int:
            break
        else:
            number = input('Percentage: ')

    return number


def fitintobox(spaces, a):
    taken = len(a)
    sure =[]
    leftover = []
    if taken > spaces:
        words = a.split(" ")
        lenght = 0 
        for i in words:
            lenght = lenght + len(i)
            if lenght < spaces:
                sure.append(i)           
            else:
                leftover.append(i)    

        return sure, leftover

    else: 
        return a.split(" "), []

def turntosentence(notasentence):
    sentence = ""
    for i in notasentence:
        sentence += str(i) + " "
    sentence = sentence[:-1]
    return sentence

ones = ('Zero', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine')

twos = ('Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen')

tens = ('Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety', 'Hundred')

suffixes = ('', 'Thousand', 'Million', 'Billion')

def process(number, index):
    
    if number=='0':
        return 'Zero'
    
    length = len(number)
    
    if(length > 3):
        return False
    
    number = number.zfill(3)
    words = ''
 
    hdigit = int(number[0])
    tdigit = int(number[1])
    odigit = int(number[2])
    
    words += '' if number[0] == '0' else ones[hdigit]
    words += ' Hundred ' if not words == '' else ''
    
    if(tdigit > 1):
        words += tens[tdigit - 2]
        words += ' '
        words += ones[odigit]
    
    elif(tdigit == 1):
        words += twos[(int(tdigit + odigit) % 10) - 1]
        
    elif(tdigit == 0):
        words += ones[odigit]

    if(words.endswith('Zero')):
        words = words[:-len('Zero')]
    else:
        words += ' '
     
    if(not len(words) == 0):    
        words += suffixes[index]
        
    return words
    
def getWords(number):
    length = len(str(number))
    
    if length>12:
        return 'This program supports upto 12 digit numbers.'
    
    count = length // 3 if length % 3 == 0 else length // 3 + 1
    copy = count
    words = []
 
    for i in range(length - 1, -1, -3):
        words.append(process(str(number)[0 if i - 2 < 0 else i - 2 : i + 1], copy - count))
        count -= 1

    final_words = ''
    for s in reversed(words):
        temp = s + ' '
        final_words += temp
    
    return final_words


def rating_prf(coverage):
    if coverage < 100000:
        pass
        count = 0
        for i in range(10000, 100001, 5000):
            if coverage > i:
                count += 1
                continue
            
            else:
                rating_amount = coverage * ((0.01968 - (0.00048*count)))

        
    elif coverage > 100000 and coverage < 300000:
        rating_amount = ((coverage - 100000) *0.00288) +1104
    
    else:
        rating_amount = coverage *0.0055

    return rating_amount

def rating_sty(coverage):
    if coverage < 200000:
        count = 0
        for i in range(10000, 200001, 5000):
            if coverage > i:
                count += 1
                continue
            
            else:
                rating_amount = coverage * ((0.03888 - (0.00048*count)))

        
    elif coverage > 200000 and coverage < 500000:
        rating_amount = ((coverage - 200000) *0.00288) + 4128
    
    else:
        rating_amount = coverage *0.0060

    return rating_amount
    

def getpayment(rating_amount):
    payment = rating_amount + (rating_amount *0.125) + (rating_amount * 0.002) + (rating_amount * 0.12) + 600

    return payment

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

def get_bond_type():
    bond = input('STY or PRF or WARR: ').upper()


    while True:
        try:
            bond == 'STY' or bond == 'PRF' or bond == 'WARR'
        except:
            print('Invalid type')

        if bond == 'STY' or bond == 'PRF' or bond == 'WARR':
            break
        else:
            bond = input('STY or PRF or WARR: ').upper()

    return bond

