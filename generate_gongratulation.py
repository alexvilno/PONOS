import random

def generate_triad(congrats1: list, congrats2: list, congrats3: list):
    inx1 = random.randint(0, len(congrats1) - 1)
    inx2 = random.randint(0, len(congrats2) - 1)
    inx3 = random.randint(0, len(congrats3) - 1)
    triad = congrats1[inx1] + ', ' + congrats2[inx2] + ', ты ' + congrats3[inx3]
    return triad

def generate_holiday(holidays: list):
    inx = random.randint(0, len(holidays) - 1)
    return holidays[inx]

def generate_congrat(name, holiday, triad):
    congrat = 'Дорогой(ая) ' + name + '! поздравляю тебя с ' + holiday + '! Желаю тебе ' + triad + '!'
    return congrat