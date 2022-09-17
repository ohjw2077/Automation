import random
num = random.randrange(3000, 6000)
num2 = random.randrange(num, 10000)
num3 = random.randrange(num, 6000)
num4 = random.randrange(num2, 10000)

frac = num/num2
frac2 = num3/num4
prt = "frac1: " + str(frac) + "\n" + "frac2: " + str(frac2)
print(num, "/", num2)
print(num3, "/", num4)
answer = int(input("Which one is bigger? 1 or 2"))

print(answer)
if frac > frac2:
    if answer == 1:
        print("Right!")
        print(prt)
    else:
        print("Wrong")
        print(prt)
else:
    if answer == 2:
        print("Right!")
        print(prt)
    else:
        print("Wrong")
        print(prt)