temp = input("a?:")
guess = int(temp)

while guess > 8:
    print('big')
    temp = input('again:')
    guess = int(temp)    

while guess < 8:
    print('small')
    temp = input('again:')
    guess = int(temp)

if guess == 8: 
    print("OK")

print("over")