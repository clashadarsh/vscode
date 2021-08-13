print('Enter a string :')
s = input().lower()
print(f"Space idena?: {s.isspace()}")
vowels = {'a', 'e', 'i', 'o', 'u'}
for item in s:
    if item in vowels:
        print(item)
