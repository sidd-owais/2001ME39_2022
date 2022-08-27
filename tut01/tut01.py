def factorial(n):
    if(n == 1 or n == 0):
        return 1
    return (n*factorial(n-1))

number = int(input("Enter a number whose need to be calculated:"))
if(number < 0):
    print("Factorial doesn't exist")
else:
    print("Factorial of ",number, " is ",factorial(number))