#Global variable
x = "Awesome"

def myFunction():
    x = "Amazing"
    print("I am " + x)

myFunction()

print("I am " + x)

#Variable With global keyword

def myfunc():
    global x
    x = "fantastic"

myfunc()

print("Python is " + x)

