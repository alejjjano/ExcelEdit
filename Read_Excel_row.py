def countDiffValues(texttoread):
    # Elaborates a list from numbers in texttoread
    valuelist = [num for num in texttoread.split("	")]

    # Elaborates a list from different numbers

    items = []
    for i in valuelist:
        if i not in items:
            items.append(i)

    # Prints all different values in the row
    print(items)

    # Prints the total amount of different values in the row
    print(len(items))


x = input("Insert Data")
countDiffValues(x)