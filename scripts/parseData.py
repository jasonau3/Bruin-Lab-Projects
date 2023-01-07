raw_x_data = input("Enter x:")
raw_y_data = input("Enter y:")

# enter it like "y: [76.87, 89.99, 52.95, 68.61, 109.17, .... , 64.79],"
x_data = raw_x_data.replace(" ", "").replace("[", "").replace("]", "").split(",")
x_data[0] = x_data[0][2:] # delete the "x:" for the first element
y_data = raw_y_data.replace(" ", "").replace("[", "").replace("]", "").split(",")
y_data[0] = y_data[0][2:] # delete the "y:" for the first element

print() # empty line

# print to excel
print("Time(min) Insulin(μU/ml)")
for i in range(len(x_data)):
    print(x_data[i] + " " + y_data[i])