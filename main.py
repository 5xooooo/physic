from openpyxl import load_workbook
from matplotlib import pyplot as plt


#刪除不需要的行、工作表
wb = load_workbook("data.xlsx")
sheet = wb["Raw Data"]
wb.remove(wb["Metadata Device"])
wb.remove(wb["Metadata Time"])
sheet.delete_cols(5)
sheet.delete_cols(3)
sheet.delete_cols(2)

#讀取time
time = [0]*1886
for i in range(2, 1886):
    time[i] = sheet["A"+str(i)].value

#讀z軸角速度
omega = [0]*1886
for i in range(2, 1886):
    omega[i] = sheet["B"+str(i)].value

#寫d(theta)
sheet["C1"] = "d(theta)"
sheet["C2"] = 0
d_theta = [0]*1886
for i in range(3, 1886):
    d_theta[i] = sheet["B"+str(i)].value * (sheet["A"+str(i)].value - sheet["A"+str(i-1)].value) 
    sheet["C"+str(i)] = d_theta[i]

#寫theta
sheet["D1"] = "theta"
sheet["D2"] = 0
theta = [0]*1886
for i in range(3, 1886):
    theta[i] = sheet["D"+str(i-1)].value + sheet["C"+str(i)].value
    sheet["D"+str(i)] = theta[i]

#寫alpha
sheet["E1"] = "alpha"
sheet["E2"] = 0
alpha = [0]*1886
for i in range(3, 1886):
    alpha[i] = (sheet["B"+str(i)].value - sheet["B"+str(i-1)].value)/(sheet["A"+str(i)].value - sheet["A"+str(i-1)].value)
    sheet["E"+str(i)] = alpha[i]

y_index = {"ω (rad/s)":omega, "d_θ (rad)":d_theta, "θ (rad)":theta, "α (rad/s^2)":alpha}

def get_key (dict, value):
    return [k for k, v in dict.items() if v == value][0]

def output_chart(y_data, save_title):
    title = get_key(y_index, y_data) +  "-t" + " chart"
    plt.plot(time, y_data)
    plt.title(title)
    plt.xlabel("time(s)")
    plt.ylabel(get_key(y_index, y_data))
    plt.savefig(save_title + ".png")

output_chart(y_index["d_θ (rad)"], "d_theta")
output_chart(y_index["ω (rad/s)"], "omega")
output_chart(y_index["θ (rad)"], "theta")
output_chart(y_index["α (rad/s^2)"], "alpha")

wb.save("output.xlsx")
