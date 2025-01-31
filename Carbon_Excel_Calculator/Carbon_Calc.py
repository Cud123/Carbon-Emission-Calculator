from gettext import install

import numpy as np
import pandas
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image

#Inputs

#Gas type input and error handling

try:
    transport_car = input("What type of fuel do you use for your car? Gas or Diesel? ")
except ValueError:
    print("Invalid input. Please enter gas or diesel.")

transport_car = transport_car.strip()
while(True):
    if transport_car.lower() == "gas" or transport_car.lower() == "diesel":
        break
    else:
        transport_car = input("Invalid input. Please enter gas or diesel: ")
    

#Distance traveled by car input and error handling

try:
    transportation = float(input("What is the distance you drove per day in miles? ")) * 365
except ValueError:
    print("Invalid input. Please enter a valid number of miles.")
# Inputs miles/day driven and converts it to miles / year

while(True):
    if transportation > 0 and (transportation/365) < 10000:
        break
    else:
        transportation = float(input("Invalid input. Please enter a valid number of miles: ") * 365)
        continue


#Electricity Input and error handling

try:
    electricity = float(input("What is your monthly energy consumption in kWh? ")) * 12
except ValueError:
    print("Invalid input. Please enter a valid number of kWh.")

while(True):
    if electricity > 0 and electricity < 1000000:
        break
    else:
        electricity = float(input("Invalid input. Please enter a valid number of kWh: ")*12)


#Flight input and error handling

try:
    flights = float(input("How much miles have you traveled on flights this year? "))
except ValueError:
    print("Invalid input. Please enter a valid number of miles.")

while(True):
    if flights > 0 and flights < 100000000:
        break
    else:
        flights = float(input("Invalid input. Please enter a valid number of miles: "))
        #inputs miles flown per year


#Iteration for accuracy input and error handling

while True:
    try:
        n = int(input("How much iterations would you like to run? The larger the better the estimate.(min is 1000 and max is 10000) "))
        if n < 1000 or n > 10000:
            print("Invalid input. Please enter a number in specified range.")

        else:
            break
    except ValueError:
        print("Invalid input. Please enter a valid number.")




# Average constants:

# From EPA vehicle emissions
avg_transport_emission_rate = {"Gas": 0.367, "Diesel": 0.337} # in kg CO2/gallon
std_dev_transport_emission_rate = 0.03 #in kgCO2/km(an estimate)


# From phoenix.gov:

avg_electricity_emission_rate = 0.385 # in kgCO2/ kWh
std_dev_electricity_emission_rate = 0.07 # in kgCO2/kWh (an estimate)


# From the icct.gov
avg_flight_emission_rate = 0.144 # in kgCO2 / mile
avg_std_dev_flight_emission_rate = 0.04 # in kgCO2 / mile


CO2_transport_list = []
CO2_electricity_list = []
CO2_flight_list = []

#Monte Carlo Simulation of CO2 emissions

for i in range(n):
    norm_transport_emission_factor = np.random.normal(avg_transport_emission_rate[transport_car],
                                                         std_dev_transport_emission_rate)
    norm_electricity_emission_factor = np.random.normal(avg_electricity_emission_rate,
                                                           std_dev_electricity_emission_rate)

    norm_flight_emission_factor = np.random.normal(avg_flight_emission_rate, avg_std_dev_flight_emission_rate)

    CO2_transport_list.append(transportation * norm_transport_emission_factor)
    CO2_electricity_list.append(electricity * norm_electricity_emission_factor)
    CO2_flight_list.append(flights * norm_flight_emission_factor)


# Mean, standard deviation, and 95th Percentile values for car travel

mean_CO2_transport = np.mean(CO2_transport_list)
std_dev_CO2_transport = np.std(CO2_transport_list)
transport_95th_percentile_CO2 = np.percentile(CO2_transport_list, 95)

#Mean, standard deviation, and 95th percentile values for electricity usage

mean_CO2_electricity = np.mean(CO2_electricity_list)
std_dev_CO2_electricity = np.std(CO2_electricity_list)
electricity_95th_percentile_CO2 = np.percentile(CO2_electricity_list, 95)

#Mean, standard deviation, and 95th percentile values for flight miles
mean_CO2_flight = np.mean(CO2_flight_list)
std_dev_CO2_flight = np.std(CO2_flight_list)
flight_95th_percentile_CO2 = np.percentile(CO2_flight_list, 95)

#Dictionary for panda excel integration. First sheet which displays all values of monte carlo simulation
data = {
    "Iteration": np.arange(1,n+1),
    "CO2 Transport (kgCO2)": CO2_transport_list,
    "CO2 Electricity (kgCO2)": CO2_electricity_list,
    "CO2 Flight (kgCO2)": CO2_flight_list,
    "Total Carbon Emission (kgCO2)": np.array(CO2_transport_list) + np.array(CO2_electricity_list) + np.array(CO2_flight_list)
}

df = pandas.DataFrame(data)

#Dictionary for displaying mean, standard deviation, and 95th percentile of monte carlo simulation values in second excel sheet
summary_data = {
    "Categories": ["Transport", "Electricity", "Flight", "Total"],
    "Mean": [mean_CO2_transport, mean_CO2_electricity, mean_CO2_flight, mean_CO2_transport + mean_CO2_electricity + mean_CO2_flight],
    "Std Dev": [std_dev_CO2_transport, std_dev_CO2_electricity, std_dev_CO2_flight, std_dev_CO2_transport + std_dev_CO2_electricity + std_dev_CO2_flight],
    "95th Percentile": [transport_95th_percentile_CO2, electricity_95th_percentile_CO2, flight_95th_percentile_CO2, transport_95th_percentile_CO2 + electricity_95th_percentile_CO2 + flight_95th_percentile_CO2]
}

summary_df = pandas.DataFrame(summary_data)

#Histogram plot for monte carlo values for first excel sheet
plt.figure(figsize=(10, 6))
plt.hist(data["Total Carbon Emission (kgCO2)"], bins=50, color='violet')
plt.title("total CO2 Emissions")
plt.xlabel("CO2 Emissions (kgCO2)")
plt.ylabel("iterations")
plt.grid(axis = 'y', linestyle = '--')



hist_image = "hist_chart.png"
plt.savefig(hist_image)


#Bar plot for standarized carbon emissions in second excel sheet
filtered_df = summary_df[summary_df["Categories"] != "Total"]

plt.figure(figsize=(10, 6))
plt.pie(filtered_df["Mean"], labels = filtered_df["Categories"], autopct = '%1.1f%%', colors = ['green','red','purple',], startangle = 90, shadow = True, explode = (0.01,0.01,0.01))
plt.title("Average Carbon Emissions by Category")



pie_chart = "pie_chart.png"
plt.savefig(pie_chart, bbox_inches="tight")
plt.close()

excel_file = "Carbon_Calc_Results.xlsx"

#Pandas excel integration

with pandas.ExcelWriter(excel_file) as writer:
    df.to_excel(writer, index = False, sheet_name="Simulation Results")
    summary_df.to_excel(writer, index = False, sheet_name="Summary_data")

    simulation_sheet = writer.sheets['Simulation Results']
    summary_sheet = writer.sheets['Summary_data']

    #Set column widths for sheet 1 of excel file
    simulation_sheet.column_dimensions['A'].width = 12  # Iteration
    simulation_sheet.column_dimensions['B'].width = 20  # CO2 Transport
    simulation_sheet.column_dimensions['C'].width = 20  # CO2 Electricity
    simulation_sheet.column_dimensions['D'].width = 20  # CO2 Flights
    simulation_sheet.column_dimensions['E'].width = 30  # Total CO2


    #Set column widths for sheet 2 of excel file
    summary_sheet.column_dimensions['A'].width = 20  # Category
    summary_sheet.column_dimensions['B'].width = 20  # Mean
    summary_sheet.column_dimensions['C'].width = 20  # Std Dev
    summary_sheet.column_dimensions['D'].width = 20  # 95th Percentile


    #Display histogram on Excel file page 1
    img_1 = Image(hist_image)
    simulation_sheet.add_image(img_1, 'G2')

    #Display bar plot on Cxcel file page 2
    img_2 = Image(pie_chart)
    summary_sheet.add_image(img_2, 'G2')


print(f"Results saved to {excel_file}")




























