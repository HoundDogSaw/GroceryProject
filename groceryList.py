#!/usr/bin/env python
import PySimpleGUI as sg
import csv, os
import pandas as pd
import openpyxl as op
from openpyxl import load_workbook

blt = {'bacon': 1, 'lettuce': 1, 'tomato': 2}
steak = {'steak': 1, 'side vegetable': 1}
burgers = {'ground beef': 1, 'side vegetable': 1, 'cheddar': 1}
bean_dish = {'cannelini beans': 2, 'kidney beans': 1, 'spinach': 1, 'smoked sausage': 1, 'diced tomatoes': 1, 'garlic': 1, 'chicken boullion': 1}
enchiladas = {'corn tortillas': 1, 'ground beef': 1, 'black beans': 1, 'enchilada sauce': 1, 'diced chiles': 1, 'monterrey jack': 1, 'cilantro': 1, 'jalapeno': 1}
black_bean_soup = {'black beans': 1, 'salsa': 1, 'garlic': 1, 'chicken boullion': 1}
bbq = {'ground beef': 1, 'chicken gumbo soup': 1}
potato_soup = {'potatoes': 6, 'celery': 1, 'onion': 1, 'bacon': 1, 'ham': 1}
pot_pie = {'chicken breast': 1, 'carrots': 1, 'peas': 1, 'potatoes': 3, 'onion': 1, 'chicken boullion': 1}
steak_stir_fry = {'garlic': 1, 'sirloin': 1, 'onion': 1, 'cilantro': 1, 'soy sauce': 1, 'broccoli': 1, 'snow peas': 1}
french_onion_soup = {'onion': 3, 'beef boullion': 1, 'worchestershire sauce': 1, 'swiss': 1}
white_chicken_chili = {'chicken breast': 1, 'onion': 1, 'garlic': 1, 'chicken boullion': 1, 'great northern beans': 2, 'diced chiles': 2, 'corn': 1, 'cilantro': 1, 'cream cheese': 1, 'heavy whipping cream': 1}
zuppa_toscana_soup = {'bacon': 1, 'italian sausage': 1, 'garlic': 1, 'onion': 1, 'chicken boullion': 1, 'potatoes': 3, 'baby spinach': 1, 'heavy whipping cream': 1}
stuffed_pepper_soup = {'ground beef': 1, 'onion': 1, 'garlic': 1, 'beef boullion': 1, 'tomato sauce(29oz.)': 2, 'green peppers': 4, 'pre-cooked quinoa': 1, 'cheddar': 1}
chicken = {'chicken': 1, 'side vegetable': 1}
pasta = {'gf noodles': 1, 'ground beef': 1, 'pasta sauce': 1, 'mozzarella': 1}
beef_stew = {'stew meat': 2, 'beef boullion': 1, 'onion': 1, 'carrot': 1, 'peas': 1, 'potatoes': 3}
tacos = {'ground beef': 1, 'tomatoes': 2, 'cilantro': 1, 'cheddar': 1, 'jalapeno': 1, 'taco shells': 1}
burrito_bowls = {'sirloin': 1, 'garlic': 1, 'limes': 4, 'soy sauce': 1, 'tomatoes': 2, 'cilantro': 1, 'black beans': 1, 'cheddar': 1, 'jalapeno': 1, 'avocado': 4}
pork_chops = {'pork chops': 1, 'side vegetable': 1}
steak_fajitas = {'sirloin': 1, 'red pepper': 1, 'green pepper': 1, 'yellow pepper': 1, 'onion': 1, 'fajita mix': 1, 'corn tortillas': 1}
chicken_fajitas = {'chicken breast': 1, 'red pepper': 1, 'green pepper': 1, 'yellow pepper': 1, 'onion': 1, 'fajita mix': 1, 'corn tortillas': 1}

layout = [[sg.Text('SOUP', text_color='black')],
          [sg.Checkbox('Black Bean Soup', key='black_bean_soup', size=(18, 1)), sg.Checkbox('Potato Soup', key='potato_soup', size=(18, 1)), sg.Checkbox('French Onion Soup', key='french_onion_soup', size=(18, 1))],
          [sg.Checkbox('Zuppa Toscana Soup', key='zuppa_toscana_soup', size=(18, 1)), sg.Checkbox('Stuffed Pepper Soup', key='stuffed_pepper_soup', size=(18, 1))],
          [sg.Text('MEAT', text_color='black')],
          [sg.Checkbox('Steak', key='steak', size=(12, 1)), sg.Checkbox('Burgers', key='burgers', size=(12, 1)), sg.Checkbox('Pork Chops', key='pork_chops', size=(12, 1)), sg.Checkbox('Chicken', key='chicken', size=(12, 1))],
          [sg.Text('MEXICAN', text_color='black')],
          [sg.Checkbox('Steak Fajitas', key='steak_fajitas', size=(18, 1)), sg.Checkbox('Chicken Fajitas', key='chicken_fajitas', size=(18, 1)), sg.Checkbox('Tacos', key='tacos', size=(18, 1))],
          [sg.Checkbox('Burrito Bowls', key='burrito_bowls', size=(18, 1)), sg.Checkbox('Enchiladas', key='enchiladas', size=(18, 1))],
          [sg.Text('SANDWICHES', text_color='black')],
          [sg.Checkbox('BLT', key='blt', size=(18, 1)), sg.Checkbox('BBQ', key='bbq', size=(18, 1))],
          [sg.Text('DISHES', text_color='black')],
          [sg.Checkbox('Bean Dish', key='bean_dish', size=(18, 1)), sg.Checkbox('Pot Pie', key='pot_pie', size=(18, 1)), sg.Checkbox('White Chicken Chili', key='white_chicken_chili', size=(18, 1))],
          [sg.Checkbox('Pasta', key='pasta', size=(18, 1)), sg.Checkbox('Steak Stir Fry', key='steak_stir_fry', size=(18, 1)), sg.Checkbox('Beef Stew', key='beef_stew', size=(18, 1))],
          [sg.Submit(tooltip='Submit'), sg.OK()]]


window = sg.Window('Select item:', layout, default_element_size=(40, 1), size=(500, 400))
new_dict = {}
while True:
    event, values = window.read()
    if event == 'Submit':
        for value in values:
            if values[value] == True:
                print(value)
                if value == 'enchiladas':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in enchiladas.items():
                            writer.writerow([key, value])
                elif value == 'blt':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in blt.items():
                            writer.writerow([key, value])
                elif value == 'steak':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in steak.items():
                            writer.writerow([key, value])
                elif value == 'burgers':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in burgers.items():
                            writer.writerow([key, value])
                elif value == 'black_bean_soup':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in black_bean_soup.items():
                            writer.writerow([key, value])
                elif value == 'bean_dish':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in bean_dish.items():
                            writer.writerow([key, value])
                elif value == 'bbq':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in bbq.items():
                            writer.writerow([key, value])
                elif value == 'potato_soup':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in potato_soup.items():
                            writer.writerow([key, value])
                elif value == 'pot_pie':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in pot_pie.items():
                            writer.writerow([key, value])
                elif value == 'steak_stir_fry':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in steak_stir_fry.items():
                            writer.writerow([key, value])
                elif value == 'french_onion_soup':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in french_onion_soup.items():
                            writer.writerow([key, value])
                elif value == 'white_chicken_chili':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in white_chicken_chili.items():
                            writer.writerow([key, value])
                elif value == 'zuppa_toscana_soup':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in zuppa_toscana_soup.items():
                            writer.writerow([key, value])
                elif value == 'stuffed_pepper_soup':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in stuffed_pepper_soup.items():
                            writer.writerow([key, value])
                elif value == 'chicken':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in chicken.items():
                            writer.writerow([key, value])
                elif value == 'pasta':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in pasta.items():
                            writer.writerow([key, value])
                elif value == 'beef_stew':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in beef_stew.items():
                            writer.writerow([key, value])
                elif value == 'tacos':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in tacos.items():
                            writer.writerow([key, value])
                elif value == 'burrito_bowls':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in burrito_bowls.items():
                            writer.writerow([key, value])
                elif value == 'pork_chops':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in pork_chops.items():
                            writer.writerow([key, value])
                elif value == 'steak_fajitas':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in steak_fajitas.items():
                            writer.writerow([key, value])
                elif value == 'chicken_fajitas':
                    with open('output.csv', 'a+') as output:
                        writer = csv.writer(output)
                        for key, value in chicken_fajitas.items():
                            writer.writerow([key, value])
    elif event == 'OK':
        window.close()
        break
    else:
        window.close()
        break


def return_contents(file_name):
    with open(file_name) as infile:
        reader = csv.reader(infile)
        return list(reader)


groceryList = return_contents('output.csv')
#print(groceryList)

#This converts the count to an integer instead of string
for item in groceryList:
    item[1] = int(item[1])
#print(groceryList)

title = ['ingredient', 'count']
df = pd.DataFrame(groceryList, columns = title)
'''df = df.sort_values(by=['ingredient'])'''


#This groups the ingredients by name and sums the counts
df = df.groupby(by=['ingredient'])['count'].sum().reset_index()
df.to_csv(r'/home/pi/Programs/ingredients.csv', index=True)
print(df)

'''df = pd.DataFrame(df, columns= ['ingredient', 'count'])'''
'''ingredients = df.values.tolist()
print(ingredients)'''

with pd.ExcelWriter('test.xlsx', engine='openpyxl', mode='a') as writer:
    df.to_excel(writer, sheet_name='Sheet1', header=False, index=False)

#Uncomment this to hard print the list of ingredients
'''os.system("lpr -P HP_Officejet_Pro_8620_728EA3_ ingredients.csv")'''
os.remove('output.csv')
os.remove('ingredients.csv')

