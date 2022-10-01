from bs4 import BeautifulSoup
import pandas as pd
import os

def join_leave(text):
    if 'left' in text:
        return int(text[10:14]), int(text[24:28])
    else:
        return int(text[10:14]), 'N/A'


def main():
    TITLE = 'Data.xlsx'

    names = []
    sites = []
    countries = []
    genders = []
    parties = []
    joins = []
    leaves = []

    tableDict = {
                'Delegate Name' : names,
                'Link to Delegate Site': sites,
                'Country Name' : countries,
                'Mr/Ms' : genders,
                'Political group' : parties,
                'Joined' : joins,
                'Left' : leaves
                }

    for filename in os.listdir(os.getcwd() + '\\page_sources'):
        with open(os.getcwd() + '\\page_sources\\' + filename) as file:
            soup = BeautifulSoup(file, 'html.parser')

        for person in soup.find_all('a', 'small'):
            names.append(person.contents[1].b.get_text())

            sites.append('https://pace.coe.int/en' + person['href'])

            countries.append(person.contents[3].get_text())

            if person.contents[1].b.get_text()[0:2] == 'Mr':
                genders.append(1)
            else:
                genders.append(0)

            if len(person.contents) == 9:
                parties.append(person.contents[5].get_text())
                (join, leave) = join_leave(person.contents[7].get_text())
                joins.append(join)
                leaves.append(leave)
            else:
                parties.append('N/A')
                (join, leave) = join_leave(person.contents[5].get_text())
                joins.append(join)
                leaves.append(leave)
    
    dataFrame = pd.DataFrame(tableDict)
    dataFrame.to_excel(os.getcwd() + '\\data\\' + TITLE)

if __name__ == "__main__":
    main()