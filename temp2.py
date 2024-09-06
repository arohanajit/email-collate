import requests
from bs4 import BeautifulSoup
import csv

url = "https://poole.ncsu.edu/group/people/faculty/"

response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

faculty_data = []

# Find all faculty members
faculty_members = soup.find_all('div', class_='views-row')

for member in faculty_members:
    name = member.find('h3', class_='field-content').text.strip()
    designation = member.find('div', class_='views-field-field-faculty-designation').find('div', class_='field-content').text.strip()
    email_div = member.find('div', class_='views-field-field-email')
    email = email_div.find('div', class_='field-content').text.strip() if email_div else "N/A"
    
    faculty_data.append([name, email, designation])

# Write data to CSV file
with open('faculty_data.csv', 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Name', 'Email', 'Designation'])  # Write header
    writer.writerows(faculty_data)

print("Data has been scraped and saved to faculty_data.csv")