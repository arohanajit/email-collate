import requests
from bs4 import BeautifulSoup
import csv

url = "https://ccee.ncsu.edu/group/faculty/"

response = requests.get(url)
soup = BeautifulSoup(response.content, 'html.parser')

faculty_data = []

# Find all faculty members
faculty_members = soup.find_all('div', class_='person-card')

for member in faculty_members:
    name = member.find('a', class_='name').text.strip()
    email = member.find('a', href=lambda x: x and x.startswith('mailto:')).text.strip()
    title = member.find('p', class_='title').text.strip()
    designation = title.split(',')[0].strip()  # Extract designation
    
    faculty_data.append([name, email, designation])

# Write data to CSV file
with open('faculty_data.csv', 'w', newline='', encoding='utf-8') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['Name', 'Email', 'Designation'])  # Write header
    writer.writerows(faculty_data)

print("Data has been scraped and saved to faculty_data.csv")