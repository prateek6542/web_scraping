import requests
from bs4 import BeautifulSoup
import openpyxl


file_path = 'Web Scraper Intern Data Fields.xlsx'
wb = openpyxl.load_workbook(file_path)
sheet = wb.active

def update_excel(row, col, data):
    sheet.cell(row=row, column=col, value=data if data else 'Not Available')

# Function to dynamically scrape university details
def scrape_university_info(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Scraping university information dynamically from the website
    university_name = soup.find('meta', {'property': 'og:site_name'})['content'] if soup.find('meta', {'property': 'og:site_name'}) else 'Not Available'
    contact_section = soup.find('footer')  

 
    university_email = contact_section.find('a', href=True).text if contact_section and contact_section.find('a', href=True) else 'Not Available'
    university_phone = contact_section.find('p').text if contact_section and contact_section.find('p') else 'Not Available'

    return {
        'name': university_name,
        'email': university_email,
        'phone': university_phone
    }

# Function to scrape course information
def scrape_courses(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    courses = []
    course_sections = soup.find_all('div', class_='course-item') 

    for course_section in course_sections:
        course_name = course_section.find('h3').text if course_section.find('h3') else 'Not Available'
        course_level = 'Undergraduate' if 'Undergraduate' in course_name else 'Postgraduate'  
        study_major = course_section.find('p', class_='major').text if course_section.find('p', class_='major') else 'Not Available'
        course_url = course_section.find('a', href=True)['href'] if course_section.find('a', href=True) else 'Not Available'
        tuition_fee = course_section.find('span', class_='tuition-fee').text if course_section.find('span', class_='tuition-fee') else 'Not Available'
        duration = course_section.find('span', class_='duration').text if course_section.find('span', class_='duration') else 'Not Available'
        ielts = course_section.find('span', class_='ielts').text if course_section.find('span', class_='ielts') else 'Not Available'
        toefl = course_section.find('span', class_='toefl').text if course_section.find('span', class_='toefl') else 'Not Available'

        courses.append({
            'course_name': course_name,
            'course_level': course_level,
            'study_major': study_major,
            'course_url': course_url,
            'tuition_fee': tuition_fee,
            'duration': duration,
            'ielts': ielts,
            'toefl': toefl
        })
    return courses

# Function to scrape scholarship information
def scrape_scholarships(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'html.parser')

    scholarships = []
    scholarship_sections = soup.find_all('div', class_='scholarship-item') 

    for scholarship_section in scholarship_sections:
        scholarship_name = scholarship_section.find('h3').text if scholarship_section.find('h3') else 'Not Available'
        description = scholarship_section.find('p').text if scholarship_section.find('p') else 'Not Available'

        scholarships.append({
            'scholarship_name': scholarship_name,
            'description': description
        })
    return scholarships

# Update Excel file with dynamically scraped data
def update_excel_with_data():
    university_url = "https://www.cam.ac.uk/"   
    courses_url = "https://www.cam.ac.uk/courses"  
    scholarships_url = "https://www.cam.ac.uk/scholarships" 

   
    university_info = scrape_university_info(university_url)
    
    row_index = 3  
    update_excel(row_index, 2, university_info['name'])  
    update_excel(row_index, 9, university_info['email'])  
    update_excel(row_index, 10, university_info['phone'])  

    # Scrape courses
    courses = scrape_courses(courses_url)
    for course in courses:
        update_excel(row_index, 21, course['course_name']) 
        update_excel(row_index, 23, course['course_level']) 
        update_excel(row_index, 24, course['study_major']) 
        update_excel(row_index, 35, course['course_url']) 
        update_excel(row_index, 37, course['tuition_fee'])  
        update_excel(row_index, 36, course['duration'])  
        update_excel(row_index, 41, course['ielts'])  
        update_excel(row_index, 42, course['toefl'])  
        row_index += 1

    # Scrape scholarships
    scholarships = scrape_scholarships(scholarships_url)
    row_index = 3
    for scholarship in scholarships:
        update_excel(row_index, 61, scholarship['scholarship_name'])  
        update_excel(row_index, 62, scholarship['description']) 
        row_index += 1

# Run the scraping and update the Excel file
update_excel_with_data()

# Save the updated Excel file
wb.save('Updated_Web_Scraper_Intern_Data.xlsx')

print('Scraping complete. File saved as Updated_Web_Scraper_Intern_Data.xlsx')
