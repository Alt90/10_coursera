import sys
import requests
import json


from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_xml(course_url):
    return requests.get(course_url).content


def get_courses_list_links(coursera_xml):
    return [element[0].text for element in etree.XML(coursera_xml)[:20]]


def get_courses_list(course_list_links):
    courses_list = []
    for course_url in course_list_links:
        course_info = get_course_info(course_url)
        if course_info != {}:
            courses_list.append(course_info)
    return courses_list


def get_course_info(course_url):
    html = requests.get(course_url)
    if (html.url != course_url):
        return {}
    print(u'Parse url: %s' % html.url)
    info = BeautifulSoup(html.content, 'html.parser')
    name = info.find("div", attrs={"class": "title"}).text
    script_json_data = info.find("script",
                                 attrs={"type": "application/ld+json"})
    if script_json_data:
        json_date = json.loads(script_json_data.text)
        course_instance = json_date['hasCourseInstance'][0]
        start_date = course_instance.get('startDate', '')
        language = course_instance.get('inLanguage', '')
    else:
        language = start_date = ''
    rating_info = info.find("div", attrs={"class": "ratings-text"})
    if rating_info:
        rating = rating_info.text[:rating_info.text.find(' ')]
    else:
        rating = ''
    count_weeks = len(info.find_all("div", attrs={"class": "week"}))
    return {'cource_name': name,
            'language': language,
            'start_date': start_date,
            'count_weeks': count_weeks,
            'rating': rating}


def output_courses_info_to_xlsx(course_list):
    work_book = Workbook()
    work_sheet = work_book.active
    work_sheet.column_dimensions['A'].width = 70
    work_sheet.append(['cource_name',
                       'language',
                       'start_date',
                       'count_weeks',
                       'rating'])
    work_sheet.append([])
    for course in course_list:
        work_sheet.append([course['cource_name'],
                           course['language'],
                           course['start_date'],
                           course['count_weeks'],
                           course['rating']])
    work_book.save(filename='courses.xlsx')


if __name__ == '__main__':
    if (len(sys.argv) < 2):
        print("File don`t enter. We use default file from coursera.")
        course_url = 'https://www.coursera.org/sitemap~www~courses.xml'
    else:
        course_url = sys.argv[1]
    coursera_xml = get_xml(course_url)
    course_list_links = get_courses_list_links(coursera_xml)
    course_list = get_courses_list(course_list_links)
    output_courses_info_to_xlsx(course_list)
