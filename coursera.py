import sys
import requests
import json

from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_list_links(coursera_xml_file):
    etree_root = etree.parse(coursera_xml_file).getroot()
    return [element[0].text for element in etree_root]


def get_courses_list(course_list_links):
    courses_list = []
    for course_url in course_list_links:
        course_info = get_course_info(course_url)
        if course_info != {}:
            courses_list.append(course_info)
    return courses_list


def get_course_info(course_url):
    html = requests.get(course_url, allow_redirects=False)
    if (html.status_code != requests.codes.ok):
        return {}
    print(u'Parse url: %s' % html.url)
    info = BeautifulSoup(html.content, 'html.parser')
    name = info.find("div", attrs={"class": "title"}).text
    script_json_data = info.find("script",
                                 attrs={"type": "application/ld+json"})
    if script_json_data:
        json_date = json.loads(script_json_data.text)
        course_instance = json_date['hasCourseInstance'][0]
    else:
        course_instance = {}
    start_date = course_instance.get('startDate', None)
    language = course_instance.get('inLanguage', None)
    rating_info = info.find("div", attrs={"class": "ratings-text"})
    rating = rating_info.text if rating_info else None
    rating = rating[:rating.find(' ')] if rating is not None else None
    count_weeks = len(info.find_all("div", attrs={"class": "week"}))
    return {'cource_name': name,
            'language': language,
            'start_date': start_date,
            'count_weeks': count_weeks,
            'rating': rating}


def output_courses_info_to_xlsx(course_list, max_cource_name_width=70):
    work_book = Workbook()
    work_sheet = work_book.active
    work_sheet.column_dimensions['A'].width = max_cource_name_width
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
        print("File don`t enter.")
        exit()
    else:
        coursera_xml_file = sys.argv[1]
    course_list_links = get_courses_list_links(coursera_xml_file)
    course_list = get_courses_list(course_list_links)
    output_courses_info_to_xlsx(course_list)
