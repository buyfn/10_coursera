import requests
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_list(courses_num=20):
    xml_url = 'https://www.coursera.org/sitemap~www~courses.xml'
    content = requests.get(xml_url).content
    tree = etree.fromstring(content)
    links = tree.xpath("string()").split()
    return links[:courses_num]


def get_soup(url):
    page = requests.get(url).content
    soup = BeautifulSoup(page, 'html.parser')
    return soup


def get_title(course_soup):
    title = course_soup.find('h1', attrs={'class': 'title'})
    if title:
        return title.text
    else:
        return 'no title found'


def get_lang(course_soup):
    language = course_soup.find('div', attrs={'class': 'rc-Language'})
    if language:
        primary_lang = next(language.stripped_strings)
        return primary_lang
    else:
        return 'no language found'


def get_start_date(course_soup):
    start_date = course_soup.find('div', attrs={'class': 'startdate'})
    if start_date:
        return start_date.text
    else:
        return 'no start date found'


def get_duration(course_soup):
    duration = course_soup.find('div', attrs={'class': 'rc-WeekView'})
    if duration:
        weeks = sum(1 for _ in duration.children)
        return weeks
    else:
        return 'no duration found'


def get_rating(course_soup):
    rating = course_soup.find('div', attrs={'class': 'ratings-text'})
    if rating:
        return rating.text
    else:
        return 'no rating found'


def get_course_info(course_slug):
    soup = get_soup(course_slug)
    return (get_title(soup),
            get_lang(soup),
            get_start_date(soup),
            get_duration(soup),
            get_rating(soup))


def output_courses_info_to_xlsx(course_list, filepath='courses.xlsx'):
    wb = Workbook()
    ws = wb.active
    ws.append(['Course title',
               'Language',
               'Nearest start date',
               'Duration, weeks',
               'Rating'])

    for course in course_list:
        ws.append(course)

    wb.save(filepath)


if __name__ == '__main__':
    print('\nGetting links to courses...', end=' ')
    links = get_courses_list()
    courses_num = len(links)
    print('got {} links'.format(courses_num))

    courses = []
    for counter, link in enumerate(links, 1):
        print('[{}/{}] Parsing {}...'.format(counter, courses_num, link))
        course_info = (get_course_info(link))
        courses.append(course_info)

    print('Writing data to file...')
    output_courses_info_to_xlsx(courses)
