import requests
from lxml import etree
from bs4 import BeautifulSoup
from openpyxl import Workbook


def get_courses_list(courses_num=20):
    xml_url = 'https://www.coursera.org/sitemap~www~courses.xml'
    r = requests.get(xml_url)
    content = r.content
    tree = etree.fromstring(content)
    links = tree.xpath("string()").split()
    return links[:courses_num]


def get_soup(url):
    r = requests.get(url)
    page = r.content
    soup = BeautifulSoup(page, 'html.parser')
    return soup


def get_course_info(course_slug):
    def get_title(course_soup):
        title = course_soup.find('h1', attrs={'class': 'title'})
        return title.text

    def get_lang(course_soup):
        language = course_soup.find('div', attrs={'class': 'rc-Language'})
        primary_lang = next(language.stripped_strings)
        return primary_lang
    
    def get_start_date(course_soup):
        start_date = course_soup.find('div', attrs={'class': 'startdate'})
        return start_date.text

    def get_duration(course_soup):
        duration = course_soup.find('div', attrs={'class': 'rc-WeekView'})
        try:
            weeks = sum(1 for _ in duration.children)
            return weeks
        except AttributeError:
            return None
    
    def get_rating(course_soup):
        rating = course_soup.find('div', attrs={'class': 'ratings-text'})
        return (rating.text if rating else None)

    soup = get_soup(course_slug)
    return (get_title(soup),
            get_lang(soup),
            get_start_date(soup),
            get_duration(soup),
            get_rating(soup))


def output_courses_info_to_xlsx(course_list):
    wb = Workbook()
    ws = wb.active
    ws.append(['Course title',
               'Language',
               'Nearest start date',
               'Duration, weeks',
               'Rating'])
    
    for course in course_list:
        ws.append(course)
        
    wb.save('courses.xlsx')


if __name__ == '__main__':
    print('\nGetting links to courses...', end=' ')
    links = get_courses_list()
    courses_num = len(links)
    print('got {} links'.format(courses_num))

    courses = []
    for counter, link in enumerate(links, 1):
        print('\n[{}/{}] Parsing {}...'.format(counter, courses_num, link))
        try:
            course_info = (get_course_info(link))
            print('parsed successfully')
        except AttributeError:
            course_info = ('Error', link)
            print('parsing error')
        courses.append(course_info)

    print('\nWriting data to file...')
    output_courses_info_to_xlsx(courses)
