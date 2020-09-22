import argparse
from urllib.request import urlopen
import json
import xlwt

stepik_url = "https://stepik.org/api"
min_session_size = 10
max_session_size = 30
align = "align: vert center, horiz center;"
borders = "borders: left thin, right thin, bottom thin, top thin;"
# thick_borders = "borders: left thick, right thick, bottom thick, top thick;"

course_row_style = xlwt.Style.easyxf(
    align
    + borders
    + "pattern: pattern solid, fore_color 0x08;"
)
section_row_style = xlwt.Style.easyxf(
    align
    + borders
    + "pattern: pattern solid, fore_color 0x09;"
)
lesson_style = xlwt.Style.easyxf(
    align
    + borders
    + "pattern: pattern solid, fore_color 0x0A;"
)
border_style = xlwt.Style.easyxf(
    borders
)
unused_style = xlwt.Style.easyxf(
    align
    + borders
    + "pattern: pattern solid, fore_color 0x0B;"
)

def get_course(id):
    res = json.load(urlopen(stepik_url + "/courses/{}".format(id)))
    return res['courses'][0]
def get_section(id):
    res = json.load(urlopen(stepik_url + "/sections/{}".format(id)))
    return res['sections'][0]
def get_unit(id):
    res = json.load(urlopen(stepik_url + "/units/{}".format(id)))
    return res['units'][0]
def get_lesson(id):
    res = json.load(urlopen(stepik_url + "/lessons/{}".format(id)))
    return res['lessons'][0]

class CourseTree:
    def __init__(self, course_id):
        course = get_course(course_id)
        self.children = []
        self.name     = course['title']
        self.length   = 0
        self.max_section_length = 0
        for section_id in course['sections']:
            self.children.append(SectionTree(section_id))
            l = self.children[-1].length
            self.length += l
            if l > self.max_section_length:
                self.max_section_length = l


    def generate_table(self, file_name, x, y):
        book  = xlwt.Workbook()
        sheet = book.add_sheet("Progress")
        book.set_colour_RGB(0x08, 162, 208, 142)
        book.set_colour_RGB(0x09, 198, 224, 180)
        book.set_colour_RGB(0x0A, 226, 239, 218)
        book.set_colour_RGB(0x0B, 217, 217, 217)
        max_days_to_pass = 0
        section_x = x + 1
        for section in self.children:
            days_to_pass = section._generate_table(sheet, section_x, y, self.max_section_length)
            if max_days_to_pass < days_to_pass:
                max_days_to_pass = days_to_pass
            section_x += 3 # Space for section name, lesson number and progress rows
        
        sheet.write_merge(x, x, y, y + self.max_section_length - 1, self.name,
            style = course_row_style)

        for i in range(0, self.max_section_length):
            sheet.col(y + i).width //= 3

        book.save(file_name if file_name else self.name + ".xls")

class SectionTree:
    def __init__(self, section_id):
        section = get_section(section_id)
        self.children = []
        self.name     = section['title']
        self.length   = 0
        for unit_id in section['units']:
            unit   = get_unit(unit_id)
            lesson = LessonLeaf(unit['lesson'])
            self.children.append(lesson)
            self.length += lesson.days_to_pass
            

    def _generate_table(self, sheet, x, y, max_section_length):
        curr_less_x = x + 1
        curr_less_y = y
        days_to_pass = 0
        # print(self.name)
        for index, lesson in enumerate(self.children):
            lesson_length = lesson._generate_table(sheet, curr_less_x, curr_less_y, index)
            curr_less_y += lesson_length
            days_to_pass += lesson_length

        sheet.write_merge(x, x, y, y + max_section_length - 1, self.name,
            style = section_row_style)
        if days_to_pass != max_section_length:
            sheet.write_merge(x + 1, x + 2, y + days_to_pass , y + max_section_length-1, "", style = unused_style)
        # print(self.name, x, y, days_to_pass)

        return days_to_pass

class LessonLeaf:
    def __init__(self, lesson_id):
        lesson = get_lesson(lesson_id)
        self.name = lesson['title']
        self.time_to_pass = lesson['time_to_complete']//60 + (lesson['time_to_complete']%60>0)
        t = self.time_to_pass
        self.days_to_pass = t // max_session_size + (t % max_session_size > min_session_size) + (t <= min_session_size)

    def _generate_table(self, sheet, x, y, index):
        # print(" "*3, x, y, self.days_to_pass)
        sheet.write_merge(x, x, y, y + self.days_to_pass - 1, index + 1,
            style = lesson_style)
        for i in range(0, self.days_to_pass):
            sheet.write(x+1, y+i, style = border_style)

        return self.days_to_pass

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description = "Generates progress table for stepik's course.")
    parser.add_argument('course_id', metavar='id', type=int, help='an id of a course')
    parser.add_argument('--max', type=int, default=30, help='max session size (default: 30)')
    parser.add_argument('--min', type=int, default=10, help='min session size (default: 10)')
    parser.add_argument('--out', type=str, default=None, help='output file name. Defaults to course\'s title')
    args = parser.parse_args()
    max_session_size = args.max
    min_session_size = args.min
    course_id        = args.course_id
    out_filename     = args.out

    print("Receiving data")
    course = CourseTree(course_id)
    print("Generating table")
    course.generate_table(None, 3, 3)
    print("Table is generated")
