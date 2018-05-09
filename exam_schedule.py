import sys
import os
from time import sleep
try:
    from ics import Calendar, Event
    import arrow
    from dateutil import tz
    import shutil
    import requests
    import numpy as np
    import datetime
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import PyPDF2
except ImportError:
    print("Warning: Various packages missing.  Installing...", file=sys.stderr)
    import subprocess
    subprocess.call([sys.executable] + "-m pip install ics requests numpy selenium PyPDF2".split())
    from ics import Calendar, Event
    import arrow
    from dateutil import tz
    import shutil
    import requests
    import numpy as np
    import datetime
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import PyPDF2


def login_with_andrew_id(driver, andrew_id, password):
    driver.find_element_by_id("j_username").send_keys(andrew_id)
    driver.find_element_by_id("j_password").send_keys(password)
    driver.find_elements_by_css_selector(".loginbutton")[0].click()

def get_courses(andrew_id="", password="", manual_login=False):
    """ Uses your andrew_id and password to retrieve the classes you are taking
        """

    course_ids = []
    course_friendly_names = {}
    
    driver = webdriver.Chrome()

    driver.get("https://s3.andrew.cmu.edu/sio/") # SSO login to here
    if manual_login:
        WebDriverWait(driver, 60 * 60).until(
            EC.title_is("CMU Student Information Online")
        )
    else:
        login_with_andrew_id(driver, andrew_id, password)
    driver.get("https://s3.andrew.cmu.edu/sio/#schedule-registration")
    sleep(2)
    driver.find_element_by_id("Yes").click()
    table = driver.find_element_by_xpath(
        "//div[contains(text(),'Official Schedule')]/../../../" +
        "div[contains(@class, 'portlet-body')]/div/table/tbody"
    )
    rows = table.find_elements_by_tag_name("tr")
    for row in rows:
        cell = row.find_elements_by_tag_name("td")[0]
        if cell.text.strip() == "TITLE / NUMBER & SECTION" or \
                cell.text.strip() == "TOTAL UNITS:":
            continue
        else:
            course_ids += [(
                cell.text.strip().split()[-2],
                cell.text.strip().split()[-1],
            )]
            course_friendly_names[cell.text.strip().split()[-2]] = \
                " ".join(cell.text.strip().split()[:-2])
    return course_ids, course_friendly_names

def get_final_exam_times(course_ids):
    res = {}
    r = requests.get("https://www.cmu.edu/hub/docs/final-exams.pdf", stream=True)
    r.raw.decode_content = True
    with open("./final_schedule.pdf", "wb") as f:
        shutil.copyfileobj(r.raw, f)
    pdf = PyPDF2.PdfFileReader("./final_schedule.pdf")
    collected_lines = {}
    for i in range(pdf.getNumPages()):
        anticipate_number = None
        anticipate_section = None
        seen_section = False
        assume_next_is_room = False
        for line in pdf.getPage(i).extractText().splitlines():
            if anticipate_number is None:
                for course_number, section_id in course_ids:
                    if anticipate_number is None and course_number == line.strip():
                        anticipate_number = course_number
                        anticipate_section = section_id
            elif not seen_section:
                if anticipate_section == line:
                    seen_section = True
            elif line.strip().startswith(
                    ("Monday", "Tuesday", "Wednesday", 
                     "Thursday", "Friday", "Saturday", "Sunday")):
                collected_lines[anticipate_number, anticipate_section, "date"] = line.strip()
                if (anticipate_number, anticipate_section, "time") in collected_lines:
                    assume_next_is_room = True # no real way to make a regex or anything for this
            elif line.strip().endswith(("p.m.","a.m.")):
                collected_lines[anticipate_number, anticipate_section, "time"] = line.strip()
                if (anticipate_number, anticipate_section, "date") in collected_lines:
                    assume_next_is_room = True # no real way to make a regex or anything for this
            elif assume_next_is_room:
                collected_lines[anticipate_number, anticipate_section, "room"] = line.strip()
                assume_next_is_room = False
                anticipate_number = None
                anticipate_section = None
                seen_section = False
    absent_course_numbers = []
    for course_number, section_id in course_ids:
        if course_number in res:
            continue 
            # SIO has multiple entries, lecture number and section number.  
            # Expect exactly one of these to correspond to a final exam time entry
        try:
            date = collected_lines[course_number, section_id, "date"]
            time = collected_lines[course_number, section_id, "time"]
            room = collected_lines[course_number, section_id, "room"]
        except KeyError:
            absent_course_numbers += [course_numbers]
            continue
        first_time, second_time = time.split(" - ")[0], time.split(" - ")[1]
        first_time = first_time[:-3] + "m" # replace a.m. with am and p.
        second_time = second_time[:-3] + "m" # replace a.m. with am and p.
        res[course_number] = (
            datetime.datetime.strptime(
                    date + " " + first_time, "%A, %B %d, %Y %I:%M %p"
                ),
            datetime.datetime.strptime(
                    date + " " + second_time, "%A, %B %d, %Y %I:%M %p"
                ),
            section_id,
            room
        )
    absent_course_numbers = [
        course_number for course_number in absent_course_numbers
        if course_number not in res # ya know, multiple entries and shit
    ]
    if len(absent_course_numbers) != 0:
        print(
            "Warning: could not find final exam times for courses: %s" % 
                ",".join(absent_course_numbers),
            file=sys.stderr
        )
    return res

def create_ics_file(course_friendly_names, exam_times, path):
    c = Calendar()
    for course_number in exam_times:
        dt_begin, dt_end, section, room = exam_times[course_number]
        e = Event(
            name="%s %s Final Exam" % (course_number, section),
            location=room,
            description="Final Exam for %s" % (course_friendly_names[course_number])
        )
        e.begin = arrow.get(dt_begin, tz.gettz('US/Eastern'))
        e.end = arrow.get(dt_end, tz.gettz('US/Eastern'))
        c.events.append(e)
    with open(path, "w") as f:
        f.writelines(c)

def export_google_calendar(ics_filename, andrew_id="", password="", manual_login=False):
    driver = webdriver.Chrome()
    driver.get("https://calendar.google.com/")
    def patient_find_el_by_xpath_clk(path):
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, path)))
        return driver.find_element_by_xpath(path)
    def patient_find_el_by_xpath(path):
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, path)))
        return driver.find_element_by_xpath(path)
    if manual_login:
        WebDriverWait(driver, 60 * 60).until(
            EC.url_contains("accounts.google.com")
        )
        WebDriverWait(driver, 60 * 60).until_not(
            EC.url_contains("accounts.google.com")
        )
        WebDriverWait(driver, 60 * 60).until(
            EC.url_contains("calendar.google.com")
        )
    else:
        patient_find_el_by_xpath('//input[@type="email"][@name="identifier"]') \
            .send_keys(andrew_id + "@andrew.cmu.edu")
        patient_find_el_by_xpath(
            '//div[@role="button"][./content/span[contains(text(), "Next")]]'
        ).click()
        WebDriverWait(driver, 10).until(
            EC.url_contains("login.cmu.edu")
        )
        login_with_andrew_id(driver, andrew_id, password)
    # Create Final Exam Schedule Calendar
    patient_find_el_by_xpath_clk('//div[@role="button"][@aria-label="Add other calendars"]').click()
    patient_find_el_by_xpath_clk('//content[@aria-label="New calendar"]').click()
    patient_find_el_by_xpath('//input[@type="text"][@aria-label="Name"]').send_keys("Final Exam Schedule")
    patient_find_el_by_xpath('//textarea[@aria-label="Description"]').send_keys("Auto-Generated by CMU Final Exam Schedule Google Calendar Improrter.")
    patient_find_el_by_xpath('//div[@role="button"][./content/span[contains(text(), "Create")]]').click()

    sleep(1)

    driver.get("https://calendar.google.com/")

    # Import it!
    patient_find_el_by_xpath_clk('//div[@role="button"][@aria-label="Add other calendars"]').click()
    patient_find_el_by_xpath_clk('//content[@aria-label="Import"][@role="menuitem"]').click()
    patient_find_el_by_xpath('//input[@type="file"][@name="filename"]').send_keys(os.path.abspath(ics_filename))
    patient_find_el_by_xpath_clk('//div[@role="listbox"][@aria-label="Add to calendar"]').click()
    patient_find_el_by_xpath_clk('//div[@role="option"][@aria-label="Final Exam Schedule"]').click()
    patient_find_el_by_xpath_clk('//div[@role="button"][./content/span[contains(text(), "Import")]]').click()





if __name__ == "__main__":
    if len(sys.argv) < 3:
        course_ids, course_friendly_names = get_courses(manual_login)
        exam_times = get_final_exam_times(course_ids)
        create_ics_file(course_friendly_names, exam_times, "./final_schedule.ics")
        export_google_calendar("./final_schedule.ics", manual_login=True)
    else:
        andrew_id = sys.argv[1]
        password = sys.argv[2]
        course_ids, course_friendly_names = get_courses(andrew_id, password)
        exam_times = get_final_exam_times(course_ids)
        create_ics_file(course_friendly_names, exam_times, "./final_schedule.ics")
        export_google_calendar("./final_schedule.ics", andrew_id, password)
