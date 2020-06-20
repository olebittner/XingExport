import contextlib
import json

from selenium.webdriver.support import ui
from seleniumrequests import Firefox
import xlsxwriter

start_url = 'https://www.xing.com/settings'
xing_api = 'https://www.xing.com/xing-one/api'
profile_url = 'https://www.xing.com/profile/{}/cv'

contact_list_query = {"operationName": "contactsNetwork",
                      "variables": {"limit": 0, "offset": 0, "orderBy": "LAST_NAME"},
                      "query": "query contactsNetwork($offset: Int, $limit: Int, $orderBy: ContactsListOrderBy, $filters: ContactsFilters) {\n  viewer {\n    contactsNetwork(offset: $offset, limit: $limit, orderBy: $orderBy, filters: $filters) {\n      total\n      collection {\n        id\n        contactCreatedAt\n        normalizedInitialOfFirstName\n        normalizedInitialOfLastName\n        memo\n        tagList {\n          id\n          name\n          __typename\n        }\n        xingId {\n          firstName\n          lastName\n          ...UserInfoWithOccupation\n          __typename\n        }\n        __typename\n      }\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment UserInfoWithOccupation on XingId {\n  ...UserInfo\n  profileOccupation {\n    occupationOrg\n    occupationTitle\n    __typename\n  }\n  __typename\n}\n\nfragment UserInfo on XingId {\n  userFlags {\n    displayFlag\n    __typename\n  }\n  displayName\n  gender\n  pageName\n  profileImage(size: SQUARE_64) {\n    url\n    __typename\n  }\n  __typename\n}\n"}


def load_contact_list(driver: Firefox, limit: int):
    contacts = []
    total_contacts = 0
    offset = 0
    first = True
    while offset < total_contacts or first:
        first = False
        data = contact_list_query
        data["variables"]["limit"] = limit
        data["variables"]["offset"] = offset
        response = driver.request('POST', xing_api, data=json.dumps(data),
                                  headers={'content-type': 'application/json', 'Accept': '*/*'})
        if response.status_code == 200:
            data = json.loads(response.content)
            network = data['data']['viewer']['contactsNetwork']
            total_contacts = network['total']
            contacts = contacts + network['collection']
            offset = len(contacts)
        else:
            print('Contact fetch failed')
            exit(1)
        print(f"{offset}/{total_contacts}")
    return contacts


def parse_contacts(raw_contacts: []):
    results = []
    for raw in raw_contacts:
        xing_id = raw['xingId']
        contact = {
            'username': xing_id['pageName'],
            'firstname': xing_id['firstName'],
            'lastname': xing_id['lastName'],
            'contact_since': raw['contactCreatedAt'],
            'note': raw['memo'],
            'org': xing_id['profileOccupation']['occupationOrg'],
            'title': xing_id['profileOccupation']['occupationTitle']
        }
        if contact['note'] is None:
            contact['note'] = ''
        results.append(contact)
    return results


all_contacts = []

with contextlib.closing(Firefox()) as driver:
    wait = ui.WebDriverWait(driver, 300)
    driver.get('https://www.xing.com/settings')
    wait.until(lambda driver: str(driver.current_url).startswith(start_url))
    raw_contacts = load_contact_list(driver, 25)
    all_contacts = parse_contacts(raw_contacts)

workbook = xlsxwriter.Workbook('XingNetwork.xlsx')
contacts_sheet = workbook.add_worksheet('Contacts')

col_width = [0] * 5 + [12]

cols = []
for col in ['Name', 'First Name', 'Organisation', 'Title', 'Note', 'Xing-Profile']:
    cols.append({'header': col})

contacts_sheet.add_table(0, 0, len(all_contacts), 5, {'columns': cols})
for row, contact in enumerate(all_contacts, start=1):
    for column, field in enumerate(['lastname', 'firstname', 'org', 'title', 'note']):
        contacts_sheet.write(row, column, contact[field])
        col_width[column] = max(col_width[column], len(contact[field]) * 1.1)
    contacts_sheet.write_url(row, 5, profile_url.format(contact['username']), string='Link')

for col, width in enumerate(col_width):
    contacts_sheet.set_column(col, col, width)
workbook.close()
