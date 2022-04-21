import requests
import sys
import urllib3
from bs4 import BeautifulSoup
from tabulate import tabulate
from docx import Document
import time

# import pandas as pd
# import warnings
# warnings.filterwarnings("ignore")

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# def get_hb(sys_args):

#   ## THIS DOESN'T WORK BECAUSE THE FILE IS PROTECTED/RESTRICTED. NOT SURE HOW TO GET AROUND THIS ISSUE
#   # wb = load_workbook(r'.\QF1200F - Test Stand Equipment List.xlsx')
#   # print(type(wb))
#   pass

def get_hb():

  hb = []
  start = input('Start Date: ')
  end = input('End Date: ')
  loc = input('Location: ')
  tmp = input('Enter HB number of equipment (type "done" when finished): ')

  while tmp.lower() != 'done':
    hb.append(tmp.upper())
    tmp = input('Enter HB number of equipment (type "done" when finished): ')

  return hb,start,end,loc

def get_cal_dates(hb_lst,start,end,loc):

  cal_dates = []
  manufact = []
  desc = []
  start_lst = [start]*len(hb_lst)
  end_lst = [end]*len(hb_lst)
  loc_lst = [loc]*len(hb_lst)

  for hb in hb_lst:

    # print(hb)
    url = 'https://geets.gm.com/webapps/metex1/met_ex.exe?INPUTFORM=RUNREPORT&REPORTNAME=INVENTORY%3A+BY+EQUIPMENT+INFORMATION&PROMPT1=%25&PROMPT2=%25&PROMPT3=%25&PROMPT4=%25&PROMPT5=%25&PROMPT6=%25&PROMPT7=%25&PROMPT8=%25&PROMPT9=%25&PROMPT10={}&PROMPT11=%25&PROMPT12=%25'.format(hb)

    html = requests.get(url,verify=False)
    time.sleep(3)
    # print(html.text)

    html1 = '''<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
    <html><head><title>MET/EX INVENTORY: BY EQUIPMENT INFORMATION</title>
    <script>if (top != self){alert('Attention:\nThis application cannot run in a framed window.\n Attempting to reload..');top.location.replace(self.location.href);};</script>
    <script type='text/javascript' src='/metex1/script/mx_script.js'></script>
    <link rel='stylesheet' href='/metex1/mx_base_style.css' type='text/css' />
    <!--[if lt IE 11]><link rel='stylesheet' href='/metex1/mx_base_style_IE.css' type='text/css' /><![endif]-->
    <link rel="stylesheet" href="/metex1/default/mx_reportlist.css" type="text/css" />
    </head>
    <body bgcolor="" background="/">
    <div id='pagecontainer'>
    <h2>Metrology Xplorer | INVENTORY: BY EQUIPMENT INFORMATION</h2>
    <div id='formcontainer'>
    <table id='sqlreport' summary=''>
    <thead>
    <tr><th>Category</th><th>Sub Category</th><th>Equipment Number</th><th>Manufacturer</th><th>Model</th><th>Description</th><th>Serial Number</th><th>Segregate</th><th>System Role</th><th>Parent Equip#</th><th>Country of Origin</th><th>PTS Administrator</th><th>PTS Storage Location</th><th>PTS</th><th>Asset Tag#</th><th>Purchase Cost</th><th>Capitalized value</th><th>PO# / ISR#</th><th>Acquision Date</th><th>In Service Date</th><th>Owning Group Name</th><th>Status</th><th>Status Date</th><th>Trace</th><th>Last Inventory</th><th>Location record created by:</th><th>Location</th><th>Details</th><th>Loc Status</th><th>Test# Project# Workorder #</th><th>Event Date</th><th>Est. Return Date</th><th>Last Action Code</th><th>Cal-date</th><th>Due-Date</th>
    </thead><tbody>
    <tr><td>POWER</td><td>ANALYZER</td><td>HB002592</td><td>ANDERSON ELECTRIC</td><td>AC2600PD/XT2640-2CH  </td><td>POWER PROCESSING SYSTEM</td><td>26421908004</td><td>US_HB</td><td>NON-SYSTEM</td><td>&nbsp;</td><td>Y</td><td>Leah Mapletoft</td><td>BSL TC06</td><td>W27A</td><td>100030201823</td><td>494897</td><td>0</td><td>4300786822</td><td>11/19/2018</td><td>6/5/2020</td><td>Global Battery Systems Lab</td><td>HOLD</td><td>5/10/2021</td><td>F</td><td>5/12/2021</td><td>Veronica E Mapletoft</td><td>B2-7</td><td>PTS</td><td>HOLD</td><td>GMF3LRXIOMA</td><td>5/12/2021</td><td>&nbsp;</td><td>&nbsp;</td><td>8/12/2019</td><td>8/12/2020</td>
    </tbody><tfoot><tr><td></td><tr></tfoot>
    </table>
    <p>Total number of records for this report: 1</p>
    </div>
    <div class='footer'>
    <p><a href='/webapps/metex1/met_ex.exe'>Home</a></p>
    <p>
      Monday, December 6, 2021 15:15:03 |
      User: METEX |
      Serial#:   Expiration Date:  |
      CGI Version: 1.3017.1025
    </p>
    </div>
    </div>
    </body></html>'''

    # df = pd.read_html(html.text)
    # print(df[0]['Due-Date'][0])

    soup = BeautifulSoup(html.text, 'html.parser')
    # print(soup.prettify())
    table = soup.find(id='sqlreport')
    # print(table)
    # exit()

    try:
      rows = table.find_all('tr')
    except AttributeError:
      print(f'Trouble processing {hb}. Most likely a timeout. Moving on to next HB number.')
      continue

    print(f'Processing {hb}')
    for row in rows:
        #print(row)
        columns_h = row.find_all('th')
        columns_d = row.find_all('td')

        if len(columns_h) > 0:

            str_cols_h = [str(i) for i in columns_h]
            idx = str_cols_h.index('<th>Due-Date</th>')
            idx1 = str_cols_h.index('<th>Manufacturer</th>')
            idx2 = str_cols_h.index('<th>Description</th>')

        if len(columns_d) > idx:

            tmp = str(columns_d[idx]).split('<td>')[1].split('</td>')[0]
            cal_date = tmp if tmp != '\xa0' else 'N/A'
            cal_dates.append(cal_date)
            manufact.append(str(columns_d[idx1]).split('<td>')[1].split('</td>')[0])
            desc.append(str(columns_d[idx2]).split('<td>')[1].split('</td>')[0])

  data_dict = {'Equipment Number':hb_lst,'Location':loc_lst,'Manufacturer':manufact,'Description':desc,
                  'Cal Date':cal_dates,'Start Date':start_lst,'End Date':end_lst}
  return data_dict

def make_table(data_dict):

  # print(tabulate(data_dict,headers='keys'))
  doc = Document()
  table = doc.add_table(rows=len(data_dict['Equipment Number'])+1,cols=len(data_dict))

  for i,name in enumerate(data_dict.keys()):
    for j in range(len(table.rows)):

      if j==0:
        table.cell(j,i).text = name
      else:
        table.cell(j,i).text = data_dict[name][j-1] if data_dict[name][j-1] != '' else 'N/A'
  table.style = 'Table Grid'
  doc.save('Equipment table.docx')

if __name__ == '__main__':

  hb,start,end,loc = get_hb()
  data = get_cal_dates(hb,start,end,loc)
  make_table(data)
