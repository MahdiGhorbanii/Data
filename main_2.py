# Importing Libraries
import json
import jira.client
from jira.client import JIRA
import pandas as pd
import xlsxwriter
import os
from datetime import datetime, timedelta
import numpy as np
from slack_sdk import WebClient
import sys


API_TOCKEN = "API TOKEN"
EMAIL = "EMAIL"
server = "JIAR SERVER"
jql = "FILTERS"
MY_SLACK_ID = "SLACK ID"
CHANNEL = "CHANNEL NAME"

# Authentication
client = WebClient(SLACK_BOT_TOKEN)
auth_test = client.auth_test()


# Importing Json files
with open('domains.json', "r") as business_domains:
    business_domains = json.load(business_domains)
with open('owners.json', "r") as component_owners:
    component_owners = json.load(component_owners)
with open('ticket_links.json', "r") as ticket_links:
    ticket_links = json.load(ticket_links)
with open('total_tickets.json', "r") as total_tickets:
    total_tickets = json.load(total_tickets)
with open('priority_links.json', "r") as priority_links:
    priority_links = json.load(priority_links)
with open('priority_stats.json', "r") as priority_stats:
    priority_stats = json.load(priority_stats)


# Importing the Dataset from Jira through API


jira = JIRA(options={'server': server}, basic_auth=(EMAIL, API_TOCKEN))
jira_issues = jira.search_issues(jql, maxResults=0)

issues = pd.DataFrame()
for issue in jira_issues:
    d = {
        'created': issue.fields.created,
        'components': issue.fields.components,
        'priority': issue.fields.priority,
        'status': issue.fields.status.name
    }

    issues = issues.append(d, ignore_index=True)

# Keeping only the first element of the components
data_set = pd.DataFrame(issues)

Comps = []
for i in data_set['components']:
    Comps.append(i[0].name)

data_set['Components'] = Comps
data_set.drop("components", axis=1, inplace=True)

# Formatting the date appearance
for i in range(len(data_set['created'])):
    data_set['created'][i] = data_set['created'][i].split('T')[0]

try:
    component_leaders = []

    for j in data_set['Components']:
        for i in range(len(component_owners['component_owners'].values())):
            if j in list(component_owners['component_owners'].values())[i]:
                component_leaders.append(list(component_owners['component_owners'].keys())[i])

    data_set['Component Lead'] = component_leaders

    domains = []

    for j in data_set['Components']:
        for i in range(len(business_domains['business_domains'].keys())):
            if j in list(business_domains['business_domains'].values())[i]:
                domains.append(list(business_domains['business_domains'].keys())[i])

    data_set['Domains'] = domains

except:
    message = "Hii,\nNew or Unknown Doamians,\nThe file can not be created,\nPlease fix the code and run it again"
    client.chat_postMessage(
        channel=MY_SLACK_ID,
        text=message
    )

    sys.exit()

# TREND

time_now = datetime.now()
today = time_now.date()

today1 = datetime.today()
yesterday = today1 - timedelta(days=1)
yesterday = yesterday.date()

today1 = datetime.today()
yesterday = today1 - timedelta(days=1)
yesterday = yesterday.date()

# Data accumulation for TODAY
today_trend = []

for i in range(len(data_set['created'])):
    if data_set['created'][i] == str(today):
        today_trend.append(data_set['Component Lead'][i])

frequency_today = {}
for item in today_trend:
   # checking the element in dictionary
    if item in frequency_today:
      # incrementing the count
      frequency_today[item] += 1
    else:
      # initializing the count
      frequency_today[item] = 1

# data accumulation for YESTERDAY
yesterday_trend = []


for i in range(len(data_set['created'])):
    if data_set['created'][i] == str(yesterday):
        yesterday_trend.append(data_set['Component Lead'][i])

frequency_yesterday = {}
for item in yesterday_trend:
   # checking the element in dictionary
    if item in frequency_yesterday:
      # incrementing the counr
      frequency_yesterday[item] += 1
    else:
      # initializing the count
      frequency_yesterday[item] = 1

Com_leaders = np.unique(component_leaders)

Com_leaders = Com_leaders.tolist()

list_of_trends = [0] * len(Com_leaders)

for i in Com_leaders:
    if i in frequency_today.keys():
        list_of_trends[Com_leaders.index(i)] = frequency_today[i]

list_of_trends_yes = [0] * len(Com_leaders)

for i in Com_leaders:
    if i in frequency_yesterday.keys():
        list_of_trends_yes[Com_leaders.index(i)] = frequency_yesterday[i]

Total_trend = np.subtract(list_of_trends, list_of_trends_yes)
Total_trend = Total_trend.tolist()

# Creating the DataFrame
component_leaders = []

for j in data_set['Components']:
    for i in range(len(component_owners['component_owners'].values())):
        if j in list(component_owners['component_owners'].values())[i]:
            component_leaders.append(list(component_owners['component_owners'].keys())[i])

data_set['Domains'] = domains

new_set = data_set

data_set['Component Lead'] = component_leaders

new_set = new_set.reset_index().groupby(['Component Lead', 'status'])['status'].count().unstack()

column_list = ['Groomed', 'In Progress', 'Open', 'Review', 'Validation']

for i in column_list:
    new_set[i] = new_set[i].fillna(0).astype(int)

new_set['Open'] = new_set['Open'] + new_set['Groomed']

new_set['Work In Progress'] = new_set['In Progress'] + new_set['Review']

new_set['Total'] = new_set['Open'] + new_set['Work In Progress'] + new_set['Validation']

new_set['Trend'] = Total_trend

final_columns = ['Open', 'Work In Progress','Validation', 'Total', 'Trend']
for i in new_set.columns:
    if i not in final_columns:
        new_set = new_set.drop(i, axis=1)

columns_titles = ["Open", "Work In Progress","Validation", "Total", "Trend"]
new_set = new_set.reindex(columns=columns_titles)

final_list = []
for i in range(len(new_set.index)):
    each_person = []
    each_person.append(new_set.index[i])
    for j in range(5):
        each_person.append(new_set.iloc[i][j])
    final_list.append(each_person)

# Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('PATH')
worksheet = workbook.add_worksheet(name=f'Report_{today}')

# Add a sample alternative link format.

mid_cell = workbook.add_format({
    'bg_color': '#b7b7b7'
})

font_format_0 = workbook.add_format({
    'font_color': 'black',
    'border': 1,
    'bold':      1,
    'underline':  1,
    'font_size':  12,
    'bg_color': '#b7b7b7'
})
font_format_0.set_align('center')
font_format_0.set_align('vcenter')


font_format = workbook.add_format({
    'font_color': 'black',
    'bold':      1,
    'border': 1,
    'underline':  1,
    'font_size':  12,
    'bg_color': '#b7b7b7'
})
font_format.set_align('center')
font_format.set_align('vcenter')

font_format000 = workbook.add_format({
    'font_color': 'black',
    'bold':      1,
    'border': 1,
    'underline':  1,
    'font_size':  12,
    'bg_color': '#cfe2f3'
})
font_format000.set_align('center')
font_format000.set_align('vcenter')

font_format_special = workbook.add_format({
    'border': 1,
    'font_color': 'black',
    'bold':      1,
    'font_size':  12,
    'bg_color': '#b7b7b7'
})

font_format_special.set_align('center')
font_format_special.set_align('vcenter')


font_format0 = workbook.add_format({
    'bold': 1,
    'border': 1,
    'bg_color': '#efefef'
})

font_format0.set_align('center')
font_format0.set_align('vcenter')

font_format00 = workbook.add_format({
    'bold': 1,
    'border': 1,
    'bg_color': '#ffffff'
})

font_format00.set_align('center')
font_format00.set_align('vcenter')

font_format1 = workbook.add_format({
    'border': 1,
    'bg_color': '#efefef'
})

font_format1.set_align('center')
font_format1.set_align('vcenter')

font_format2 = workbook.add_format({
    'border': 1,
    'bg_color': '#ffffff'
})

font_format2.set_align('center')
font_format2.set_align('vcenter')

color_range = ['#57bb8a', '#67bf8b', '#78c38d', '#89c78e', '#9acb90',
               '#abd091', '#bbd493', '#cceadb', '#ddf1e7', '#eef8f3',
               '#ffffff', '#fdf2f1', '#fbe5e4', '#f8d8d5', '#f6cbc8',
               '#f3beb9', '#f0b1ab', '#eea49d', '#eb978f', '#e98a82',
               '#e67c73']
color_num = [-10, -9, -8, -7, -6, -5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]

cell_format = workbook.add_format()

cell_format.set_align('left')

cell_format1 = workbook.add_format()

cell_format1.set_align('center')
cell_format1.set_align('vcenter')

header_color = '#d9d2e9'
body_color = '#ffe599'

data_format1 = workbook.add_format()
data_format1.set_align('center')
data_format1.set_align('vcenter')

font_format3 = workbook.add_format({
    'border': 1,
    'bg_color': '#d0e0e3',
    'font_color': '#0000ff'
})

font_format3.set_align('center')
font_format3.set_align('vcenter')

font_format4 = workbook.add_format({
    'border': 1,
    'bg_color': '#d0e0e3'
})

font_format4.set_align('center')
font_format4.set_align('vcenter')

font_format5 = workbook.add_format({
    'border': 6,
    'bg_color': '#d0e0e3'
})

font_format5.set_align('center')
font_format5.set_align('vcenter')

font_format6 = workbook.add_format({
    'border': 1,
    'bg_color': '#eef7ff'
})
font_format6.set_align('center')
font_format6.set_align('vcenter')

font_format7 = workbook.add_format({
    'border': 1,
    'font_color': '#0000ff',
    'bg_color': '#efefef'
})
font_format7.set_align('center')
font_format7.set_align('vcenter')


font_format7_a = workbook.add_format({
    'border': 1,
    'font_color': '#0000ff',
    'bg_color': '#efefef'
})
font_format7_a.set_align('center')
font_format7_a.set_align('vcenter')

font_format7_b = workbook.add_format({
    'border': 1,
    'font_color': '#0000ff',
    'bg_color': '#ffffff'
})
font_format7_b.set_align('center')
font_format7_b.set_align('vcenter')

font_format8 = workbook.add_format({
    'border': 1,
    'bg_color': '#d0e0e3',
    'bold': 1
})

font_format8.set_align('center')
font_format8.set_align('vcenter')

# transforming the data to the excel file
row = 1
for item, first_value, second_value, third_value, fourth_value, fifth_value in final_list:
    if row % 2 == 0:
        col = 0
        worksheet.write(row, col, item, font_format0)
        col += 4
        worksheet.write(row, col, fourth_value, font_format1)
        col += 1
        if fifth_value in color_num:

            conditional_form = workbook.add_format({'border': 1, 'bg_color': color_range[color_num.index(fifth_value)]})
            conditional_form.set_align('center')
            conditional_form.set_align('vcenter')

            worksheet.write(row, col, fifth_value, conditional_form)

        elif fifth_value < -10:
            conditional_form = workbook.add_format({'border': 1, 'bg_color': '#57bb8a'})
            conditional_form.set_align('center')
            conditional_form.set_align('vcenetr')

            worksheet.write(row, col, fifth_value, conditional_form)

        elif fifth_value > 10:
            conditional_form = workbook.add_format({'border': 1, 'bg_color': '#e98a82'})
            conditional_form.set_align('cenetr')
            conditional_form.set_align('vcenter')

            worksheet.write(row, col, fifth_value, )
    else:
        col = 0
        worksheet.write(row, col, item, font_format00)
        col += 4
        worksheet.write(row, col, fourth_value, font_format2)
        col += 1
        if fifth_value in color_num:

            conditional_form = workbook.add_format({'border': 1, 'bg_color': color_range[color_num.index(fifth_value)]})
            conditional_form.set_align('center')
            conditional_form.set_align('vcenter')

            worksheet.write(row, col, fifth_value, conditional_form)

        elif fifth_value < -10:
            conditional_form = workbook.add_format({'border': 1, 'bg_color': '#57bb8a'})
            conditional_form.set_align('center')
            conditional_form.set_align('vcenetr')

            worksheet.write(row, col, fifth_value, conditional_form)

        elif fifth_value > 10:
            conditional_form = workbook.add_format({'border': 1, 'bg_color': '#e98a82'})
            conditional_form.set_align('cenetr')
            conditional_form.set_align('vcenter')

            worksheet.write(row, col, fifth_value, )

    row += 1

worksheet.set_column(0, 0, 30, cell_format1)

worksheet.set_column(1, 7, 18, cell_format)

worksheet.write(len(new_set.index)+1, 0, "Total", font_format8)
worksheet.write_url(f'B{len(new_set.index)+2}', total_tickets["Total_open"], string=str(sum(new_set['Open'])), cell_format=font_format3)
worksheet.write_url(f'C{len(new_set.index)+2}', total_tickets["Total_wip"], string=str(sum(new_set['Work In Progress'])), cell_format=font_format3)
worksheet.write_url(f'D{len(new_set.index)+2}', total_tickets["Total_validation"], string=str(sum(new_set['Validation'])), cell_format=font_format3)
worksheet.write_url(f'E{len(new_set.index)+2}', total_tickets["total"], string=str(sum(new_set['Total'])), cell_format=font_format3)
worksheet.write(len(new_set.index)+1, 5, '-', font_format4)


columns = ['B', 'C', 'D']

status = 0
values = 1
for j in columns:
    for i in range(len(new_set.index)):
        if i%2!=0:
            worksheet.write_url(f'{j}{i + 2}', ticket_links['links'][new_set.index[i]][status], string=str(final_list[i][values]),
                                cell_format=font_format7_a)
        else:
            worksheet.write_url(f'{j}{i + 2}', ticket_links['links'][new_set.index[i]][status],
                                string=str(final_list[i][values]),
                                cell_format=font_format7_b)

    status += 1
    values += 1

row = 0
col = 0

columns_titles = ["Component Leaders", "Open", "Work In Progress","Validation", "Total", "Trend"]

for i in columns_titles:
    worksheet.write(row, col, i, font_format_0)
    col += 1

# Pie-Chart
chart = workbook.add_chart({'type': 'pie'})

pie_chart_total = new_set['Total'].values
pie_chart_total = pie_chart_total.tolist()
data = [
    [item[0] for item in final_list],
    pie_chart_total,
]


chart.add_series({
    'name': 'Component Leaders vs Total',
    'categories': f'=Report_{today}!$A$2:$A${len(new_set.index)}',
    'values':     f'=Report_{today}!$E$2:$E${len(new_set.index)}',
    'data_labels': {'percentage': True}
})

chart.set_size({'width': 650, 'height': 500})

chart.set_style(15)

worksheet.insert_chart('H1', chart)

# Report based on priorities
new_set1 = data_set

data_set['priority'] = [str(item) for item in data_set['priority']]

new_set1 = new_set1.reset_index().groupby(['Component Lead', 'priority'])['priority'].count().unstack()

column_list = ['P0', 'Team Blocker', 'Company Blocker']

for i in column_list:
    new_set1[i] = new_set1[i].fillna(0).astype(int)


try:
    new_set1 = new_set1.drop(['P1', 'P2', 'P3'], axis=1)
except:
    pass

new_set1['Total'] = new_set1['Company Blocker'] + new_set1['Team Blocker'] + new_set1['P0']

columns_titles = ["Company Blocker", "Team Blocker","P0", "Total"]
new_set1 = new_set1.reindex(columns=columns_titles)

row = 24
col = 0

columns_titles = ["Component Leaders", "Company Blocker", "Team Blocker", "P0", "Total"]

for i in columns_titles:
    worksheet.write(row, col, i, font_format_0)
    col += 1

final_list1 = []
for i in range(len(new_set1.index)):
    each_person = []
    each_person.append(new_set1.index[i])
    for j in range(4):
        each_person.append(new_set1.iloc[i][j])
    final_list1.append(each_person)

row = 25
for item, a, b, c, total in final_list1:
    if row % 2 == 0:
        col = 0
        worksheet.write(row, col, item, font_format0)
        col += 4
        worksheet.write(row, col, total, font_format1)
    else:
        col = 0
        worksheet.write(row, col, item, font_format00)
        col += 4
        worksheet.write(row, col, total, font_format2)
    row += 1


columns = ['B', 'C', 'D']
status = 0
values = 1

for j in columns:

    for i in range(len(new_set1.index)):
        if i%2!=0:
            worksheet.write_url(f'{j}{i + 26}', priority_links['priority_links'][new_set1.index[i]][status],
                                string=str(final_list1[i][values]), cell_format=font_format7_a)
        else:
            worksheet.write_url(f'{j}{i + 26}', priority_links['priority_links'][new_set1.index[i]][status],
                                string=str(final_list1[i][values]), cell_format=font_format7_b)
    status += 1
    values += 1

worksheet.write(len(new_set1.index)+25, 0, "Total", font_format8)
worksheet.write_url(f'B{len(new_set1.index)+26}', total_tickets["Total_company_blocker"], string=str(sum(new_set1['Company Blocker'])), cell_format=font_format3)
worksheet.write_url(f'C{len(new_set1.index)+26}', total_tickets["Total_team_blocker"], string=str(sum(new_set1['Team Blocker'])), cell_format=font_format3)
worksheet.write_url(f'D{len(new_set1.index)+26}', total_tickets["Total_p0"], string=str(sum(new_set1['P0'])), cell_format=font_format3)
worksheet.write_url(f'E{len(new_set1.index)+26}', total_tickets["Total_blockers"], string=str(sum(new_set1['Total'])), cell_format=font_format3)


# Blockers second slide
data_set['priority'] = [str(item) for item in data_set['priority']]

for i in range(len(data_set['status'])):
    if data_set['status'][i] == 'Groomed':
        data_set['status'][i] = 'Open'

for i in range(len(data_set['status'])):
    if data_set['status'][i] == 'Review':
        data_set['status'][i] = 'In Progress'

status_list = ['Open', 'In Progress', 'Validation']

for i in range(len(data_set['status'])):
    if data_set['status'][i] not in status_list:
        data_set.drop(index=i, axis=0, inplace=True)
data_set.reset_index(inplace=True)
data_set.drop(columns='index', inplace=True, axis=1)

column_list = ['P0', 'Team Blocker', 'Company Blocker']
for i in range(len(data_set['priority'])):
    if data_set['priority'][i] not in column_list:
        data_set.drop(index=i, axis=0, inplace=True)
data_set.reset_index(inplace=True)
data_set.drop(columns='index', inplace=True, axis=1)

name_lst = list(data_set['Component Lead'].unique())
prios = ['Company Blocker', 'Team Blocker', 'P0']
stats = ['Open', 'In Progress', 'Validation']

final_prio_lst = [[] for _ in range(len(name_lst))]
for i in range(len(name_lst)):
    final_prio_lst[i].append(name_lst[i])

n = 0
for j in name_lst:
    for w in range(3):
        for q in range(3):
            for i in range(len(data_set)):
                if (data_set.iloc[i, :][4] == j) and (data_set.iloc[i, :][1] == prios[w]) and (
                        data_set.iloc[i, :][2] == stats[q]):
                    n += 1
            final_prio_lst[name_lst.index(j)].append(n)
            n = 0

worksheet2 = workbook.add_worksheet("Priority Status")

worksheet2.set_column(0, 0, 30, cell_format1)
worksheet2.set_column(1, 11, 18, cell_format)

row = 2

for item in final_prio_lst:
    if row % 2 == 0:
        col = 0
        worksheet2.write(row, col, item[0], font_format0)
    else:
        col = 0
        worksheet2.write(row, col, item[0], font_format00)

    row += 1


columns_2 = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']

value_1 = 0
value_2 = 1
for j in columns_2:
    if j == 'E':
        worksheet2.set_column('E:E', 0.25)
        for z in range(len(final_prio_lst)+1):
            worksheet2.write(z+1, 4, '', mid_cell)
    elif j == 'I':
        worksheet2.set_column('I:I', 0.25)
        for s in range(len(final_prio_lst)+1):
            worksheet2.write(s+1, 8, '', mid_cell)
    else:
        for i in range(len(final_prio_lst)):
            if i%2==0:
                 worksheet2.write_url(f'{j}{i + 3}', priority_stats['priority_stats'][final_prio_lst[i][0]][value_1], string=str(final_prio_lst[i][value_2]),
                                     cell_format=font_format7_a)
            else:
                worksheet2.write_url(f'{j}{i + 3}', priority_stats['priority_stats'][final_prio_lst[i][0]][value_1],
                                     string=str(final_prio_lst[i][value_2]),
                                     cell_format=font_format7_b)

        value_2 += 1
        value_1 += 1


n = 1
column_1 = [2, 6, 10]
list_1 = ['Company Blocker', 'Team Blocker', 'P0']
list_2 = ['Open', 'WIP', 'Validation']

for i in range(3):
     worksheet2.write(0, column_1[i], list_1[i], font_format000)
     for j in range(3):
         if (n != 4) and (n != 8):
             worksheet2.write(1, n, list_2[j], font_format_special)
             n += 1
         else:
             worksheet2.write(1, n+1, list_2[j], font_format_special)
             n += 2

worksheet2.write(1, 0, 'Component Leader', font_format)
worksheet2.write(0, 0, '', font_format)
worksheet2.write(0, 1, '', font_format000)
worksheet2.write(0, 3, '', font_format000)
worksheet2.write(0, 4, '', font_format_special)
worksheet2.write(0, 5, '', font_format000)
worksheet2.write(0, 7, '', font_format000)
worksheet2.write(0, 8, '', font_format_special)
worksheet2.write(0, 9, '', font_format000)
worksheet2.write(0, 11, '', font_format000)

#  Third worksheet
worksheet3 = workbook.add_worksheet(name='TPMs')

for i in range(len(component_owners['component_owners'].keys())):
    worksheet3.set_column(0, i, 40)

row=0
col=0

for i in component_owners['component_owners'].keys():
    worksheet3.write(row, col, i, font_format5)
    col += 1

component_owners_values = []
for i in component_owners['component_owners'].values():
    component_owners_values.append(i)

length_list = []
for i in component_owners_values:
    length_list.append(len(i))
max_length = max(length_list)

col = 0
row = 1
for i in range(len(component_owners['component_owners'].keys())):
    for j in range(max_length + 5):
        worksheet3.write(row, col, '', font_format6)
        row += 1
    col += 1
    row = 1

row=1
col=0


for i in component_owners_values:
    for j in i:
        worksheet3.write(row, col, j, font_format6)
        row += 1
    row = 1
    col += 1

# Fourth Worksheet
worksheet4 = workbook.add_worksheet(name='Business Domains')
for i in range(len(business_domains['business_domains'].keys())):
    worksheet4.set_column(0, i, 40)

row=0
col=0

for i in business_domains['business_domains'].keys():
    worksheet4.write(row, col, i, font_format5)
    col += 1

business_domains_values = []
for i in business_domains['business_domains'].values():
    business_domains_values.append(i)

# Max length of the lists for formatting the cells
length_list1 = []
for i in business_domains_values:
    length_list1.append(len(i))
max_length1 = max(length_list1)

col = 0
row = 1
for i in range(len(business_domains['business_domains'].keys())):
    for j in range(max_length1 + 5):
        worksheet4.write(row, col, '', font_format6)
        row += 1
    col += 1
    row = 1

row=1
col=0


for i in business_domains_values:
    for j in i:
        worksheet4.write(row, col, j, font_format6)
        row += 1
    row = 1
    col += 1


workbook.close()

upload_text_file = client.files_upload(
      channels=CHANNEL,
      title="Daily Report",
      file='PATH',
      initial_comment="Here is the daily Report")


