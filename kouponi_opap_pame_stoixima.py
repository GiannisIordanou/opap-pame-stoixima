# -*- coding: utf-8 -*-

# Imports
import csv
import re
import sys
import urllib2
import json
import time
import datetime
import xlwt



# Data Functions
def read_opap_json():
        # Get matches
        try:
            url = 'http://www.opap.gr/web/services/rs/betting/availableBetGames/sport/program/4100/0/sport-1.json?localeId=el_GR'        
            data = json.load(urllib2.urlopen(url))
            print 'Data retrieved successfully.'
        except:
            data = ''
            print 'Error: Please try again later.'
        return data

def generate_bet_categories_dictionary(checkboxes_dict):
    bet_categories_dict = {0: u'1', 1: u'Χ', 2: u'2',
                                   6: u'ΠΓ', 7: u'ΠΧ', 8: u'ΠΦ'}
    for i,j  in zip(range(25, 27), checkboxes_dict[u'Under/Over']):
        bet_categories_dict[i] = j
        
    for i,j  in zip(range(29, 31), checkboxes_dict[u'Goal/No Goal']):
        bet_categories_dict[i] = j
    
    for i,j  in zip(range(21, 25), checkboxes_dict[u'Σύνολο Τερμάτων']):
        bet_categories_dict[i] = j
        
    for i,j  in zip(range(9, 12), checkboxes_dict[u'Ημίχρονο']):
        bet_categories_dict[i] = j
        
    for i,j  in zip(range(12, 21), checkboxes_dict[ u'Hμίχρονο/Τελικό']):
        bet_categories_dict[i] = j

    for i,j  in zip(range(3, 6), checkboxes_dict[u'Διπλή Ευκαιρία']):
        bet_categories_dict[i] = j
    
    for i,j  in zip(range(36, 72), checkboxes_dict[u'Ακριβές σκορ']):
        bet_categories_dict[i] = j
    return bet_categories_dict

def get_matches(data, standard_headers, checked_headers, bet_categories_dict):
    global matches
    matches = [] 
    matches_all = {}
    for ind, x in enumerate(data['betGames']):
        match_dict = {}
        tournament_key = x['tournamentId']
        hometeam_key = [i['value'] for i in x['properties']['prop'] if i['id'] == 30][0]
        awayteam_key = [i['value'] for i in x['properties']['prop'] if i['id'] == 31][0]
        div_key = [i['value'] for i in x['properties']['prop'] if i['id'] == 46][0]
        code_key = [i['value'] for i in x['properties']['prop'] if i['id'] == 6][0]
        match_dict[u'Α'] = code_key
        ee_key = [i['value'] for i in x['properties']['prop'] if i['id'] == 28][0]

        for i, j in zip([u'ΓΗΠ' , u'ΦΙΛΟΞ', u'Πρωτάθλ', u'Δ'], [hometeam_key, awayteam_key, tournament_key, div_key]):
            match_dict[i] = x[u'lexicon']['resources'][j]
    
        match_date_time = x['betEndDate']/1000.
        match_date, match_time = datetime.datetime.fromtimestamp(match_date_time).strftime('%d-%m-%Y %H:%M').split(' ')
        match_dict[u'Ημερ.'] = match_date
        match_dict[u'ΩΕ'] = match_time
        match_dict[u'Ε/Ε'] = ee_key


        for index, i in enumerate(x['codes']):
            #print '>', index, i
            i = dict(i)
            bet_category = i['code']['value']
            bet_odds = i['odd']
            if bet_category in bet_categories_dict.keys():
                    #print bet_categories_dict[i['code']['value']], ':', i['odd']
                    value = i['odd']
                    if value == 0.0:
                        value = '-'
                    else:
                        value = str(i['odd'])
                        
                    match_dict[bet_categories_dict[i['code']['value']]] = value
#         for i in match_dict.keys():
#             print 'Dict key:', i 
            
#             print u'Α' in match_dict.keys()
        matches_all[match_dict[u'Α']] = match_dict
    return matches_all

def get_data():
    
    data = read_opap_json()
    print 'data OK'
    
    checkboxes_dict = {u'Under/Over': ['U', 'O'], u'Goal/No Goal': ['G', 'NG'], u'Σύνολο Τερμάτων': ['G0-1', 'G2-3', 'G4-6', 'G7+'],
                       u'Διπλή Ευκαιρία': [u'1Χ', u'12', u'Χ2'], u'Ημίχρονο': [u'Ημιχρ. 1', u'Ημιχρ. Χ', u'Ημιχρ. 2'],
                       u'Hμίχρονο/Τελικό': [u'1-1', u'Χ-1', u'2-1', u'1-Χ', u'Χ-Χ', u'2-Χ', u'1-2', u'Χ-2', u'2-2'],
                       u'Ακριβές σκορ': [u'Σκορ 0-0', u'Σκορ 1-0', u'Σκορ 2-0', u'Σκορ 3-0', u'Σκορ 4-0', u'Σκορ 5+-0',
                                         u'Σκορ 2-1', u'Σκορ 3-1', u'Σκορ 4-1', u'Σκορ 5+-1', u'Σκορ 3-2', u'Σκορ 4-2',
                                         u'Σκορ 5+-2', u'Σκορ 4-3', u'Σκορ 5+-3', u'Σκορ 5+-4', u'Σκορ 1-1', u'Σκορ 2-2',
                                         u'Σκορ 3-3', u'Σκορ 4-4', u'Σκορ 5+-5+', u'Σκορ 4-5+', u'Σκορ 3-4', u'Σκορ 3-5+',
                                         u'Σκορ 2-3', u'Σκορ 2-4', u'Σκορ 2-5+', u'Σκορ 1-2', u'Σκορ 1-3', u'Σκορ 1-4',
                                         u'Σκορ 1-5+', u'Σκορ 0-1', u'Σκορ 0-2', u'Σκορ 0-3', u'Σκορ 0-4', u'Σκορ 0-5+']}
    
    bet_categories_dict = generate_bet_categories_dictionary(checkboxes_dict)
    print 'bet_categories_dict OK'
    
    standard_headers = [u'Ημερ.', u'Δ', u'ΩΕ', u'Α', u'Ε/Ε', u'1', u'ΠΓ', u'ΓΗΠ', u'Χ', u'ΦΙΛΟΞ', u'ΠΦ', u'2']
    checked_headers = checkboxes_dict.keys()
    
    matches_all = get_matches(data, standard_headers, checked_headers, bet_categories_dict)
    print 'matches_all OK'
    excel_headers = [u'Ημερ.', u'Δ', u'ΩΕ', u'Α', u'Ε/Ε', u'1', u'ΠΓ', u'ΓΗΠ', u'Χ', u'ΦΙΛΟΞ', u'ΠΦ', u'2']
    matches_all2 = []
    for i in sorted(matches_all.keys()):
        tmp = []
        for j in matches_all[i].keys():
            for headers in excel_headers:
                tmp.append(matches_all[i][headers])
        matches_all2.append(tmp)
#         print '>>', i, matches_all[i]
    return matches_all, matches_all2



# Get data 
checkboxes_dict = {u'Under/Over': ['U', 'O'], u'Goal/No Goal': ['G', 'NG'],
                       u'Σύνολο Τερμάτων': ['G0-1', 'G2-3', 'G4-6', 'G7+'], 
                       u'Διπλή Ευκαιρία': [u'1Χ', u'12', u'Χ2'], u'Ημίχρονο': [u'Ημιχρ. 1', u'Ημιχρ. Χ', u'Ημιχρ. 2'],
                       u'Hμίχρονο/Τελικό': [u'1-1', u'Χ-1', u'2-1', u'1-Χ', u'Χ-Χ', u'2-Χ', u'1-2', u'Χ-2', u'2-2'],
                       u'Ακριβές σκορ': [u'Σκορ 0-0', u'Σκορ 1-0', u'Σκορ 2-0', u'Σκορ 3-0', u'Σκορ 4-0', u'Σκορ 5+-0',
                                         u'Σκορ 2-1', u'Σκορ 3-1', u'Σκορ 4-1', u'Σκορ 5+-1', u'Σκορ 3-2', u'Σκορ 4-2',
                                         u'Σκορ 5+-2', u'Σκορ 4-3', u'Σκορ 5+-3', u'Σκορ 5+-4', u'Σκορ 1-1', u'Σκορ 2-2',
                                         u'Σκορ 3-3', u'Σκορ 4-4', u'Σκορ 5+-5+', u'Σκορ 4-5+', u'Σκορ 3-4', u'Σκορ 3-5+',
                                         u'Σκορ 2-3', u'Σκορ 2-4', u'Σκορ 2-5+', u'Σκορ 1-2', u'Σκορ 1-3', u'Σκορ 1-4',
                                         u'Σκορ 1-5+', u'Σκορ 0-1', u'Σκορ 0-2', u'Σκορ 0-3', u'Σκορ 0-4', u'Σκορ 0-5+']}


matches_all, matches_all2 = get_data()
print 'Data retrieved.'



# Get proper headers    
excel_headers = [u'Ημερ.', u'Δ', u'ΩΕ', u'Α', u'Ε/Ε', u'1', u'ΠΓ', u'ΓΗΠ', u'Χ', u'ΦΙΛΟΞ', u'ΠΦ', u'2', u'Διπλή Ευκαιρία', u'Ημίχρονο', u'Hμίχρονο/Τελικό', u'Σύνολο Τερμάτων', 'Under/Over', 'Goal/No Goal', u'Ακριβές σκορ']
matches_all_excel = []
for i in sorted(matches_all.keys()):
    tmp = []
    for headers in excel_headers:
        if headers in checkboxes_dict.keys():
            #print headers
            if len(checkboxes_dict[headers]) > 1:
                for sub_head in checkboxes_dict[headers]:
                    tmp.append(matches_all[i][sub_head])
            else:
                tmp.append(matches_all[i][headers])
        else:
            tmp.append(matches_all[i][headers])
    matches_all_excel.append(tmp)    
print 'Excel Data retrieved.'

start_date = matches_all_excel[0][0].replace('-', '/')
end_date = matches_all_excel[-1][0].replace('-', '/')
date_kouponi = ' - '.join([start_date, end_date])

# Prepare Excel file
book = xlwt.Workbook()

# Sheet 1
sheet_name_1 = unicode(u'Κουπόνι ' + date_kouponi).replace('/', '.')
sheet1 = book.add_sheet(sheet_name_1)
print 'Created sheet:', sheet_name_1

sheet1.set_panes_frozen(True)
sheet1.set_horz_split_pos(7)

sheet1.set_vert_split_pos(13) 
sheet1.set_remove_splits(True)

# Add custom colours
xlwt.add_palette_colour("custom_colour_1", 0x28)
book.set_colour_RGB(0x28, 0, 112, 192)

xlwt.add_palette_colour("custom_colour_2", 0x21)
book.set_colour_RGB(0x21, 0, 146, 0)

xlwt.add_palette_colour("custom_colour_3", 0x22)
book.set_colour_RGB(0x22, 234, 244, 233)

xlwt.add_palette_colour("custom_colour_4", 0x23)
book.set_colour_RGB(0x23, 255, 255, 255)

xlwt.add_palette_colour("custom_colour_5", 0x24)
book.set_colour_RGB(0x24, 209, 231, 208)

xlwt.add_palette_colour("custom_colour_6", 0x25)
book.set_colour_RGB(0x25, 255, 255, 203)

xlwt.add_palette_colour("custom_colour_7", 0x26)
book.set_colour_RGB(0x26, 234, 254, 169)

print 'Custom colours added to palette.'

# EasyXF Styles

# Title
style_1 = xlwt.easyxf("align: horiz center, vert bottom; font: height 440, name Times New Roman, colour_index white; \
                      pattern: pattern solid, fore_color custom_colour_1, back_color white; borders: top medium, bottom medium, left medium, right medium;")

# Style : Ημερ., Σύνολο Τερμάτων, Under/Over, Goal/No Goal
style_2 = xlwt.easyxf("align: horiz center, vert bottom; font: height 240, name Times New Roman, colour_index white; \
                      pattern: pattern solid, fore_color custom_colour_2, back_color white; borders: top medium, bottom medium, left medium, right medium;")

# Style : Δ, ΓΗΠ,  ΦΙΛΟΞ
style_3 = xlwt.easyxf("align: horiz center, vert bottom; font: height 240, name Times New Roman; \
                      pattern: pattern solid, fore_color custom_colour_3, back_color white; borders: top medium, bottom medium, left medium, right medium;")

# Style : ΩΕ, Ε/Ε
style_4 = xlwt.easyxf("align: horiz center, vert bottom; font: height 240, name Times New Roman; \
                      pattern: pattern solid, fore_color custom_colour_4, back_color white; borders: top medium, bottom medium, left medium, right medium;")

# Style : Α
style_5 = xlwt.easyxf("align: horiz center, vert bottom; font: height 240, name Times New Roman; \
                      pattern: pattern solid, fore_color custom_colour_5, back_color white; borders: top medium, bottom medium, left medium, right medium;")

# Style : 1, Χ, 2
style_6 = xlwt.easyxf("align: horiz center, vert bottom; font: height 240, name Times New Roman; \
                      pattern: pattern solid, fore_color custom_colour_6, back_color white; borders: top medium, bottom medium, left medium, right medium;")

# Style : ΠΓ, ΠΦ
style_7 = xlwt.easyxf("align: horiz center, vert bottom; font: height 240, name Times New Roman; \
                      pattern: pattern solid, fore_color custom_colour_7, back_color white; borders: top medium, bottom medium, left medium, right medium;")

# Sheet 1 
start_row = 4
star_col = 1

# Write cells

# Title
sheet1.write_merge(4,4,1,71, u'Κουπόνι ΟΠΑΠ ΠΑΜΕ ΣΤΟΙΧΗΜΑ ' + date_kouponi, style_1)

print 'Title has been written.'

#Headers
sheet1.write_merge(5, 6, 1, 1, u'Ημερ.', style_2)
sheet1.write_merge(5, 6, 2, 2, u'Δ', style_3)
sheet1.write_merge(5, 6, 3, 3, u'ΩΕ', style_4)
sheet1.write_merge(5, 6, 4, 4, u'Α', style_5)
sheet1.write_merge(5, 6, 5, 5, u'Ε/Ε', style_4)
sheet1.write_merge(5, 6, 6, 6, 1, style_6)
sheet1.write_merge(5, 6, 7, 7, u'ΠΓ', style_7)
sheet1.write_merge(5, 6, 8, 8, u'ΓΗΠΕΔΟΥΧΟΣ', style_3)
sheet1.write_merge(5, 6, 9, 9, u'Χ', style_6)
sheet1.write_merge(5, 6, 10, 10, u'ΦΙΛΟΞΕΝΟΥΜΕΝΟΣ', style_3)
sheet1.write_merge(5, 6, 11, 11, u'ΠΦ', style_7)
sheet1.write_merge(5, 6, 12, 12, 2, style_6)
sheet1.write_merge(5, 5, 13, 15, u'Διπλή Ευκαιρία', style_2)
sheet1.write_merge(5, 5, 16, 18, u'Ημίχρονο', style_2)
sheet1.write_merge(5, 5, 19, 27, u'Hμίχρονο/Τελικό', style_2)
sheet1.write_merge(5, 5, 28, 31, u'Σύνολο Τερμάτων', style_2)
sheet1.write_merge(5, 5, 32, 33, u'Under/Over', style_2)
sheet1.write_merge(5, 5, 34, 35, u'Goal/No Goal', style_2)
sheet1.write_merge(5, 5, 36, 71, u'Ακριβές σκορ', style_2)

print 'Headers have been written.'
#Subheaders
col_ind = 12
for x in checkboxes_dict[u'Διπλή Ευκαιρία']:
    col_ind += 1
    try:
        x = unicode(x)
    except:
        print x
    sheet1.write(6, col_ind, x, style_6)

for x in checkboxes_dict[u'Ημίχρονο']:
    col_ind += 1
    x = unicode(x)
    sheet1.write(6, col_ind, x, style_6)
    
for x in checkboxes_dict[u'Hμίχρονο/Τελικό']:
    col_ind += 1
    x = unicode(x)
    sheet1.write(6, col_ind, x, style_6)
    
for x in checkboxes_dict[u'Σύνολο Τερμάτων']:
    col_ind += 1
    x = unicode(x)
    sheet1.write(6, col_ind, x, style_6)

for x in checkboxes_dict[u'Under/Over']:
    col_ind += 1
    x = unicode(x)
    sheet1.write(6, col_ind, x, style_6)
    
for x in checkboxes_dict[u'Goal/No Goal']:
    col_ind += 1
    x = unicode(x)
    sheet1.write(6, col_ind, x, style_6)
    
for x in checkboxes_dict[u'Ακριβές σκορ']:
    col_ind += 1
    try:
        x = unicode(x).strip(u'Σκορ ')
    except:
        print x
    sheet1.write(6, col_ind, x, style_6)

print 'Subheaders have been written.'
    
# Column widths
sheet1.col(1).width = 256 * 12
sheet1.col(2).width = 256 * 7
sheet1.col(3).width = 256 * 7
sheet1.col(4).width = 256 * 5
sheet1.col(5).width = 256 * 5
for i in range(6, 13, 3) + range(13, 21):
    sheet1.col(i).width = 256 * 9
sheet1.col(7).width = 256 * 5
sheet1.col(8).width = 256 * 25
sheet1.col(10).width = 256 * 25
sheet1.col(11).width = 256 * 5

print 'Column widths adjusted.'

print 'Excel preparation completed.'

# Column styles dictionary
style_dict = {1: style_2, 4: style_5}
for i in [2, 8, 10]:
    style_dict[i] = style_3

for i in [3, 5]:
    style_dict[i] = style_4
    
for i in range(6, 13,3) + range(13, 72):
    style_dict[i] = style_6    

for i in [7, 11]:
    style_dict[i] = style_7    


text_headers = [u'Ημερ.', u'Δ',  u'ΩΕ',u'ΓΗΠ',  u'ΦΙΛΟΞ']

start_row = 7
start_col = 1

# Write data to excel file
for row_ind, each_row in enumerate(matches_all_excel):
    row_ind = row_ind + start_row
    for col_ind, each_col in enumerate(each_row):
        col_ind = col_ind + start_col
        if col_ind not in [1, 2, 3, 8, 10]:
            if each_col != '-':
                each_col = float(each_col)
        sheet1.write(row_ind, col_ind, each_col, style_dict[col_ind])
        
print 'Data inserted to cells.'

book.save('kouponi_opap.xls')
print 'Excel file saved.'
