import xlrd 
from cElementTree import *
import pymongo
import time
import math
import json
import Levenshtein
from operator import itemgetter

#To-Do: Move to pubschls class
#Only works for California
def open_db(file_name):
    book = None
    try:
         book = xlrd.open_workbook(str(file_name))
    except IOError:
         print str(file_name) + " doesn't exist"

    if book is not None:
        sheet_names = book.sheet_names()
        sheets = []

        for n in sheet_names:
            sheet = book.sheet_by_name(n)
            sheets.append(sheet)

        return sheets
    else:
        return []

#Moves to pubschls class
#Works for California
def get_column_names(sheet):
    curr_cell = -1
    column_names = []
    num_cells = sheet.ncols - 1

    while curr_cell < num_cells:
        curr_cell += 1
        column_name = sheet.cell_value(0, curr_cell)
        column_names.append(column_name)

    return [ c.encode('ascii', 'ignore') for c in column_names ]

#Move to pubschls Class
def convert_list_to_object(sheet, l):
    column_names = get_column_names(sheet)

    if len(column_names) == len(l):
        converted = {}
        for i in range(len(l)):
            converted[column_names[i]] = l[i]

        return converted
    else:
        print "List must be same length as sheet.ncells-1"
        return 

#To-Do: Move to Pubschls class
#for ca, school status is the 4th column
def get_ca_schools():
    sheet = open_db("pubschls.xls")[0]
    i = 0
    num_rows = sheet.nrows - 1
    #closed = unicode("Closed")
    #closed_schools = []
    '''
    school_index = {
        "Elementary Schools (Public)": [],
        "Elemen Schools In 1 School Dist. (Public)": [],
        "Intermediate/Middle Schools (Public)": [],
        "Junior High Schools (Public)": [],
        "K-12 Schools (Public)": [],
        "High Schools (Public)": [],
        "High Schools In 1 School Dist. (Public)": []
    }
    '''
    school_index = {}

    while i < num_rows:
        i += 1
        row = sheet.row(i)
        #status = sheet.cell_value(i,3)
        #date = sheet.cell_value(i,21).encode("ascii", "ignore")
        #SOCType = sheet.cell_value(i,27).encode("ascii", "ignore")
        school = convert_list_to_object(sheet,row)
        school_index[i] = school
        #if len(SOCType) > 0 and ( SOCType in school_index.keys()):
        #    school_object = convert_list_to_object(sheet, row)
        #    school_index[SOCType].append(school_object)

        '''

        if (date != ""):
            y = int(date.split("/")[2])
            if y < year:
                if len(SOCType) > 0 and ( SOCType in school_index.keys() ):
                    school_object = convert_list_to_object(sheet,row)
                    school_index[SOCType].append(school_object)
                    #closed_schools.append(school_object)

        '''
    return school_index

#To-Do: Move to pubschls class
def populate_ca_public_schools_db():
    conn = pymongo.Connection()
    count = 0
    db = conn.ca_public_schools 
    index = get_ca_schools()
    keys = [ k for k in index.keys() ]

    for key in keys:
        school = index[key]
        data = {k:parse_xls_value(school[k]) for k in school.keys()}
        db.test.insert(data)

#To-Do: Move to utils class
#take out the text or empty parts of values
def parse_xls_value(v):
    if str(v).find("text") != -1:
        end_index = len(str(v)) - 1
        return str(v)[7: end_index]
    if str(v).find("empty") != -1:
        return ""

#To-Do: Move to oscuridad class
def collect_test_scores(state, year):
    schools = []

    if state == "ca":
        if year == 2009:
            schools = collect_ca_2009_test_scores()

    return schools

def get_all_ca_2009_test_schools():
    LIM  = 100
    schools = collect_ca_2009_school_objects(LIM)
    converged = False

    while not converged:
        LIM = 2*LIM
        more_schools = collect_ca_2009_school_objects(LIM)
        if ( len(more_schools) <= len(schools) ):
            converged = True
        else:
            schools = more_schools

    return schools

def populate_ca_2009_test(schools):
    conn = pymongo.Connection()
    db = conn.ca_2009_test

    for school in schools:
        db.test.insert(school)

    print "finished"



#Move to ca_2009_test class
def collect_ca_2009_school_objects(LIM):
    count = 0
    f = "ca2009_1_csv_v3.txt"
    tags = [ "County Code","District Code","School Code","Charter Number","Test Year","Subgroup ID","Test Type", "CAPA Assessment Level","Total STAR Enrollment","Total Tested At Entity Level","Total Tested At Subgroup Level","Grade","Test Id","STAR Reported Enrollment/CAPA Eligible","Students Tested","Percent Tested","Mean Scale Score","Percentage Advanced","Percentage Proficient","Percentage At Or Above Proficient","Percentage Basic","Percentage Below Basic","Percentage Far Below Basic","Students with Scores","CMA/STS Average Percent Correct" ]
    schools = []

    with open(f) as infile:
        for line in infile:
            if (count < LIM):
                count += 1
                ts = clean_csv_line(line)
                if len(ts) == len(tags):
                    school = { tags[i]:ts[i] for i in range(len(ts)) }
                    schools.append(school)
                else:
                    print "Line " + count +" in text file is wrong length"
            else:
                break

    return schools

#Move to Utils class
def clean_csv_line(line):
    raw = line.split(",")
    clean = []

    for r in raw:
        if ( r.find('"') != -1 ):
            if len(r) > 2:
                clean.append(r[1:len(r)-1])
            else:
                clean.append('')
        else:
            clean.append(r)

    return clean

#Move to ca2009_test class
def get_ca_district_code(school_dict):
    DC = "District Code"
    try:
        DC = school_dict[DC]
    except KeyError:
        print "School dictionary has no district code key"

    if DC == "District Code":
        return ""
    else:
        return DC

#Move to ca2009_class
def get_ca_school_code(school_dict):
    SC = "School Code"

    try:
        SC = school_dict[SC]
    except KeyError:
        print "School dictionary has no school code key"

    if SC == "School Code":
        return ""
    else:
        return SC

def get_ca_county_code(school_dict):
    CC = "County Code"

    try:
        CC = school_dict[CC]
    except KeyError:
        print "School dictionary has no county code key"

    if CC == "County Code":
        return ""
    else:
        return CC


def collect_ca_2009_entities(LIM):
    count = 0
    f = "ca2009entities_csv.txt"
    tags = [ "County Code","District Code","School Code","Charter Number","Test Year","Type Id","County Name","District Name","School Name","Zip Code" ]
    entities = []

    with open(f) as infile:
        for line in infile:
            if (count < LIM):
                count += 1
                ts = clean_csv_line(line)
                entity = {}
                for i in range(len(tags)):
                    if i < len(ts):
                        entity[tags[i]] = ts[i]
                    else:
                        entity[tags[i]] = ""
                entities.append(entity)
            else:
                break

    return entities

#id consists of a tuple (School Code, District Code, County Code)
def collect_ca_2009_test_scores(LIM):
    school_ids = []
    entity_ids = []
    schools = collect_ca_2009_school_objects(LIM)
    entities = collect_ca_2009_entities(LIM)
    scores = {}
    entity_dict = {}

    for s in schools:
        school_id = (get_ca_school_code(s), get_ca_district_code(s), get_ca_county_code(s))
        if school_id not in school_ids:
            scores[school_id] = [s]
            school_ids.append(s)
        else:
            scores[school_id].append(s)

    for e in entities:
        entity_id = (get_ca_school_code(e), get_ca_district_code(e), get_ca_county_code(e))
        if entity_id not in entity_ids:
            entity_dict[entity_id] = [e]
            entity_ids.append(entity_id)
        else:
            entity_dict[entity_id].append(e)

    return scores, entity_dict

def collate_entities_scores_2009(LIM):
    scores, entities = collect_ca_2009_test_scores(LIM)
    matched = []

    for e in entities.keys():
        if e in scores.keys():
            matched.append( combine_s_e( entities[e][0], scores[e][0] ) )

    return matched


def combine_s_e(s,e):
    combo = {}
    common_keys = [ k for k in e.keys() if (k in s.keys())]
    e_keys = [ k for k in e.keys() if (k not in s.keys())]
    s_keys = [ k for k in s.keys() if (k not in e.keys())]

    for k in common_keys:
        if (len(s[k]) == 0) and (len(e[k]) == 0):
            combo[k] = ""
        else:
            if ( len(s[k]) != 0 ):
                combo[k] = s[k]
            else:
                combo[k] = e[k]

    e_dict = { e_keys[i]:e[e_keys[i]] for i in range(len(e_keys))}
    s_dict = { s_keys[i]:s[s_keys[i]] for i in range(len(s_keys))}
    combo.update(e_dict)
    combo.update(s_dict)

    return combo

def match_pubschls_ca2009(LIM):
    conn = pymongo.Connection()
    db = conn.ca_public_schools
    public_schools = list(db.test.find({}).limit(LIM))
    scores = { s["_id"]:[] for s in public_schools }

    ca2009_schools = collate_entities_scores_2009(LIM)
    while ( len(ca2009_schools) < len(public_schools) ):
        LIM = 2*LIM
        ca2009_schools = collate_entities_scores_2009(LIM)

    for school in public_schools:
        object_id = school["_id"]

        for s in ca2009_schools:
            scored = { "score": score_schools(school,s),
                       "school": s }
            scores[object_id].append(scored)

        scores[object_id] = sorted(scores[object_id],key=itemgetter("score")).reverse()


    return scores

def find_match():
    lim = 100


def drop_object_id(obj):
    if (obj.pop(u'_id', None) != None):
        return obj.pop(u'_id', None)
    else:
        return obj

#school1 is from pubschls.xls and school2 is from ca2009 documents
def score_schools(school1, school2):
    tags_pubschls= [ u'Zip', u'CharterNum', u'School', u'County', u'District']
    tags_ca2009 = ['Zip Code', 'Charter Number', 'School Name', 'County Name', 'District Name']

    score = 0
    lev = Levenshtein

    for i in range(len(tags_ca2009)):
        key1 = school1[tags_pubschls[i]].encode("ascii","ignore")
        key2 = str(school2[tags_ca2009[i]])
        score += lev.ratio(key1, key2)

    return score


def score_zip(zip1, zip2):


    if (zip1 == zip2):
        return 1
    else:
        return 0

def score_charter_num(num1, num2):

    if ( num1 == num2 ):
        return 1
    else:
        return 0


def get_all_test_scores_ca_2009():
    LIM = 10000
    scores = collate_entities_scores_2009(LIM)
    converged = False
    count = 1
    performance = {"start":time.time()}

    while not converged:
        #add in loop performance data
        performance["loop: " + str(count)] =  time.time() 

        LIM = 2*LIM
        more_scores = collate_entities_scores_2009(LIM)
        if math.fabs(len(scores)-len(more_scores))  == 0:
            converged = True
            performance["loop: " + str(count)] = time.time()-performance["loop: " + str(count)]
        else:
            performance["loop: " + str(count)] = time.time()-performance["loop: " + str(count)]
            count += 1
            scores = more_scores

    performance["end"] = time.time()
    performance["total"] =  performance["end"] - performance["start"]
    return performance, scores



'''
def get_ca_test_scores_2009():
    tags = ['CountyCode', 'DistrictCode', 'SchoolCode', 'CharterNumber', 'TestYear', 'SubgroupID', 'TestType', 'CAPAAssessmentLevel', 'TotalSTAREnrollment', 'TotalTestedEntityLevel', 'TotalTestedSubgrpLevel', 'Grade', 'TestID', 'STARRptdEnroll_CAPAEligible', 'StudentsTested', 'PercentTested', 'MeanScaleScore', 'PercentAdvanced', 'PercentProficient', 'PercentBasic', 'PercentBelowBasic', 'PercentFarBelowBasic', 'StudentsWithScores']
    filename = "ca2009_all_xml_v3/ca2009_all_xml_v3.xml"
    schools = parse_schools(filename, tags)

    return schools
    context = iterparse(open("ca2009_all_xml_v3/ca2009_all_xml_v3.xml"),events=("start","end"))
    context = iter(context)
    event, root = context.next()
    keys = ['ResearchData', 'CountyCode', 'DistrictCode', 'SchoolCode', 'CharterNumber', 'TestYear', 'SubgroupID', 'TestType', 'CAPAAssessmentLevel', 'TotalSTAREnrollment', 'TotalTestedEntityLevel', 'TotalTestedSubgrpLevel', 'Grade', 'TestID', 'STARRptdEnroll_CAPAEligible', 'StudentsTested', 'PercentTested', 'MeanScaleScore', 'PercentAdvanced', 'PercentProficient', 'PercentBasic', 'PercentBelowBasic', 'PercentFarBelowBasic', 'StudentsWithScores']
    ks = []
    schools = []
    school = {k:[] for k in keys}
    c = 0

    for event, elem in context:
        #if we haven't associated value to key
        count = len(ks)
        if c < LIM:
            if count < len(keys) and len(school[keys[count]]) == 0:
                if elem.tag not in keys:
                    print "element tag not in known tags"
                else:
                    school[keys[count]] = elem.text
                ks.append(elem.tag)
            else:
                #when we've add date for all keys
                if len(ks) == len(keys):
                    schools.append(school)
                    school = {k:[] for k in keys}
                    ks = []
                    c += 1
        else:
            break

    return schools


def get_ca_entities_2009():
    #context = iterparse(open("ca2009entities_xml.xml"),events=("start","end"))
    #context = iter(context)
    #event, root = context.next()
    tags = ['CountyCode', 'DistrictCode', 'SchoolCode', 'CharterNumber', 'TestYear', 'TypeID', 'CountyName', 'DistrictName', 'SchoolName', 'ZipCode']
    filename = "ca2009_all_xml_v3/ca2009entities_xml.xml"
    entities = parse_schools(filename, tags)

    return entities

#works for ca_2009 entities and test files
def parse_schools(filename, tags):
    context = iterparse(open(filename),events=("start","end"))
    context = iter(context)
    event, root = context.next()
    #keys = ['ResearchData', 'CountyCode', 'DistrictCode', 'SchoolCode', 'CharterNumber', 'TestYear', 'SubgroupID', 'TestType', 'CAPAAssessmentLevel', 'TotalSTAREnrollment', 'TotalTestedEntityLevel', 'TotalTestedSubgrpLevel', 'Grade', 'TestID', 'STARRptdEnroll_CAPAEligible', 'StudentsTested', 'PercentTested', 'MeanScaleScore', 'PercentAdvanced', 'PercentProficient', 'PercentBasic', 'PercentBelowBasic', 'PercentFarBelowBasic', 'StudentsWithScores']
    ts = []
    schools = []
    school = {t:[] for t in tags}
    c = 0
    LIM = 600

    for event, elem in context:
        #if we haven't associated value to key
        count = len(ts)
        if c < LIM:
            if count < len(tags) and len(school[tags[count]]) == 0:
                if elem.tag not in tags:
                    print "element tag not in known tags"
                else:
                    school[tags[count]] = elem.text
                ts.append(elem.tag)
            else:
                #when we've added data for all keys
                if len(ts) == len(tags):
                    schools.append(school)
                    school = {t:[] for t in tags}
                    ts = []
                    c += 1
        else:
            break

    return schools

def match_scores_and_entities():
    tags = ['DistrictCode', 'CountyCode', 'SchoolCode']
    schools = get_ca_test_scores_2009()
    entities = get_ca_entities_2009()
    matches = []

    t1 = time.time()
    for s in schools:
        s_values = [s[t] for t in tags]
        for e in entities:
            e_values = [e[t] for t in tags]
            if s_values == e_values:
                result = s
                for k in e.keys():
                    if k not in result.keys():
                        result[k] = e[k]
                matches.append(result)
    t2 = time.time()
    print t2-t1

    return matches
'''


