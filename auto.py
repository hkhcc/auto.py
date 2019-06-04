#This is the newest version of auto.py and supercedes 1.02 with effective date 11/09/2017.
#!/usr/bin/env python3
import datetime
import math
import os
import platform
import random
import re
import sys
import win32com.client
from numpy import percentile
from scipy import stats
xlApp = win32com.client.Dispatch("Excel.Application")
workbook = xlApp.Workbooks.Open(os.path.expanduser('~') + "/Desktop/result.xls", False, True, None, "Lis*****")
sheet1 = workbook.Sheets(1).UsedRange.Value
try:
    PHI_list = xlApp.Workbooks.Open(os.path.expanduser('~') + "/Desktop/PHI list.xls", False, True, None, "Phi*****")
    sheet2 = PHI_list.Sheets(1).UsedRange.Value
except:
    print('PHI_list.xls NOT loaded!', file=sys.stderr)




def bootstrap_CI(input_list, lower_percentile=5, upper_percentile=95, replicates=100):
    print('#', 'Constructing bootstrap CI for list with length', len(input_list), 'with', replicates, 'replicates', file=sys.stderr)
    lower_bound, upper_bound = [], []
    for i in range(0, replicates):
        replicate = []
        for j in range(0, len(input_list)):
            replicate.append(random.choice(input_list))
        replicate = sorted(replicate)
        lower_bound.append(replicate[int(len(replicate)*lower_percentile/100)])
        upper_bound.append(replicate[int(len(replicate)*upper_percentile/100)])
    return [sum(lower_bound)/len(lower_bound), sum(upper_bound)/len(upper_bound)]

def within_valid_period(test_date, days_valid=185):
    date = datetime.datetime.strptime(str(test_date).split(' ')[0], '%Y-%m-%d')
    time_difference = datetime.datetime.now() - date
    valid_period = datetime.timedelta(days=days_valid)
    if time_difference > valid_period:
        return False
    else:
        return True

def ewma(array, w=0.2, s=None):
    '''Calculate EWMA and return an array'''
    ewma = []
    if s:
        ewma.append(s)
    else:
        ewma.append(array[0])
    for i in range(1,len(array)):
        past = ewma[-1]
        present = array[i]
        ewma.append(past * (1 - w) + present * w)
    return ewma

def swks(array, target_mean, target_sd, k=20):
    '''Calculate the sliding-window KS p-value'''
    swks = []
    if k >= len(array):
        return [0.5] * len(array)
    else:
        swks = [0.5] * (k - 1)
    # calculate the KS p-value
    for i in range(k-1, len(array)):
        d = array[i-k+1:i+1]
        stat = stats.kstest(d, stats.norm.cdf, args=(target_mean, target_sd))
        print(target_mean, target_sd, d, stat[1], file=sys.stderr)
        swks.append(stat[1])
    return swks

def mswks(array, target_mean, target_sd, k=3):
    '''Calculate the scaled swks p-value'''
    result = swks(array, target_mean, target_sd, k)
    result = [x * 4 * target_sd + target_mean - 2 * target_sd for x in result]
    return result


def time_object(time_string):
    time = datetime.datetime.strptime(time_string, '%H:%M:%S').time()
    return time

def to_time(py_time_object):
    time = datetime.datetime.strptime(str(py_time_object).split('+')[0], '%Y-%m-%d %H:%M:%S')
    return time

class SPE_Patient:
    '''A patient undergoing SPE testing'''
    def __init__(self, name, pid, reqno, sex, age, location, total_protein, clinical_details, dob):
        self.name = name
        self.pid = pid
        self.reqno = reqno
        self.sex = sex
        self.age = age
        self.location = location
        if total_protein == None:
            total_protein = -1
        self.total_protein = round(total_protein, 2)
        self.tests = {'SPE': {},
                      'IF': {},
                      'BJP': {}}
        self.diagnosis = clinical_details
        self.dob = dob

    def new_test(self, code, reqno, date, result):
        if code in self.tests:
            self.tests[code][reqno] = {'date': date,
                                       'result': result}
        else:
            raise ValueError(code)

    def organize_results(self, BJP_mode=False):
        text_report = ''
        # SPE & IF results
        if len(self.tests['SPE']) == 1 and BJP_mode == False:  # LIS only contains the current SPE request
            text_report += 'A. SPE: current request only, no previous SPE results in LIS\n'
        else:
            text_report += 'A. SPE & IF results\n'
            sorted_SPE_results = sorted(self.tests['SPE'].items(), key=lambda x: x[1]['date'], reverse=True)
            for SPE_result in sorted_SPE_results:
                key = SPE_result[0]  # the request number
                text_report += key + '\t' + str(self.tests['SPE'][key]['date']).replace('+00:00', '') + '\t' + str(self.tests['SPE'][key]['result']) + '\n'
                if key in self.tests['IF']:
                    text_report += key + '(IF)' + '\t' + str(self.tests['IF'][key]['date']).replace('+00:00', '') + '\t' + str(self.tests['IF'][key]['result']) + '\n'
        # BJP results
        if len(self.tests['BJP']) == 0:
            text_report += 'B. BJP: no previous BJP results in LIS\n'
        else:
            text_report += 'B. BJP results\n'
            sorted_BJP_results = sorted(self.tests['BJP'].items(),
                                        key=lambda x: x[1]['date'],
                                        reverse=True)
            for BJP_result in sorted_BJP_results:
                key = BJP_result[0]  # the request number
                text_report += key + '\t' \
                    + str(self.tests['BJP'][key]['date']).replace('+00:00', '') \
                    + '\t' + str(self.tests['BJP'][key]['result']) + '\n'

        return text_report

'''class SPE_Patient_interlab:
    A patient undergoing SPE testing
    def __init__(self, reqno, total_protein, location, name, sex, dob, pid):
        self.reqno = reqno
        if total_protein == None:
            total_protein = -1
        self.total_protein = round(total_protein, 2)
        self.location = location
        self.name = name
        self.sex = sex
        self.dob = dob
        self.pid = pid
        self.tests = {'SPE': {},
                      'IF': {},
                      'BJP': {}}

    def new_test(self, code, reqno, date, result):
        if code in self.tests:
            self.tests[code][reqno] = {'date': date,
                                       'result': result}
        else:
            raise ValueError(code)'''
    
class ADNA_Patient:
    '''A patient undergoing anti-dsDNA testing'''
    def __init__(self, name, pid, reqno, past_adna, age):
        self.name = name
        self.pid = pid
        self.reqno = reqno
        self.past_adna = past_adna
        self.age = age
        self.highest_titre = 0
        self.highest_titre_is_recent = False
        self.negative_ana_reqno = ''
        self.pending_pattern_reqno = []
        self.pending_titre_reqno = []
        self.ana_reqno = []

    def new_titre(self, titre, date):
        try:
            dilution = int(titre.split(':')[1])
        except:
            if titre == 'Quantity Insufficient':
                print('# QI in titre field.', file=sys.stderr)
                return None
            else:
                raise ValueError('Unexpected titre:', titre)
        if dilution < 40 or dilution > 2560:
            raise ValueError(titre)
        if dilution >= self.highest_titre:
            self.highest_titre = dilution
            if within_valid_period(date):
                self.highest_titre_is_recent = True
                print('# highest titre is recent', file=sys.stderr)
            print('# updating titre for', self.name, file=sys.stderr)

    def new_pattern(self, pattern, date, reqno):
        if pattern == 'Negative':
            if within_valid_period(date):
                if self.negative_ana_reqno == '':
                    self.negative_ana_reqno = reqno
                else:
                    self.negative_ana_reqno += ';' + reqno
        else:
            raise ValueError(pattern)

    def new_pending_pattern_reqno(self, reqno):
        self.pending_pattern_reqno.append(reqno)

    def new_pending_titre_reqno(self, reqno):
        self.pending_titre_reqno.append(reqno)

    def new_ana_reqno(self, date, reqno):
        if within_valid_period(date):
            if reqno not in self.ana_reqno:
                self.ana_reqno.append(reqno)


class TFT_Patient:
    '''A patient undergoing TFT (TSH, FT4, FT3)'''
    def __init__(self, reqno, pid):
        self.reqno = reqno
        self.pid = pid
        self.TSH = None
        self.FT4 = None
        self.FT3 = None
        self.ref_interval = {'TSH': [None, None],
                             'FT4': [None, None],
                             'FT3': [None, None]
                             }
        self.flag = None
        self.result_status = []

    def new_result(self, test_code, value, test_ref_interval, result_status):
        key = None
        if test_code == 4273 or test_code == 6312:  # TSH or TSH-B
            self.TSH = value
            key = 'TSH'
        elif test_code == 4458 or test_code == 6313:  # FT4 or FT4-B
            self.FT4 = value
            key = 'FT4'
        elif test_code == 5025 or test_code == 6314:  # FT3 or FT3-B
            self.FT3 = value
            key = 'FT3'
        self.ref_interval[key][0] = float(test_ref_interval[0])
        self.ref_interval[key][1] = float(test_ref_interval[1])
        self.result_status.append(int(result_status))

    def interpret(self):
        if self.TSH and self.FT4:
            if self.TSH >= self.ref_interval['TSH'][1] and self.FT4 >= self.ref_interval['FT4'][1]:
                self.flag = 'High TSH / High FT4'
            elif self.TSH <= self.ref_interval['TSH'][0] and self.FT4 <= self.ref_interval['FT4'][0]:
                self.flag = 'Low TSH / Low FT4'
            elif self.TSH >= self.ref_interval['TSH'][0] and self.FT4 >= self.ref_interval['FT4'][1]:
                self.flag = 'Inapp. norm. TSH'

class MPRL_Patient:
    '''A patient undergoing macroprolactin testing'''
    def __init__(self, reqno, pid, new_mprl_count, old_mprl_count):
        self.reqno = reqno
        self.pid = pid
        self.new_mprl_count = int(new_mprl_count)
        self.old_mprl_count = int(old_mprl_count)
        self.serial_mprl = []
        self.decision = None

    def new_mprl_result(self, reqno, date, recovery):
        '''A simple method to append test results'''
        if recovery:
            try:
                recovery = int(recovery)
            except:
                print('# Rejecting:', 'reqno', reqno, 'date', date, 'recovery',
                      recovery)
                return
            entry = {
                'reqno': reqno,
                'date': date,
                'recovery': int(recovery)
            }
            self.serial_mprl.append(entry)

    def decide(self):
        '''Decide whether to test for macroprolactin again'''
        # test if no previous testing done
        if self.new_mprl_count + self.old_mprl_count == 0:
            self.decision = 'Proceed (no previous MPRL within 1 year)'
            return

        # test if recovery all along below 70%
        # or falling below 70%
        true_prl = []
        other_prl = []
        for i in range(0, len(self.serial_mprl)):
            if self.serial_mprl[i]['recovery'] >= 70:
                true_prl.append(i)
            else:
                other_prl.append(i)
        # now compare the indices of the last measurements
        if len(true_prl) > 0 and len(other_prl) > 0:
            if true_prl[-1] > other_prl[-1]:
                self.decision = 'Cancel (latest recovery >= 70%)'
            else:
                self.decision = 'Proceed (falling recovery %)'

        elif len(true_prl) > 0 and len(other_prl) == 0:
            self.decision = 'Cancel (single recovery >= 70%)'

        elif len(true_prl) == 0:
            self.decision = 'Proceed (previous recovery values < 70%)'
        else:
            raise ValueError

class PHI:
    '''A patient undergoing PHI testing'''
    def __init__(self, reqno, pid, pname, phi_past, psa):
        self.reqno = reqno
        self.pid = pid
        self.pname = pname
        self.phi_past = int(phi_past)
        self.psa= psa

class QC:
    '''A QC object'''

    def __init__(self, machine, test_code, qc_number, lower_limit, upper_limit, target_mean, target_sd):
        self.name = self.get_name(machine, test_code, qc_number)
        self.machine = machine
        self.test_code = test_code
        self.qc_number = qc_number
        self.lower_limit = lower_limit
        self.upper_limit = upper_limit
        self.target_mean = target_mean
        self.target_sd = target_sd
        self.readings = []

    def new_reading(self, value, date):
        self.readings.append([datetime.datetime.strptime(str(date).split('+')[0], '%Y-%m-%d %H:%M:%S'), value])

    def get_name(self, machine, test_code, qc_number):
        return '-'.join([machine, test_code, str(int(qc_number))])


# requires format "Excel with headers"
tag = None

# select / detect output type

if len(sys.argv) > 1:
    tag = sys.argv[1]
else:
    if sheet1[0][0] == '_TFT_Collected_Date':
        tag = 'TFT'
    elif sheet1[0][0] == 'TFT':
        tag = 'TFT2'
    elif sheet1[0][0] == 'TFT3':
        tag = 'TFT3'
    elif sheet1[0][0].startswith('DNA_Name'):
        tag = 'DNA1'
    elif sheet1[0][0].startswith('_SPE_'):
        tag = 'SPE'
    elif sheet1[0][0].startswith('_BJP_'):
        tag = 'BJP'
    elif sheet1[0][1].startswith('MPRL_result'):
        tag = 'MPRL'
    elif sheet1[0][0].startswith('machine'):
        tag = 'QC'
    elif sheet1[0][0].startswith('testrslt_reqno'):
        tag = 'GEN'
    elif sheet1[0][1].startswith('month'):
        tag = 'TAT'
    elif sheet1[0][0] == 'Request_number' and sheet1[0][5] == 'Flag':
        tag = 'xTFT'
    elif sheet1[1][0] == 'T3Tox':
        tag = 'T3Tox'
    elif sheet1[0][0] == 'Request_Number_PHI':
        tag = 'PHI'
    else:
        raise ValueError('Tag not valid!')

print('#', tag, "input was identified", file=sys.stderr)

# initialize the patient list
patients = []
this_patient = None

if tag == 'PHI':
    wl = os.stat(os.path.expanduser('~') + "/Desktop/PHI list.xls")
    wltime = datetime.datetime.fromtimestamp(wl.st_mtime)
    print("The last modified date and time of the PHI List:", wltime)
    id_list = []
    Cancel = 0
    Proceed = 0
    decision = ''
    total = 0
    for r in sheet2[1:]:
        if r[2] != None:
            id_list.append(r[2])    
    for row in sheet1[1:]:
        this_patient = PHI(row[0], row[1], row[2], row[3], row[4])
        patients.append(this_patient)
        if this_patient.pid in id_list:
            decision = 'Proceed. The patient is in PWH study. Please check if tPSA is done.'
            Proceed += 1
        elif this_patient.phi_past > 0:
            decision = 'Cancel. '+ str(this_patient.phi_past) + ' test(s) have been done in previous year.'
            Cancel += 1
        elif this_patient.psa == None:
            decision = 'T/F. tPSA not done.'
        elif float(this_patient.psa) > 20 or float(this_patient.psa) < 2:
            decision = 'Cancel. The tPSA result is ' + str(this_patient.psa)
            Cancel += 1
        else:
            decision = 'Proceed. tPSA = ' + str(this_patient.psa)
            Proceed += 1
        print (this_patient.reqno + ' ' + this_patient.pname + ': ' + decision + '\n')
        total +=1
    print ('Total number of samples:', total)
    print ('Total number of proceeds:', Proceed)
    print ('Total number of cancels:', Cancel)
    print ('Total number of T/F:', total - Proceed - Cancel)
    print ('Total number of patients in PWH study:', len(id_list))
            
if tag == 'TFT':
    output_cache, ft4_list, tsh_list = [], [], []
    for row in sheet1[1:]:
        lab_no, ft4, ft4_status, tsh, tsh_status = row[2], row[3], row[5], row[6], row[8]
        if ft4 == None and tsh == None:
            continue
        output_cache.append([lab_no, ft4, ft4_status, tsh, tsh_status])
        if ft4 != None:
            ft4_list.append(ft4)
        if tsh != None:
            tsh_list.append(tsh)
    ft4_range = [9.5, 18.1]
    tsh_range = [0.35, 3.80]
    print('#', len(output_cache), 'lines ready for filtering...')
    for row in output_cache:
        if row[1] and row[1] > ft4_range[1] and row[3] and row[3] > tsh_range[1]:
            print('\t'.join(str(x)for x in row))
        if row[1] and row[1] > ft4_range[1] and row[3] and row[3] > tsh_range[0] and row[3] <= tsh_range[1]:
            print('\t'.join(str(x) for x in row), 'inappropriately normal TSH', sep='\t')

elif tag == 'TFT2':
    '''
    SELECT
    "TFT" "TFT",
    * FROM testrslt
    WHERE
    testrslt_status_date >= DATEADD(dd, -7, GETDATE()) AND
    (testrslt_member_ckey = 4273 OR
    testrslt_member_ckey = 4458 OR
    testrslt_member_ckey = 5025)
    ORDER BY testrslt_reqno
    '''
    for row in sheet1[1:]:
        if this_patient == None or this_patient.reqno != row[1]:
            this_patient = TFT_Patient(row[1], row[37])
            patients.append(this_patient)
        this_patient.new_result(row[3], row[10], [row[13], row[14]], row[16])
    print('# Finished adding', len(patients), 'patients...', file=sys.stderr)

    patients = sorted(patients, key= lambda x: x.reqno)

    for p in patients:
        p.interpret()
        if p.flag != None and p.flag != 'Normal':
            print(p.reqno, p.flag, p.TSH, p.FT4, p.result_status)

elif tag == 'TFT3':
    '''
    SELECT
    "TFT3" "TFT3",
    * FROM testrslt
    WHERE
    testrslt_status_date >= DATEADD(dd, -7, GETDATE()) AND
    (testrslt_member_ckey = 6312 OR
    testrslt_member_ckey = 6313 OR
    testrslt_member_ckey = 6314 OR
    testrslt_member_ckey = 4273 OR
    testrslt_member_ckey = 4458 OR
    testrslt_member_ckey = 5025
    )
    ORDER BY testrslt_reqno
    '''
    for row in sheet1[1:]:
        if this_patient == None or this_patient.reqno != row[1]:
            this_patient = TFT_Patient(row[1], row[37])
            patients.append(this_patient)
        this_patient.new_result(row[3], row[10], [row[13], row[14]], row[16])
    print('# Finished adding', len(patients), 'patients...', file=sys.stderr)

    patients = sorted(patients, key= lambda x: x.reqno)

    for p in patients:
        p.interpret()
        if p.flag != None and p.flag != 'Normal':
            print(p.reqno, p.flag, p.TSH, p.FT4, p.result_status)

elif tag == 'xTFT':
    '''
    Generate the SQL query for selecting TFT results from  all possible T3 toxicosis cases
    '''
    t3_toxic_reqno = []
    for row in sheet1[1:]:
        if row[5] == '? T3 toxicosis (consider add FT3-B)':
            t3_toxic_reqno.append(row[0])
    pid_list = "SELECT testrslt_pid_group FROM testrslt WHERE testrslt_reqno = '" + "' OR testrslt_reqno = '".join(t3_toxic_reqno) + "'"
    sql = 'SELECT "T3Tox", * FROM testrslt tr2 WHERE tr2.testrslt_pid_group IN \n(' + pid_list + ') AND (testrslt_member_ckey = 6312 OR testrslt_member_ckey = 6313 OR testrslt_member_ckey = 6314 OR testrslt_member_ckey = 4273 OR testrslt_member_ckey = 4458 OR testrslt_member_ckey = 5025)'
    sql += ' ORDER BY tr2.testrslt_add_date'
    print('# Please use the following SQL query to retrieve the historical TFT results:')
    print(sql.replace('OR ', 'OR \n'))

elif tag == 'T3Tox':
    '''
    Use heuristics to screen for genuine T3 toxicosis cases
    '''
    import matplotlib.pyplot as plt
    pid = {}
    for row in sheet1[1:]:
        # correct the LIS data structure
        row = list(row)
        if row[50] == None:
            row[50] = 'FT3'
        if not row[37] in pid:
            pid[row[37]] = {'TSH-B':[], 'FT4-B':[], 'FT3-B':[], 'TSH':[], 'FT4':[], 'FT3':[]}
        pid[row[37]][row[50]].append([row[27], row[10], row[1]])
    for p in pid:
        print('#####')
        # added try-except to prevent error on empty or invalid results - CCHO 20180615
        try:
            for ft4 in pid[p]['FT4']:
                print(to_time(ft4[0]), ft4[1], ft4[2])
            for ft4 in pid[p]['FT4-B']:
                print(to_time(ft4[0]), ft4[1], ft4[2])
            y = [x[1] for x in pid[p]['FT4']] + [x[1] for x in pid[p]['FT4-B']]
            x = [to_time(x[0]) for x in pid[p]['FT4']] + [to_time(x[0]) for x in pid[p]['FT4-B']]
            plt.scatter(x=x, y=y)
        except:
            pass
        plt.show()

elif tag == 'DNA1':
    '''
    SELECT * FROM 
	(
		SELECT
			(SELECT MAX(req_name) FROM request WHERE req_pid_group = adna_table.testrslt_pid_group) "DNA_Name",
			adna_table.testrslt_reqno "REQNO",
			ana_table.testrslt_result "ANA_result",
			CASE WHEN ana_table.testrslt_member_ckey = 5152 THEN "pattern" ELSE "titre" END "Type",
			ana_table.testrslt_reqno "ANA_REQNO",
			ana_table.testrslt_status "ANA_status",
			ana_table.testrslt_status_date "ANA_date",
			ana_table.testrslt_pid_group "pid_group",
			(SELECT COUNT(*) FROM testrslt 
			WHERE 
			(testrslt_member_ckey = 5952 OR testrslt_member_ckey =  6334)
			AND 
			testrslt_pid_group = adna_table.testrslt_pid_group
			AND
			testrslt_status = 6
			) "Prev_ADNA_count",
			(SELECT req_age FROM request WHERE req_reqno = adna_table.testrslt_reqno) "Age"
		FROM
			(SELECT * FROM testrslt 
			WHERE testrslt_reqno IN (SELECT req_reqno FROM request_detail WHERE req_reqno IN (select req_reqno from request where req_complete <= 1) AND req_alpha_code = 'DNA2')
			AND testrslt_member_ckey = 6334 AND testrslt_status <= 1) adna_table
		LEFT JOIN
			(SELECT testrslt_result, testrslt_member_ckey, testrslt_status, testrslt_status_date, testrslt_reqno, testrslt_pid_group FROM testrslt WHERE testrslt_member_ckey = 5152 or testrslt_member_ckey = 5160) ana_table
		ON
			adna_table.testrslt_pid_group = ana_table.testrslt_pid_group
	) pre_filter ORDER BY pid_group, REQNO, ANA_REQNO
    '''
    for row in sheet1[1:]:
        if this_patient == None or this_patient.reqno != row[1]: # new request
            this_patient = ADNA_Patient(row[0], row[7], row[1], row[8], row[9])
            patients.append(this_patient)
        if row[3] == 'titre' and (row[5] == 5 or row[5] == 6): # changed to >= 5 on 2017-09-05 to allow pre-authorized result screening
            this_patient.new_titre(row[2], row[6])
        if row[3] == 'titre' and row[5] == 0:
            this_patient.new_pending_titre_reqno(row[4])
        if row[3] == 'pattern' and row[5] == 0:
            this_patient.new_pending_pattern_reqno(row[4])
        if row[3] == 'pattern' and row[2] == 'Negative':
            this_patient.new_pattern('Negative', row[6], row[4])
        if row[3] == 'pattern':
            this_patient.new_ana_reqno(row[6], row[4])

    print('# Total', len(patients), 'patients added.', file=sys.stderr)
    patients = sorted(patients, key=lambda x: x.reqno)
    print('# Total', len(patients), 'patients sorted.', file=sys.stderr)

    patient_counter = 1

    for patient in patients:
        decision = None
        # add ANA rules
        if patient.ana_reqno == []:
            decision = 'No ANA result within past 6 months. Add ANA.'
        # cancel rules
        if patient.negative_ana_reqno != '':
            decision = 'Cancel (Negative ANA within 6 months: ' + patient.negative_ana_reqno + ')'
        if patient.age >= 40 and patient.highest_titre == 80:  # new case >= 40 y.o. ANA 1:160 or higher
            if patient.highest_titre_is_recent:
                decision = 'Cancel R3 (1:' + str(patient.highest_titre) + '), age ' + str(int(patient.age))
            else:
                decision = 'Past ANA result present, but no negative ANA result within past 6 months. Add ANA.'
        # test rules
        if patient.past_adna >= 3: # old case for serial monitoring
            decision = 'Proceed R1 (Anti-dsDNA tests thus far = ' + str(int(patient.past_adna)) + ')'
        if patient.past_adna == 1 or patient.past_adna == 2:
            decision = 'Proceed R1# (Anti-dsDNA tests thus far = ' + str(int(patient.past_adna)) + ')'
        if patient.age < 40 and patient.highest_titre >= 80: # new case < 40 y.o. ANA 1:80 or higher
            decision = 'Proceed R2 (1:' + str(patient.highest_titre) + '), age ' + str(int(patient.age))
        if patient.age >= 40 and patient.highest_titre >= 160: # new case >= 40 y.o. ANA 1:160 or higher
            decision = 'Proceed R3 (1:' + str(patient.highest_titre) + '), age ' + str(int(patient.age))
        # wait rules
        if set(patient.pending_pattern_reqno) & set(patient.pending_titre_reqno) != set():
            decision = 'ANA T/F ' + ';'.join(list(set(patient.pending_pattern_reqno) & set(patient.pending_titre_reqno)))
        print('(' + str(patient_counter) + ') ' + patient.name, patient.reqno, decision, sep='\t')
        patient_counter += 1
    print('Total', str(patient_counter - 1), 'patients.')
    print('Please check this ECSearch-derived list carefully against actual LIS job sheets.')
    print('# For patients marked with R1#, please double-check previous dsDNA tests are appropriate!')

elif tag == 'SPE' or tag == 'BJP':
    '''
    # SQL query for SPE batch
    SELECT
    (SELECT req_name FROM request WHERE req_reqno = rd.req_reqno) "_SPE_Name",
    (SELECT req_pid FROM request WHERE req_reqno = rd.req_reqno) "ID",
    rd.req_reqno,
    rd.req_alpha_code,
    rd.req_registered_date,
    CASE
        WHEN rd.req_alpha_code = "IF" THEN
        (SELECT testrslt_varchar FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 4907)
        WHEN rd.req_alpha_code = "BJP" THEN
        (SELECT testrslt_varchar FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 5559)
        ELSE
        (SELECT req_comment FROM request WHERE req_reqno = rd.req_reqno)
    END "Comment",
    (SELECT req_sex FROM request WHERE req_reqno = rd.req_reqno) "Sex",
    (SELECT req_age FROM request WHERE req_reqno = rd.req_reqno) "Age",
    (SELECT req_locn_hosp FROM request WHERE req_reqno = rd.req_reqno) "Hosp",
    (SELECT office_alpha FROM office WHERE office_ckey = (SELECT req_unit FROM request WHERE req_reqno = rd.req_reqno) AND office_hosp_code = (SELECT req_locn_hosp FROM request WHERE req_reqno = rd.req_reqno)) "Unit",
    (SELECT office_alpha FROM office WHERE office_ckey = (SELECT req_locn FROM request WHERE req_reqno = rd.req_reqno) AND office_hosp_code = (SELECT req_locn_hosp FROM request WHERE req_reqno = rd.req_reqno)) "Location",
    (SELECT testrslt_numeric FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 4172) "TP",
    (SELECT req_cdetail FROM request WHERE req_reqno = rd.req_reqno) "Clinical_detail",
    (SELECT COUNT(*) FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 4010) "Cancelled",
    (SELECT DISTINCT batch_code FROM batch_order WHERE rd.req_reqno = batch_order.reqno AND batch_order.worksheet = "SPE") "Batch"

    FROM request_detail rd
    WHERE (req_alpha_code = "BJP" OR req_alpha_code = "SPE" OR req_alpha_code = "IF")
    AND req_reqno IN (
    SELECT req_reqno FROM
    request
    WHERE req_pid_group IN (SELECT req_pid_group FROM request WHERE
    req_reqno IN (SELECT DISTINCT reqno FROM batch_order WHERE batch_code = '150' AND worksheet = 'SPE')))
    ORDER BY ID, req_alpha_code DESC, Batch DESC

    # SQL query for BJP batch
    SELECT
    (SELECT req_name FROM request WHERE req_reqno = rd.req_reqno) "_BJP_Name",
    (SELECT req_pid FROM request WHERE req_reqno = rd.req_reqno) "ID",
    rd.req_reqno,
    rd.req_alpha_code,
    rd.req_registered_date,
    CASE
        WHEN rd.req_alpha_code = "IF" THEN
        (SELECT testrslt_varchar FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 4907)
        WHEN rd.req_alpha_code = "BJP" THEN
        (SELECT testrslt_varchar FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 5559)
        ELSE
        (SELECT req_comment FROM request WHERE req_reqno = rd.req_reqno)
    END "Comment",
    (SELECT req_sex FROM request WHERE req_reqno = rd.req_reqno) "Sex",
    (SELECT req_age FROM request WHERE req_reqno = rd.req_reqno) "Age",
    (SELECT req_locn_hosp FROM request WHERE req_reqno = rd.req_reqno) "Hosp",
    (SELECT office_alpha FROM office WHERE office_ckey = (SELECT req_unit FROM request WHERE req_reqno = rd.req_reqno) AND office_hosp_code = (SELECT req_locn_hosp FROM request WHERE req_reqno = rd.req_reqno)) "Unit",
    (SELECT office_alpha FROM office WHERE office_ckey = (SELECT req_locn FROM request WHERE req_reqno = rd.req_reqno) AND office_hosp_code = (SELECT req_locn_hosp FROM request WHERE req_reqno = rd.req_reqno)) "Location",
    (SELECT testrslt_numeric FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 4172) "TP",
    (SELECT req_cdetail FROM request WHERE req_reqno = rd.req_reqno) "Clinical_detail",
    (SELECT COUNT(*) FROM testrslt WHERE testrslt_reqno = rd.req_reqno AND testrslt_member_ckey = 4010) "Cancelled"
    FROM request_detail rd
    WHERE (req_alpha_code = "BJP" OR req_alpha_code = "SPE" OR req_alpha_code = "IF")
    AND req_reqno IN (
    SELECT req_reqno FROM
    request
    WHERE req_pid_group IN (SELECT req_pid_group FROM request WHERE
    req_reqno IN (SELECT DISTINCT reqno  FROM batch_order WHERE batch_code = '121' AND worksheet = 'BJP')))
    ORDER BY ID, req_alpha_code, req_reqno DESC
    '''
    for row in sheet1[1:]:
        if row[13] == 1:  # skip cancelled requests
            continue
        if row[14] is not None and int(row[14]) == 999: # do not add patient to batch unless batch no. is valid
            print('#', row[0], row[1], row[2], 'was not added due to invalid batch number!', file=sys.stderr)
            continue
        if this_patient == None or this_patient.pid != row[1]: # "new" i.e. different SPE patient
            this_patient = SPE_Patient(row[0], row[1], row[2], row[6], row[7], '/'.join([row[8], row[9], row[10]]), row[11], row[12], row[15])
            patients.append(this_patient)
        this_patient.new_test(row[3], row[2], row[4], row[5])
    print('#', len(patients), 'patients added...', file=sys.stderr)
    patients = sorted(patients, key=lambda x: x.reqno)
    print('# Now sorted by reqno', [p.reqno for p in patients], file=sys.stderr)

    print('# Now printing batch summary...', file=sys.stderr)
    if tag == 'SPE': # print the SPE summary for input into Sebia system
        position_count = 1
        for p in patients:
            print("", re.split("[A-Z]", p.reqno)[1], str(p.name).replace(",", ", "), p.sex, str(int(p.age)) + " year" , p.location, p.total_protein, sep='\t')
            position_count += 1

    print("@") # the terminator character for VBScript loader.vbs
    print("---For InterLab SPE, please copy the following statements, including the first blank line, to new notepad and save as 'SPE_import.asc'.---")
    print()
    for p in patients:
            print(str('''""'''+ re.split("[A-Z]", p.reqno)[1]), p.total_protein, p.location, str('"'+ str(p.name).replace(",", ", ")+ '"'), p.sex, str('"'+ p.dob+ '"'), p.pid, '''"\x03"''', sep=',')
    '''for p in patients:
            print('"', re.split("[A-Z]", p.reqno)[1], p.total_protein, p.location, str(p.name).replace(",", ", "), p.sex, p.dob, p.pid, '"\x03"', sep='","')'''
    print('InterLab SPE ends. do not copy this sentence.\n')
    print('# Now printing past SPE/BJP/IF results...', file=sys.stderr)
    patient_counter = 1
    p_names = []
    for p in patients:
        if p.name not in p_names:
            p_names.append(p.name)
        else:
            print('** Warning: duplicate patient name:', p.name, file=sys.stderr)
            p.name = '*' + p.name
        print('[' + str(patient_counter) + ']', p.name, p.pid, p.reqno, sep='\t')
        print('Clinical details:\t' + str(p.diagnosis).replace('\r\n', '\r\n\t'))
        if tag == 'BJP':
            print(p.organize_results(BJP_mode=True))
        else:
            print(p.organize_results())
        patient_counter += 1

elif tag == 'MPRL':
    '''
    # SQL query for pending MPRL batch
    SELECT
        mprl_table.testrslt_reqno "REQNO",
        CASE
            WHEN mprl_table.testrslt_reqno = old_mprl_table.testrslt_reqno THEN "<current test>"
            ELSE old_mprl_table.testrslt_result
        END "MPRL_result",
        old_mprl_table.testrslt_reqno "MPRL_REQNO",
        old_mprl_table.testrslt_status "MPRL_status",
        old_mprl_table.testrslt_status_date "MPRL_date",
        old_mprl_table.testrslt_pid_group "pid_group",
        (SELECT COUNT(*) FROM testrslt
        WHERE
            testrslt_member_ckey = 6120
            AND
            testrslt_pid_group = mprl_table.testrslt_pid_group
            AND
            testrslt_status = 6
        ) "MPRL_count_(new)",
        (SELECT COUNT(*) FROM testrslt
        WHERE
            testrslt_member_ckey = 6007
            AND
            testrslt_pid_group = mprl_table.testrslt_pid_group
            AND
            testrslt_status = 6
        ) "MPRL_count_(old)"

    FROM
        (SELECT * FROM testrslt
        WHERE testrslt_reqno IN (SELECT req_reqno FROM request_detail WHERE req_reqno IN (select req_reqno from request where req_complete <= 1) AND req_alpha_code = 'MPRL')
        AND testrslt_member_ckey = 6120 AND testrslt_status = 0) mprl_table

    LEFT JOIN
        (SELECT testrslt_result, testrslt_status, testrslt_status_date, testrslt_reqno, testrslt_pid_group FROM testrslt WHERE testrslt_member_ckey = 6007 or testrslt_member_ckey = 6120) old_mprl_table

    ON    mprl_table.testrslt_pid_group = old_mprl_table.testrslt_pid_group

    ORDER BY REQNO, MPRL_date
    '''

    for row in sheet1[1:]:
        if this_patient == None or this_patient.pid != row[5]: # new patient
            this_patient = MPRL_Patient(row[0], row[5], row[6], row[7])
            patients.append(this_patient)
        if row[1] != '<current test>':
            this_patient.new_mprl_result(row[0], row[4], row[1])
    print('# Total', len(patients), ' patients added! Now sorting by request number...')
    patients = sorted(patients, key=lambda x: x.reqno)
    for p in patients:
        p.decide()
        print(p.reqno, p.decision)

elif tag == 'QC':
    '''
    SELECT
    (SELECT analyser_name FROM analyser a WHERE non_null.analyser_no = a.analyser_no) "machine",
    non_null.test_alpha_code "test_code",
    non_null.test_no "test_number",
    non_null.qc_type_no "qc_number",
    non_null.qc_value "qc_value",
    CASE
        WHEN non_null.qc_value = non_null.target_mean THEN "Passed (target SD = 0)"
        WHEN non_null.qc_value < non_null.target_mean AND non_null.target_SD = 0 THEN "Failed: - (target SD = 0)"
        WHEN non_null.qc_value > non_null.target_mean AND non_null.target_SD = 0 THEN "Failed: + (target SD = 0)"
        WHEN non_null.qc_value < non_null.lower_limit THEN "*Failed: - " + CONVERT(varchar(10), CAST(ABS((non_null.qc_value - non_null.target_mean) / non_null.target_SD) AS decimal(10,2))) + " target SD"
        WHEN non_null.qc_value > non_null.upper_limit THEN "*Failed: + " + CONVERT(varchar(10), CAST(ABS((non_null.qc_value - non_null.target_mean) / non_null.target_SD) AS decimal(10,2))) + " target SD"
        ELSE "Passed"
    END "result",
    non_null.lower_limit "lower_limit",
    non_null.upper_limit "upper_limit",
    non_null.target_mean "target_mean",
    non_null.target_SD "target_SD",
    non_null.create_datetime "date",
    non_null.deleted "deleted"

    FROM

    (
        SELECT * FROM
        (
        SELECT
            CASE
                WHEN QC_range.u_target_mean <> NULL THEN QC_range.u_target_mean
                WHEN QC_range.target_mean <> NULL THEN QC_range.target_mean
            END "target_mean",
            CASE
                WHEN QC_range.u_upper_limit <> NULL THEN QC_range.u_upper_limit
                WHEN QC_range.upper_limit <> NULL THEN QC_range.upper_limit
            END "upper_limit",
            CASE
                WHEN QC_range.u_lower_limit <> NULL THEN QC_range.u_lower_limit
                WHEN QC_range.lower_limit <> NULL THEN QC_range.lower_limit
            END "lower_limit",
            CASE
                WHEN QC_range.u_target_SD <> NULL THEN QC_range.u_target_SD
                WHEN QC_range.target_SD <> NULL THEN QC_range.target_SD
            END "target_SD",
            qc_data.*
        FROM QC_range
        RIGHT JOIN
            (
            SELECT * FROM test_dict
            RIGHT JOIN
            QC
            ON
            test_dict.test_ckey = QC.test_no
            WHERE QC.create_datetime > DATEADD(dd, -31, GETDATE())
            AND QC.create_datetime <= GETDATE()
            ) qc_data
        ON QC_range.qc_type_no = qc_data.qc_type_no AND QC_range.test_no = qc_data.test_no
        ) raw_data WHERE
        raw_data.target_mean <> NULL
    ) non_null
    ORDER by test_alpha_code, qc_number, date

    '''
    QCs = []
    # filter QC values by SD from mean (if provided)
    sigma_filter = None
    try:
        sigma_filter = float(sys.argv[1])
    except:
        print('# No QC filter in place', file=sys.stderr)

    for row in sheet1[1:]:
        if row[11] == 'Y': # neglect if deleted == 'Y'
            print('#', row[0], row[1], row[3], row[4], 'target', row[7], 'rejected!', file=sys.stderr)
            continue
        this_QC = QC(row[0], row[1], row[3], row[6], row[7], row[8], row[9])
        if len(QCs) == 0 or this_QC.name != QCs[-1].name:
            QCs.append(this_QC)
        else:
            this_QC = QCs[-1]
        if sigma_filter:
            if abs(row[4] - row[8]) > sigma_filter * row[9]:
                print('**', this_QC.name, row[4], row[8], row[9], row[10])
                continue
        this_QC.new_reading(row[4], row[10])

    import matplotlib.pyplot as plt
    # determine the number of rows and columns
    figures_per_page = 80
    ncols = int(figures_per_page ** 0.5)
    nrows = math.ceil(figures_per_page / ncols)


    QC_batches = [QCs[i:i+figures_per_page] for i in range(0, len(QCs), figures_per_page)]

    if len(sys.argv) < 2:
        sys.argv.append('show')

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    batch_count = 1
    for QC_batch in QC_batches:
        plot_count = 1
        for q in QC_batch:
            plt.subplot(ncols, nrows, plot_count)
            X, y = [x[0] for x in q.readings], [y[1] for y in q.readings]
            plt.plot(X, y, '.b-')
            # the EWMA plot
            plt.plot(X, ewma(y, s=q.target_mean), '+r-')
            plt.plot(X, mswks(y, q.target_mean, q.target_sd), '+g-')
            plt.axhline(y=q.target_mean, color='black', linestyle='--')
            # the target lines
            plt.axhline(y=q.target_mean + 1*q.target_sd, color='#00FF00', linestyle='-')
            plt.axhline(y=q.target_mean - 1*q.target_sd, color='#00FF00', linestyle='-')
            plt.axhline(y=q.target_mean + 2*q.target_sd, color='r', linestyle='-')
            plt.axhline(y=q.target_mean - 2*q.target_sd, color='r', linestyle='-')
            plt.title(q.machine + '\n' + q.name.split('-')[1] + ' ' + q.name.split('-')[2] , fontsize=6)
            # minX, maxX = min([x[0] for x in q.readings]), max([x[0] for x in q.readings])
            # plt.xlim(minX - datetime.timedelta(days=1), maxX + datetime.timedelta(days=1))
            plt.axis('off')
            print('\t# Plotting chart', plot_count, file=sys.stderr)
            plot_count += 1
        if sys.argv[1] == 'show':
            plt.show()
        elif sys.argv[1] == 'save':
            if not os.path.exists('./QC'):
                os.mkdir('QC')
            plt.savefig('./QC/' + timestamp + '_' + str(batch_count) + '.png', dpi=300)
            print('# Saving batch', batch_count, file=sys.stderr)
        batch_count += 1

    # calculate the per-analyser QC stats
    print('# Collecting machine statistics...')
    machines = {}
    for QC in QCs:
        if QC.machine not in machines:
            machines[QC.machine] = {'pass': 0,
                                    'fail': 0}
    for QC in QCs:
        for reading in QC.readings:
            if QC.lower_limit <= reading[1] <= QC.upper_limit:
                machines[QC.machine]['pass'] += 1
            else:
                machines[QC.machine]['fail'] += 1
    print('# QC scores:')
    for machine in sorted(machines.keys()):
        print(machine, end='\t')
        print('Pass:', machines[machine]['pass'], end='\t')
        print('Fail:', machines[machine]['fail'], end='\t')
        score = machines[machine]['pass'] / (machines[machine]['pass'] + machines[machine]['fail'])
        print(score)



elif tag == 'GEN':
    label = str(int(sheet1[1][2]))
    series = []
    for row in sheet1[1:]:
        series.append({
            'x': datetime.datetime.strptime(str(row[20]).split('+')[0], '%Y-%m-%d %H:%M:%S'),
            'y': row[9]
            })
    import matplotlib.pyplot as plt
    x, y = [], []
    for point in series:
        if point['x'] and point['y']:
            x.append(point['x'])
            y.append(point['y'])
    plt.scatter(x=x, y=y)
    plt.title('Test code: ' + label)
    plt.plot()
    plt.show()
    points = [point['y'] for point in series]
    points = [point for point in points if point]
    plt.hist(x=points)
    plt.title('Histogram of ' + label)
    plt.plot()
    plt.show()

elif tag == 'TAT':
    '''
    SELECT
        tat_data.hosp "hosp",
        tat_data.month "month",
        tat_data.day "day",
        DATEPART(dw, tat_data.arrival_date) "dow",
        DATEPART(hh, tat_data.arrival_date) "hh",
        DATEPART(mi, tat_data.arrival_date) "mi",
        (SELECT public_holiday_desc FROM public_holiday WHERE CAST(public_holiday AS DATE) = tat_data.arrival_date) "ph",
        tat_data.arrival_date "arrival_date",
        tat_data.reqno "reqno",
        tat_data.code "test_code",
        tat_data.time "tat",
        tat_data.urgency "urgency"

        FROM
        (
                SELECT
                DATEPART(mm, raw_data.arrival_date) "month",
                DATEPART(dd, raw_data.arrival_date) "day",
                arrival_date,
                raw_data.test_alpha_code "code",
                raw_data.request_no "reqno",
                DATEDIFF(ss, raw_data.arrival_date, raw_data.auth_date) "time",
                raw_data.hospital "hosp",
                raw_data.urgency "urgency"

                FROM
                        (
                        SELECT

                        (SELECT td.test_alpha_code FROM test_dict td WHERE td.test_ckey = tr.testrslt_member_ckey) "test_alpha_code",
                        tr.testrslt_member_ckey "test_code",
                        tr.testrslt_auth_date "auth_date",
                        rq.*


                        FROM

                        testrslt tr

                        RIGHT JOIN

                                (
                                SELECT
                                request.req_reqno "request_no",
                                request.req_arrived_date "arrival_date",
                                request.req_hospital "hospital",
                                request.req_urgency "urgency"
                                FROM request
                                WHERE
                                request.req_arrived_date >= '2016-11-01 00:00' AND
                                request.req_arrived_date < '2016-12-01 00:00'
                                ) rq

                        ON
                                tr.testrslt_reqno = rq.request_no
                        ) raw_data

                WHERE raw_data.test_alpha_code = 'NA'
                AND raw_data.hospital = 'RHT'
                AND raw_data.urgency = 2
        ) tat_data

    GROUP BY tat_data.month, tat_data.day
    ORDER BY tat_data.month, tat_data.day
    '''
    A = []
    P = []
    N = []
    H = []
    W = []
    X = []
    Y = []
    Z = []
    for row in sheet1[1:]:
        if not row[10]:
            continue
        time_part = datetime.datetime.strptime(str(row[7]).split('+')[0], '%Y-%m-%d %H:%M:%S').time()
        date_part = datetime.datetime.strptime(str(row[7]).split('+')[0], '%Y-%m-%d %H:%M:%S').date()
        if row[0] == 'RHT':
            if row[3] == 1 or row[6]:  # is Sunday or public holiday
                if time_part >= time_object('09:01:00') and time_part <= time_object('17:00:00'):
                    H.append([row[10], date_part])
                elif time_part >= time_object('17:01:00') and time_part <= time_object('19:00:00'):
                    W.append([row[10], date_part])
                elif time_part >= time_object('19:01:00') and time_part <= time_object('21:00:00'):
                    Z.append([row[10], date_part])
                elif time_part >= time_object('21:01:00') or time_part <= time_object('09:00:00'):
                    N.append([row[10], date_part])
            elif row[3] >= 2 and row[3] <= 6:  # Mon - Fri
                if time_part >= time_object('09:01:00') and time_part <= time_object('17:00:00'):
                    A.append([row[10], date_part])
                elif time_part >= time_object('17:01:00') and time_part <= time_object('19:00:00'):
                    W.append([row[10], date_part])
                elif time_part >= time_object('19:01:00') and time_part <= time_object('21:00:00'):
                    X.append([row[10], date_part])
                elif time_part >= time_object('21:01:00') or time_part <= time_object('09:00:00'):
                    N.append([row[10], date_part])
            elif row[3] == 7:  # Sat
                if time_part >= time_object('09:01:00') and time_part <= time_object('13:00:00'):
                    A.append([row[10], date_part])
                elif time_part >= time_object('13:01:00') and time_part <= time_object('17:00:00'):
                    Y.append([row[10], date_part])
                elif time_part >= time_object('17:01:00') and time_part <= time_object('21:00:00'):
                    Z.append([row[10], date_part])
                elif time_part >= time_object('21:01:00') or time_part <= time_object('09:00:00'):
                    N.append([row[10], date_part])
        elif row[0] == 'PYN':
            if row[3] == 1 or row[6]:  # is Sunday or public holiday
                if time_part >= time_object('09:01:00') and time_part <= time_object('17:00:00'):
                    H.append([row[10], date_part])
                elif time_part >= time_object('17:01:00') and time_part <= time_object('21:30:00'):
                    P.append([row[10], date_part])
                elif time_part >= time_object('21:31:00') or time_part <= time_object('09:00:00'):
                    N.append([row[10], date_part])
            elif row[3] >= 2 and row[3] <= 6:  # Mon - Fri
                if time_part >= time_object('09:01:00') and time_part <= time_object('17:00:00'):
                    A.append([row[10], date_part])
                elif time_part >= time_object('17:01:00') and time_part <= time_object('21:30:00'):
                    P.append([row[10], date_part])
                elif time_part >= time_object('21:31:00') or time_part <= time_object('09:00:00'):
                    N.append([row[10], date_part])
            elif row[3] == 7:  # Sat
                if time_part >= time_object('09:01:00') and time_part <= time_object('17:00:00'):
                    A.append([row[10], date_part])
                elif time_part >= time_object('17:01:00') and time_part <= time_object('21:30:00'):
                    P.append([row[10], date_part])
                elif time_part >= time_object('21:31:00') or time_part <= time_object('09:00:00'):
                    N.append([row[10], date_part])
        else:
            raise ValueError(row[0], 'is not recognized')
    chart_title = 'TAT90 estimates for ' + sheet1[1][0] + ' ' + sheet1[1][9]
    if sheet1[1][11] == 2:
        chart_title +=' ROUTINE'
    elif sheet1[1][11] == 1:
        chart_title +=' URGENT'
    print(chart_title)
    shifts = {'A': A, 'P': P,
              'N': N, 'H': H,
              'W': W, 'X': X,
              'Y': Y, 'Z': Z
              }
    import matplotlib.pyplot as plt
    for shift_code in 'APNHWXYZ':
        if len(shifts[shift_code]) > 0:
            print(shift_code, percentile([x[0] for x in shifts[shift_code]], 90, interpolation='higher')/3600, 'n=', len(shifts[shift_code]))
            daily_tat = {}
            for point in shifts[shift_code]:
                day = point[1]
                tat = point[0]/3600
                if day not in daily_tat:
                    daily_tat[day] = []
                daily_tat[day].append(tat)
            box_plot_data = [[x[0]/3600 for x in shifts[shift_code]]]
            x_labels = ['Monthly average n = ' + str(len(shifts[shift_code]))]
            for day in sorted(daily_tat.keys()):
                box_plot_data.append(daily_tat[day])
                x_labels.append(day)
            plt.boxplot(box_plot_data, whis=[10,90])
            plt.xticks(range(1, len(x_labels)+1), x_labels, rotation='vertical')
            plt.title(chart_title + ' ' + shift_code)
            plt.show()

elif tag == 's':
    result = {}
    test_code_dict = {
                       '5191': 'IGA',
                       '5192': 'IGG',
                       '5193': 'IGM',
                       '4178': 'Ca',
                       '4179': 'PO4',
                       '4175': 'ALP',
                       '4176': 'ALT',
                       '4173': 'ALB',
                       '4258': 'GLOB',
                       '4171': 'CRE'
                       }
    for item in test_code_dict:
        result[test_code_dict[item]] = 'n.a.'

    for row in sheet1[1:]:
        key = str(int(row[2]))
        test_code = test_code_dict[key]
        result[test_code] = row[9]

    for item in sorted(result):
        print(item, ':', result[item])


print('# Done.', file=sys.stderr)









