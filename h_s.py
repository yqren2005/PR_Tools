from datetime import datetime, date
import pandas as pd
import sys

DATE_FORMAT = "%m/%d/%Y"
TODAY = datetime.strptime(date.today().strftime(DATE_FORMAT), DATE_FORMAT)


def convert_date(s):
    return datetime.strptime(str(s).split(' ')[0], DATE_FORMAT)


def calculate_date(d):
    return (TODAY - d).days


def contain_words(s, l):
    for w in l:
        if w in s:
            return True
    return False


def filtering(df, keywords, columns, flag):
    msg, rules = '', []
    lower_columns = [col.lower() for col in df.columns]
    while flag not in ['i', 'e']:
        print "\n%s not recognized" % flag
        flag = raw_input("Enter either i to include or e to exclude: ")
    for k, v in enumerate(columns):
        while v.lower() not in lower_columns:
            print "\n%s doesn't exist" % v
            v = raw_input("Re-enter this column: ")
        columns[k] = df.columns[lower_columns.index(v.lower())].encode('utf8')
        rules.append(df[columns[k]].str.contains('(?i)' + str(keywords), na=False))
    msg = "%s PRs matching %s in %s" % (('Omit:' if flag == 'e' else 'Include:'),
                                        (keywords if len(keywords) < 10 else keywords[:10] + '...'),
                                        ', '.join(columns))
    return reduce(lambda x, y: x | y, rules), msg


try:
    filename, customer = sys.argv[1:3]
except ValueError:
    print '''\nUsage: python h_s.py inputfile customer
Example: python h_s.py piir-hs-report-results.xls google
Check result in output-<timestamp>.xlsx'''
    exit(2)

try:
    df = pd.read_excel(filename, skiprows=8, sheetname=0, converters={'Bug ID': str, 'Score (BCF)': str})
except IOError as e:
    print str(e).split('] ')[1]
    exit(2)

df.insert(0, 'Analysis', '')
df['Created'] = df['Arrival-Date'].map(convert_date)
df['Updated'] = df['Last-Modified'].dropna().map(convert_date)
df_include = df_exclude = pd.DataFrame(columns=df.columns)
count = df.shape[0]

print "\nHack and slash in process..."

rule_1 = df['Customer'].str.contains('(?i)' + str(customer), na=False)

rule_2 = (df['Created'] == df['Updated']) | \
         df.apply(lambda x: calculate_date(x['Created']) >= 30 and pd.isnull(x['Last-Modified']), axis=1)

rule_3 = df.apply(lambda x: pd.isnull(x['Fixed In (BCF)']) and pd.notnull(x['Updated']) and
                            calculate_date(x['Updated']) >= 180, axis=1)

rule_4 = df.apply(lambda x: x['Submitter-Id'] in ('beta', 'development', 'jtac', 'other', 'systest') and
                            pd.isnull(x['JTAC-Case-Id']) and pd.isnull(x['Customer']) and pd.notnull(x['Updated']) and
                            calculate_date(x['Updated']) >= 30 and pd.isnull(x['Fixed In (BCF)']), axis=1)

rule_5 = df.apply(lambda x: x['State'] in ('closed', 'feedback', 'info') and
                            pd.isnull(x['Fixed In (BCF)']) and pd.notnull(x['Updated']) and
                            calculate_date(x['Updated']) >= 30, axis=1)

rule_6 = df.apply(lambda x: x['State'] in ('monitored', 'suspended') and
                            pd.isnull(x['Fixed In (BCF)']), axis=1)

rule_7 = df.apply(lambda x: contain_words(x['Synopsis'], ['core', 'crash', 'panic', 'assert']) and
                            pd.isnull(x['Fixed In (BCF)']) and pd.notnull(x['Updated']) and
                            calculate_date(x['Updated']) >= 30, axis=1)

rule_8 = df.apply(lambda x: x['Problem-Level'] == '6-IL4', axis=1)

rule_9 = df.apply(lambda x: float(x['Score (BCF)'].replace(',', '')) < 10.0, axis=1)

rules = {1: rule_1, 2: rule_2, 3: rule_3, 4: rule_4, 5: rule_5, 6: rule_6, 7: rule_7, 8: rule_8, 9: rule_9}
msgs = {1: 'Include: Bug encountered by customer',
        2: 'Omit: PR without updated since initially created',
        3: 'Omit: PR last updated more than 6 months ago with no Committed-Release or Conf-Committed-Release',
        4: 'Omit: More than 1 month old Development bug with no Committed-Release or Conf-Committed-Release',
        5: 'Omit: Closed, Feedback, Info with no Committed-Release or Conf-Committed-Release',
        6: 'Omit: Monitor / Suspend PR with no Committed-Release or Conf-Committed-Release',
        7: 'Omit: Non-reproducible crash related bug',
        8: 'Omit: Problem-Level = 6',
        9: 'Omit: PR Score is less than 10'}

while True:
    ans_1 = raw_input("\nAdd more rules? (y/n): ")
    if ans_1 == 'y':
        ans_2 = raw_input("Enter your criteria (For example: i,bgp|ospf,External-Title|Synopsis): ")
        while len(ans_2.split(',')) != 3:
            ans_2 = raw_input("\nInvalid criteria, please re-enter: ")
        flag, keywords, cols = [i for i in ans_2.split(',')]
        flag, cols = flag.strip(), [c.strip() for c in cols.split('|')]
        rules[max(rules.iterkeys()) + 1], msgs[max(msgs.iterkeys()) + 1] = filtering(df, keywords, cols, flag)
    elif ans_1 == 'n':
        break

for key in sorted(rules):
    df.loc[rules[key], 'Analysis'] = msgs[key]
    rows = df.loc[rules[key], :]
    if msgs[key][0] == 'I':
        df_include = df_include.append(rows, ignore_index=True)
    if msgs[key][0] == 'O':
        df_exclude = df_exclude.append(rows, ignore_index=True)
    print '\n', rows.shape[0], '\t-->', msgs[key]
    df.drop(rows.index, inplace=True)

writer = pd.ExcelWriter('output-' + str(date.today()) + '.xlsx', engine='xlsxwriter')
workbook = writer.book
format1 = workbook.add_format()
format1.set_text_wrap()
format1.set_align('left')
format1.set_align('vcenter')

for k, v in sorted({'Include': df_include, 'Exclude': df_exclude, 'Final': df}.items(), reverse=True):
    v.to_excel(writer, sheet_name=k, index=False, columns=df.columns[:-2])
    worksheet = writer.sheets[k]
    worksheet.set_zoom(100)
    worksheet.set_column('A:Z', 30, format1)
    worksheet.freeze_panes(1, 2)
    worksheet.autofilter(0, 0, 0, len(df.columns) - 1)

writer.save()

print "\nPR count: %i, Included: %i, Excluded: %i, Final: %i" \
      % (count, df_include.shape[0], df_exclude.shape[0], df.shape[0])
