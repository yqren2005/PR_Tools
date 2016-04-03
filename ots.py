import pandas as pd
import sys


def main():
    try:
        filename, keywords, cols, flag = sys.argv[1:]
    except ValueError:
        print "Usage: python ots.py <ots_file> <keywords> <columns> <include|exclude>"
        print '''Example: python ots.py ots_report.xls "ospf|isis|bgp" "External-Title|Synopsis" i'''
        exit(2)

    try:
        ots_data = pd.ExcelFile(filename)
    except IOError as e:
        print str(e).split('] ')[1]
        exit(2)

    cols = [c.strip() for c in cols.split('|')]

    while flag not in ['i', 'e']:
        flag = raw_input("Enter either i to include or e to exclude: ")

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('result.xlsx', engine='xlsxwriter')
    workbook = writer.book
    format1 = workbook.add_format()
    format1.set_text_wrap()
    format1.set_align('left')
    format1.set_align('vcenter')

    for sheet in ots_data.sheet_names:
        df = ots_data.parse(sheet)
        old = df.shape[0]
        # Create an empty rule set
        rules = []

        for c in range(len(cols)):
            while cols[c] not in df.columns:
                print "\n%s doesn't exist" % cols[c]
                cols[c] = raw_input("Re-enter this column: ")
            rules.append(df[cols[c]].str.contains('(?i)' + str(keywords), na=False))
        condition = reduce(lambda rule_1, rule_2: rule_1 | rule_2, rules)
        if flag == 'e':
            df = df[~condition]
            # Remove rows with no external title and description
            df = df[~(pd.isnull(df.ix[:, 1]) & pd.isnull(df.ix[:, 5]))]
        elif flag == 'i':
            df = df[condition]

        print '\n' + sheet
        print "PR count: %s, new PR count: %s, filtered on %s" % (old, df.shape[0], str(list(set(cols))))

        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(writer, sheet_name=sheet, index=False)

        # Formatting
        worksheet = writer.sheets[sheet]
        worksheet.set_zoom(100)
        worksheet.set_column('A:M', 15, format1)
        worksheet.set_column('B:B', 45, format1)
        worksheet.set_column('D:D', 45, format1)
        worksheet.set_column('F:H', 45, format1)

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    print "\n%sCheck the result in result.xlsx" \
          % ("PRs with empty external title and description are also excluded. " if flag == 'e' else '')


if __name__ == '__main__':
    main()
