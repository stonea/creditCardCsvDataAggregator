import csv, os, re, sys, pyodbc, datetime, dateutil 
from decimal import *
from os import listdir
from dateutil import parser as du_parser

# This script requires pyodbc.  To install use 'pip install pyodbc'
# Also make sure the 64 bit version of Microsoft's ACE driver is
# installed.  This ccan be found at
#   https://www.microsoft.com/en-US/download/details.aspx?id=13255

# This script requires dateutil.  Install with pip:
# pip install python-dateutil

# This script is probably pretty fragile.  One problem is figuring out
# how to identify what account a given file or row in a file corresponds to.
# What we do is we use heuristics to try and determine the account from
# rows in the file.  If that fails we pattern match the name of the file.
# See the patterns dictionary in processAccount.  This will likely have
# to be updated and made more robust.

def getDownloadPath():
    """Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


def findCsvFilenames(path_to_dir, suffix=".csv"):
    filenames = list(map(lambda x: x.lower(), listdir(path_to_dir)))
    return ['%s/%s' % (path_to_dir, filename) for filename in filenames if filename.endswith(suffix)]


def queryYesNo(question, default="yes"):
    """Ask a yes/no question via raw_input() and return their answer.

    "question" is a string that is presented to the user.
    "default" is the presumed answer if the user just hits <Enter>.
        It must be "yes" (the default), "no" or None (meaning
        an answer is required of the user).

    The "answer" return value is True for "yes" or False for "no".
    """
    valid = {"yes": True, "y": True, "ye": True,
             "no": False, "n": False}
    if default is None:
        prompt = " [y/n] "
    elif default == "yes":
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/N] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        sys.stdout.write(question + prompt)
        choice = input().lower()
        if default is not None and choice == '':
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            sys.stdout.write("Please respond with 'yes' or 'no' "
                             "(or 'y' or 'n').\n")


def determineSheetFormat(topRow):
    # Returns a dictionary indicating indices for columns of interest in the spreadsheet.
    # Some institutions will have separate columns for debit vs credit transactions.
    # Others will put everything into the same column.
    #
    # Some spreadsheets will not have a separate account column.  In these cases
    # we will return 'None' for the account and later on we will need to try and infer
    # the account from the filename.
    #
    # Some institutions split the date out into a transaction date and a post
    # date.  In such cases we prefer to get the transaction date.
    
    def determineDateField():
        # Get a list of all column indices that have the word 'date' in them
        dateFields = []
        for colIdx in range(0, len(topRow)):
            col = topRow[colIdx]
            if "date" in col:
                dateFields.append(colIdx)

        assert len(dateFields) > 0, 'Could not find date field in spreadsheet.'

        # If there is more than one datefield choose the one that says 'post' or 'posted'
        if len(dateFields) > 1:
            for potentialDateFieldIdx in dateFields:
                potentialDateField = topRow[potentialDateFieldIdx]
                if 'post' in potentialDateField:
                    return potentialDateFieldIdx

        assert len(dateFields) == 1, 'Multiple date fields in CSV, can not determine which one to use'
        return dateFields[0]


    dateColumnIdx = determineDateField()
    payeeColumnIdx = None
    creditColumnIdx = None
    debitColumnIdx = None
    accountColumnIdx = None
    
    for colIdx in range(0,len(topRow)):
        col = topRow[colIdx]
        if "payee" in col or "description" in col:
            payeeColumnIdx = colIdx
        if "debit" in col:
            debitColumnIdx = colIdx
        if "credit" in col:
            creditColumnIdx = colIdx
        if "amount" in col:
            debitColumnIdx = colIdx
            creditColumnIdx = colIdx
        if "account" in col or 'card no' in col:
            accountColumnIdx = colIdx

    return {'dateColumnIdx': dateColumnIdx,
            'payeeColumnIdx': payeeColumnIdx,
            'creditColumnIdx': creditColumnIdx,
            'debitColumnIdx': debitColumnIdx,
            'accountColumnIdx': accountColumnIdx}


def getNextRow(readCsv):
    return list(map(lambda x:x.lower(), next(readCsv)))


def findTopRow(readCsv):
    row = []
    while len(row) <= 3:
        row = getNextRow(readCsv)
    return row


def parseAndProcessRow(row, sheetFormat, fileName):
    # Will return None for items that are not credit card transactions
    
    def processAmount(amount):
        # Coming out of this function amounts should be positive for debits and negative for
        # credits.
        #
        # Some spreadsheets will have a separate credit column for debits vs credits.
        # Coming out of this function we want a debit transaction to be POSITIVE
        # and a credit transaction to be NEGATIVE.
        #
        # I operate under the assumption that any given transaction will either be a
        # debit or a credit.
        #
        # The following table shows how different banks format things in their CSV files:
        #
        # bank		debit (charge)		credit (refund\payoff)	format
        # -----		--------------		-----------		--------
        # barclays	NEGATIVE                POSITIVE		single column
        # capital one	POSITIVE                POSITIVE		separate columns
        # chase		NEGATIVE                POSITIVE		single column
        # boa		NEGATIVE                POSITIVE		single column
        # pmcu		NEGATIVE (in parens)	POSITIVE		single column

        # Extract amount from either debit or credit column.  If we get it
        # from a separate debit column (see capital one) then prefix the
        # string with a minus sign
        if(sheetFormat['debitColumnIdx'] != sheetFormat['creditColumnIdx']):
            # We're in the tricky capital one separate column case here
            debitAmount = row[sheetFormat['debitColumnIdx']].strip()
            creditAmount = row[sheetFormat['creditColumnIdx']].strip()
            assert creditAmount == '' or debitAmount == '', 'transaction expected to either be uniquely a credit or debit'
            if debitAmount != '':
                amount = debitAmount
                assert amount[0] != '-', 'debit transactions expected to be positive coming out of CSV file with separate debit and credit columns'
                amount = '-' + amount # Make it negative so it matches all the other CSV formats
            else:
                amount = creditAmount
                assert amount[0] != '-', 'credit transactions expected to be positive coming out of CSV filewith separate debit and credit columns'
        else:
            # We're in the normal case where there is one column for all transactions
            amount = row[sheetFormat['debitColumnIdx']].strip()

        # Strip out dollar signs, commas, replace amounts in parens with a negative sign
        amount = amount.replace('$', '')
        amount = amount.replace(',', '')
        if '(' in amount and ')' in amount:
            amount = amount.replace('(', '')
            amount = amount.replace(')', '')
            amount = '-%s' % amount

        return -Decimal(amount)

    def processPayee(payee):
        payee = re.sub(' +', ' ', payee)
        payee = payee.replace('"', '')
        payee = payee.replace("'", '')
        payee = payee.replace(";", '')
        return payee

    def processAccount(account):
        if account is None:
            return 'NOT A CC'
        account = account.lower()

        patterns = [
            (r'^bank of america.*cash rewards visa.*', 'boa_cash'),
            (r'^bank of america.*amtrak world.*', 'boa_amtrak'),
            (r'^bank of america.*travel rewards visa.*', 'boa_travel'),
            (r'^bank of america.*alaska airlines.*', 'boa_alaska'),
            (r'.*/chase9999_.*', 'chase_hyatt'),
            (r'.*/chase9998_.*', 'chase_csp'),
            (r'^9997$', 'c1_venture'),
            (r'.*/creditcard_.*', 'bk_jetblue'),
            (r'.*/export_\d+\.csv$', 'pmcu_visa'),
            (r'^bank of america.*adv plus banking.*', 'NOT A CC')]
        foundMatch = False
        for (pattern, substitute) in patterns:
            if re.match(pattern, account):
                foundMatch = True
                account = substitute
                break
        
        assert foundMatch, "Was not able to determine how to classify account: %s" % account
        return account

    date = du_parser.parse(row[sheetFormat['dateColumnIdx']].strip())
    payee = processPayee(row[sheetFormat['payeeColumnIdx']].strip())
    amount = processAmount(processAmount(row))
    if sheetFormat['accountColumnIdx'] == None:
        account = processAccount(fileName)
    else:
        account = processAccount(row[sheetFormat['accountColumnIdx']].strip())

    if account == 'NOT A CC':
       return None

    return {'date': date, 'amount': amount, 'payee': payee,
            'account': account}


def processSpreadsheet(fileName, readCsv):
    topRow = findTopRow(readCsv)
    sheetFormat = determineSheetFormat(topRow)

    data = []
    for row in readCsv:
        parsedRow = parseAndProcessRow(row, sheetFormat, fileName)
        if(parsedRow is not None):
            data.append(parsedRow)
    return data


def importDataFromCsvFiles():
    data = []
    for spreadsheetFileName in findCsvFilenames(getDownloadPath()):
        readCsv = csv.reader(open(spreadsheetFileName), delimiter=',')
        data.extend(processSpreadsheet(spreadsheetFileName, readCsv))
    return data


def printData(aliasToAccountIdMap, data):
    sortedData = \
        sorted(data, key = lambda x : (x['account'], x['amount']), reverse=True)
    accountIdToAliasMap = {v: k for k, v in aliasToAccountIdMap.items()}

    for entry in sortedData:
        date = entry['date']
        amount = entry['amount']
        payee = entry['payee']
        account = accountIdToAliasMap[entry['account']]
        entry['account']

        duplicateChar = ' '
        #if 'isDuplicate' in entry and entry['isDuplicate']:
        #    duplicateChar = '*'
        #    if skipDuplicates:
        #        continue
        
        print('%s%10s %8.2f %42s   %17s' % (duplicateChar, date, amount, payee, account))


def entryTuple(entry):
    date = "#%s#" % str(entry['date'])
    amount = str(entry['amount'])
    payee = "'%s'" % entry['payee']
    account = "'%s'" % entry['account']

    return (date, amount, payee, account)


def openDatabase():
    path = '%s/%s' % (os.getcwd(), 'finances_db.accdb')
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s;' % path)
    cursor = conn.cursor()
    cursor.execute('select * from cc_transactions')
    return cursor


def filterOutDuplicateEntries(db, data):
    newData = []
    sqlCheckIfExistsCmd = """\
        SELECT count(*) FROM cc_transactions WHERE
            transaction_date = %s AND
            amount = %s AND
            payee = %s AND
            account = %s;"""

    for dataEntry in data:
        dataTuple = entryTuple(dataEntry)
    
        # Check if entry already exists
        val = db.execute(sqlCheckIfExistsCmd % dataTuple)
        containsValue = val.fetchall()[0][0] > 0

        if not containsValue:
            newData.append(dataEntry)

    return newData


def commitDataToDatabase(data, db):
    sqlCmd = """\
        INSERT
            INTO cc_transactions (transaction_date, amount, payee, account)
            VALUES (%s, %s, %s, %s);"""
    
    for dataEntry in data:
        dataTuple = entryTuple(dataEntry)
        db.execute(sqlCmd % dataTuple)

    db.commit()


def commitIfUserAgrees(data, db):
    print()
    shouldCommit = queryYesNo("Commit Transactions?")
    if shouldCommit:
        commitDataToDatabase(data, db)
        print()
        print("Transactiosn committed")
    else:
        print("Submit canceled")
        sys.exit(1)


def retreiveAccountTypesEnum(db):
    sql = "SELECT ID, alias FROM ENUM_AccountTypes;"
    val = db.execute(sql)
    result = {}
    for (id, alias) in val.fetchall():
        result[alias] = id
    return result


def retreiveAliasToAccountIdMap(db):
    creditCardID = retreiveAccountTypesEnum(db)['Credit Card']
    sql = "SELECT ID, alias FROM accounts WHERE Type = %d" % creditCardID
    val = db.execute(sql)
    aliasToAccountIdMap = {}
    for (id, alias) in val.fetchall():
        aliasToAccountIdMap[alias] = id
    return aliasToAccountIdMap

    
def assignAccountIds(aliasToAccountIdMap, data):
    # We have a table of credit cards in cc_accounts.
    # We want to refer to the primary keys in there rather
    # than storing the payee as a string.
    for line in data:
        line['account'] = aliasToAccountIdMap[line['account']]
    return data

# ---------------------------------------
                           
def processTransactions():
    data = importDataFromCsvFiles()
    db = openDatabase()  
    aliasToAccountIdMap = retreiveAliasToAccountIdMap(db)
    data = assignAccountIds(aliasToAccountIdMap, data)
    data = filterOutDuplicateEntries(db, data)
    printData(aliasToAccountIdMap, data)
    commitIfUserAgrees(data, db)

def displaySpreadsheetFormatInfo():
    # This is a useful function to call if you've opened a new bank account.
    # It will display information about the format this script is assuming
    # the bank is using.  You'll want to double check this before processing
    # a bunch of data.
    for spreadsheetFileName in findCsvFilenames(getDownloadPath()):
        print()
        print()        
        readCsv = csv.reader(open(spreadsheetFileName), delimiter=',')

        print("SPREADSHEET:\n\t%s" % spreadsheetFileName)
        
        topRow = findTopRow(readCsv)
        print("TOP ROW:\n\t%s" % "\n\t".join("%d: %s" % (num,val) for num,val in zip(range(len(topRow)), topRow)))
        
        sheetFormat = determineSheetFormat(topRow)
        print("SHEET FORMAT:\n\t%s" % "\n\t".join("%20s: %s" % (k,v) for k,v in sheetFormat.items()))

        nextRow = next(readCsv)
        print ("NEXT ROW:\n\t%s" % "\n\t".join("%d: %s" % (num,val) for num,val in zip(range(len(topRow)), nextRow)))

        processedRow = parseAndProcessRow(nextRow, sheetFormat, spreadsheetFileName)
        print ("PROCESSED ROW:\n\t%s" % "\n\t".join("%10s: %s" % (k,v) for k,v in processedRow.items()))

# ---------------------------------------
processTransactions()
#displaySpreadsheetFormatInfo()
