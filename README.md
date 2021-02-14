# Credit Card CSV Data Aggregator
Most banking websites let you export the history of your credit card
transactions into a CSV file.  This script combines data from these files
(across multiple institutions) into a single, locally stored, database.  

--

I have several credit cards and I wanted to aggregate my transaction history
into a single database I could use to analyze my spending habits.

There are a lot of convenient online services that exist to do this like YNAB
and Mint but these services rely on you giving the username and password to
each bank you want to import data from. I'm uneasy doing this and would rather
do everything locally so I wrote this script.

Essentially, the intent is you go to your various banks’ websites and download
your data, as CSV files, into your operating system's default download
directory.  The script then processes those files, extracts whatever
transactions haven't already been added to the database, then dumps those.

Every bank formats the data in the CSV file differently so the script tries to
infer the format. Currently it’s been configured to process CSV files dumped
from: Barclays, Capital One, Chase, BOA, as well as a local credit union I use.

It also uses the filename to try and infer what credit card the CSV file refers
to.  Again, this is driven by a table that would have to be updated for any
individual’s specific use case.

--

# Assumptions

This script is pretty much designed around my specific use case.  So if you
care to hack this for your use case be aware about the following assumptions:

* You've downloaded the CSV files you want to process to your Operating
  System's download directory.  Currently I assume you're either using Windows
  or your download directory is `~/downloads`.  Hack the getDownloadPath()
  if you want to change this behavior.
* The name of the credit card can be extracted from the filename using the
  patterns in the processAccount() function.
* The determineSheetFormat() function tries to infer the format of the CSV
  file.  I know this works for Barclays, Capital One, Chase, BOA, and a local
  credit union I use but it will likely have to be adapted to serve other institutions.

Additionally, the script assumes you already have a database set up to export
the data.

* You have an Access database named `finances_db.accdb1 in whatever directory you’re invoking the script from.
* This database has a table named cc_transactions with the fields: `ID`, `transaction_date`, `amount`, `payee`, and `account` where the field types are like this:
  * `ID` is an `Autonumber`
  * `Transaction_date` is a `Date/Time`
  * `Amount` is a  `Currency`
  * `Payee` is a `Short Text`
  * `Account` is a `Number`

Ideally, I should set up a script to set this all up for the user (and
generalize it to work with other types of databases). I haven't done this due
to personal time constraints.

