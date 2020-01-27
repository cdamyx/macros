# macros
Handy VBA macros 

### extractMonthData.SortByDay.VBA
Very specific macro. If you have a worksheet (named "Birthday") of employee data, including headers called "Birth Date", "Hire Date", and "Rehire Date", this program will filter by month on those columns, create 2 new worksheets, then sort by day (disregarding year). Here are the headers from the original excel file that I was working with: Branch,	Dept,	Emp Name,	Address 1, Suite,	Address 2,	City,	State/Province,	ZIP/Postal Code,	Birth Date,	Age,	Hire Date,	Rehire Date,	Length of Service. If you have a worksheet with those headers (and populated data - works best if appropriate fields are completed), have named the worksheet "Birthday", then run the main() subroutine, you should successfully see this program run.


### pullDayAndSort.vba
snippet originally created before extractMonthData.SortByDay.vba


### saveAttachmentsLocally.vba
Grab all attachments of emails in a particular Outlook folder and save them to a local path
