# BankBot
Banking automation in MS Excel showcasing Investec Beta transfer API functionality

# Summary
BankBot is an automated personal banking solution leveraging the Investec Programmable Banking environment.

BankBot is housed in a Macro enabled Microsoft Excel file, with extensive VBA (Visual Basic for Applications) coding for automation elements.

This submission is put forward as part of the Q42021 Hackathon.
It should be noted that only the **Transfer** functionality is under assessment from a Hackathon point of view - the underlying base of BankBot, using the existing Investec banking API's, was built **prior** to the Hackathon and only acts as a foundation for the Transfer functionality to be leveraged.

# Pre-requsites
1. MS Excel 2016+, with the ability to run Macros enabled.
2. Scheduler system such as Windows Task scheduler.
3. Investec bank account, enabled for programmable banking.

# Additional disclaimers
1. The author is a hobbyist programmer, with no formal programming training - accordingly there may be some deficits in code etiquette or conventions that will need to be excused.
2. This solution was developed rapidly with focus on achieving operation as swiftly as possible - it is intended to be useable only by advanced Programmers at this point and is by no means considered "polished".
3. Note that there is currently no solution for securing client credentials or bank account numbers - the User would thus need to exercise extreme caution not to expose this information unintentionally (consider password protecting the Excel Workbook to prevent unauthorised access).
