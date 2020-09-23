Programming.exe
By:        Jørgen E. Levesen and numerous others, credit is given in the frmAbout
MailTo:    jorgen@levesen.com / jorgen@sandviks.com
Internet:  http://www.levesen.com


This program is first and foremost developed for my own use as a helping hand in my spare time programming, using Visual Basic. It is now made public domain because I found it just appropriate to offer this complete program in return for all the help I have got from various programmers.

The program got 6 sections, which are:

1.	Code: Code snippets, Code Zip-files
2.	Programming: Program hours, Projects, Licence, Mail, Customers, Invoice, Payments
3.	Database: Passwords, Database fields, Repair database, User record, Screen text, Country
4.	Visual Basic: Code statistic, Spell VB code
5.	Pictures: Picture viewer, Icons 
6.	About: About, Write to me, Supplier, Register, Update program

As well as Internet Links, Colors, Computer Password and Key ASCII.

Start up:
Please go to sub-system: "Database/User Record" and change the fields as you wish, please observe that the ComboBox: "Screen Language" on Tab-page 3 controls the actual text language which is shown in the various forms. All text is stored in the MS Access database "CodeLang.mdb". As is there are at the moment three languages in the database: English, Norwegian and Italian. If you select a language not in the database, the text will be shown in English - you will then have to go to "Database/Screen Text" and manually change all the text-fields shown for the selected language (or make the canges directly in the database).

At startup the system ask you for a CodeSnippet database name, and rename the database (MyOwnSnippets.mdb) acordingly with the given name. This database is your own personal in which you can store your own code snippets. This is done that way because you can download an update CodeSnippet database from my WEB-site any time you wish. Upon downloading this new database will not owerwrite whichever data you may have written yourself.

The "Copy fields names with apostrophes" are for use with "the Database/Database fields" sub-system where you can copy recordset field-names for use in your programming.

I have started a simple Help-system; if you click the help-button in top of the screen a small help-screen will appear giving a short textual help to the actual shown form. If you click the form-name caption shown below the TextBox, a special edit-form appears where you can edit/add the help-text.

I have submitted all the complete code snippets I have collected (approximately 1800 snippets); the Zip-codes are only submitted as an example because of the database size. At the moment I got approximately 900 zip-files in the database contributing to a 80 MB database-size, which is impossible for me to upload (I only got a 50 Mb web-site).
In the CodeSnippets section you can click with right hand mouse button for formating of the code text (bold, italic, Uline, font, font size, font color as well as copy and paste functions).

In the sub-system "About/Register Program" you can (if you wish) email me a registration form, I will then inform you of changes made, new database updates etc.

In the program the following refferences are made to:

1.	Visual Basic for Applications		(msvbvm60.dll)
2.	Visual Basic Runtime objects and procedures
3.	Visual Basic objects and procedures	(vb6.olb)
4.	Ole Automation				(stdole2.tlb)
5.	Microsoft Word 9.0 object library	(msword9.olb)
6.	Microsoft DAO 3.6 object library	(dao360.dll)
7.	Microsoft Access 9.0 object library	(msacc9.olb)
8.	Microsoft Outlook 9.0 object library	(msoutl9.olb)
9.	SMTP Send Mail for VB6.0		(vbSendMail.dll)

I have used the following controls:

1.	Microsoft common Dialog Control 6.0 (SP3)	(comdlg32.ocx)
2.	Microsoft Databound Grid Control 5.0 (SP3)	(dbgrid32.ocx)
3.	Microsoft FlexGrid Control 6.0 (SP3)		(msflxgrd.ocx)
4.	Microsoft Masked Edit Control 6.0 (SP3)		(msmask32.ocx)
5.	Microsoft Rich TextBox control 6.0 (SP4)	(richtx32.ocx)
6.	Microsoft Tabbed Dialog Control 6.0 (SP5)	(tabctl32.ocx)
7.	Microsoft Windows Common Controls 6.0 (SP4)	(mscomctl.ocx)
8.	Microsoft Winsock Control 6.0 (SP5)		(mswinsck.ocx)

The download comes includes six (6) various databases:

1.	CodeIco.mdb		Icon Database
2.	CodeLang.mdb		Language database, where the form text is stored.
3.	CodeMaster.mdb		The master database where most of the recordset are stored.
4.	CodeSnippets.mdb	Database for various code snippets, -type and -language.
5.	CodeZip.mdb		Database for zip-files
6.	MyOwnSnippets.mdb	Your own code snippets database (renamed at first run)

As previous stated, this source code is free - you can alter it at your hearts wish - but I would be very grateful for any feed-back, good or bad, you may want to share with me.

Have a nice day
Jørgen Levesen - jorgen@levesen.com - http://www.levesen.com