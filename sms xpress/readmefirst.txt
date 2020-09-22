SMS Xpress Verson 1.0.0

This is an updated version of my SMS Enterprise application as posted here. I decided to work on a single more comprehensive approach into this application and also add more functionality to the application, including its own ocx for sending smses to the phones.

A navigation screen has been made available now for ease of executing actions between the computer, the phone and the sim card. Each of these is treated separately though the actions that can be done are similar.

On Computer there is...
Modems - list installed modems on the computer, you can select the default modem
Groups - list names of groups existing in the computer database, you can add, delete and send an sms to a group more groups here
Contacts - list of contacts that are in the computer database, you can add, delete, send sms, export, copy from computer to sim, copy from computer to phone
Messages - messages stored in the computer database, you can create a new sms, resend, reply, delete messages stored in the computer. You can also view the inbox, outbox, sent box and recycled messages

On the Phone there is...
Contacts - list of contacts that are in the mobile equipment/cellphone database, you can add, delete, send sms, export, copy from phone to sim or copy from phone to computer
Messages - messages stored in the mobile equipment/cellphone database, you can forward, resend, reply, delete messages stored in the cellphone/modem memory. You can also view the inbox, outbox, sent box of the modem/cellphone.

On the Sim there is...
Contacts - list of contacts that are in the sim card database, you can add, delete, send sms, export, copy from sim card to phone or copy from sim card to computer
Messages - messages stored in the sim card database, you can forward, resend, reply, delete messages stored in the cellphone/modem memory. You can also view the inbox, outbox, sent box of the modem/cellphone.

GSM.OCX

This ocx works with the phone and sim card.

The ocx includes an mscomm control that will interface with your modem/cellphone. Currently this code has been tested with the Fastlink E620 data card and it works perfectly.
CommPort - this indicates the port that the connection will be made to
LogFile - if this is specified, a log will be kept of everything sent and received from the modem
Settings - connection settings of the modem, specify the speed property and the settings will be saved automarically, for this application, the maximum speed of the modem is used.
Speed - the speed of the connection, setting this will set the settings property of the comm port.
PortOpen - is the port open or closed
Connect - connect to the modem by specifying the comm port and the speed of the connection, this will return OK if the modem is ready
ManufacturerIdentification - returns the Manufacturer Identification
ModemSerialNumber - returns the Modem Serial Number i.e imei number
RevisionIdentification - returns the Revision Identification of the modem
ModelIdentification - returns the Model Identification of the mobile modem
SMS_NewMessageIndicate - set true or false if you want new messages to be sent to the terminal (Please note that I have not experienced with the outcome of this yet)
SignalQualityMeasure - returns the signal strength as a percentage eg 31
PhoneBook_MemoryStorage - set / return the phone book memory to use
PhoneBook_ReadEntry - return the details of a phone book entry using the location
PhoneBook_FindEntry - return the details of the phonebook entry searched using the name
PhoneBook_EntryExists - returns the location of the phonebook entry if found by searching the phone book using the name and cellnumber
PhoneBook_WriteEntry - write an entry to the phonebook by specifying the location, cellphone and name
PhoneBook_DeleteEntry - delete an entry from the phonebook using the location of the entry
PhoneBook_AvailableIndexes - list available indexes from the phonebook using an at command
Echo - sets echo to off or on. When echo is off, less traffic result however for this project it has been turned on.
Request - returns the result of an at commnd request. This uses a timer to check the result of the comm port. For this project the ExpectedResult variable has been left blank as unexpected results are obtained when OK or ERROR is specified
DescriptiveError - return a readable error message from phone errors, eg "+CME ERROR: 3" which means "Operation not allowed"
PhoneBook_ListView - list all phonebook entries on a listview
PhoneBook_AvailableIndex - returns the next available index on the phonebook. This is used when adding new contact details to the phonebook
PhoneBook_AddEntry - add an entry to the phonebook
SMS_MessageFormat - sets / returns the text format mode for the phone
SMS_MemoryStorage - sets / returns the memory storage for smses
SMS_CentreNumber - sets / returns the centre number for the phone
SubscriberNumber - returns the Subscriber Number (Currently this is not returning anything)
InternationalMobileSubscriberIdentity - returns the International Mobile Subscriber Identity i.e a unique number that identifies you in the world for your cellphone. I recommend you hide this number from your program as selfish people can use it to send smses on your phone. I also want to find out how.
SMS_ReadMessageEntry - read a sms using the index location from sim/phone
SMS_ListView - list all smses to listview, you can select which messages to view, sent, unsent, received
SMS_DeleteEntry - delete an sms from sim/phone using the index location
SMS_ReadMessages - returns a collection of all messages on the sim/phone
SMS_Send - send a sms to a particular number to the maximum of 320 characters only
PhoneBook_Export - export the phonebook to a csv file
PhoneBook_Import - import phonebook details from a csv file

To do
Copying of messages from sim, phone to computer and vise versa.
Complete backup of phone data including sim card data to computer and also restore
Import contacts from csv file to phone/sim card

To Note
Due to the time it takes to execute commands, the mscomm port is difficult to figure, however, one is cautioned to be patient with the time the application returns commands. From reading, a maximum of 20 seconds has been indicated as a maximum period. FOr this application, I run a loop 10 times to wait for 500 milliseconds for results of the comm port.