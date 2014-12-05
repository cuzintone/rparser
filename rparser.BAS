'
'
' Route Parser
' Written by Antony Gordon
'
' 10/13/2010
'
'------------------------------------------------------------------------
'
' Route Parser takes the electronically generated route sheets from Comcast
' and parses them into a comma separated values (CSV) file that can be
' easily imported and analyzed using tools like Microsoft Excel, Microsoft
' Access, or any other database or spreadsheet application.
'
' Requirements:
' --------------
'
' REP1.TXT - The electronic version of the Comcast routes




DEFINT A-Z
' $DYNAMIC

' SUB/FUNCTION Declarations
DECLARE SUB QBMain ()
DECLARE SUB Splash ()
DECLARE SUB Split (theString AS STRING, theDelim AS STRING, strOut() AS STRING)
DECLARE FUNCTION ReadUntil$ (theString AS STRING, theDelim AS STRING)
DECLARE SUB ParseString (theString AS STRING, Buf() AS STRING, Delimiter AS STRING)
DECLARE FUNCTION GetJob$ (s AS STRING, strPos AS INTEGER)
DECLARE FUNCTION GrabTechNumber (s AS STRING)
DECLARE SUB ParseLine (theString AS STRING, theLine AS INTEGER, retval AS ANY)
DECLARE FUNCTION GetFileCount& (filespec$)
' Variable Declarations

COMMON SHARED TechNumber AS INTEGER
CONST FALSE = 0
CONST TRUE = NOT (FALSE)
CONST theDataFile$ = "REP1.TXT"


TYPE TechJob
  JobNumber AS STRING * 8
  JobTimeFrame AS STRING * 11
  JobType AS STRING * 8
  JobCust AS STRING * 30
  JobAddress AS STRING * 41
  JobTasks(5) AS STRING * 41
  JobCity AS STRING * 30
  JobState AS STRING * 2
  JobPostalCode AS STRING * 10
  JobPhone(2) AS STRING * 10
  JobAcctNo AS STRING * 9
  JobWIP AS STRING * 41
END TYPE

DIM FileCount AS LONG

Splash                  ' Show intro screen

'
' This block of code ensures that we have the data file for processing before
' we actually try to open it.
'
' If the file isn't there, gracefully exit instead of showing the user a
' cryptic error message

FileCount = GetFileCount(theDataFile$)
IF FileCount > 0 THEN
  QBMain
ELSE
  PRINT : PRINT "Unable to locate data file "; theDataFile$
  PRINT "Processing unable to continue."
END IF

REM $STATIC
FUNCTION GetFileCount& (filespec$)
'This function evaluates a file specification and returns the
'number of files that match the specification.  Wild card characters
'("*" and "?") are permitted. Drive and directory path specifications
'may also be included in filespec$.
DIM FileCount AS LONG
 
    IF LEN(DIR$(filespec$)) = 0 THEN     'Ensure filespec is valid.
	FileCount& = 0                   'It's not.
    ELSE
	FileCount = 1                    'It is, so count files.
	DO WHILE LEN(DIR$) > 0
	    FileCount& = FileCount& + 1
	LOOP
    END IF
    GetFileCount = FileCount&
END FUNCTION

FUNCTION GetJob$ (s AS STRING, strPos AS INTEGER)
	GetJob$ = MID$(s, 1, strPos)
END FUNCTION

FUNCTION GrabTechNumber (s AS STRING)
'
' Grab the tech number from a string variable

  theTech = VAL(MID$(s, 14, 4))
  PRINT "Currently processing route for Tech #"; theTech
  GrabTechNumber = theTech
END FUNCTION

SUB OldParseString (theString AS STRING, Buf() AS STRING, Delimiter AS STRING)

  DIM Num, OldNum, Strlen, Count AS INTEGER
  'Initialize variables
  Num = 0
  OldNum = 1
  Count = 1

  ' Break down the string
  DO
    ' REDIM PRESERVE Buf(Count) AS STRING
    Num = INSTR(Num + 1, theString, Delimiter)
    Strlen = Num - OldNum
    IF Num > 0 THEN
      Buf(Count) = MID$(theString, OldNum, Strlen)
    ELSE
      Buf(Count) = MID$(theString, OldNum)
    END IF
    Count = Count + 1
    OldNum = Num + 1
  LOOP UNTIL Num = 0
  ' REDIM PRESERVE Buf(Count) AS STRING
END SUB

SUB ParseLine (theString AS STRING, theLine AS INTEGER, retval AS TechJob)
'
'  This subroutine will process the input line read from the green bar file
'  anpopu
DIM JobBuffer(1) AS STRING

  IF theLine = 1 THEN
    
    ' Break apart the first line from the file

    Split theString, ":", JobBuffer()
    

    retval.JobNumber = LTRIM$(STR$(TechNumber) + "-" + JobBuffer(0))
    retval.JobType = JobBuffer(1)
    retval.JobAddress = JobBuffer(2)
    retval.JobTasks(theLine - 1) = JobBuffer(3)
  END IF
  IF theLine = 2 THEN
   ' Break apart the next line
   Split theString, ":", JobBuffer()

   retval.JobTimeFrame = LEFT$(JobBuffer(1), 5) + "-" + RIGHT$(JobBuffer(1), 5)
   retval.JobCity = RTRIM$(LEFT$(JobBuffer(2), 17))
   retval.JobState = MID$(JobBuffer(2), 18, 2)
   retval.JobPostalCode = MID$(JobBuffer(2), 21, 5)
   retval.JobPhone(theLine - 2) = RIGHT$(JobBuffer(2), 10)
   retval.JobTasks(theLine - 1) = JobBuffer(3)
  END IF
  IF theLine = 3 THEN
    Split theString, ":", JobBuffer()
    retval.JobAcctNo = LEFT$(JobBuffer(2), 9)
    retval.JobPhone(theLine - 2) = RIGHT$(JobBuffer(2), 10)
    retval.JobCust = MID$(JobBuffer(2), 14, 15)
    retval.JobTasks(theLine - 1) = JobBuffer(3)
  END IF
  IF theLine = 4 THEN
    Split theString, ":", JobBuffer()
    retval.JobWIP = JobBuffer(1)
    retval.JobTasks(theLine - 1) = JobBuffer(2)
  END IF
END SUB

SUB QBMain
'
'
' The main execution portion of the program dir$
'
DIM theJob AS STRING * 3
DIM JobInfo(1) AS STRING
DIM theJobs AS TechJob
DIM theInFile AS STRING, theRoute AS STRING, aline AS STRING
DIM InFileHandle, OutFileHandle AS INTEGER


  CLS
  
  PRINT : PRINT "Processing cover sheets..."
  theInFile = theDataFile$
  theRoute = "ROUTE.TXT"

  ' Get next available file handles
  InFileHandle = FREEFILE
  

  ' Open the files
  OPEN theInFile FOR INPUT AS InFileHandle
  OutFileHandle = FREEFILE
  OPEN theRoute FOR OUTPUT AS OutFileHandle
  WRITE #OutFileHandle, "Tech-Job", "Job Type", "Time Frame", "Account #", "Customer Name", "Address", "City", "State", "Postal Code", "Phone 1", "Phone 2", "WIP", "Task 1", "Task 2", "Task 3", "Task 4"
  ' Process the Cover Sheets
  PRINT "Now processing route..."

  DO
  'FOR i = 1 TO 32
    INPUT #InFileHandle, aline
    IF LEFT$(aline, 3) = "INS" THEN
      ' Grab the tech number
      TechNumber = GrabTechNumber(aline)
    END IF
    SELECT CASE VAL(LEFT$(aline, 1))
      CASE 1 TO 9
	linecount = 1: ShowRecord = FALSE
	DO UNTIL linecount > 4
	  ParseLine aline, linecount, theJobs
	  linecount = linecount + 1
	  LINE INPUT #InFileHandle, aline
	LOOP
	IF linecount > 4 THEN
	  ShowRecord = TRUE
	END IF
      CASE ELSE
    END SELECT

    IF ShowRecord THEN
      ' Write the data to the file

      WRITE #OutFileHandle, theJobs.JobNumber, theJobs.JobType, theJobs.JobTimeFrame, theJobs.JobAcctNo, theJobs.JobCust, theJobs.JobAddress, theJobs.JobCity, theJobs.JobState, theJobs.JobPostalCode, theJobs.JobPhone(0), theJobs.JobPhone(1), theJobs _
.JobWIP, theJobs.JobTasks(0), theJobs.JobTasks(1), theJobs.JobTasks(2), theJobs.JobTasks(3)
      'FOR k = LBOUND(theJobs.JobTasks) TO UBOUND(theJobs.JobTasks)
      '  WRITE #OutFileHandle, theJobs.JobTasks(k)
      'NEXT
      ShowRecord = FALSE
    END IF
  LOOP UNTIL EOF(InFileHandle)
  ' NEXT
  PRINT "Route processing complete."

  CLOSE InFileHandle, OutFileHandle

END SUB

FUNCTION ReadUntil$ (theString AS STRING, theDelim AS STRING)

  DIM nPos AS INTEGER
  nPos = INSTR(1, theString, theDelim)

  IF nPos > 0 THEN
    ReadUntil$ = LEFT$(theString, nPos - 1)
    theString = MID$(theString, nPos + LEN(theDelim))
  END IF
END FUNCTION

SUB Splash
  PRINT "Auto Router Route Processing Program"
  PRINT "Version 0.01"
  PRINT "Written by Antony Gordon": PRINT

END SUB

SUB Split (theString AS STRING, theDelim AS STRING, strOut() AS STRING)
  DIM strRead AS STRING, nC AS INTEGER

  REDIM PRESERVE strOut(nC) AS STRING
  IF theDelim = "" THEN
    strOut(nC) = theString
  END IF

  strRead = ReadUntil$(theString, theDelim)
  DO
    REDIM PRESERVE strOut(nC) AS STRING
    strOut(nC) = strRead
    nC = nC + 1
    strRead = ReadUntil$(theString, theDelim)
  LOOP WHILE strRead <> ""
  REDIM PRESERVE strOut(nC) AS STRING

END SUB

