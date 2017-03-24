Option Explicit

Dim oFSO
Dim LogFile 			'logfile object
Dim conn 			'db connection
Dim cst				'db connection string
Dim rs, rs1, rs2, datelist	'recordset object
Dim objExcel			'Excel object
Dim SQL 			'SQL query statement
Dim x,y				'counter variables
Dim strContent	 		'SQL clause statements
Dim strLogFile, strOutputFile	'Files created by this script
Dim strWorkingDirectory
Dim strLastTeam
Dim intYear
Dim intSeasonLength
Dim arrTeamInfo()

' #############################################################
' #
' # Define Objects
' # 

'define universal variables
intYear = 2017
strWorkingDirectory = "c:\Users\mjbernha\Box Sync\python\"
strLogFile = "impact_tables.txt"
strOutputFile = "MLS Impact Table " & intYear & ".xlsx"

' File System Object
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Open log file
Set LogFile = oFSO.CreateTextFile(strLogFile,true)
LogText "Started"

' Open Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = false
objExcel.Workbooks.Add
LogText "Excel initialized"

' Database connection
cst = "Driver={MySQL ODBC 5.2 Unicode Driver};" & _ 
        "Server=SERVER;" & _ 
        "Port=3306;" & _
        "Database=DATABASE;" & _ 
        "User Id=USERNAME;" & _ 
        "Password=PASSWORD" 
set conn = CreateObject("ADODB.Connection")
conn.open cst
LogText "Opened db connection"

' #############################################################
' #
' # Start work
' # 

' #
' # Get overall team information
' #

LogText "Compiling overall team information"

' SQL = "SELECT TeamID, MatchTypeID, SUM(GameLength)
' FROM (
' 	SELECT TeamID, MatchTypeID, MAX(TimeOff) AS GameLength
'	FROM tbl_gameminutes m
'	INNER JOIN tbl_games g ON m.GameID = g.ID
'	WHERE YEAR(MatchTime) = 2015 AND MatchTypeID = 21 
'	GROUP BY TeamID, GameID
') AS GameLengths
'INNER JOIN lkp_matchtypes ON GameLengths.MatchTypeID = lkp_matchtypes.ID
'	GROUP BY GameLengths.TeamID, MatchTypeID"

SQL = "SELECT tbl_games.MatchTime, tbl_teams.teamname, tbl_gameminutes.TeamID, MAX(TimeOff) AS GameLength " & _
	"FROM tbl_gameminutes " & _
	"LEFT OUTER JOIN tbl_games ON tbl_gameminutes.GameID = tbl_games.ID " & _
	"LEFT OUTER JOIN tbl_teams ON tbl_gameminutes.TeamID = tbl_teams.ID " & _
	"WHERE YEAR(MatchTime) = " & intYear & " AND tbl_games.MatchTypeID = 21 " & _
	"GROUP BY tbl_gameminutes.GameID, tbl_gameminutes.TeamID " & _
	"ORDER BY teamname, MatchTime"
LogText SQL
set rs = conn.execute(SQL)
x = 0
intSeasonLength = 0
ReDim arrTeamInfo(2,0)
strLastTeam = rs("teamname")
arrTeamInfo(0,x) = rs("teamname")
arrTeamInfo(1,x) = rs("TeamID")
arrTeamInfo(2,x) = intSeasonLength

do while not rs.eof
	if strLastTeam <> rs("teamname") then
		'store the season length so far
		arrTeamInfo(2,x) = intSeasonLength
		
		'reset variables, increment X
		x = x + 1
		ReDim Preserve arrTeamInfo(2,x)
		
		intSeasonLength = 0
		arrTeamInfo(0,x) = rs("teamname")
		arrTeamInfo(1,x) = rs("TeamID")
		arrTeamInfo(2,x) = intSeasonLength
		
	end if

	intSeasonLength = intSeasonLength + CInt(rs("GameLength"))

	strLastTeam = rs("teamname")	
	rs.movenext
loop	

arrTeamInfo(2,x) = intSeasonLength
ArrayContents

' #
' # Get player information for all players who have appeared
' #

SQL = "SELECT CONCAT(FirstName,' ',LastName) AS PlayerName, teamname, tbl_gameminutes.PlayerID, tbl_gameminutes.TeamID, SUM(TimeOff-TimeOn) AS Minutes " & _
	"FROM tbl_gameminutes " & _
	"LEFT OUTER JOIN tbl_games ON tbl_gameminutes.GameID = tbl_games.ID " & _
	"LEFT OUTER JOIN tbl_teams ON tbl_gameminutes.TeamID = tbl_teams.ID " & _
	"LEFT OUTER JOIN tbl_players ON tbl_gameminutes.PlayerID = tbl_players.ID " & _
	"WHERE YEAR(MatchTime) = " & intYear & " AND tbl_games.MatchTypeID = 21 " & _
	"GROUP BY PlayerID, TeamID " & _
	"ORDER BY TeamName ASC, LastName ASC, FirstName ASC"
	LogText SQL
set rs = conn.execute(SQL)
x = 1
y = 0

'First Line
' Basic info
objExcel.Cells(x,1).Value = "##"
objExcel.Cells(x,2).Value = "ID"
objExcel.Cells(x,3).Value = "Player Name"
objExcel.Cells(x,4).Value = "Team"
' Playing time
objExcel.Cells(x,5).Value = "Minutes"
objExcel.Cells(x,6).Value = "Team Minutes"
objExcel.Cells(x,7).Value = "Pctg"
' Goals for
objExcel.Cells(x,8).Value = "Plus"
objExcel.Cells(x,9).Value = "PlusRate"
objExcel.Cells(x,10).Value = "TeamPlus"
objExcel.Cells(x,11).Value = "PctgPlus"
' Goals against
objExcel.Cells(x,12).Value = "Minus"
objExcel.Cells(x,13).Value = "MinusRate"
objExcel.Cells(x,14).Value = "TeamMinus"
objExcel.Cells(x,15).Value = "PctgMinus"
' Total goals
objExcel.Cells(x,16).Value = "Scoring"
objExcel.Cells(x,17).Value = "ScoringRate"
objExcel.Cells(x,18).Value = "Difference"
x = x + 1

do while not rs.eof
	if arrTeamInfo(0,y) <> rs("TeamName") then y = y + 1
	objExcel.Cells(x,1).Value = x
	objExcel.Cells(x,2).Value = rs("PlayerID")
	objExcel.Cells(x,3).Value = rs("PlayerName")
	objExcel.Cells(x,4).Value = rs("TeamName")
	objExcel.Cells(x,5).Value = rs("Minutes")
	objExcel.Cells(x,6).Value = arrTeamInfo(2,y)
	
	SQL = "SELECT Sum(Plus) AS Plus, Sum(Minus) AS Minus " & _
	"FROM tbl_gamestats " & _
	"LEFT OUTER JOIN tbl_games ON tbl_gamestats.GameID = tbl_games.ID " & _
	"WHERE PlayerID = " & rs("PlayerID") & " AND TeamID = " & rs("TeamID") & " AND Year(MatchTime) = " & intYear & " AND MatchTypeID = 21 " & _
	"GROUP BY PlayerID"
	LogText SQL
	set rs1 = conn.execute(SQL)
	
	if rs1.bof and rs1.eof then		
	else

		GoalsAndRate x,8,9,CInt(rs1("Plus")),CInt(rs("Minutes"))
		GoalsAndRate x,12,13,CInt(rs1("Minus")),CInt(rs("Minutes"))

		SQL = "SELECT SUM(IF(HTeamID=" & rs("TeamID") & ",HScore,AScore)) AS HomeGoals, SUM(IF(HTeamID=" & rs("TeamID") & ",AScore,HScore)) AS OppGoals " & _
		"FROM tbl_games " & _
		"WHERE YEAR(MatchTime) = " & intYear & " AND MatchTypeID = 21 " & _
		"AND MatchTime < NOW() " & _
		"AND (HTeamID = " & rs("TeamID") & " OR ATeamID = " & rs("TeamID") & ")"
		LogText SQL
		set rs2 = conn.execute(SQL)

		if rs2.bof and rs2.eof then
		else
			objExcel.Cells(x,10).Value = rs2("HomeGoals")
			objExcel.Cells(x,14).Value = rs2("OppGoals")
		end if

	end if
	
	x = x + 1
	rs.movenext
loop


objExcel.ActiveWorkbook.SaveAs(strWorkingDirectory & strOutputFile)

' #############################################################
' #
' # Close objects, wrap up
' # 

'wipe dbEventLogs
conn.close
set conn = Nothing
LogText "Closed db connection"

'Close the logfile
LogText "Finished"
objExcel.Quit
Set LogFile = Nothing
Set oFSO = Nothing
'wscript.echo("Finished!")

' #############################################################
' #
' # Subroutines
' # 

Sub LogText(strMessage)
	LogFile.WriteLine(now() & ":" & vbTab & strMessage)
End Sub

Sub ArrayContents()
	LogText "Array Contents:"
	for y = 0 to ubound(arrTeamInfo,2)
		LogText y & vbTab & arrTeamInfo(0,y) & vbTab & arrTeamInfo(1,y) & vbTab & arrTeamInfo(2,y)
	next
	LogText ""
End Sub

Sub GoalsAndRate(x,intX,intY,intGoals,intMinutes)
	objExcel.Cells(x,intX).Value = intGoals
	if intGoals=0 then 
		objExcel.Cells(x,intY).Value = ""
	else 
		objExcel.Cells(x,intY).Value = intMinutes/intGoals
	end if
End Sub

