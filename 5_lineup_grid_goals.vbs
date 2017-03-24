Option Explicit

Dim strWorkingDir
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
Dim strOpponent
Dim intTeamID 

' #############################################################
' #
' # Define Objects
' # 

'define universal variables
intTeamID = 11
strLogFile = "lineup_grid.txt"
strOutputFile = "lineup_grid_" & intTeamID & ".xlsx"

strWorkingDir = "C:\Users\mattj_000\Box Sync\Soccer\"

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
' # Get list of game dates
' #

LogText "Getting list of game dates"
SQL = "SELECT tbl_games.ID, MatchTime, lkp_matchtypes.MatchType, HTeamID, hometeam.teamname AS Hteam, ATeamID, awayteam.teamname AS ATeam " & _
	"FROM scouting.tbl_games " & _
	"INNER JOIN scouting.tbl_gameminutes ON tbl_games.ID = tbl_gameminutes.GameID " & _
	"INNER JOIN tbl_teams hometeam ON tbl_games.HTeamID = hometeam.ID " & _
	"INNER JOIN tbl_teams awayteam ON tbl_games.ATeamID = awayteam.ID " & _
	"LEFT OUTER JOIN lkp_matchtypes ON tbl_games.MatchTypeID = lkp_matchtypes.ID " & _
	"WHERE HTeamID = " & intTeamID & " OR ATeamID = " & intTeamID & " " & _
	"GROUP BY tbl_games.ID " & _
	"ORDER BY MatchTime ASC"
set datelist = conn.execute(SQL)
LogText SQL

LogText ""

LogText "Building table header"
x = 3
do while not datelist.eof
	objExcel.Cells(x,1).Value = datelist("ID")
	objExcel.Cells(x,2).Value = datelist("MatchTime")
	if datelist("HTeamID") = intTeamID then
		strOpponent = datelist("ATeam")
	else
		strOpponent = "@ " & datelist("Hteam")
	end if
	objExcel.Cells(x,3).Value = strOpponent
	'objExcel.Cells(x,4).Value = datelist("MatchType")
	x = x + 1
	datelist.movenext
loop
LogText "Table Header Complete"

LogText ""

' #
' # Get list of players, by first appearance
' #

LogText "Getting list of players"
SQL = "SELECT tbl_players.ID, Concat(FirstName,' ',LastName) AS PlayerName " & _
	"FROM scouting.tbl_players " & _
	"INNER JOIN scouting.tbl_gameminutes ON tbl_players.ID = tbl_gameminutes.PlayerID " & _
	"INNER JOIN scouting.tbl_games ON tbl_gameminutes.GameID = tbl_games.ID " & _
	"WHERE TimeOff > 0 AND TeamID = " & intTeamID & " " & _
	"GROUP BY tbl_players.ID " & _
	"ORDER BY Min(MatchTime), TimeOn ASC "
set rs = conn.execute(SQL)
LogText SQL

LogText "Looping through players"
y = 4
do while not rs.eof
	objExcel.Cells(1,y).Value = rs("ID")
	objExcel.Cells(2,y).Value = rs("PlayerName")
	
'	#
'	# Get list of appearances for this player
'	#

	LogText "Player " & rs("ID") & ": " & rs("PlayerName")
	
	SQL = "SELECT GameID, TimeOn, TimeOff - TimeOn AS Minutes " & _
	"FROM scouting.tbl_gameminutes " & _
	"INNER JOIN scouting.tbl_games ON tbl_gameminutes.GameID = tbl_games.ID " & _
	"WHERE TimeOff > 0 AND TeamID = " & intTeamID & " AND PlayerID = " & rs("ID") & " " & _
	"ORDER BY MatchTime ASC"
	'LogText SQL
	set rs1 = conn.execute(SQL)
	
	x = 1
	do while not rs1.eof
		'LogText rs1("GameID")
		
		do while objExcel.Cells(x,1).Value <> rs1("GameID")
			x = x + 1
		loop
		objExcel.Cells(x,y).Value = rs1("Minutes")
		
		'Did the player start?
		if CInt(rs1("TimeOn")) = 0 then
			objExcel.Cells(x,y).Font.Bold = True
		end if

		'Look up if player scored
		SQL = "SELECT COUNT(ID) AS Goals FROM tbl_gameevents WHERE GameID = " & rs1("GameID") & " AND Event = 1 AND PlayerID = " & rs("ID")
		set rs2 = conn.execute(SQL)
		if CInt(rs2("Goals")) > 0 then
			objExcel.Cells(x,y).Interior.Color = RGB(200,160,35)
			objExcel.Cells(1,y).Interior.Color = RGB(200,160,35)
		end if
		
		rs1.movenext
	loop
	
	y = y + 1
	rs.movenext
loop

objExcel.ActiveWorkbook.SaveAs(strWorkingDir & strOutputFile)

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
wscript.echo("Finished!")

' #############################################################
' #
' # Subroutines
' # 

Sub LogText(strMessage)
	LogFile.WriteLine(now() & ":" & vbTab & strMessage)
End Sub