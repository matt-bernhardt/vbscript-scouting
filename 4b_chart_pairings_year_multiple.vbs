Option Explicit

Dim strWorkingDir
Dim oFSO
Dim LogFile 			'logfile object
Dim conn 			'db connection
Dim cst				'db connection string
Dim rs				'recordset object
Dim SQL 			'SQL query statement
Dim x,y,z			'counter variables
Dim strLogFile			'Files created by this script
Dim objExcel			'Excel object
Dim strOutputFile		'output Excel file
Dim strJSONfile, JSONfile	'output JSON file for use by d3.js

Dim intGameID
Dim intPlayerID, intPlayerCount
Dim intTimeOn, intTimeOff
Dim intYear

Dim players			'recordset

Dim strTeam, intTeamID

Dim strOrderBy

' #############################################################
' #
' # Define Objects
' # 

'define universal variables
intYear = 2017
intTeamID = 11

strWorkingDir = "C:\Users\mattj_000\Box Sync\Soccer\pairings\" & intYear & "\"

' File System Object
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Open log file
strLogFile = "chart_pairings_multiple_" & intYear & ".txt"
Set LogFile = oFSO.CreateTextFile(strLogFile,true)
LogText "Started"

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

BuildPlot 11

BuildPlot 12

BuildPlot 13

BuildPlot 14

BuildPlot 15

BuildPlot 16

BuildPlot 17

BuildPlot 18

BuildPlot 19

BuildPlot 20

BuildPlot 42

BuildPlot 43

BuildPlot 44

BuildPlot 45

BuildPlot 340

BuildPlot 427

BuildPlot 463

BuildPlot 479

BuildPlot 506

BuildPlot 521

BuildPlot 547

BuildPlot 599

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

Sub BuildPlot(intTeamID)
	LogText "_-=-_"
	LogText "Starting Work"

	strOutputFile = "chart_pairings_" & intTeamID & "_" & intYear & ".xlsx"
	strJSONfile = "chart_pairings_" & intTeamID & "_" & intYear & ".json"

	' Open JSON file
	Set JSONfile = oFSO.CreateTextFile(strWorkingDir & strJSONfile,true)
	JSONfile.Write("{""nodes"":[")

	' Open Excel file
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = false
	objExcel.Workbooks.Add
	LogText "Excel initialized"

	intPlayerCount = 1
	x = 0
	y = 0

	'get list of players who appeared that year
	SQL = "SELECT PlayerID, concat(FirstName,' ',LastName) AS PlayerName, year(min(MatchTime))-1995 AS Class, SUM(TimeOff-TimeOn) AS Minutes " & _
		"FROM tbl_gameminutes " & _
		"INNER JOIN tbl_games ON tbl_gameminutes.GameID = tbl_games.ID " & _
		"INNER JOIN tbl_players ON tbl_gameminutes.PlayerID = tbl_players.ID " & _
		"WHERE TeamID = " & intTeamID & " AND Year(MatchTime) = " & intYear & " " & _
		"GROUP BY PlayerID " & _
		"HAVING Max(TimeOff) > 0 " & _
		"ORDER BY SUM(TimeOff-TimeOn) DESC"
	LogText SQL	
	set players = conn.execute(SQL)

	x = 3
	players.movefirst
	do while not players.eof
		objExcel.Cells(x,1).Value = players("PlayerName")
		objExcel.Cells(x,2).Value = players("Minutes")
		objExcel.Cells(2,x).Value = players("PlayerID")
		objExcel.Cells(1,x).Value = players("PlayerName")
		objExcel.Cells(1,x).Orientation = 90
		
		' {"name":"Myriel","group":1},
		JSONfile.Write("{""name"":""" & players("PlayerName") & """,""group"":" & players("Class") & "},")
		
		' Build order clause for next SQL statement
		if x > 3 then
			strOrderBy = strOrderBy & ","
		end if
		strOrderBy = strOrderBy & players("PlayerID")

		x = x + 1
		intPlayerCount = intPlayerCount + 1
		players.movenext
	loop

	LogText intPlayerCount & " players all time"
	LogText strOrderBy

	JSONfile.Write("],""links"":[")

	x = 3
	players.movefirst
	do while not players.eof
		LogText ""
		LogText ""
		LogText "Looking at teammates of " & players("PlayerID")

		'pairings are built based on a single player, intPlayerID

		'get gameID, time on and time off for a single player in a single year
		'join tbl_games to itself to find other players from that team who played in that game 
		'filter teammate timeon and timeoff via if() to find their pairing info

		'competition information isn't relevant right now, but may be an option later

		SQL = 	"SELECT target.GameID, target.TimeOn, target.TimeOff, " & _
			"  teammate.PlayerID, teammate.TimeOn, teammate.TimeOff, " & _
			"  if(teammate.TimeOn < target.TimeOn, target.TimeOn, teammate.TimeOn) AS AdjustedOn, " & _
			"  if(teammate.TimeOff > target.TimeOff, target.TimeOff, teammate.TimeOff) AS AdjustedOff, " & _
			"  sum(if(teammate.TimeOff > target.TimeOff, target.TimeOff, teammate.TimeOff) - if(teammate.TimeOn < target.TimeOn, target.TimeOn, teammate.TimeOn)) AS AdjustedMin " & _
			"FROM scouting.tbl_gameminutes target " & _
			"INNER JOIN scouting.tbl_gameminutes teammate ON target.GameID = teammate.GameID " & _
			"INNER JOIN tbl_games ON target.GameID = tbl_games.ID " & _
			"WHERE Year(MatchTime) = " & intYear & " AND target.PlayerID = " & players("PlayerID") & " AND target.TimeOff > 0 " & _
			"  AND teammate.TimeOff > 0 AND target.PlayerID <> teammate.PlayerID " & _
			"  AND target.TeamID = teammate.TeamID " & _
			"  AND teammate.TimeOff > target.TimeOn " & _
			"  AND teammate.TimeOn < target.TimeOff " & _
			"  AND target.TeamID = " & intTeamID & " " & _
			"GROUP BY teammate.PlayerID " & _
			"ORDER BY FIND_IN_SET(teammate.PlayerID,'" & strOrderBy & "')"
		LogText SQL
		set rs = conn.execute(SQL)

		z = 0
		y = 3
		rs.movefirst
		do while not rs.eof
			
			do while objExcel.Cells(2,y).Value <> rs("PlayerID")
				LogText "skipping cell with " & objExcel.Cells(2,y).Value & " not " & rs("PlayerID")
				if x = y then objExcel.Cells(x,y).Interior.Color = RGB(192,192,192)
				y = y + 1
			loop

			LogText rs("PlayerID") & vbTab & rs("AdjustedMin")
			objExcel.Cells(x,y).Value = rs("AdjustedMin")
			
			if x-3 > y-3 then
				' {"source":1,"target":0,"value":1},
				JSONfile.Write("{""source"":" & x-3 & ",""target"":" & y-3 & ",""value"":" & CInt(rs("AdjustedMin"))/100 & "},")
			end if
			
			z = z + 1
			y = y + 1
			rs.movenext
		loop

		x = x + 1
		players.movenext
	loop

	objExcel.ActiveWorkbook.SaveAs(strWorkingDir & strOutputFile)
	objExcel.Quit

	Set JSONfile = Nothing
End Sub

Function StoppageTimeAdjust(intTime)
	if instr(intTime,"+") > 0 then
		intTime = CInt(replace(intTime,"+",""))
		if intTime > 90 then
			intTime = 89
		else
			intTime = 45
		end if
	elseif intTime = 46 then
		intTime = 45
	end if
	
	StoppageTimeAdjust = intTime
End Function