Option Explicit

Dim oFSO
Dim LogFile 			'logfile object
Dim conn 			'db connection
Dim cst				'db connection string
Dim rs, rs1, datelist		'recordset object
Dim players
Dim SQL 			'SQL query statement
Dim x,y				'counter variables
Dim strLogFile			'Files created by this script
Dim strPlayerName, strFirstName, strLastName
Dim intPlayerID, intGameID
Dim intPlus, intMinus

' #############################################################
' #
' # Define Objects
' # 

'define universal variables
strLogFile = "plus_minus.txt"

' File System Object
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Open log file
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

intPlus = 0
intMinus = 0

' #############################################################
' #
' # Start work
' # 

LogText "Here we go..."

' Get list of players
SQL = "SELECT ID, concat(FirstName,' ',LastName) AS PlayerName " & _
	"FROM scouting.tbl_players " & _ 
	"ORDER BY LastName, FirstName"
set players = conn.execute(SQL)

' Loop over each player in the list...
players.movefirst
do while not players.eof
	LogText "Player " & players("ID") & ": " & players("PlayerName")

	' Get list of games in which that players appeared
	SQL = "SELECT GameID, TeamID, TimeOn, TimeOff " & _
	"FROM scouting.tbl_gameminutes " & _
	"WHERE PlayerID = " & players("ID") & " " & _
	"ORDER BY GameID ASC"
	set rs = conn.execute(SQL)
	
	' Loop over each game in which the player appeared...
	if rs.bof and rs.eof then
		LogText vbTab & "No Appearances"
	else
		rs.movefirst
		do while not rs.eof
			LogText vbTab & "Game " & rs("GameID") & ": On at " & rs("TimeOn") & ", off at " & rs("TimeOff") & " for " & rs("TeamID")
			
			' Make sure the player actually appeared (TimeOff is not zero - unused subs are stored for some games)
			if rs("TimeOff") > 0 then

				' Get list of goals that were scored in this game
				SQL = "SELECT MinuteID, TeamID, lkp_gameevents.Event " & _
				"FROM scouting.tbl_gameevents " & _
				"INNER JOIN scouting.lkp_gameevents ON tbl_gameevents.Event = lkp_gameevents.ID " & _
				"WHERE GameID = " & rs("GameID") & " AND (lkp_gameevents.Event = 'Goal' OR lkp_gameevents.Event = 'Own Goal') " & _
				"ORDER BY MinuteID ASC"
				set rs1 = conn.execute(SQL)
			
				' If there were goals scored...
				if rs1.bof and rs1.eof then
					LogText vbTab & vbTab & "No goals"
				else
			
					intPlus = 0
					intMinus = 0
				
					' Loop over list of goals, incrementing the number of goals scored and conceded while this player was on the field
					rs1.movefirst
					do while not rs1.eof
						if rs("TimeOn") < rs1("MinuteID") and rs("TimeOff") >= rs1("MinuteID") then
							LogText vbTab & vbTab & rs1("MinuteID") & ": " & rs1("Event") & " by " & rs1("TeamID")
							if rs("TeamID") = rs1("TeamID") then
								'Same team
								if rs1("Event") = "Goal" then
									'goal = good
									intPlus = intPlus + 1
								else 
									intMinus = intMinus + 1
								end if
							else
								'Opposite team
								if rs1("Event") = "Goal" then
									'goal = bad
									intMinus = intMinus + 1
								else
									intPlus = intPlus + 1
								end if
							end if
						else
							LogText vbTab & vbTab & "Not on field during minute " & rs1("MinuteID")
						end if					
	
						rs1.movenext
					loop
					
					' Summarize our findings for this player, game, team, and goal counts
					Summarize players("ID"),rs("GameID"),rs("TeamID"),intPlus,intMinus
				end if

			end if		
			
			LogText ""
			rs.movenext
		loop
	end if

	LogText ""
	players.movenext
loop

LogText "End of Loop"

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
'wscript.echo("Finished!")

' #############################################################
' #
' # Subroutines
' # 

Sub LogText(strMessage)
	LogFile.WriteLine(now() & ":" & vbTab & strMessage)
End Sub

Sub Summarize(intPlayerID,intGameID,intTeamID,intPlus,intMinus)
	Dim stats
	
	LogText vbTab & vbTab & "Summarizing: " & intPlus & " / " & intMinus
	
	' Check if this player/game/team stat line has been created
	SQL = "SELECT ID " & _
	"FROM scouting.tbl_gamestats " & _
	"WHERE PlayerID = " & intPlayerID & " AND GameID = " & intGameID & " AND TeamID = " & intTeamID
	set stats = conn.execute(SQL)
	
	if stats.bof and stats.eof then
		' Insert new record for this player/game/team 
		SQL = "INSERT INTO scouting.tbl_gamestats (PlayerID, GameID, TeamID, Plus, Minus) VALUES (" & _
		intPlayerID & ", " & intGameID & ", " & intTeamID & ", " & intPlus & ", " & intMinus & ")"
		conn.execute(SQL)
	else
		' Update record for this player/game/team
		SQL = "UPDATE scouting.tbl_gamestats " & _
		"SET Plus = " & intPlus & ", Minus = " & intMinus & " " & _
		"WHERE ID = " & stats("ID")
		conn.execute(SQL)
	end if
End Sub