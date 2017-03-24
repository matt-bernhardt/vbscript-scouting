Option Explicit

Dim oFSO
Dim LogFile 			'logfile object
Dim conn 			'db connection
Dim cst				'db connection string
Dim rs, rs1, rs2, rs3, rs4		'recordset object
Dim SQL 			'SQL query statement
Dim x,y,z			'counter variables
Dim strLogFile, strInputFile	'Files created by this script
Dim strWorkingDir		'Directory with worksheet files
Dim strOutputFile

Dim intYear
Dim intTeamID
Dim intGameCount
Dim intMinuteCount
Dim intComboID
Dim arrPlayers()
Dim arrGames()
Dim intOn
Dim intOff

Dim intOn1, intOn2
Dim intOff1, intOff2

Dim intThisNone
Dim intThisOne
Dim intThisTwo
Dim intThisTogether

Dim intNone
Dim intOne
Dim intTwo
Dim intTogether

Dim intAplus
Dim intAminus
Dim intBplus
Dim intBminus
Dim intABplus
Dim intABminus
Dim intNplus
Dim intNminus

Dim boolA
Dim boolB

Dim strDescription

' #############################################################
' #
' # Define Objects
' # 

'define universal variables
intYear = 2017
intTeamID = 11

strWorkingDir = "C:\Users\mattj_000\Box Sync\Soccer\python\"
strLogFile = strWorkingDir & "logs\compile_combos_" & intYear & "_" & intTeamID & ".txt"

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

' #############################################################
' #
' # Start work
' # 

' Get list of players to appear this year
SQL = "SELECT m.PlayerID, CONCAT(FirstName,' ',LastName) AS PlayerName " & _
"FROM tbl_gameminutes m " & _
"INNER JOIN tbl_games g ON m.GameID = g.ID " & _
"INNER JOIN tbl_players p ON m.PlayerID = p.ID " & _
"INNER JOIN lkp_matchtypes t on g.MatchTypeID = t.ID " & _
"WHERE YEAR(MatchTime) = " & intYear & " AND m.TeamID = " & intTeamID & " AND TimeOff > 0 AND g.MatchTypeID = 21 " & _ 
"GROUP BY m.PlayerID " & _
"ORDER BY m.PlayerID ASC"
LogText "List of Players: " & SQL
set rs = conn.execute(SQL)

' Get list of goals
SQL = "SELECT g.ID AS GameID, g.MatchTime, e.MinuteID, e.TeamID, e.Event, e.PlayerID " & _
"FROM tbl_gameevents e " & _
"INNER JOIN tbl_games g ON e.GameID = g.ID " & _
"WHERE (Event = 1 OR Event = 6) AND e.GameID IN (SELECT DISTINCT g.ID FROM tbl_games g INNER JOIN lkp_matchtypes t ON g.MatchTypeID = t.ID WHERE g.MatchTypeID = 21 AND YEAR(MatchTime) = " & intYear & " AND (HTeamID = " & intTeamID & " OR ATeamID = " & intTeamID & ")) " & _
"ORDER BY g.MatchTime ASC, e.MinuteID"
LogText "List of Goals: " & SQL
set rs4 = conn.execute(SQL)

x = 1
do while not rs.eof
	
	ReDim preserve arrPlayers(x)
	arrPlayers(x) = rs("PlayerID")
	LogText x & ": " & arrPlayers(x)

	x = x + 1
	rs.movenext
loop

' Get list of games
SQL = "SELECT m.GameID, MAX(m.TimeOff) AS GameLength " & _
"FROM tbl_gameminutes m " & _
"INNER JOIN tbl_games g ON m.GameID = g.ID " & _
"INNER JOIN lkp_matchtypes t ON g.MatchTypeID = t.ID " & _
"WHERE t.Official = 1 AND YEAR(MatchTime) = " & intYear & " AND (m.TeamID = " & intTeamID & ") AND g.MatchTypeID = 21 " & _
"GROUP BY m.GameID " & _
"ORDER BY g.MatchTime ASC"
LogText "List of Games: " & SQL
set rs2 = conn.execute(SQL)

intMinuteCount = 0
x = 1
do while not rs2.eof

	ReDim preserve arrGames(x)
	arrGames(x) = rs2("GameID")

	intMinuteCount = intMinuteCount + rs2("GameLength")

	x = x + 1
	rs2.movenext
loop
intGameCount = x
LogText "Team has played " & intGameCount & " games, with " & intMinuteCount & " minutes"

x = 1
LogText "arrPlayers holds " & UBound(arrPlayers) & " players"

' Loop through players array twice, building comparisons
' This loop only concerns minutes played - goals are handled next
for x = 1 to UBound(arrPlayers)
	for y = 1 to x-1

		LogText x & "(" & arrPlayers(x) & ") , " & y & "(" & arrPlayers(y) & ")"

		intOne = 0
		intTwo = 0
		intNone = 0
		intTogether = 0

		' Loop through games array for this pair of players
		for z = 1 to UBound(arrGames)

			' Get minutes for Player 1
			intOn1 = 0
			intOff1 = 0
			SQL = "SELECT TimeOn, TimeOff " & _
			"FROM tbl_gameminutes m " & _
			"INNER JOIN tbl_games g ON m.GameID = g.ID " & _
			"WHERE GameID = " & arrGames(z) & " AND PlayerID = " & arrPlayers(x) & " AND g.MatchTypeID = 21"
			set rs2 = conn.execute(SQL)
			if not (rs2.bof and rs2.eof) then
				intOn1 = rs2("TimeOn")
				intOff1 = rs2("TimeOff")
			end if

			' Get minutes for Player 2
			intOn2 = 0
			intOff2 = 0
			SQL = "SELECT TimeOn, TimeOff " & _
			"FROM tbl_gameminutes m " & _
			"INNER JOIN tbl_games g ON m.GameID = g.ID " & _
			"WHERE GameID = " & arrGames(z) & " AND PlayerID = " & arrPlayers(y) & " AND g.MatchTypeID = 21"
			set rs2 = conn.execute(SQL)
			if not (rs2.bof and rs2.eof) then
				intOn2 = rs2("TimeOn")
				intOff2 = rs2("TimeOff")
			end if

			intThisOne = 0
			intThisTwo = 0
			intThisNone = 0
			intThisTogether = 0

			LogText vbTab & arrGames(z) & vbTab & arrPlayers(x) & ": " & intOn1 & " - " & intOff1 & vbTab & vbTab & arrPlayers(y) & ": " & intOn2 & " - " & intOff2

			' Break down available 90 minutes into four categories
			if intOn1 = intOn2 then ' On together
				intThisNone = intOn1

				if intOff1 = intOff2 then
					intThisTogether = intOff1 - intOn1
					intThisOne = 0
					intThisTwo = 0
					intThisNone = intThisNone + (90 - intOff1) 'What about games longer than 90 minutes?
					if intThisNone < 0 then LogText "### Look at game length"

				elseif intOff1 < intOff2 then
					intThisTogether = intOff1 - intOn1
					intThisOne = 0
					intThisTwo = intOff2 - intOff1
					intThisNone = 90 - intOff2
					if intThisNone < 0 then LogText "### Look at game length"

				elseif intOff1 > intOff2 then
					intThisTogether = intOff2 - intOn1
					intThisOne = intOff1 - intOff2
					intThisTwo = 0
					intThisNone = 90 - intOff1
					if intThisNone < 0 then LogText "### Look at game length"

				else
					LogText "### Look at intOff1 and intOff2"

				end if

			elseif intOn1 < intOn2 then ' P1 on first
				intThisNone = intOn1
				if intOff1 <= intOn2 then ' Check that they appeared together at all
					intThisTogether = 0
					intThisOne = intOff1 - intOn1
					intThisNone = intThisNone + (intOn2 - intOff1)
					intThisTwo = intOff2 - intOn2
					intThisNone = intThisNone + (90 - intOff2)
					if (90-intOff2) < 0 then LogText "### Look at time off"

				else ' they did appear together
					intThisOne = intOn2 - intOn1

					if intOff1 = intOff2 then
						intThisTogether = intOff1 - intOn2
						intThisTwo = 0
						intThisNone = intThisNone + (90 - intOff1)
						if (90-intOff1) < 0 then LogText "### Look at time off"

					elseif intOff1 < intOff2 then
						intThisTogether = intOff1 - intOn2
						intThisTwo = intOff2 - intOff1
						intThisNone = intThisNone + (90 - intOff2)
						if (90-intOff2) < 0 then LogText "### Look at time off"

					elseif intOff1 > intOff2 then
						intThisTogether = intOff2 - intOn2
						intThisOne = intOff1 - intOff2
						intThisNone = intThisNone + (90 - intOff1)
						if (90-intOff1) < 0 then LogText "### Look at time off"

					else
						LogText "### Look at intOff1 and intOff2"

					end if

				end if

			elseif intOn1 > intOn2 then ' P2 on first
				intThisNone = intOn2
				if intOff2 <= intOn1 then ' Check that they appeared together at all
					intThisTogether = 0
					intThisTwo = intOff2 - intOn2
					intThisNone = intThisNone + (intOn1 - intOff2)
					intThisOne = intOff1 - intOn1
					intThisNone = intThisNone + (90 - intOff1)
					if (90-intOff1) < 0 then LogText "### Look at time off"

				else ' they did appear together
					intThisTwo = intOn1 - intOn2

					if intOff1 = intOff2 then
						intThisTogether = intOff1 - intOn1
						intThisOne = 0
						intThisNone = intThisNone + (90 - intOff1)
						if (90-intOff1) < 0 then LogText "### Look at time off"

					elseif intOff1 < intOff2 then
						intThisTogether = intOff1 - intOn1
						intThisOne = 0
						intThisTwo = intThisTwo + (intOff2 - intOff1)
						intThisNone = intThisNone + (90 - intOff2)
						if (90-intOff2) < 0 then LogText "### Look at time off"

					elseif intOff1 > intOff2 then
						intThisTogether = intOff2 - intOn1
						intThisOne = intOff1 - intOff2
						intThisNone = intThisNone + (90 - intOff1)
						if (90-intOff1) < 0 then LogText "### Look at time off"

					else
						LogText "### Look at intOff1 and intOff2"

					end if

				end if

			else
				LogText "### Look at intOn1 and intOn2"
			end if

			' Status check: one two none together
			LogText vbTab & intThisOne & vbTab & intThisTwo & vbTab & intThisNone & vbTab & intThisTogether
			if intThisNone + intThisOne + intThisTwo + intThisTogether <> 90 then 
				LogText "### Something doesn't add up"
			end if

			intOne = intOne + intThisOne
			intTwo = intTwo + intThisTwo
			intNone = intNone + intThisNone
			intTogether = intTogether + intThisTogether

		next

		if (intOne+intTwo+intNone+intTogether) < intMinuteCount then intNone = intMinuteCount - (intOne+intTwo+intTogether)
		' one two none together
		LogText vbTab & intOne & vbTab & intTwo & vbTab & IntNone & vbTab & intTogether & vbTab & vbTab & (intOne+intTwo+intNone+intTogether)

		' Now we store the four combination states
		' Get comboIDs
		' One
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_0," & arrPlayers(x) & "_1'"
		set rs3 = conn.execute(SQL)
		if rs3.bof and rs3.eof then
			LogText "One create"
			RegisterCombo arrPlayers(y),0,arrPlayers(x),1,intOne
		else
			LogText "One exists"
			UpdateCombo rs3("ComboID"),intOne,intYear
		end if

		' Two
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_1," & arrPlayers(x) & "_0'"
		set rs3 = conn.execute(SQL)
		if rs3.bof and rs3.eof then
			LogText "Two create"
			RegisterCombo arrPlayers(y),1,arrPlayers(x),0,intTwo
		else
			LogText "Two exists"
			UpdateCombo rs3("ComboID"),intTwo,intYear
		end if

		' None
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_1," & arrPlayers(x) & "_1'"
		set rs3 = conn.execute(SQL)
		if rs3.bof and rs3.eof then
			LogText "None create"
			RegisterCombo arrPlayers(y),1,arrPlayers(x),1,intNone
		else
			LogText "None exists"
			UpdateCombo rs3("ComboID"),intNone,intYear
		end if

		' Together
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_0," & arrPlayers(x) & "_0'"
		set rs3 = conn.execute(SQL)
		if rs3.bof and rs3.eof then
			LogText "Together create"
			RegisterCombo arrPlayers(y),0,arrPlayers(x),0,intTogether
		else
			LogText "Together exists"
			UpdateCombo rs3("ComboID"),intTogether,intYear
		end if

		LogText ""

		' Loop through goals, building up plus/minus figures
		' Reset
		intAplus = 0
		intAminus = 0
		intBplus = 0
		intBminus = 0
		intABplus = 0
		intABminus = 0
		intNplus = 0
		intNminus = 0
		rs4.movefirst
		do while not rs4.eof

			LogText vbTab & rs4("GameID")

			'Look up time on / time off for this game
			SQL = "SELECT TimeOn, TimeOff FROM tbl_gameminutes WHERE GameID = " & rs4("GameID") & " AND PlayerID = " & arrPlayers(x)
			set rs = conn.execute(SQL)
			if rs.bof and rs.eof then
				intOn1 = 0
				intOff1 = 0
			else 
				intOn1 = rs("TimeOn")
				intOff1 = rs("TimeOff")
			end if

			SQL = "SELECT TimeOn, TimeOff FROM tbl_gameminutes WHERE GameID = " & rs4("GameID") & " AND PlayerID = " & arrPlayers(y)
			set rs = conn.execute(SQL)
			if rs.bof and rs.eof then
				intOn2 = 0
				intOff2 = 0
			else 
				intOn2 = rs("TimeOn")
				intOff2 = rs("TimeOff")
			end if

			boolA = false
			boolB = false

			if intOn1 < rs4("MinuteID") AND intOff1 >= rs4("MinuteID") then
				boolA = true
			end if

			if intOn2 < rs4("MinuteID") AND intOff2 >= rs4("MinuteID") then
				boolB = true
			end if

			'Goal for or against?
			if (rs4("TeamID") = intTeamID AND rs4("Event") = 1) OR (rs4("TeamID") <> intTeamID AND rs4("Event") = 6) then
				' Goal for

				if boolA = true then
					if boolB = true then
						intABplus = intABplus + 1
					else
						intAplus = intAplus + 1
					end if
				else
					if boolB = true then
						intBplus = intBplus + 1
					else
						intNplus = intNplus + 1
					end if
				end if

			else
				' Goal against

				if boolA = true then
					if boolB = true then
						intABminus = intABminus + 1
					else
						intAminus = intAminus + 1
					end if
				else
					if boolB = true then
						intBminus = intBminus + 1
					else
						intNminus = intNminus + 1
					end if
				end if
				
			end if


			rs4.movenext
		loop

		' Store plus/minus for each of the four combinations

		LogText vbTab & "Plus/Minus"
		' One
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_0," & arrPlayers(x) & "_1'"
		set rs3 = conn.execute(SQL)
		SQL = "UPDATE tbl_combos_stats SET Plus = " & intAPlus & ", Minus = " & intAMinus & " WHERE ComboID = " & rs3("ComboID") & " AND Year = " & intYear & " AND TeamID = 11 AND CompetitionID = 21"

		LogText vbTab & intAPlus & vbTab & intAMinus
		conn.execute(SQL)

		' Two
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_1," & arrPlayers(x) & "_0'"
		set rs3 = conn.execute(SQL)
		SQL = "UPDATE tbl_combos_stats SET Plus = " & intBPlus & ", Minus = " & intBMinus & " WHERE ComboID = " & rs3("ComboID") & " AND Year = " & intYear & " AND TeamID = 11 AND CompetitionID = 21"
		LogText vbTab & intBPlus & vbTab & intBMinus
		conn.execute(SQL)

		' None
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_1," & arrPlayers(x) & "_1'"
		set rs3 = conn.execute(SQL)
		SQL = "UPDATE tbl_combos_stats SET Plus = " & intNPlus & ", Minus = " & intNMinus & " WHERE ComboID = " & rs3("ComboID") & " AND Year = " & intYear & " AND TeamID = 11 AND CompetitionID = 21"
		LogText vbTab & intNPlus & vbTab & intNMinus
		conn.execute(SQL)

		' Together
		SQL = "SELECT l.ComboID, GROUP_CONCAT(PlayerID,'_',Exclude ORDER BY PlayerID) AS Players " & _
		"FROM lnk_players_combos l " & _
		"GROUP BY ComboID " & _
		"HAVING Players = '" & arrPlayers(y) & "_0," & arrPlayers(x) & "_0'"
		set rs3 = conn.execute(SQL)
		SQL = "UPDATE tbl_combos_stats SET Plus = " & intABPlus & ", Minus = " & intABMinus & " WHERE ComboID = " & rs3("ComboID") & " AND Year = " & intYear & " AND TeamID = 11 AND CompetitionID = 21"
		LogText vbTab & intABPlus & vbTab & intABMinus
		conn.execute(SQL)

		LogText ""

	next

	LogText ""
next


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

Sub RegisterCombo(one,oneX,two,twoX,minutes)

	' Register combination itself
	strDescription = one & "_" & oneX & "," & two & "_" & twoX
	SQL = "INSERT INTO tbl_combos (Description) VALUES ('" & strDescription & "')"
	LogText vbTab & vbTab & SQL
	conn.execute(SQL)

	' Recall created combination ID
	set rs = conn.execute("SELECT LAST_INSERT_ID() AS ComboID")
	intComboID = rs("ComboID")

	' Affiliate player/states with combo
	conn.execute("INSERT INTO lnk_players_combos (ComboID, PlayerID, Exclude) VALUES (" & intComboID & ", " & one & ", " & oneX & ")")
	conn.execute("INSERT INTO lnk_players_combos (ComboID, PlayerID, Exclude) VALUES (" & intComboID & ", " & two & ", " & twoX & ")")

	' Register combination with year
	conn.execute("INSERT INTO tbl_combos_stats (ComboID, Year, CompetitionID, TeamID, GP, Min, Plus, Minus) VALUES (" & intComboID & ", " & intYear & ", 21, 11, 0, " & minutes & ", 0, 0)")

End Sub

Sub UpdateCombo(comboID,minutes,year)
	Dim tempRS 

	SQL = "SELECT ID FROM tbl_combos_stats WHERE ComboID = " & comboID & " AND Year = " & year & " AND TeamID = 11 AND CompetitionID = 21"
	set tempRS = conn.execute(SQL)
	if tempRS.bof and tempRS.eof then
		SQL = "INSERT INTO tbl_combos_stats (ComboID, Year, TeamID, CompetitionID, Min, GP, Plus, Minus) VALUES (" & comboID & ", " & year & ", 11, 21, " & minutes & ", 0, 0, 0)"
	else 
		SQL = "UPDATE tbl_combos_stats SET Min = " & minutes & " WHERE ComboID = " & comboID & " AND Year = " & year & " AND TeamID = 11 AND CompetitionID = 21"
	end if
	conn.execute(SQL)
End Sub