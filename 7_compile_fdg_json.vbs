Option Explicit

Dim strWorkingDir
Dim strLogFile
Dim strOutputFile
Dim oFSO
Dim LogFile
Dim OutputFile
Dim cst
Dim conn
Dim SQL
Dim rs
Dim rs1
Dim rs2

Dim x,y
Dim intYear
Dim intTeamID
Dim intPlayerCount
Dim intMinutes

' ###############################################
' #
' # Define Objects
' #

intTeamID = 11
intYear = 2017

strWorkingDir = "C:\Users\mattj_000\Box Sync\Soccer\python\"
strLogFile = "compile-fdg-json.txt"
strOutputFile = "crew" & intYear & ".json"

' File System Object
Set oFSO = CreateObject("Scripting.FileSystemObject")

' Open log file
Set LogFile = oFSO.CreateTextFile(strWorkingDir & "logs\" & strLogFile,true)
Set OutputFile = oFSO.CreateTextFile(strWorkingDir & "output\" & strOutputFile,true)
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

' ###############################################
' #
' # Start Work
' #

OutputFile.WriteLine "{""nodes"":["

' 1-Find scope of file, build top half of JSON file

SQL = "SELECT p.ID, CONCAT(FirstName,' ',LastName) AS PlayerName " & _
	"FROM scouting.tbl_players p " & _
	"LEFT OUTER JOIN scouting.tbl_gameminutes m ON p.ID = m.PlayerID " & _
	"LEFT OUTER JOIN scouting.tbl_games g ON m.GameID = g.ID " & _
	"WHERE g.MatchTime < NOW() AND YEAR(g.MatchTime) = " & intYear & " AND m.TeamID = " & intTeamID & " " & _
	"GROUP BY p.ID"
LogText SQL	
set rs = conn.execute(SQL)
intPlayerCount = 0
do while not rs.eof
	OutputFile.WriteLine "{""name"":""" & trim(rs("PlayerName")) & """,""group"":2},"
	intPlayerCount = intPlayerCount + 1
	rs.movenext
loop
rs.movefirst

SQL = "SELECT g.ID, DATE_FORMAT(MatchTime,'%c/%e/%y') AS Date, CONCAT(IF(HTeamID=" & intTeamID & ",'vs ','@ '), IF(HTeamID=" & intTeamID & ",a.teamname,h.teamname)) AS Opponent " & _
	"FROM tbl_games g " & _
	"LEFT OUTER JOIN tbl_teams a ON g.ATeamID = a.ID " & _
	"LEFT OUTER JOIN tbl_teams h ON g.HTeamID = h.ID " & _
	"WHERE YEAR(MatchTime) = " & intYear & " AND (HTeamID = " & intTeamID & " OR ATeamID = " & intTeamID & ") AND MatchTime < NOW()"
LogText SQL	
set rs1 = conn.execute(SQL)
do while not rs1.eof
	OutputFile.WriteLine "{""name"":""" & rs1("Date") & " " & rs1("Opponent") & """,""group"":1},"
	rs1.movenext
loop
rs1.movefirst

OutputFile.WriteLine "],""links"":["

x=0
do while not rs.eof

	y = 0
	do while not rs1.eof

		SQL = "SELECT TimeOff-TimeOn AS Minutes FROM tbl_gameminutes WHERE GameID = " & rs1("ID") & " AND PlayerID = " & rs("ID") & " AND TimeOff > 0"
		LogText SQL
		set rs2 = conn.execute(SQL)
		if not (rs2.bof and rs2.eof) then
			intMinutes = CInt(rs2("Minutes"))
			OutputFile.WriteLine "{""source"":" & x & ", ""target"":" & (intPlayerCount + y) & ", ""value"":" & intMinutes/10 & "},"
		end if

		y = y + 1
		rs1.movenext
	loop

	rs1.movefirst
	LogText ""

	x = x + 1
	rs.movenext
loop

OutputFile.WriteLine "]}"

' ###############################################
' #
' # Close Objects, wrap up
' #

'close db connection
conn.close
set conn = Nothing
LogText "Closed db connection"

LogText "Finished"
Set OutputFile = Nothing
Set LogFile = Nothing
Set oFSO = Nothing

wscript.echo("Finished!")

' ###############################################
' #
' # Subroutines and Functions
' #

Sub LogText(strMessage)
	LogFile.WriteLine(now() & vbTab & strMessage)
End Sub