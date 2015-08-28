<%
  SET ProcessSP = Server.CreateObject("ADODB.Command")
  SET ProcessRS = Server.CreateObject("ADODB.Recordset")
    WITH ProcessSP
      .ActiveConnection = Dconn 'Connection to the server. This is a variable that is set in the config file.
      .CommandText = "spTS_ShowInfo_New" 'Stored Proc Name
      .Parameters.Append .CreateParameter("@ShowID", 200, 1, 25, left(showid,4)) 'Parameters for stored proc (VARCHAR)
      .Parameters.Append .CreateParameter("@ShowID", 3, 1, , ) 'Parameters for stored proc (INT)
      .CommandType = 4 'Tells the server it is a Stored Proc
      .CommandTimeout = 7 'Resets the timeout to be shorter
      .Prepared = true
      SET ProcessRS = .Execute
    END WITH
  SET ProcessSP = Nothing
  
    if not ProcessRS.EOF then 'This is say, if there is Data than show
%>

<%
    DO WHILE NOT ProcessRS.EOF 'This is the LOOP
%>
  

<%
    ProcessRS.MoveNext() 'VERY IMPORTANT, tells the loop to go to the next record
    LOOP
%>

<%
    end if
  ProcessRS.Close 'Closes the recordSet
  SET ProcessRS = Nothing 'Sets the RecordSet to nothing
%>

