<% @language = "VBSCRIPT" codePage = "65001" %>
<% Option Explicit %>
<%
  '**** Declare all variables, just like with JS. It is better to have the your DIMs at the top of the page ****
  
  DIM showId
  DIM dConn
  DIM processRS, processSP
  DIM dataArray
%>

<%
  dConn = "dsn=Lando;uid=reports;pwd="
  showId = "toms"

  SET processSP = Server.CreateObject("ADODB.Command")
  SET processRS = Server.CreateObject("ADODB.Recordset")
    WITH processSP
      .ActiveConnection = dConn 
      .CommandText = "spTS_ShowInfo_New" 
      .Parameters.Append .CreateParameter("@ShowID", 200, 1, 25, left(showid,4)) 
      .CommandType = 4 
      .CommandTimeout = 7 
      .Prepared = true
      SET processRS = .Execute
    END WITH
  SET processSP = Nothing

  dataArray = processRS.GetRows()


  'Just like any other Array, starts at 0 for the first record
  response.write(dataArray(5, 0))

  processRS.Close 
  SET processRS = Nothing 
%>
