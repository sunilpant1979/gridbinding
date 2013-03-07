<%@Language=JavaScript%>
<!--#include file=misc.asp-->
<!--#include file=b.dropdown.asp-->
<!--#include file=b.grid.asp-->

<%
	Response.Expires = 0;
	Response.Buffer = true;
	PageHeader("Employee Territories");

	Out("<form method=post>");
	Out("<a href=\"default.asp\">Top</a> | ");
	Out("<a href=\"employeeterritories.asp\">Employee Territories</a><hr>");
	
	var Conn = CreateConnection();
	var rsEmployees = Conn.Execute("select EmployeeID, LastName + ', ' + FirstName as DisplayName from Employees order by LastName, FirstName");
	var rsTerritories = Conn.Execute("Territories");
	rsTerritories.Sort = "TerritoryDescription";

	var rs = Server.CreateObject("ADODB.Recordset");
	rs.Open("EmployeeTerritories",Conn,adOpenStatic,adLockOptimistic,adCmdTable);
	
	var Grid = new BGrid(rs);
	Grid.SetLookup("EmployeeID",rsEmployees,"EmployeeID","DisplayName");
	Grid.SetLookup("TerritoryID",rsTerritories,"TerritoryID","TerritoryDescription");
	Grid.Process();
	Grid.Display();
	Grid = null;

	rsEmployees.Close();
	rsEmployees = null;
	rsTerritories.Close();
	rsTerritories = null;

	Conn.Close();
	Conn = null;

	Out("</form>");

	PageFooter(false);
%>
