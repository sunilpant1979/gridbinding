<%@Language=JavaScript%>
<!--#include file=misc.asp-->
<!--#include file=b.dropdown.asp-->
<!--#include file=b.grid.asp-->

<%
	Response.Expires = 0;
	Response.Buffer = true;
	PageHeader("Territories");

	Out("<form method=post>");
	Out("<a href=\"default.asp\">Top</a> | ");
	Out("<a href=\"territories.asp\">Territories</a><hr>");
	
	var Conn = CreateConnection();
	var rsRegions = Conn.Execute("Region");

	var rs = Server.CreateObject("ADODB.Recordset");
	rs.Open("Territories",Conn,adOpenStatic,adLockOptimistic,adCmdTable);
	
	var Grid = new BGrid(rs);
	Grid.SetLookup("RegionID",rsRegions,"RegionID","RegionDescription");
	Grid.Process();
	Grid.Display();
	Grid = null;

	rsRegions.Close();
	rsRegions = null;

	Conn.Close();
	Conn = null;

	Out("</form>");

	PageFooter(false);
%>
