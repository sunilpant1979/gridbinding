<%@Language=JavaScript%>
<!--#include file=misc.asp-->
<!--#include file=b.grid.asp-->
<!--#include file=b.dropdown.asp-->

<%
	Response.Expires = 0;
	Response.Buffer = true;
	if(Request.QueryString("table").Count!=1) Server.Transfer("default.asp");
	PageHeader("Table Editor",true);
	
	Out("<form method=post>");
	Out("<a href=\"default.asp\">Top</a> | ");
	Out("<a href=\"edittable.asp?table=" + Request.QueryString("table").Item + "\">" + Request.QueryString("table").Item + "</a>");
	Out("<hr>");

	var Conn = CreateConnection();
	var rs = Server.CreateObject("ADODB.Recordset");
	rs.Open(Request.QueryString("table").Item,Conn,adOpenStatic,adLockOptimistic,adCmdTable);
	
	var Grid = new BGrid(rs);
	Grid.SetOption("truncate",25);
	Grid.Process();
	Grid.Display();
	Grid = null;

	rs.Close();
	rs = null;

	Conn.Close();
	Conn = null;
	Out("</form>");

	PageFooter(false);
%>
