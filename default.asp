<%@language=javascript%>
<!--#include file=misc.asp-->

<%
	PageHeader("dblib");
	
	Out("<h3>Tables</h3>");
	Out("<ul>");

  	var db = CreateConnection();
  	var rs = db.OpenSchema(adSchemaTables);

  	while(!rs.EOF) {
  		if(rs.Fields("TABLE_TYPE").Value == "TABLE")
      		Out("<li><a href=\"edittable.asp?table=" + rs.Fields("TABLE_NAME").Value + "\">" + rs.Fields("TABLE_NAME").Value + "</a>")
    	rs.MoveNext
  	}

	rs.Close();
  	rs = null;
  	db.Close();
  	db = null;
	Out("</ul>");

	Out("<h3>Customized Tables</h3>");
	Out("<ul>");
	Out("<li><a href=\"territories.asp\">Territories</a>");
	Out("<li><a href=\"employeeterritories.asp\">Employee Territories</a>");
	Out("</ul>");

	PageFooter();
%>
