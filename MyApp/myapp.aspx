<%@ Page Language="C#" Debug="true" %>
<%@ Import Namespace="System.Data"%>
<%@ Import Namespace="System.Data.SqlClient"%>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Web" %>
<%@ Import Namespace="System.Web.UI" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="Spire.DataExport" %>
<%@ Import Namespace="ClosedXML.Excel" %>

<script runat="server">

string searchtext = String.Empty ;
string mySQL = String.Empty ;
string connStr = @"Data Source=PC\MSSQLEXPRESS;Initial Catalog=mybase;Integrated Security=true";

void Page_Load(object sender, EventArgs e)
{
	if (!IsPostBack)
	{
        LogAction(this, "get", "");
        Session["SearchText"] = string.Empty;
        Session["SortedView"] = new DataView(BindGridView());
        GridView1.DataSource = Session["SortedView"]; 
        GridView1.DataBind();
        GridView1.Visible = true;
	}
}

private DataTable BindGridView()
{
    if (Session["SearchText"] == string.Empty)
    {
        mySQL = "SELECT a.aid,a.adate,b.bdate,c.cdate,a.anumber,b.bnumber,c.cnumber,a.avalue,b.bvalue,c.cvalue,a.atext + b.btext + c.ctext AS text " +
                "FROM atable AS A JOIN btable AS B ON a.aid=b.bid JOIN ctable AS C ON a.aid=c.cid";
    }
    else{
        mySQL = "SELECT  a.aid,a.adate,b.bdate,c.cdate,a.anumber,b.bnumber,c.cnumber,a.avalue,b.bvalue,c.cvalue,a.atext + b.btext + c.ctext AS text " +
                "FROM atable AS A JOIN btable AS B ON a.aid=b.bid JOIN ctable AS C ON a.aid=c.cid WHERE a.atext + b.btext + c.ctext Like @text";
    }
    SqlConnection cnn = new SqlConnection(connStr);
    SqlDataAdapter sda = new SqlDataAdapter(mySQL, cnn);
    sda.SelectCommand.Parameters.Add("@text", SqlDbType.VarChar).Value = Convert.ToString("%" + Session["SearchText"] + "%");
    DataSet ds = new DataSet();
    sda.Fill(ds);
    return ds.Tables[0];
}

private DataTable BindGridViewAll()
{
    mySQL = "SELECT a.aid,a.adate,b.bdate,c.cdate,a.anumber,b.bnumber,c.cnumber,a.avalue,b.bvalue,c.cvalue,a.atext + b.btext + c.ctext AS text " +
            "FROM atable AS A JOIN btable AS B ON a.aid=b.bid JOIN ctable AS C ON a.aid=c.cid";
    SqlConnection cnn = new SqlConnection(connStr);
    SqlDataAdapter sda = new SqlDataAdapter(mySQL, cnn);
    DataSet ds = new DataSet();
    sda.Fill(ds);
    return ds.Tables[0];
}

protected void GridView1_Sorting(object sender, GridViewSortEventArgs e)
{
    string sortingDirection = string.Empty;
    if (direction == SortDirection.Ascending)
    {
        direction = SortDirection.Descending;
        sortingDirection = "Desc";
    }
    else
    {
        direction = SortDirection.Ascending;
        sortingDirection = "Asc";
    }
    DataView sortedView = new DataView(BindGridView());
    sortedView.Sort = e.SortExpression + " " + sortingDirection;
    LogAction(this, "sort", e.SortExpression + "_" + sortingDirection);
    Session["SortedView"] = sortedView;
    GridView1.DataSource = sortedView;
    GridView1.DataBind();

    foreach (DataControlField col in GridView1.Columns)
    {
        GridView1.Columns[GridView1.Columns.IndexOf(col)].HeaderStyle.CssClass = "";
         
        if (col.SortExpression == e.SortExpression)
        {
            int index = GridView1.Columns.IndexOf(col);
            if (sortingDirection == "Asc")
            {
                GridView1.Columns[index].HeaderStyle.CssClass = "sortasc";
                //GridView1.Columns[index].HeaderText = "<i class='asc'></i>";
            }
            else
            {
                GridView1.Columns[index].HeaderStyle.CssClass = "sortdesc";
            }
        }
    }

}

protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
{
    GridView1.PageIndex = e.NewPageIndex;
    if (Session["SortedView"] != null)
    {
        GridView1.DataSource = Session["SortedView"];
        GridView1.DataBind();
    }
    else
    {
        GridView1.DataSource = BindGridView();
        GridView1.DataBind();
    }
}

public SortDirection direction
{
    get
    {
        if (ViewState["directionState"] == null)
        {
            ViewState["directionState"] = SortDirection.Ascending;
        }
        return (SortDirection)ViewState["directionState"];
    }
    set
    {
        ViewState["directionState"] = value;
    }
}


void Run(object o, EventArgs e)
{	
	searchtext = Request.Form["CmdSearch"]; 
	CmdSearch.Text = "";
	if(searchtext.ToString().Trim() == "")
	{
        Session["SearchText"] = string.Empty;
        Session["SortedView"] = new DataView(BindGridView());
        GridView1.DataSource = Session["SortedView"];
        GridView1.DataBind();
        GridView1.Visible = true;
        LogAction(this, "get", "");
	}
    else
	{
        Session["SearchText"] = searchtext.ToString().Trim();
        Session["SortedView"] = new DataView(BindGridView());
        GridView1.DataSource = Session["SortedView"]; //((DataView)Session["SortedView"]).ToTable();
        GridView1.DataBind();
        
        if (GridView1.Rows.Count == 0)
        {
            GridView1.Visible = false;
            pnResult.Visible = true;
            LogAction(this, "search", searchtext.ToString().Trim()+ ": not found");
        }
        else
        {
            GridView1.Visible = true;
            pnResult.Visible = false;
            LogAction(this, "search", searchtext.ToString().Trim() + ": found");
        }     
	}
}

protected void ExportToXLS(object sender, EventArgs e)
{
    Spire.DataExport.XLS.WorkSheet workSheet1 = new Spire.DataExport.XLS.WorkSheet();
    Spire.DataExport.XLS.CellExport cellExport1 = new Spire.DataExport.XLS.CellExport();
    workSheet1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
    workSheet1.DataTable = ((DataView)Session["SortedView"]).ToTable();
    workSheet1.StartDataCol = ((System.Byte)(0));
    cellExport1.Sheets.Add(workSheet1);
    Response.AddHeader("Transfer-Encoding", "identity");
    cellExport1.SaveToHttpResponse("filtered_result.xls", Response);
    LogAction(this, "download", "filtered_xls");
}
protected void ExportToXLSX(object sender, EventArgs e)
{
    XLWorkbook wb = new XLWorkbook();
    wb.Worksheets.Add(((DataView)Session["SortedView"]).ToTable(), "Sheet1");
    Response.Clear();
    Response.Buffer = true;
    Response.Charset = "";
    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    Response.AddHeader("content-disposition", "attachment;filename=filtered_result.xlsx");
    MemoryStream MyMemoryStream = new MemoryStream();
    wb.SaveAs(MyMemoryStream);
    MyMemoryStream.WriteTo(Response.OutputStream);
    Response.Flush();
    LogAction(this, "download", "filtered_xlsx");
    Response.End();
    
}
protected void ExportToCSV(object sender, EventArgs e)
{
    Spire.DataExport.TXT.TXTExport txtExport1 = new Spire.DataExport.TXT.TXTExport();
    txtExport1.ExportType = Spire.DataExport.TXT.TextExportType.CSV;
    txtExport1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
    txtExport1.DataTable = ((DataView)Session["SortedView"]).ToTable();
    Response.AddHeader("Transfer-Encoding", "identity");
    txtExport1.SaveToHttpResponse("filtered_result.csv", Response);
    LogAction(this, "download", "filtered_csv");
}
protected void ExportToDBF(object sender, EventArgs e)
{
    Spire.DataExport.DBF.DBFExport DBFExport = new Spire.DataExport.DBF.DBFExport();
    DBFExport.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
    DBFExport.DataTable = ((DataView)Session["SortedView"]).ToTable();
    Response.AddHeader("Transfer-Encoding", "identity");
    DBFExport.SaveToHttpResponse("filtered_result.dbf", Response);
    LogAction(this, "download", "filtered_dbf");
}
protected void ExportAllToXLS(object sender, EventArgs e)
{
    Spire.DataExport.XLS.WorkSheet workSheet1 = new Spire.DataExport.XLS.WorkSheet();
    Spire.DataExport.XLS.CellExport cellExport1 = new Spire.DataExport.XLS.CellExport();
    workSheet1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
    workSheet1.DataTable = BindGridViewAll() as DataTable;
    workSheet1.StartDataCol = ((System.Byte)(0));    
    cellExport1.Sheets.Add(workSheet1);
    Response.AddHeader("Transfer-Encoding", "identity");
    cellExport1.SaveToHttpResponse("result.xls", Response);
    LogAction(this, "download", "all_xls");
}

protected void ExportAllToXLSX(object sender, EventArgs e)
{
    XLWorkbook wb = new XLWorkbook();
    wb.Worksheets.Add(BindGridViewAll(), "Sheet1");
    Response.Clear();
    Response.Buffer = true;
    Response.Charset = "";
    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    Response.AddHeader("content-disposition", "attachment;filename=result.xlsx");
    MemoryStream MyMemoryStream = new MemoryStream();
    wb.SaveAs(MyMemoryStream);
    MyMemoryStream.WriteTo(Response.OutputStream);
    Response.Flush();
    LogAction(this, "download", "all_xlsx");
    Response.End();
}
protected void ExportAllToCSV(object sender, EventArgs e)
{
    Spire.DataExport.TXT.TXTExport txtExport1 = new Spire.DataExport.TXT.TXTExport();
    txtExport1.ExportType = Spire.DataExport.TXT.TextExportType.CSV;
    txtExport1.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
    txtExport1.DataTable = BindGridViewAll() as DataTable;
    Response.AddHeader("Transfer-Encoding", "identity");
    txtExport1.SaveToHttpResponse("result.csv", Response);
    LogAction(this, "download", "all_csv");
}
protected void ExportAllToDBF(object sender, EventArgs e)
{
    Spire.DataExport.DBF.DBFExport DBFExport = new Spire.DataExport.DBF.DBFExport();
    DBFExport.DataSource = Spire.DataExport.Common.ExportSource.DataTable;
    DBFExport.DataTable = BindGridViewAll() as DataTable;
    Response.AddHeader("Transfer-Encoding", "identity");
    DBFExport.SaveToHttpResponse("result.dbf", Response);
    LogAction(this, "download", "all_dbf");
}

public void LogAction(Page page, string action, string value)
{
    string sql = "INSERT INTO log ([dt],[ip],[browser],[action],[value]) VALUES (@date,@ip,@useragent,@action,@values)";
    using (SqlConnection conn = new SqlConnection(connStr))
    {
        conn.Open();
        SqlCommand cmd = new SqlCommand(sql, conn);
        cmd.Parameters.Add("@date", SqlDbType.DateTime).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        cmd.Parameters.Add("@ip", SqlDbType.VarChar).Value = page.Request.UserHostAddress;
        cmd.Parameters.Add("@useragent", SqlDbType.VarChar).Value = page.Request.UserAgent;
        cmd.Parameters.Add("@action", SqlDbType.VarChar).Value = action;
        cmd.Parameters.Add("@values", SqlDbType.VarChar).Value = value;
        int num = cmd.ExecuteNonQuery();
    }
}

</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  lang="ru" xml:lang="ru">
<head>
<meta http-equiv="Content-Language" content="ru">
<meta http-equiv="Content-Type" content="text/html;charset=windows-1251">
<title>Веб приложение фильтрации данных из БД</title>

<style>

a {
	color: #0078ff;
}
#nav {
	margin: 0;
	line-height: 100%;
}
#nav li {
	margin: 0 5px;
    padding: 0 0 8px;
	float: left;
	position: relative;
	list-style: none;
}
#nav a {
	display: block;
	padding:  8px 20px;
	margin: 0;
}
#nav a:hover {
	background: #000;
	color: #fff;
}
#nav ul li:hover a, #nav li:hover li a {
	background: none;
	border: none;
	color: #000;
    text-decoration:none;
}
#nav ul a:hover {
	background: #0078ff  0 -100px !important;
	color: #fff !important;
}

#nav li:hover > ul {
	display: block;
}

#nav ul {
	display: none;
	margin: 0;
	padding: 0;
	width: 125px;
	position: absolute;
	top: 35px;
	left: 0;
	background: #ddd;
}
#nav ul li {
	float: none;
	margin: 0;
	padding: 0;
}

#nav ul a {
	font-weight: normal;
}
/* clearfix */
#nav:after {
	content: ".";
	display: block;
	clear: both;
	visibility: hidden;
	line-height: 0;
	height: 0;
}
#nav {
	display: inline-block;
}

.sortasc {
    background-origin:content-box !important;
    background-position-y:center !important;
    background-position-x:right !important;
    cursor:pointer;
    background-size: 12px !important;
    background: url('data:image/gif;base64,R0lGODlhGQANAPceAGVlZc7OzsXFxdDQ0L6+vsjIyMnJyaioqMrKysPDw8DAwNLS0o+PjwEBAQICArS0tNTU1MfHx8/Pz8HBwaysrBEREb29vcLCws3NzU9PT05OTrKysszMzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB4ALAAAAAAZAA0AAAhTABkAGEiwoMGDAxl4aPCvocOHECP+8+BhgQOJGCEOoEhRQkaMATiK/ChRpEkEJBsaMMky5USWJguQjAATpgCJCWrqLKlTpwKHBHoKdSi0KIWiAQEAOw==') no-repeat;
}
.sortdesc{
    background-origin:content-box  !important;
    background-position-y:center !important;
    background-position-x:right !important;
    cursor:pointer;
    background-size: 12px !important;
    background: url('data:image/gif;base64,R0lGODlhGQANAPceAGVlZc7OzsXFxdDQ0L6+vsjIyMnJyaioqMrKysPDw8DAwNLS0o+PjwEBAQICArS0tNTU1MfHx8/Pz8HBwaysrBEREb29vcLCws3NzU9PT05OTrKysszMzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACH5BAEAAB4ALAAAAAAZAA0AAAhTAD0IHEhQIIWCCBMW/MdQoUOCBBj+U/DQocSLFREmuChRQEaBETiKLJBRpMmKBkyKRGBRpcqEAVyqlEBwgEyZDhYIvHmzgQcGAIIKHUq0aFAGAQEAOw==') no-repeat;
   }

</style>
</head>
<body>
    <div>
        <form id="form1" method="post" runat="server" defaultbutton="CmdOk" defaultfocus="CmdSearch" >
	        <br />
	        <div>Фильтр по столбцу "Текст":&nbsp;&nbsp;<asp:TextBox id="CmdSearch" size="60" runat="server" />&nbsp;
                <asp:Button OnClick="Run" ID="CmdOk" runat="server" Text="Отфильтровать" Width="120px" />
	        </div>
	        <br />
	        <hr />
	        <ul id="nav">
	            <li><a href="#">Скачать отфильтрованное</a>
	                <ul>
	                    <li><asp:LinkButton ID="LinkButton1" runat="server" onclick="ExportToXLS">XLS</asp:LinkButton></li>
	                    <li><asp:LinkButton ID="LinkButton2" runat="server" onclick="ExportToXLSX">XLSX</asp:LinkButton></li>
	                    <li><asp:LinkButton ID="LinkButton3" runat="server" onclick="ExportToCSV">CSV</asp:LinkButton></li>
	                    <li><asp:LinkButton ID="LinkButton4" runat="server" onclick="ExportToDBF">DBF</asp:LinkButton></li>
	                </ul>
	            </li>
	            <li><a href="#">Скачать все</a>
	                <ul>
	                    <li><asp:LinkButton ID="LinkButton5" runat="server" onclick="ExportAllToXLS">XLS</asp:LinkButton></li>
	                    <li><asp:LinkButton ID="LinkButton6" runat="server" onclick="ExportAllToXLSX">XLSX</asp:LinkButton></li>
	                    <li><asp:LinkButton ID="LinkButton7" runat="server" onclick="ExportAllToCSV">CSV</asp:LinkButton></li>
	                    <li><asp:LinkButton ID="LinkButton8" runat="server" onclick="ExportAllToDBF">DBF</asp:LinkButton></li>
	                </ul>
	            </li>
	        </ul>
	        <asp:panel id="pnResult" runat="server" visible="false">
		        <h2>Ничего не найдено по вашему запросу: <b><%=searchtext%></b></h2>
	        </asp:panel>		
	        <table>
		        <tr>
                    <td>
                        <asp:GridView 
                            ID="GridView1" 
                            runat="server" 
                            Font-Names="Arial"
                            CellPadding="4"
                            Font-Size="10pt"
                            EnableTheming="true"
                            HeaderStyle-BackColor="Gray"
                            HeaderStyle-ForeColor="White"
                            AlternatingRowStyle-BackColor="#dddddd" 
                            SortedAscendingHeaderStyle-CssClass="sortasc"
                            SortedDescendingHeaderStyle-CssClass="sortdesc"
                            AutoGenerateColumns="false" 
                            EmptyDataText="Нет данных" 
                            AllowSorting="true"
                            AllowPaging="true" 
                            ShowHeader="true"
                            OnSorting="GridView1_Sorting" 
                            OnPageIndexChanging="GridView1_PageIndexChanging"
                            PageSize="20">
                            <Columns>
                                <asp:BoundField DataField="aid" HeaderText="№" SortExpression="aid"  HeaderStyle-Width="40px"/>
                                <asp:BoundField DataField="adate" HeaderText="Дата1" DataFormatString="{0:yyyy-MM-dd}" HtmlEncode="false" SortExpression="adate"/>
                                <asp:BoundField DataField="bdate" HeaderText="Дата2" DataFormatString="{0:yyyy-MM-dd}" HtmlEncode="false" SortExpression="bdate"/>
                                <asp:BoundField DataField="cdate" HeaderText="Дата3" DataFormatString="{0:yyyy-MM-dd}" HtmlEncode="false" SortExpression="cdate"/>
                                <asp:BoundField DataField="anumber" HeaderText="Целое1" SortExpression="anumber" HeaderStyle-Width="70px"/>
                                <asp:BoundField DataField="bnumber" HeaderText="Целое2" SortExpression="bnumber" HeaderStyle-Width="70px"/>
                                <asp:BoundField DataField="cnumber" HeaderText="Целое3" SortExpression="cnumber" HeaderStyle-Width="70px"/>
                                <asp:BoundField DataField="avalue" HeaderText="Дробное1"  SortExpression="avalue"/>
                                <asp:BoundField DataField="bvalue" HeaderText="Дробное2"  SortExpression="bvalue"/>
                                <asp:BoundField DataField="cvalue" HeaderText="Дробное3"  SortExpression="cvalue"/>
                                <asp:BoundField DataField="text" HeaderText="Текст" SortExpression="text"/>
                            </Columns>

<SortedDescendingHeaderStyle CssClass="sortdesc"></SortedDescendingHeaderStyle>
                        </asp:GridView>
                    </td>
                </tr>
                <tr>
                    <td colspan="9" align="center">
                        <asp:Label ID="lblmsg" runat="server" Width="500px"></asp:Label>
                    </td>
                
                </tr>
            </table>
        </form>
    </div >
</body>
</html>