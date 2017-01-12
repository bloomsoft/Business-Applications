<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Resultcard.aspx.vb" Inherits="Whiterose.Resultcard"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML>
	<HEAD>
		<title>White Rose School System (Teacher Remarks)</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<LINK href="Styles.css" type="text/css" rel="stylesheet">
			<style type="text/css">.style1 { COLOR: #55b136 }
	.style5 { FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #ffffff }
	.style6 { FONT-WEIGHT: bold; COLOR: #ffffff }
	.style7 { FONT-WEIGHT: bold; FONT-SIZE: 12px; COLOR: #ffffff; FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif }
	.style8 { COLOR: #6e62a5 }
	.style10 { FONT-WEIGHT: bold; COLOR: #feb429 }
	.style13 { COLOR: #feb429 }
	.style14 { FONT-WEIGHT: bold; FONT-SIZE: 12px }
	.style15 { COLOR: #feb429 }
	.style17 { FONT-SIZE: 12px }
	.style19 { FONT-WEIGHT: bold; FONT-SIZE: 18px }
	</style>
	</HEAD>
	<body>
		<form id="Form1" name="form1" method="post" runat="server">
			<table cellSpacing="0" cellPadding="0" width="746" align="center" border="0">
				<tr>
					<td width="748" height="74"><table width="97%" height="75" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td width="62%" height="75"><img src="images/WhiteRoseBanner.jpg" width="449" height="75"></td>
								<td width="38%">
									<OBJECT codeBase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"
										height="75" width="275" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000">
										<PARAM NAME="_cx" VALUE="7276">
										<PARAM NAME="_cy" VALUE="1984">
										<PARAM NAME="FlashVars" VALUE="">
										<PARAM NAME="Movie" VALUE="images/MovieTop.swf">
										<PARAM NAME="Src" VALUE="images/MovieTop.swf">
										<PARAM NAME="WMode" VALUE="Window">
										<PARAM NAME="Play" VALUE="-1">
										<PARAM NAME="Loop" VALUE="-1">
										<PARAM NAME="Quality" VALUE="High">
										<PARAM NAME="SAlign" VALUE="">
										<PARAM NAME="Menu" VALUE="-1">
										<PARAM NAME="Base" VALUE="">
										<PARAM NAME="AllowScriptAccess" VALUE="">
										<PARAM NAME="Scale" VALUE="ShowAll">
										<PARAM NAME="DeviceFont" VALUE="0">
										<PARAM NAME="EmbedMovie" VALUE="0">
										<PARAM NAME="BGColor" VALUE="999999">
										<PARAM NAME="SWRemote" VALUE="">
										<PARAM NAME="MovieData" VALUE="">
										<PARAM NAME="SeamlessTabbing" VALUE="1">
										<PARAM NAME="Profile" VALUE="0">
										<PARAM NAME="ProfileAddress" VALUE="">
										<PARAM NAME="ProfilePort" VALUE="0">
										<PARAM NAME="AllowNetworking" VALUE="all">
										<PARAM NAME="AllowFullScreen" VALUE="false">
										<embed src="images/MovieTop.swf" width="275" height="75" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer"
											type="application/x-shockwave-flash" bgcolor="#999999"> </embed>
									</OBJECT>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td bgcolor="#6e639f">
						<OBJECT codeBase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"
							height="23" width="726" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000">
							<PARAM NAME="_cx" VALUE="19209">
							<PARAM NAME="_cy" VALUE="609">
							<PARAM NAME="FlashVars" VALUE="">
							<PARAM NAME="Movie" VALUE="images/bluedot.swf">
							<PARAM NAME="Src" VALUE="images/bluedot.swf">
							<PARAM NAME="WMode" VALUE="Window">
							<PARAM NAME="Play" VALUE="0">
							<PARAM NAME="Loop" VALUE="-1">
							<PARAM NAME="Quality" VALUE="High">
							<PARAM NAME="SAlign" VALUE="">
							<PARAM NAME="Menu" VALUE="-1">
							<PARAM NAME="Base" VALUE="">
							<PARAM NAME="AllowScriptAccess" VALUE="">
							<PARAM NAME="Scale" VALUE="ShowAll">
							<PARAM NAME="DeviceFont" VALUE="0">
							<PARAM NAME="EmbedMovie" VALUE="0">
							<PARAM NAME="BGColor" VALUE="">
							<PARAM NAME="SWRemote" VALUE="">
							<PARAM NAME="MovieData" VALUE="">
							<PARAM NAME="SeamlessTabbing" VALUE="1">
							<PARAM NAME="Profile" VALUE="0">
							<PARAM NAME="ProfileAddress" VALUE="">
							<PARAM NAME="ProfilePort" VALUE="0">
							<PARAM NAME="AllowNetworking" VALUE="all">
							<PARAM NAME="AllowFullScreen" VALUE="false">
							<embed src="images/bluedot.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer"
								type="application/x-shockwave-flash" width="726" height="23"> </embed>
						</OBJECT>
					</td>
				</tr>
				<tr>
					<td height="16" bgcolor="#feb429" align="left"><span class="style19"> Result 
							Card&nbsp;of&nbsp;
							<asp:Label id="lblStdName" runat="server" Font-Size="Medium" Font-Bold="True">Label</asp:Label></span></td>
				</tr>
				<tr>
					<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="76%" valign="top">
								</td>
							</tr>
							<TR>
								<TD vAlign="top" width="76%">
									<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<TD>
												<asp:Label id="Label1" runat="server" Font-Size="X-Small">Select Year</asp:Label></TD>
											<TD>
												<asp:DropDownList id="lstYears" runat="server" Width="155px"></asp:DropDownList></TD>
											<TD>
												<asp:Label id="Label2" runat="server" Font-Size="X-Small">Select Term</asp:Label></TD>
											<TD>
												<asp:DropDownList id="lstTerms" runat="server" Width="155px"></asp:DropDownList></TD>
											<TD>
												<asp:Button id="Button1" runat="server" Text="Show"></asp:Button></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD vAlign="top" align="center" width="76%">
									<asp:DataGrid id=DataGrid1 runat="server" AutoGenerateColumns="False" DataMember="ExamResultVW" DataSource="<%# DsResults1 %>" Width="720px" ShowFooter="True">
										<ItemStyle Font-Size="X-Small" Font-Names="Arial"></ItemStyle>
										<HeaderStyle Font-Size="Small" Font-Names="Arial" Font-Bold="True" ForeColor="#FF8000" BackColor="White"></HeaderStyle>
										<FooterStyle Font-Size="X-Small" Font-Names="Arial" Font-Bold="True" ForeColor="#FF8000" BackColor="White"></FooterStyle>
										<Columns>
											<asp:BoundColumn DataField="Subject" SortExpression="Subject" HeaderText="Subject Name">
												<HeaderStyle Width="60%"></HeaderStyle>
											</asp:BoundColumn>
											<asp:BoundColumn DataField="Outof" SortExpression="Outof" HeaderText="Out/Of">
												<HeaderStyle Width="10%"></HeaderStyle>
											</asp:BoundColumn>
											<asp:BoundColumn DataField="Obtained" SortExpression="Obtained" HeaderText="Obtained">
												<HeaderStyle Width="15%"></HeaderStyle>
											</asp:BoundColumn>
											<asp:BoundColumn DataField="Grade" SortExpression="Grade" HeaderText="Grade">
												<HeaderStyle Width="15%"></HeaderStyle>
											</asp:BoundColumn>
										</Columns>
									</asp:DataGrid></TD>
							</TR>
							<TR>
								<TD vAlign="top" width="76%"></TD>
							</TR>
						</table>
					</td>
				</tr>
				<tr>
					<td>&nbsp;
						<asp:TextBox id="txtAdmNo" runat="server" Visible="False"></asp:TextBox></td>
				</tr>
				<tr>
					<td bgcolor="#ffffff"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
							<tr>
								<td width="32%"><div align="center"><a href="http://www.bloomsoft.net" target="_blank" class="style13">Copyrights 
											© BloomSoft Technologies</a></div>
								</td>
								<td width="68%"><div align="center" class="style6"><span class="style15"><a href="default.aspx">HOME</a>
											| <a href="welcome.aspx">WELCOME</a> | <a href="Message.aspx">MESSAGE</a> | <a href="programs.aspx">
												PROGRAMS</a> | <a href="achieve1.aspx">ACHIEVEMENTS</a> | <a href="campuses.aspx">
												CAMPUS</a></span>ES
									</div>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
