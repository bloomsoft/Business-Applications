<%@ Page CodeBehind="inquiry.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="Whiterose.inquiry" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML>
	<HEAD>
		<title>White Rose School System (Send an Inquiry)</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<link href="Styles.css" rel="stylesheet" type="text/css">
			<style type="text/css">
			.style6 { FONT-WEIGHT: bold; COLOR: #ffffff }
			.style13 { COLOR: #feb429 }
			.style15 { COLOR: #feb429 }
			.style19 { FONT-WEIGHT: bold; FONT-SIZE: 18px }
			</style>
	</HEAD>
	<body>
		<form name="form1" method="post" runat="server">
			<table width="746" border="0" align="center" cellpadding="0" cellspacing="0">
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
					<td height="16" bgcolor="#feb429" align="left"><span class="style19">Send a Inquiry </span></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td vAlign="top">&nbsp;
						<TABLE id="Table1" cellSpacing="0" cellPadding="0" width="90%" align="center" border="0">
							<TR>
								<TD width="18"></TD>
								<TD>
									<asp:Label id="Label1" runat="server" Font-Size="X-Small">Your Name *</asp:Label></TD>
								<TD>
									<asp:TextBox id="TextBox1" runat="server" Width="274px"></asp:TextBox></TD>
							</TR>
							<TR>
								<TD width="18"></TD>
								<TD>
									<asp:Label id="Label3" runat="server" Font-Size="X-Small">Email Address</asp:Label></TD>
								<TD>
									<asp:TextBox id="TextBox2" runat="server" Width="274px"></asp:TextBox></TD>
							</TR>
							<TR>
								<TD width="18"></TD>
								<TD>
									<asp:Label id="Label4" runat="server" Font-Size="X-Small">Phone *</asp:Label></TD>
								<TD>
									<asp:TextBox id="TextBox3" runat="server" Width="274px"></asp:TextBox></TD>
							</TR>
							<TR>
								<TD width="18"></TD>
								<TD vAlign="top">
									<asp:Label id="Label2" runat="server" Font-Size="X-Small">Message *</asp:Label></TD>
								<TD>
									<asp:TextBox id="TextBox4" runat="server" Width="274px" Height="144px" TextMode="MultiLine"></asp:TextBox></TD>
							</TR>
							<TR>
								<TD width="18"></TD>
								<TD colSpan="2">
									<asp:Label id="Label5" runat="server" Font-Size="X-Small" Width="465px" ForeColor="Red"></asp:Label></TD>
							</TR>
							<TR>
								<TD width="18"></TD>
								<TD></TD>
								<TD>
									<asp:Button id="Button1" runat="server" Text="Send"></asp:Button>
									<asp:Button id="Button2" runat="server" Text="Cancel"></asp:Button></TD>
							</TR>
						</TABLE>
					</td>
				</tr>
				<TR>
					<TD></TD>
				</TR>
				<tr>
					<td bgcolor="#ffffff"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
							<tr>
								<td width="32%"><div align="center"><a href="http://www.bloomsoft.net" target="_blank" class="style13">Copyrights 
											© BloomSoft Technologies</a></div>
								</td>
								<td width="68%"><div align="center" class="style6"><span class="style15"><a href="default.aspx">HOME</a> | <a href="welcome.aspx">WELCOME</a> | <a href="Message.aspx">MESSAGE</a> | <a href="programs.aspx">PROGRAMS</a>								| <a href="achieve1.aspx">ACHIEVEMENTS</a> | <a href="campuses.aspx">CAMPUS</a></span>ES
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
