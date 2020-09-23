<div align="center">

## Serial Communication ASP


</div>

### Description

To communicate with any device connected to com port via an activeX provided by Microsoft called MSComm (MSCOMM32.OCX).
 
### More Info
 
I've tried using my proximity card reader connected to COM1 and was able to read data from the card once it touch the reader.

it needs to be beautify coz I only spent about 2-3 hours to come out with this piece of code.

any string feeded in by the device through COM port.

not tested


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Muhamad Kamal](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/muhamad-kamal.md)
**Level**          |Advanced
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/muhamad-kamal-serial-communication-asp__4-7763/archive/master.zip)





### Source Code

```
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<SCRIPT LANGUAGE="VBScript">
'Sub Window_OnLoad()
'	MSComm1.PortOpen =True
'end sub
Sub OpenPort()
	if Not MSComm1.PortOpen Then
		MSComm1.PortOpen =True
	else
		msgbox "Port already opened !", vbOKOnly, "Warning"
	end if
end sub
Sub ClosePort()
	if MSComm1.PortOpen Then
		MSComm1.PortOpen = False
	else
		msgbox "Port already closed !", vbOKOnly, "Warning"
	end if
end sub
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function MSComm1_OnComm() {
	var fldWeight = frmView.txtWeight
	var strInput
	strInput = MSComm1.Input;
	window.alert (strInput);
	fldWeight.Value == strInput;
	fldWeight.focus();
	return false;
}
//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=MSComm1 EVENT=OnComm>
<!--
	MSComm1_OnComm()
//-->
</SCRIPT>
</HEAD>
<BODY>
<OBJECT classid=clsid:648A5600-2C6E-101B-82B6-000000000014 id=MSComm1
style="LEFT: 54px; TOP: 14px">
<PARAM NAME="_ExtentX" VALUE="1005">
<PARAM NAME="_ExtentY" VALUE="1005">
<PARAM NAME="_Version" VALUE="393216">
<PARAM NAME="CommPort" VALUE="1">
<PARAM NAME="DTREnable" VALUE="-1">
<PARAM NAME="Handshaking" VALUE="0">
<PARAM NAME="InBufferSize" VALUE="1024">
<PARAM NAME="InputLen" VALUE="0">
<PARAM NAME="NullDiscard" VALUE="0">
<PARAM NAME="OutBufferSize" VALUE="512">
<PARAM NAME="ParityReplace" VALUE="63">
<PARAM NAME="RThreshold" VALUE="14">
<PARAM NAME="RTSEnable" VALUE="0">
<PARAM NAME="BaudRate" VALUE="9600">
<PARAM NAME="ParitySetting" VALUE="0">
<PARAM NAME="DataBits" VALUE="7">
<PARAM NAME="StopBits" VALUE="0">
<PARAM NAME="SThreshold" VALUE="0">
<PARAM NAME="EOFEnable" VALUE="-1">
<PARAM NAME="InputMode" VALUE="0"></OBJECT>
<FORM action="" method=post id=frmView name=frmView>
<P>
<INPUT id=txtWeight name=txtWeight style="LEFT: 50px; TOP: 17px"></P>
<P>&nbsp;</P>
<P>
<BUTTON OnClick="OpenPort()" value="Open Port">Open COM Port</BUTTON>
<BUTTON onclick="ClosePort()" value="Close Port">Close COM Port</BUTTON>
</P>
<P>
</P>
</FORM>
</BODY>
</HTML>
```

