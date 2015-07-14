/*  DepressedPress.com DP_SharePoint

Author: Jim Davis, the Depressed Press of Boston
Date: August 03, 2012
Contact: webmaster@depressedpress.com
Website: www.depressedpress.com

Full documentation can be found at:
http://www.depressedpress.com/

DP_SharePoint abstracts common, simple tasks in SharePoint 2007 and 2010

	- DP_Debug: Built-in Debugging requires DP_Debug (available from depressedpress.com).
	- DP_DateExtensions: All DateTime-related method requires DP_DateExtensions (also available from the depressedpress.com)


Copyright (c) 1996-2012, The Depressed Press (depressedpress.com)

All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

+) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

+) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

+) Neither the name of the THE DEPRESSED PRESS (DEPRESSEDPRESS.COM) nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

*/

	// Manage other, optional, DP components
	//

	// Set up DP_Debug (prevents errors if DP_Debug is not present
try {  DP_Debug.isEnabled()  } catch (e) {  DP_Debug = new Object(); DP_Debug.isEnabled = function() { return false }  };


	// Create the Root DP_SharePoint object
	//
DP_SharePoint = {};



	// Add/Edit Template field methods.
	//

	// DP_SharePoint.getControlRef Method - returns references to the "OK" and "Cancel" buttons on SharePoint forms
	//
DP_SharePoint.getControlRef = function() {

		// Set an object to collect the output
	var Output = {"Submit":[], "Cancel":[]};
		// Get All input tags
	var inputTags = document.getElementsByTagName("INPUT");
		// Find the correct Tags
	for (var cnt=0; cnt < inputTags.length; cnt++) {
			// Find the buttons
		if ( inputTags[cnt].type == "button" ) {
			if ( inputTags[cnt].value == "OK" || inputTags[cnt].value == "Save" ) {
				Output.Submit.push(inputTags[cnt]);
			};
			if ( inputTags[cnt].value == "Cancel" ) {
				Output.Cancel.push(inputTags[cnt]);
			};
		};
	};
		// Return
	return Output;

};


	// DP_SharePoint.getFieldRef Method
	//
DP_SharePoint.getFieldRef = function( Title ) {

		// Get the type of form
	var FormType = "";
	if ( document.location.href.indexOf("EditForm.aspx") > -1 ) {
		FormType = "Edit";
	} else if ( document.location.href.indexOf("NewForm.aspx") > -1 ) {
		FormType = "New";
	} else if ( document.location.href.indexOf("DispForm.aspx") > -1 ) {
		FormType = "Display";
	};

		// Set an object to collect the output
	var Output = {};

		// Get All h3 tags
	var h3Tags = document.getElementsByTagName("H3");

		// Find the correct h3
	var FieldFound = false;
	for (var cnt=0; cnt < h3Tags.length; cnt++) {
			// The H3's contain all sorts of odd HTML, so we check the whole Node for the exact string bounded by other tags
		if ( h3Tags[cnt] && h3Tags[cnt].parentNode.innerHTML.indexOf(">" + Title + "<") > -1 ) {
			Output.LabelRef = h3Tags[cnt];
			FieldFound = true;
			break;
		};
	};
		// Return null (failure) if the field wasn't found
	if ( !FieldFound ) { return null };

		// Get the containing TR
	Output.RowRef = Output.LabelRef.parentNode.parentNode;

		// Parse out the internal comments  to gather data
	var CurComments = Output.RowRef.getElementsByTagName("TD")[1].firstChild.nodeValue.match(/"[^"]*"/g);
	Output.Name = CurComments[0].substring(1, CurComments[0].length - 1);
	Output.InternalName = CurComments[1].substring(1, CurComments[1].length - 1);
	Output.Type = CurComments[2].substring(1, CurComments[2].length - 1);
	Output.SubType = "";

		// Gather References to the fields, according to FieldTypes (only for "Edit" and "New" forms)
	if ( FormType == "Display" ) {
			// Pull the contents of the field
		Output.ElText = Output.RowRef.getElementsByTagName("TD")[1].innerHTML;
			// Strip out the Comments
		Output.ElText = Output.ElText.replace(/<!--(.*?)-->/gm, "");
	} else if ( ( FormType == "Edit" ) || ( FormType == "New" ) ) {
		switch( Output.Type ) {

				// Simple INPUT tags
			case "SPFieldText":
			case "SPFieldNumber":
			case "SPFieldCurrency":
			case "SPFieldBoolean":

				Output.ElRef = Output.RowRef.getElementsByTagName("INPUT")[0];

			break;
				// SELECT, RADIO and CHECKBOX tags
			case "SPFieldChoice":
			case "SPFieldMultiChoice":

				Output.ElRef = {};
				if ( Output.RowRef.getElementsByTagName("SELECT")[0] ) {
					if ( Output.RowRef.getElementsByTagName("INPUT")[0] ) {
						Output.ElRef.Select = Output.RowRef.getElementsByTagName("SELECT")[0];
						Output.SubType = "SelectWithCustom";
					} else {
						Output.ElRef = Output.RowRef.getElementsByTagName("SELECT")[0];
						Output.SubType = "Select";
					};
				};

				if ( Output.RowRef.getElementsByTagName("INPUT")[0] ) {
					var CurInputs = Output.RowRef.getElementsByTagName("INPUT");
					Output.ElRef.Choices = [];
					for (var cnt=0; cnt < CurInputs.length; cnt++) {
						if ( CurInputs[Cnt].type != "text" ) {
							Output.ElRef.Choices.push(CurInputs[Cnt]);
						};
						if ( CurInputs[Cnt].type == "text" ) {
							Output.ElRef.CustomChoice = CurInputs[Cnt - 1];
							Output.ElRef.Custom = CurInputs[Cnt];
						};
					};
					if ( !Output.SubType ) {
						if ( CurInputs[Cnt].type == "radio" ) {
							if ( Output.ElRef.Custom ) {
								Output.SubType = "RadioWithCustom";
							} else {
								Output.SubType = "Radio";
							};
						} else if ( CurInputs[Cnt].type == "checkbox" ) {
							if ( Output.ElRef.Custom ) {
								Output.SubType = "CheckboxWithCustom";
							} else {
								Output.SubType = "Checkbox";
							};
						};
					};
				};

			break;
				// Simple TEXTAREA tags
			case "SPFieldNote":

				var BaseTag = Output.RowRef.getElementsByTagName("TEXTAREA")[0];
				Output.ElRef = {};
				Output.ElRef.TextArea = BaseTag;
				Output.ElRef.EditorDoc = RTE_GetEditorDocument(BaseTag.id);

			break;
				// Can be a SELECT or an INPUT
			case "SPFieldLookup":

				Output.ElRef = Output.RowRef.getElementsByTagName("SELECT")[0];
				Output.SubType = "Select";
				if ( !Output.ElRef ) {
					Output.ElRef = Output.RowRef.getElementsByTagName("INPUT")[0];
					Output.SubType = "Input";
				};

			break;
				// A base hidden INPUT, two SELECT fields and two buttons
			case "SPFieldLookupMulti":

				var BaseTag = Output.RowRef.getElementsByTagName("INPUT")[0];
				Output.ElRef = {};
				Output.ElRef.Source = BaseTag;
				Output.ElRef.Candidates = document.getElementById(BaseTag.id.replace("MultiLookupPicker", "SelectCandidate"));
				Output.ElRef.Selected = document.getElementById(BaseTag.id.replace("MultiLookupPicker", "SelectResult"));
				Output.ElRef.AddButton = document.getElementById(BaseTag.id.replace("MultiLookupPicker", "AddButton"));
				Output.ElRef.RemoveButton = document.getElementById(BaseTag.id.replace("MultiLookupPicker", "RemoveButton"));

			break;
			case "SPFieldUser":
			case "SPFieldUserMulti":

				Output.ElRef = {};
				Output.ElRef.Display = Output.RowRef.getElementsByTagName("TEXTAREA")[0];
				Output.ElRef.InputData = Output.RowRef.getElementsByTagName("INPUT")[0];
				Output.ElRef.SpanData = Output.RowRef.getElementsByTagName("INPUT")[1];
				Output.ElRef.XMLData = Output.RowRef.getElementsByTagName("INPUT")[2];
				Output.ElRef.CheckNamesLink = Output.RowRef.getElementsByTagName("A")[0];
				Output.ElRef.BrowseNamesLink = Output.RowRef.getElementsByTagName("A")[1];
			break;
			case "SPFieldDateTime":

				var BaseTag = Output.RowRef.getElementsByTagName("INPUT")[0];
				Output.ElRef = {};
				Output.ElRef.Date = BaseTag;
				Output.ElRef.Hours = document.getElementById(BaseTag.id + "Hours");
				Output.ElRef.Minutes = document.getElementById(BaseTag.id + "Minutes");
				if ( Output.ElRef.Hours ) {
					Output.SubType = "Date";
				} else {
					Output.SubType = "DateTime";
				};
			break;
			case "SPFieldURL":

				Output.ElRef = {};
				Output.ElRef.URL = Output.RowRef.getElementsByTagName("INPUT")[0];
				Output.ElRef.Label = Output.RowRef.getElementsByTagName("INPUT")[1];

			break;

		};
	};

		// Return the output object
	return Output;

};


	// DP_SharePoint.addFieldEvent Method
	//
DP_SharePoint.addFieldEvent = function( Field, Event, Handler ) {

	if (Field.addEventListener) {
		Field.addEventListener(Event, Handler, false);
	} else if (Field.attachEvent) {
		Field.attachEvent("on" + Event, Handler);
	};

};


	// DP_SharePoint.removeFieldEvent Method
	//
DP_SharePoint.removeFieldEvent = function( Field, Event, Handler ) {

	if (Field.removeEventListener) {
		Field.removeEventListener(Event, Handler, false);
	} else if (Field.detachEvent) {
		Field.detachEvent("on" + Event, Handler);
	};

};

	// DP_SharePoint.hideField Method
DP_SharePoint.hideField = function( Field ) {

		// Hide the field
	Field.RowRef.style.display = "none";
		// Return
	return true;

};

	// DP_SharePoint.showField Method
DP_SharePoint.showField = function( Field ) {

		// Show the field
	Field.RowRef.style.display = "inline";
		// Return
	return true;

};



	// Webservice Methods
	//

	// DP_SharePoint.getElementsWithNS Method
	// Uses the getElementsByTagNameNS() method for those browers that support it and the getElementsByTagName() method with namespace for others.
	//
DP_SharePoint.getRows = function( XML, Tag, NS ) {

		// Even with getElementsByTagNameNS Chrome still ignores the Namespace, so a wildcard is used instead. This may cause problems if multiple namespaces have "row" tags.
	return XML.getElementsByTagNameNS ? XML.getElementsByTagNameNS("*", Tag) : XML.getElementsByTagName(NS + ":" + Tag);

};

	// DP_SharePoint.callListService_Get Method
	//
DP_SharePoint.callListService_Get = function( Request, ServiceURL, ListName, ViewName, Query, QueryOptions, ViewFields ) {

	if ( !Query ) {
		Query = "";
	} else {
		Query = "<query>" + Query + "</query>";
	};

	if ( !QueryOptions ) {
		QueryOptions = "";
	} else {
		QueryOptions = "<queryOptions>" + QueryOptions + "</queryOptions>";
	};

	if ( !ViewFields ) {
		ViewFields = "";
	} else {
		ViewFields = "<viewFields>" + ViewFields + "</viewFields>";
	};

		// Setup the body of the request
	var SoapBody = "";

		// Set the body of the SOAP request
	SoapBody = '<?xml version="1.0" encoding="utf-8"?>\
	<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">\
	<soap:Body>\
	  <GetListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">\
	    <listName>' + ListName + '</listName>\
	    <viewName>' + ViewName + '</viewName>\
	    ' + Query + '\
	    ' + QueryOptions + '\
	    ' + ViewFields + '\
	  </GetListItems>\
	</soap:Body>\
	</soap:Envelope>';

		// Add the Call to the passed request
	Request.addCall("SOAP", ServiceURL, SoapBody, {"SOAPAction":"http://schemas.microsoft.com/sharepoint/soap/GetListItems"});

};


	// DP_SharePoint.callListService_Update Method
	//
DP_SharePoint.callListService_Update = function( Request, ServiceURL, ListName, Batch ) {

		// Setup the body of the request
	var SoapBody = "";

	SoapBody = '<?xml version=\"1.0\" encoding=\"utf-8\"?>\
	<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">\
	<soap:Body>\
		<UpdateListItems xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\">\
			<listName>' + ListName + '</listName>\
			<updates>\
				<Batch OnError=\"Continue\">' + Batch + '</Batch>\
			</updates>\
		</UpdateListItems>\
	</soap:Body>\
	</soap:Envelope>';

		// Add the Call to the passed request
	Request.addCall("SOAP", ServiceURL, SoapBody, {"SOAPAction":"http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"});

};


	// Webservice Cleanup Methods
	//

	// DP_SharePoint.cleanSP_Numeric Method
	//
DP_SharePoint.cleanSP_Numeric = function( XML, ColName ) {

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurVal = "";
		CurVal = Rows[Cnt].getAttribute(ColName);
		if ( CurVal != "" ) {
			Rows[Cnt].setAttribute(ColName, +(CurVal));
		};

	};

		// Return a Reference to the XML
	return XML;

};

	// DP_SharePoint.cleanSP_MultiLineText Method
	//
DP_SharePoint.cleanSP_MultiLineText = function( XML, ColName ) {

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurVal = "";
		CurVal = Rows[Cnt].getAttribute(ColName);
		if ( !CurVal ) {
			Rows[Cnt].setAttribute(ColName, "");
		} else {
				// Create a Temporary Div to store the HTML
			var TempDiv = document.createElement("div");
			TempDiv.innerHTML = CurVal;
				// Pull the (now parsed) value from the div
			if ( TempDiv.innerText !== undefined ) {
				Rows[Cnt].setAttribute(ColName, TempDiv.innerText); // IE
			} else {
				Rows[Cnt].setAttribute(ColName, TempDiv.textContent); // FF
			};
				// Null out the TempDiv
			TempDiv = null;
		};

	};

		// Return a Reference to the XML
	return XML;

};


	// DP_SharePoint.cleanSP_UserList Method
	//
DP_SharePoint.cleanSP_UserList = function( XML, ColName, Seperator ) {

	if ( !Seperator ) {
		Seperator = "; ";
	};

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurVal = "";
		CurVal = Rows[Cnt].getAttribute(ColName);

			// Strip ID values
		var TempVal = CurVal.replace(/[0-9]{1,5};#/g, "");
			// Convert remaining separators
		TempVal = TempVal.replace(/;#/g, Seperator);

			// Set new Value
		Rows[Cnt].setAttribute(ColName, TempVal);

	};

		// Return a Reference to the XML
	return XML;

};


	// DP_SharePoint.cleanSP_CalculatedField Method
	//
DP_SharePoint.cleanSP_CalculatedField = function( XML, ColName ) {

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurVal = "";
		CurVal = Rows[Cnt].getAttribute(ColName);

			// Split the value on the type identifier
		var TempVal = CurVal.split(";#");

			// Set new Value
		Rows[Cnt].setAttribute(ColName, TempVal[1]);

	};

		// Return a Reference to the XML
	return XML;

};


	// DP_SharePoint.cleanSP_LookupList Method
	//
DP_SharePoint.cleanSP_LookupList = function( XML, ColName, NoValuePlaceholder, Seperator ) {

	if ( !Seperator ) {
		Seperator = "; ";
	};

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurValue = "";
		CurVal = Rows[Cnt].getAttribute(ColName);

			// If the attribute exists, clean it up
		if ( CurVal ) {
				// Strip ID values
			var TempVal = CurVal.replace(/[0-9]{1,5};#/g, "");
				// Convert remaining separators
			TempVal = TempVal.replace(/;#/g, Seperator );

				// Set new Value
			Rows[Cnt].setAttribute(ColName, TempVal);

		} else {

				// Set Value to NoValuePlaceholder
			Rows[Cnt].setAttribute(ColName, NoValuePlaceholder);

		};

	};

		// Return a Reference to the XML
	return XML;

};


	// DP_SharePoint.cleanSP_DateTime Method
	//
DP_SharePoint.cleanSP_DateTime = function( XML, ColName, TimeFormat, DateFormat, NonDateValue ) {

		// This method requires functions from DP_DateExtensions
	if ( !Date.DP_DateExtensions ) {
		throw new Error("The DP_SharePoint.cleanSP_DateTime() method requires that DP_DateExtensions (available from depressedpress.com) be loaded.");
	};

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurVal = "";
		CurVal = Rows[Cnt].getAttribute(ColName);

			// Parse the Date
		var CurDate = Date.parseFormat(CurVal, "YYYY-MM-DD HH:mm:ss");

			// Determine what to
		var TempVal = "";
		if ( CurDate != null ) {
			if ( TimeFormat != null && DateFormat == null ) {
				TempVal = CurDate.timeFormat(TimeFormat);
			} else if ( TimeFormat == null && DateFormat != null ) {
				TempVal = CurDate.dateFormat(DateFormat);
			} else {
				TempVal = CurDate.timeFormat(TimeFormat) + " " + CurDate.dateFormat(DateFormat);
			};
		} else {
			if ( NonDateValue != null ) {
				TempVal = NonDateValue;
			} else {
				TempVal = CurVal;
			};
		};

			// Set new Value
		Rows[Cnt].setAttribute(ColName, TempVal);

	};

		// Return a Reference to the XML
	return XML;

};


	// DP_SharePoint.cleanSP_Custom Method
	//
DP_SharePoint.cleanSP_Custom = function( XML, ColName, Handler, NewColName ) {

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurVal = "";
		CurVal = Rows[Cnt].getAttribute(ColName);

			// Set new Value
		if ( !NewColName ) {
			Rows[Cnt].setAttribute(ColName, Handler(CurVal));
		} else {
			Rows[Cnt].setAttribute(NewColName, Handler(CurVal));
		};

	};

		// Return a Reference to the XML
	return XML;

};




	// Webservice Utility Methods
	//

	// DP_SharePoint.getAttributes
	//
DP_SharePoint.getAttributes = function( XML, ColName ) {

		// Set an array for return
	var Attributes = [];

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML, "row", "z");

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		Attributes[Cnt] = Rows[Cnt].getAttribute(ColName);

	};

		// Return the Array
	return Attributes;

};

	// DP_SharePoint.joinResponses
	//
DP_SharePoint.joinResponses = function( XML, ColName, XML2, ColName2 ) {

		// Get Rows of the Base and inserted XML
	var Rows = DP_SharePoint.getRows(XML, "row", "z");
	var Rows2 = DP_SharePoint.getRows(XML2, "row", "z");

		// Loop over rows of the Base XML
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Base XML row information
		var CurRow = Rows[Cnt];
		var CurVal = CurRow.getAttribute(ColName);

			// Loop over rows of the inserted XML
		for ( var iCnt=0; iCnt < Rows2.length; iCnt++ ) {

				// Get Insertable XML row information
			var CurRow2 = Rows2[iCnt];
			var CurVal2 = CurRow2.getAttribute(ColName2);

				// Set new Value
			if ( CurVal == CurVal2 ) {

					// Insert the row into the Base XML
				CurRow.appendChild(CurRow2);

			};

		};

	};

		// Return the Array
	return XML;

};


	// Utility Methods
	//

	// DP_SharePoint.ReplaceLinks Method
	//
DP_SharePoint.ReplaceLinks = function() {

		// Obtain a collection of all page links
	var AllLinks = document.getElementsByTagName("a");
	var AllLinksCnt = AllLinks.length;
	var CurMatch;
	var CurLink;
	var MatchedLinks = [];

		// Set the Target Proxy
	var TargetProtocol = "news";
	var ReplacementMask = new RegExp(TargetProtocol + "://\\*([A-Z0-9]{1,15})\\*", "i");

		// Loop over all links and find those marked for replacement
	for( var cnt=0; cnt < AllLinksCnt; ++cnt ) {
		CurLink = AllLinks[cnt];
		if ( CurLink.href ) {
			CurMatch = CurLink.href.match(ReplacementMask);
			if ( CurMatch != null ) {
					// Add a new Attribute, containing the new Protocol value
				CurLink.setAttribute("DP_SharePoint_ReplacementProtocol", CurMatch[1]);
					// Add the Link to the Matched Links group for later processing
				MatchedLinks.push(CurLink);
			};
		};
 	};

		// Loop over the Collected Links
	for( var cnt=0; cnt < MatchedLinks.length; ++cnt ) {
		CurLink = MatchedLinks[cnt];
		CurProtocol = CurLink.getAttribute("DP_SharePoint_ReplacementProtocol");
		switch ( CurProtocol ) {
			case "lync":
				ReplaceLinks_Lync( CurLink );
				break;
			case "sametime":
				ReplaceLinks_Sametime( CurLink );
				break;
			default:
				ReplaceLinks_Other( CurLink );
				break;
		};
	};

	function ReplaceLinks_Sametime(CurLink) {

			// If needed, add the Sametime style sheet to the head
		if ( !document.getElementById("metReplaceLinks_SametimeStyle") ) {
			var StyleElement = document.createElement("link");
			StyleElement.setAttribute("id", "metReplaceLinks_SametimeStyle");
			StyleElement.setAttribute("rel", "stylesheet");
			StyleElement.setAttribute("href", "http://localhost:59449/stwebapi/main.css");
			StyleElement.setAttribute("type", "text/css");
			document.getElementsByTagName("head")[0].appendChild(StyleElement);
		};
			// If needed, add the Sametime script sheet to the head
		if ( !document.getElementById("metReplaceLinks_SametimeScript") ) {
			var ScriptElement = document.createElement("script");
			ScriptElement.setAttribute("id", "metReplaceLinks_SametimeScript");
			ScriptElement.setAttribute("type", "text/javascript");
			ScriptElement.setAttribute("src", "http://localhost:59449/stwebapi/getStatus.js");
			document.getElementsByTagName("head")[0].appendChild(ScriptElement);
		};


			// Determine the ID to use
		var CurID = CurLink.href.replace(ReplacementMask, "").replace("/", "");
			// Get the Content of the link (the Sametime ID) and strip any crap HTML out of it
		if ( CurID == "" ) {
			var TempID = CurLink.innerHTML;
			CurID = TempID.replace(/<.*?>/g, '');
		};

			// Modify the CurLink
		CurLink.removeAttribute("href");
		CurLink.setAttribute("class", "awareness");
		CurLink.style.fontSize = "8pt";
		CurLink.setAttribute("userId", CurID);

	};

	function ReplaceLinks_Lync(CurLink) {

			// Determine the ID to use
		var CurID = CurLink.href.replace(ReplacementMask, "").replace("/", "");
		var CurHtmlID = CurID.replace("@", "").replace(".", "");
			// Get the Content of the link (the Lync ID) and strip any crap HTML out of it
		if ( CurID == "" ) {
			var TempID = CurLink.innerHTML;
			CurID = TempID.replace(/<.*?>/g, '');
		};

			// Remove the href
		CurLink.removeAttribute("href");

			// Create the Span Node
		CurSpan = document.createElement("span");

			// If we're not in an ActiveX capable browser
		if ( !window.ActiveXObject ) {

				// Add a simple lable (nothing else to do, really)
			CurLink.innerHTML = "Lync: " + CurLink.innerHTML;
				// Move all of the content from the Link to the span
			CurSpan.appendChild(CurLink.firstChild);
				// Replace the CurLink with the Span
			CurLink.appendChild(CurSpan);

		} else {

				// Create the Image Node
			CurImg = document.createElement("img");
			CurImg.setAttribute("border", "0");
			CurImg.setAttribute("valign", "middle");
			CurImg.setAttribute("height", "12");
			CurImg.setAttribute("width", "12");
			CurImg.setAttribute("ShowOfflinePawn", "1");
			CurImg.setAttribute("id", "Lync_" + CurHtmlID);
			CurImg.setAttribute("style", "margin-right: 3px;");

				// Add the Img to the Span
			CurSpan.appendChild(CurImg);

				// Move all of the content from the Link to the span
			CurSpan.appendChild(CurLink.firstChild);

				// Replace the CurLink with the Span
			CurLink.appendChild(CurSpan);

				// Load the image and run the script
			CurImg.setAttribute("onload", "IMNRC('" + CurID + "')");
			CurImg.setAttribute("src", "/_layouts/images/imnhdr.gif");

		};

	};

	function ReplaceLinks_Other(CurLink) {

		var NewProtocol = CurLink.getAttribute("DP_SharePoint_ReplacementProtocol");
		CurLink.href = CurLink.href.replace(ReplacementMask, NewProtocol + ":\\\\");

	};

};
