/*  DepressedPress.com DP_SharePoint

Author: Jim Davis, the Depressed Press of Boston
Date: August 03, 2012
Contact: webmaster@depressedpress.com
Website: www.depressedpress.com

Full documentation can be found at:
http://www.depressedpress.com/

DP_SharePoint abstracts common, simple tasks in SharePoint 2007 and 2010.
Certain features require additional components, also available from depressedpress.com:

	- DP_AJAX: Web service calls work in conjunction with DP_AJAX
	- DP_DateExtensions: All DateTime-related method requires DP_DateExtensions
	- DP_Debug: Built-in Debugging requires DP_Debug.


Copyright (c) 1996-2012, The Depressed Press (depressedpress.com)

All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

+) Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

+) Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

+) Neither the name of the THE DEPRESSED PRESS (DEPRESSEDPRESS.COM) nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

*/

/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
/* Configuration */
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

	// Set up DP_Debug (prevents errors if DP_Debug is not present
try {  DP_Debug.isEnabled()  } catch (e) {  DP_Debug = new Object(); DP_Debug.isEnabled = function() { return false }  };


	// Create the Root DP_SharePoint object
DP_SharePoint = {};


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
/* Page Management Methods */
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

	// DP_SharePoint.isEditMode
	// Returns "true" if the page is current in edit mode, "false" if not.
	//
DP_SharePoint.isEditMode = function( ) {

		// "MSOLayout_InDesignMode" will be "1" if a publishing page is in Edit mode
	if ( document.forms[MSOWebPartPageFormName].MSOLayout_InDesignMode.value == 1 ) {
		return true;
	};
		// "_wikiPageMode" will only exist for Wiki pages, and will be "1" if a wiki page is in Edit mode
	if ( document.forms[MSOWebPartPageFormName]._wikiPageMode && document.forms[MSOWebPartPageFormName]._wikiPageMode.value == 1 ) {
		return true;
	};

		// Return
	return false;

};

	// DP_SharePoint.isDialog
	// Returns "true" if the page is being presented as a dialog, "false" if not.
	//
DP_SharePoint.isDialog = function( ) {

		// The Command line variable  IsDlg will be "1" if the page is in a dialog.
	if ( !window.location.search.match("[?&]IsDlg=1") )        {
		return true;
	};
		// Return
	return false;

};


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
/* Form (Add/Edit/Display) Methods */
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

	// DP_SharePoint.getFormType
	// Returns the type of form presented: "Edit", "New", "Display" or null (no form type).
	//
DP_SharePoint.getFormType = function( Title ) {

		// Get the type of form
	var FormType = null;
	if ( document.location.href.indexOf("EditForm.aspx") > -1 ) {
		FormType = "Edit";
	} else if ( document.location.href.indexOf("NewForm.aspx") > -1 ) {
		FormType = "New";
	} else if ( document.location.href.indexOf("DispForm.aspx") > -1 ) {
		FormType = "Display";
	};
		// Return
	return FormType;

};

	// DP_SharePoint.getControlRef
	// Returns references to the "OK" and "Cancel" buttons on SharePoint forms
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

	// DP_SharePoint.getFieldRef
	// Returns a reference (or collection of references) to a field on a form.
	//
DP_SharePoint.getFieldRef = function( Title ) {

		// Get the type of form
	var FormType = DP_SharePoint.getFormType();

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

	// DP_SharePoint.addFieldEvent
	// Add an event to a form field.
	//
DP_SharePoint.addFieldEvent = function( Field, Event, Handler ) {

	if (Field.addEventListener) {
		Field.addEventListener(Event, Handler, false);
	} else if (Field.attachEvent) {
		Field.attachEvent("on" + Event, Handler);
	};

};

	// DP_SharePoint.removeFieldEvent
	// Removes a previously added event from a form field.
	//
DP_SharePoint.removeFieldEvent = function( Field, Event, Handler ) {

	if (Field.removeEventListener) {
		Field.removeEventListener(Event, Handler, false);
	} else if (Field.detachEvent) {
		Field.detachEvent("on" + Event, Handler);
	};

};

	// DP_SharePoint.hideField
	// Hides a field on a form.
	//
DP_SharePoint.hideField = function( Field ) {

		// Hide the field
	Field.RowRef.style.display = "none";
		// Return
	return true;

};

	// DP_SharePoint.showField
	// Shows a previously hidden form field.
	//
DP_SharePoint.showField = function( Field ) {

		// Show the field
	Field.RowRef.style.display = "inline";
		// Return
	return true;

};


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
/* Web Service Methods */
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

	// DP_SharePoint.callService
	// An abstraction method that allows web services to be called more easily.
	//
DP_SharePoint.callService = function( Request, ServiceURL, ServiceMethod, Params ) {

		// Manage Parameters
	var SoapParams = "";
	for ( var Prop in Params ) {
		SoapParams = SoapParams + "<" + Prop + ">" + Params[Prop] + "</" + Prop + ">";
	};

		// Setup the body of the request
	var SoapBody = "";

		// Set the body of the SOAP request
	SoapBody = '<?xml version="1.0" encoding="utf-8"?>\
	<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">\
	<soap:Body>\
	  <' + ServiceMethod + ' xmlns="http://schemas.microsoft.com/sharepoint/soap/">' + SoapParams + '</' + ServiceMethod + '>\
	</soap:Body>\
	</soap:Envelope>';

		// Add the Call to the passed request
	Request.addCall("SOAP", ServiceURL, SoapBody, {"SOAPAction":"http://schemas.microsoft.com/sharepoint/soap/" + ServiceMethod});

};

	// DP_SharePoint.getRows
	// Uses the getElementsByTagNameNS() method for those browers that support it and the getElementsByTagName() method with namespace for others. This is a customized version of the DP_AJAX.getElementsWithNS() method available in DP_AJAX.
	//
DP_SharePoint.getRows = function( XML ) {

		// Even with getElementsByTagNameNS Chrome still ignores the Namespace, so a wildcard is used instead. This may cause problems if multiple namespaces have "row" tags.
	return XML.getElementsByTagNameNS ? XML.getElementsByTagNameNS("*", "row") : XML.getElementsByTagName("z:row");

};

	// DP_SharePoint.getColumn
	// Returns a single column of the response as an array
	//
DP_SharePoint.getColumn = function( XML, ColName ) {

		// Set an array for return
	var Col = [];

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		Col[Cnt] = Rows[Cnt].getAttribute(ColName);

	};

		// Return the Array
	return Col;

};

	// DP_SharePoint.collateResponses
	// Inserts all rows of XMl2 into the rows of XML1 where the given columns match.
	//
DP_SharePoint.collateResponses = function( XML, ColName, XML2, ColName2 ) {

		// Get Rows of the Base and inserted XML
	var Rows = DP_SharePoint.getRows(XML);
	var Rows2 = DP_SharePoint.getRows(XML2);

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


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
/* Web Service Return Cleanup Methods */
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

	// DP_SharePoint.cleanSP_Numeric
	// Cleans the SP_Numeric field type by creating JavaScript numerics from XML/SP strings.
	//
DP_SharePoint.cleanSP_Numeric = function( XML, ColName ) {

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

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

	// DP_SharePoint.cleanSP_MultiLineText
	// Cleans the SP_MultiLineText by parsing embedded HTML.
	//
DP_SharePoint.cleanSP_MultiLineText = function( XML, ColName ) {

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

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


	// DP_SharePoint.cleanSP_UserList
	// Cleans the SP_UserList field type by eliminating prepended database identification information.
	//
DP_SharePoint.cleanSP_UserList = function( XML, ColName, Seperator ) {

	if ( !Seperator ) {
		Seperator = "; ";
	};

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

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


	// DP_SharePoint.cleanSP_CalculatedField
	// Cleans the SP_CalculatedField field type by eliminating prepended database identification information.
	//
DP_SharePoint.cleanSP_CalculatedField = function( XML, ColName ) {

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

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


	// DP_SharePoint.cleanSP_LookupList
	// Cleans the SP_LookupList field type by eliminating prepended database identification information and, optionally, applying a placeholder to missing empty fields.
	//
DP_SharePoint.cleanSP_LookupList = function( XML, ColName, NoValuePlaceholder, Seperator ) {

	if ( !Seperator ) {
		Seperator = "; ";
	};

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

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


	// DP_SharePoint.cleanSP_DateTime
	// Cleans the SP_DateTime field type by formatting values and, optionally, applying a placeholder value to non-dates.
	// * Requires DP_DateExtensions, also by DepressedPress, to be loaded.
	//
DP_SharePoint.cleanSP_DateTime = function( XML, ColName, TimeFormat, DateFormat, NonDateValue ) {

		// This method requires functions from DP_DateExtensions
	if ( !Date.DP_DateExtensions ) {
		throw new Error("The DP_SharePoint.cleanSP_DateTime() method requires that DP_DateExtensions (available from depressedpress.com) be loaded.");
	};

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

		// Loop over rows
	for ( var Cnt=0; Cnt < Rows.length; Cnt++ ) {

			// Get Current Value
		var CurVal = "";
		CurVal = Rows[Cnt].getAttribute(ColName);

			// Parse the Date
		var CurDate = Date.parseFormat(CurVal, "YYYY-MM-DD HH:mm:ss");

			// Determine what to format
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


	// DP_SharePoint.cleanSP_Custom
	// Cleans any SP field type by applying a custom handler to every value in the column, optionally storing the modified values in a new column.
	//
DP_SharePoint.cleanSP_Custom = function( XML, ColName, Handler, NewColName ) {

		// Handle parameters
	if ( !NewColName ) {
		var NewColName = ColName;
	};

		// Get Rows
	var Rows = DP_SharePoint.getRows(XML);

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


/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */
/* Utility Methods */
/* - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */

	// DP_SharePoint.ReplaceLinks
	//
DP_SharePoint.ReplaceLinks = function() {

		// Do not run the function in Edit mode - it only causes problems.
	if ( DP_SharePoint.isEditMode() ) { return false };

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
			case "custom":
				ReplaceLinks_Custom( CurLink );
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
