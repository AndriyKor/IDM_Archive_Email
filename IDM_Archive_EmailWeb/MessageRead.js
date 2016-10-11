/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
  "use strict";

  var messageBanner,
      xmlEntities,
      spinner,
      apiUrl = "https://eu-be-sos05:21106/ca/secure-jsp/mds/api/",
      messageType = Object.freeze({ SUCCESS: "success", INFO: "info", WARNING: "warning", ERROR: "error" });

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
    $(document).ready(function () {
      var element = document.querySelector('.ms-MessageBanner');
      messageBanner = new fabric.MessageBanner(element);
      messageBanner.hideBanner();
      //loadProps();

      // Add event handlers
       $('#select-division select')[0].onchange = selectDivisionChanged;
       $('#select-item-type select')[0].onchange = selectItemTypeChanged;
       $('#save-button-container button')[0].onclick = archiveEmail;

      // wire up dropdown controls
      if ($.fn.Dropdown) {
          $('.ms-Dropdown').Dropdown();
      }

      getEntitesXml();

    });
  };

  function selectDivisionChanged(event) {
      // disable item type DropDown if nothing selected
      if (event.target.value == "zero") {
          disableElement(document.getElementById('select-item-type'));
          clearItemTypeList();
         // return "None";
      } else {
          enableElement(document.getElementById('select-item-type'));
      }

      // refresh item type list
      refreshItemTypeList();
  }

  function selectItemTypeChanged(event) {
      // clear the list
      clearAttributeList();

      var entitySelected = event.target.value;
      var entityNode;
      var selectedDivision = document.getElementById("select-division").getElementsByTagName("select")[0].value;
      // var regExp1 = /^[^\\/:\*\?"<>\|]+$/; // forbidden characters \ / : * ? " < > |
      // var regExp2 = /^\./; // cannot start with dot (.)
      // var regExp3 = /^(nul|prn|con|lpt[0-9]|com[0-9])(\.|$)/i; // forbidden file names

      // get 'entity' node for selected item type from xmlEntities
      var entityList = xmlEntities.getElementsByTagName("entity");
      for (var i = 0; i < entityList.length; i++) {
          var entityName = entityList[i].getElementsByTagName("name")[0].childNodes[0].nodeValue;
          if (entityName == entitySelected) {
              entityNode = entityList[i];
              break;
          }
      }

      // go through attributes,
      // create and add TextField component (Office Fabric UI) to HTML
      var buttonArchive;
      if (typeof (entityNode) != "undefined" && entityNode != null) {
          var attributeList = entityNode.getElementsByTagName("attrs")[0].getElementsByTagName("attr");

          for (var i = 0; i < attributeList.length; i++) {
              var attributeContainer = document.getElementById("attribute-container").getElementsByClassName("ms-Grid-col")[0];
              var attributeName = attributeList[i].getElementsByTagName("name")[0].childNodes[0].nodeValue;
              var attributeDescription = attributeList[i].getElementsByTagName("desc")[0].childNodes[0].nodeValue;
              var attributeType = attributeList[i].getElementsByTagName("type")[0].childNodes[0].nodeValue;
              var attributeFlag = attributeList[i].getElementsByTagName("flag")[0].childNodes[0].nodeValue;
              var attributeSize = attributeList[i].getElementsByTagName("size")[0].childNodes[0].nodeValue;

              var nodeDivMain = document.createElement("div");

              if (attributeType == "7") {
                  // Date attribute
                  nodeDivMain.classList.add("ms-DatePicker");
                  nodeDivMain.innerHTML = getDatePickerHtml(attributeDescription.replace(" (yyyy-mm-dd)", ""), attributeName, attributeType, attributeFlag);
              } else {
                  // string/number attribute
                  // usually has type '2'
                  var nodeLabel = document.createElement("label");
                  var nodeInput = document.createElement("input");
                  var nodeSpan = document.createElement("span");

                  nodeDivMain.classList.add("ms-TextField");
                  nodeDivMain.id = attributeName;
                  nodeLabel.classList.add("ms-Label");
                  nodeLabel.textContent = attributeDescription;
                  nodeInput.classList.add("ms-TextField-field");
                  nodeInput.maxLength = attributeSize;
                  nodeInput.setAttribute("data-attribute-type", attributeType);
                  nodeInput.setAttribute("data-attribute-flag", attributeFlag);
                  nodeSpan.classList.add("ms-TextField-description");

                  if (attributeFlag == "10" || attributeFlag == "26") {
                      nodeInput.required = true;
                      nodeSpan.textContent = "* Required";
                  }

                  if (attributeName == "M3_DIVI") {
                      nodeInput.value = selectedDivision;
                      nodeInput.readOnly = true;
                  }

                  nodeDivMain.appendChild(nodeLabel);
                  nodeDivMain.appendChild(nodeInput);
                  nodeDivMain.appendChild(nodeSpan);
              }

              attributeContainer.appendChild(nodeDivMain);
              

          }

          // enable save button
          //enableElement(document.getElementById("save-button-container").getElementsByTagName("button")[0]);
          buttonArchive = document.getElementById("save-button-container").getElementsByTagName("button")[0];
          if (buttonArchive.hasAttribute("disabled")) {
              buttonArchive.removeAttribute("disabled");
          }

          // wire up text field controls
          if ($.fn.TextField) {
              $('.ms-TextField').TextField();
          }

          // wire up date picker controls
          if ($.fn.DatePicker) {
              $('.ms-DatePicker').DatePicker();
          }

      } else {
          // should not cause, but... Everything can happen :)
          // showNotification("Error", "Related node was not found for selected entity!", messageType.ERROR);

          // disable save button
          //disableElement(document.getElementById("save-button-container").getElementsByTagName("button")[0]);
          buttonArchive = document.getElementById("save-button-container").getElementsByTagName("button")[0];
          if (!buttonArchive.hasAttribute("disabled")) {
              buttonArchive.setAttribute("disabled", "disabled");
          }
      }
  }

  function refreshItemTypeList() {

      // clear the list
      clearItemTypeList();

      // fill in the list - loop
      if (typeof (xmlEntities) != "undefined" && xmlEntities != null) {
          // generate item type list
          var entityList = xmlEntities.getElementsByTagName("entity");
          var selectItemType = document.getElementById("select-item-type").getElementsByTagName("select")[0];

          for (var i = 0; i < entityList.length; i++) {
              var entityName = entityList[i].getElementsByTagName("name")[0].childNodes[0].nodeValue;
              var entityDescription = entityList[i].getElementsByTagName("desc")[0].childNodes[0].nodeValue;
              var selectedDivision = $('#select-division select')[0].value;
              var entityDivision = entityName.slice(-3);

              // Filter for external item types.
              // External item type always has division number in the end of the name
              // For example: M3_M3_RMA_101, M3_Rental_docs_607, ...
              if (!isNaN(entityDivision)) {
                  // add only item types related to selected division
                  if (selectedDivision == entityDivision) {
                      var option = document.createElement("option");
                      option.text = entityDescription;
                      option.value = entityName;
                      selectItemType.add(option);
                  }
              }
              
          }

          $('#select-item-type').Dropdown("refresh");

      } else {
          showNotification("Error!", "Item Type xml is empty!", messageType.ERROR);
      }

  }

  function enableElement(elm) {
      if (elm.classList.contains("is-disabled")) {
          elm.classList.remove("is-disabled");
      }
  }

  function disableElement(elm) {
      if (!elm.classList.contains("is-disabled")) {
          elm.classList.add("is-disabled");
      }
  }

  function enableAttributes() {
      $('#attribute-container div.ms-TextField').each(function (index, element) {
          enableElement(element);
      });
  }

  function disableAttributes() {
      $('#attribute-container div.ms-TextField').each(function (index, element) {
          disableElement(element);
      });
  }

  function enableAll() {
      // enable division selector
      enableElement(document.getElementById('select-division'));

      // enable item type selector
      enableElement(document.getElementById('select-item-type'));

      // enable attributes
      enableAttributes();

      // enable archive button
      var buttonArchive = document.getElementById("save-button-container").getElementsByTagName("button")[0];
      if (buttonArchive.hasAttribute("disabled")) {
          buttonArchive.removeAttribute("disabled");
      }
  }

  function disableAll() {
      // disable division selector
      disableElement(document.getElementById('select-division'));

      // disable item type selector
      disableElement(document.getElementById('select-item-type'));

      // disable attributes
      disableAttributes();

      // disable archive button
      var buttonArchive = document.getElementById("save-button-container").getElementsByTagName("button")[0];
      if (!buttonArchive.hasAttribute("disabled")) {
          buttonArchive.setAttribute("disabled", "disabled");
      }
  }

  function getEntitesXml() {

      var getEntitiesUrl = apiUrl + "getEntities.jsp";
      var timeOut = 10000;

      //var req = "https://eu-be-sos05:21106/ca/api/items/search/item?%24query=%2FM3_OI_INV%5B%40M3_DIVI%3D%22607%22%5D";
      var httpReq = new XMLHttpRequest();
      //httpReq.withCredentials = true;
      //httpReq.open("GET", getEntitiesURL, false);
      httpReq.open("GET", getEntitiesUrl, true);//, "AK1", "Ak2@Victaulic");
      httpReq.timeout = timeOut;
      //httpReq.setRequestHeader("Authorization", "Basic bXNydmFkbTpNM0FkbWluJQ==");
      //httpReq.setRequestHeader("Authorization", "Basic " + btoa("AK1:Ak#@Victaulic001"));

      httpReq.onreadystatechange = function () {
          if (httpReq.readyState === 4) {
              if (httpReq.status === 200) {
                  //console.log(httpReq.responseText);
                  xmlEntities = httpReq.responseXML;
                  enableElement($('#select-division')[0]);
              } else {
                  //console.log("Error", httpReq.statusText);
                  showNotification("Error!", "Failed to retrieve Item Type list. Reason: Reason: " + httpReq.statusText + ".", messageType.ERROR);
              }
          }
      }

      httpReq.ontimeout = function () {
          showNotification("Error!", "Failed to retrieve Item Type list. Reason: " + httpReq.statusText, messageType.ERROR);
      }

      httpReq.send();

  }

  // Clear item type list, except the first one
  function clearItemTypeList() {
      // var select = $('#select-division select')[0]; // alternative
      var select = document.getElementById("select-item-type").getElementsByTagName("select")[0];
      var selectLenght = select.options.length;
      for (var i = 1; i < selectLenght; i++) {
          select.remove(1);
      }

      // Refresh item type list
      $('#select-item-type').Dropdown("refresh");
      
  }

  // Clear attribute list
  function clearAttributeList() {
      // remove TextField components
      var textField = document.getElementById("attribute-container").getElementsByClassName("ms-TextField");
      var textFieldLenght = textField.length;
      for (var i = 0; i < textFieldLenght; i++) {
          textField[0].parentNode.removeChild(textField[0]);
      }

      // remove DatePicker components
      var datePicker = document.getElementById("attribute-container").getElementsByClassName("ms-DatePicker");
      var datePickerLenght = datePicker.length;
      for (var i = 0; i < datePickerLenght; i++) {
          datePicker[0].parentNode.removeChild(datePicker[0]);
      }
  }

  // Helper function for displaying notifications
  function showNotification(header, content, type) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);

    switch (type) {
        case messageType.SUCCESS:
            $("#notificationContent").css("background-color", "#dff6dd");
            break;
        case messageType.INFO:
            $("#notificationContent").css("background-color", "#f4f4f4");
            break;
        case messageType.WARNING:
            $("#notificationContent").css("background-color", "#fff4ce");
            break;
        case messageType.ERROR:
            $("#notificationContent").css("background-color", "#fde7e9");
            break;
        default:
            $("#notificationContent").css("background-color", "#f4f4f4");
    }

      

    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }

  function archiveEmail(event) {
      if (validateData()) {
          // disable controls
          disableAll();

          // spin around the world
          showSpinner();

          // Get mail content from Exchange server
          getMailContent();
      }      
  }

  // validates attributes data, entered by user
  function validateData() {
      var result = true;
      var resultFailedText = "";

      // go through TextField attribute containers (input) for validation
      // please see html structure of TextField and DatePicker components to be sure that JQuery is correct
      var attributeNodeList = $("#attribute-container div.ms-TextField");
      for (var i = 0; i < attributeNodeList.length; i++) {
          var elementContainer = attributeNodeList[i],
              name = elementContainer.getElementsByTagName("label")[0].textContent,
              flag = elementContainer.getElementsByTagName("input")[0].getAttribute("data-attribute-flag"),
              type = elementContainer.getElementsByTagName("input")[0].getAttribute("data-attribute-type"),
              value = elementContainer.getElementsByTagName("input")[0].value;

          resultFailedText = "";

          if (value === null || value == undefined || value == "") {
              // check if attribute is reuired
              if (flag == "10" || flag == "26") {

                  resultFailedText += 'Attribute "' + name + '" is required!';
                  result = false;
                  attributeNodeList[i].getElementsByTagName("input")[0].focus();
                  break;
              }
          } else {
              if ((value.indexOf("<") > -1) || (value.indexOf(">") > -1) || (value.indexOf("&") > -1) || (value.indexOf("\"") > -1)) {
                  resultFailedText += 'Attribute "' + name + '" value should not contain the following characters: <, >, ", &';
                  result = false;
                  attributeNodeList[i].getElementsByTagName("input")[0].focus();
                  break;
              }
          }
      }

      if (!result) {
          showNotification("Validation warning!", resultFailedText, messageType.WARNING);
          return false;
      } else {
          return true;
      }
  }

  function getDatePickerHtml(caption, name, type, flag) {
      var requiredText = (flag == "10" || flag == "26") ? '* Required' : '';
      var stringCode =
    //    '<div class="ms-DatePicker">' +
            '<div class="ms-TextField" id="' + name + '">' +
                '<label class="ms-Label">' + caption + '</label>' +
                '<i class="ms-DatePicker-event ms-Icon ms-Icon--event"></i>' +
                '<input class="ms-TextField-field" placeholder="Select a date…" type="text" data-attribute-type=' + type + ' data-attribute-flag=' + flag + ' >' +
                '<span class="ms-TextField-description">' + requiredText + '</span>' +
            '</div>' +
            '<div class="ms-DatePicker-monthComponents">' +
                '<span class="ms-DatePicker-nextMonth js-nextMonth"><i class="ms-Icon ms-Icon--chevronRight"></i></span>' +
                '<span class="ms-DatePicker-prevMonth js-prevMonth"><i class="ms-Icon ms-Icon--chevronLeft"></i></span>' +
                '<div class="ms-DatePicker-headerToggleView js-showMonthPicker"></div>' +
            '</div>' +
            '<span class="ms-DatePicker-goToday js-goToday">Go to today</span>' +
            '<div class="ms-DatePicker-monthPicker">' +
                '<div class="ms-DatePicker-header">' +
                    '<div class="ms-DatePicker-yearComponents">' +
                        '<span class="ms-DatePicker-nextYear js-nextYear"><i class="ms-Icon ms-Icon--chevronRight"></i></span>' +
                        '<span class="ms-DatePicker-prevYear js-prevYear"><i class="ms-Icon ms-Icon--chevronLeft"></i></span>' +
                    '</div>' +
                    '<div class="ms-DatePicker-currentYear js-showYearPicker"></div>' +
                '</div>' +
                '<div class="ms-DatePicker-optionGrid">' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="0">Jan</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="1">Feb</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="2">Mar</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="3">Apr</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="4">May</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="5">Jun</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="6">Jul</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="7">Aug</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="8">Sep</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="9">Oct</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="10">Nov</span>' +
                    '<span class="ms-DatePicker-monthOption js-changeDate" data-month="11">Dec</span>' +
                '</div>' +
            '</div>' +
            '<div class="ms-DatePicker-yearPicker">' +
                '<div class="ms-DatePicker-decadeComponents">' +
                    '<span class="ms-DatePicker-nextDecade js-nextDecade"><i class="ms-Icon ms-Icon--chevronRight"></i></span>' +
                    '<span class="ms-DatePicker-prevDecade js-prevDecade"><i class="ms-Icon ms-Icon--chevronLeft"></i></span>' +
                '</div>' +
            '</div>'; // +
    //    '</div>';

      return stringCode;

  }

  function getSubjectRequest(id) {
      // Return a GetItem operation request for the subject of the specified item. 
      var result =
   '<?xml version="1.0" encoding="utf-8"?>' +
   '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
   '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
   '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
   '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
   '  <soap:Header>' +
   '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
   '  </soap:Header>' +
   '  <soap:Body>' +
   '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
   '      <ItemShape>' +
   '        <t:IncludeMimeContent>true</t:IncludeMimeContent>' +
   //'        <t:BaseShape>AllProperties</t:BaseShape>' +
   '      </ItemShape>' +
   '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
   '    </GetItem>' +
   '  </soap:Body>' +
   '</soap:Envelope>';

      return result;
  }

  function getMailContent() {
      // Create a local variable that contains the mailbox.
      var mailbox = Office.context.mailbox;

      mailbox.makeEwsRequestAsync(getSubjectRequest(mailbox.item.itemId), callback);
  }

  function callback(asyncResult) {
      var response = asyncResult.value,
          context = asyncResult.context;

      if (response == null || response == "") {
          showNotification("Error!", "Error while retrieving data from Exchange: response is empty!", messageType.ERROR);
          enableAll();
          hideSpinner();
          return;
      }

      if (asyncResult.status == "failed") {
          showNotification("Error!", "Error while retrieving data from Exchange: request failed!", messageType.ERROR);
          enableAll();
          hideSpinner();
          return;
      }

      var xmlRes = jQuery.parseXML(response),
          mimeContent = xmlRes.getElementsByTagName("t:MimeContent")[0].childNodes[0].nodeValue,
          xml,
          entitySelected = document.getElementById("select-item-type").getElementsByTagName("select")[0].value,
          entityNode,
          item = Office.context.mailbox.item;

      // get 'entity' node for selected item type from xmlEntities
      var entityList = xmlEntities.getElementsByTagName("entity");
      for (var i = 0; i < entityList.length; i++) {
          var tmpEntityName = entityList[i].getElementsByTagName("name")[0].childNodes[0].nodeValue;
          if (tmpEntityName == entitySelected) {
              entityNode = entityList[i];
              break;
          }
      }

      // GENERATE XML
      // 

      // main tag
      xml = '<item>';

      // header (name only)
      // 
      // <entityName>M3_RMA_101</entityName>
      var entityName = entityNode.getElementsByTagName("name")[0].childNodes[0].nodeValue;
      xml += '<entityName>' + entityName + '</entityName>';

      // attribute list
      // 
      // <attrs>
      //     <attr>
      //         <name>M3_EXT_DocumentReference</name>
      //         <type>2</type>
      //         <qual>M3_EXT_DocumentReference</qual>
      //         <value>21342356</value>
      //     </attr>
      //     ...
      // </attrs>
      xml += '<attrs>';

      var attributeList = entityNode.getElementsByTagName("attrs")[0].getElementsByTagName("attr");
      for (var i = 0; i < attributeList.length; i++) {
          var attributeName = attributeList[i].getElementsByTagName("name")[0].childNodes[0].nodeValue;
          var attributeType = attributeList[i].getElementsByTagName("type")[0].childNodes[0].nodeValue;
          var attributeQual = attributeList[i].getElementsByTagName("qual")[0].childNodes[0].nodeValue;
          var attributeValue = document.getElementById(attributeName).getElementsByTagName("input")[0].value;

          // convert date value to YYYY-MM-DD format
          if (attributeType == 7) {
              var attrDate = new Date(attributeValue);
              attributeValue = attrDate.getFullYear() + '-' + ('0' + (attrDate.getMonth() + 1)).slice(-2) + '-' + ('0' + attrDate.getDate()).slice(-2);
          }

          var attrTag = '<attr>';
          attrTag += '<name>' + attributeName + '</name>';
          attrTag += '<type>' + attributeType + '</type>';
          attrTag += '<qual>' + attributeQual + '</qual>';
          attrTag += '<value>' + attributeValue + '</value>';
          attrTag += '</attr>';

          xml += attrTag;
      }
      xml += '</attrs>';

      // resourse list
      //
      // <resrs>
      //     <res>
      //         <entityName>ICMBASE</entityName>
      //         <mimetype>application/vnd.ms-outlook</mimetype>
      //         <base64>file-to-BASE64-string</base64>
      //         <filename>20163602-043645-Ticket.msg</filename>
      //     </res>
      // </resrs>
      var base64 = mimeContent;
      xml += '<resrs>';
      xml += '<res>';
      xml += '<entityName>ICMBASE</entityName>';
      xml += '<mimetype>application/vnd.ms-outlook</mimetype>';
      xml += '<base64>' + base64 + '</base64>';
      xml += '<filename>' + prepareFileName(Office.context.mailbox.item.subject.toString()) + '.eml' + '</filename>';
      xml += '</res>';
      xml += '</resrs>';

      // end main tag
      xml += '</item>';

      // GENERATE XML ENDED!

      callAddItemApi(xml);
  }

  function callAddItemApi(xml) {
      // call archive API

      var addItemExUrl = apiUrl + "addItemEx.jsp",
          timeOut = 30000,
          pid = "";

      var httpReq = new XMLHttpRequest();
      httpReq.open("POST", addItemExUrl, true);
      httpReq.timeout = timeOut;
      //httpReq.setRequestHeader("Authorization", "Basic " + btoa("AK1:Ak#@Victaulic001"));
      httpReq.onreadystatechange = function () {
          if (httpReq.readyState === 4) {
              if (httpReq.status === 200) {
                  console.log(httpReq.responseText);
                  pid = httpReq.responseXML.getElementsByTagName("item")[0].getElementsByTagName("pid")[0].textContent;

                  // check in item
                  callChekcInItemApi(pid);
              } else {
                  console.log("Error", httpReq.statusText);
                  showNotification("Error!", "Add item API failed! Reason: " + httpReq.statusText, messageType.ERROR);
                  enableAll();
                  hideSpinner();
              }
          }
      }

      httpReq.ontimeout = function () {
          showNotification("Error!", "Add item API failed! Reason: " + httpReq.statusText, messageType.ERROR);
      }

      httpReq.onload = function () {
          //showIDMPage();
      }

      httpReq.send(xml);
  }

  function callChekcInItemApi(pid) {
      var checkInItemExUrl = apiUrl + "checkInItem.jsp",
          timeOut = 10000,
          xml = '<item><pid>' + pid + '</pid></item>';

      var httpReq = new XMLHttpRequest();
      httpReq.open("POST", checkInItemExUrl, true);
      httpReq.timeout = timeOut;
      //httpReq.setRequestHeader("Authorization", "Basic " + btoa("AK1:Ak#@Victaulic001"));
      httpReq.onreadystatechange = function () {
          if (httpReq.readyState === 4) {
              if (httpReq.status === 200) {
                  console.log(httpReq.responseText);
                  showNotification("Completed!", "Email archived successfully!", messageType.SUCCESS);
              } else {
                  console.log("Error", httpReq.statusText);
                  showNotification("Error!", "Check In item API failed!", messageType.ERROR);
              }
          }
      }

      httpReq.ontimeout = function () {
          showNotification("Error!", "Check In item API timeout!", messageType.ERROR);
      }

      httpReq.onload = function () {
          enableAll();
          hideSpinner();
      }

      httpReq.send(xml);
  }

  function prepareFileName(fileName) {
      fileName = fileName.replace(/[^a-z0-9]/gi,"_");
      return fileName;
  }

  function showSpinner() {
      var spinnerContainer = document.getElementById("save-button-container").getElementsByTagName("div")[0];

      if (typeof (spinner) == 'undefined') {
          spinner = new fabric.Spinner(spinnerContainer);
      } else {
          spinnerContainer.getElementsByClassName('ms-Spinner')[0].style.visibility = 'visible';
          spinner.start();
      }
  }

  function hideSpinner() {
      var spinnerContainer = document.getElementById('save-button-container').getElementsByTagName('div')[0];

      if (typeof (spinner) != 'undefined') {
          spinnerContainer.getElementsByClassName('ms-Spinner')[0].style.visibility = 'hidden';
          spinner.stop();
      }
  }

})();