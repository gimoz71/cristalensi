<script language="JavaScript">
//----- Editor Initialization ------
window.onerror = handleErrors;
function handleErrors()
{
   //----- Used For Browsers That Don't Want To Behave -----
   return true;
}
var viewMode = 1;
function loadEditor_en()
{
  //----- Modify User Controls -----
  showWYSIWYGCtrl.style.display = 'none';
  ///hr6Ctrl.style.display = 'none';
  editbox_en.document.designMode="On";
  
  var datVal_en = "<%=testo_en%>"
  editbox_en.document.open();
  editbox_en.document.write(datVal_en);
  editbox_en.document.close();

}


function eButton(cmdButton, buttonval)
{
   //----- Controls Button Behaviors ------
   if (buttonval == "over")
   {
      cmdButton.style.backgroundColor = "threedhighlight";
      cmdButton.style.borderColor = "threeddarkshadow threeddarkshadow threeddarkshadow threeddarkshadow";
   }
   else if (buttonval == "out")
   {
      cmdButton.style.backgroundColor = "threedface";
      cmdButton.style.borderColor = "threedface";
   }
   else if (buttonval == "down")
   {
      cmdButton.style.backgroundColor = "threedlightshadow";
      cmdButton.style.borderColor = "threedshadow threedshadow threedshadow threedshadow";
   }
   else if (buttonval == "up")
   {
      cmdButton.style.backgroundColor = "threedhighlight";
      cmdButton.style.borderColor = "threedshadow threedshadow threedshadow threedshadow";
      cmdButton = null;
   }
   else
   {
      return;
   }
}
function newDocument()
{
   //----- Creates An Empty Workspace ------
   if (editbox_en.document.body.innerHTML == "")
   {
      editbox_en.document.execCommand('refresh', false, null);
   }
   else
   {
      if (confirm("Would you like to save your entry?"))
      {
         //var saveInsructions = "Click the back eButton in your browser \n" +
                               "once your input has been saved inorder \n" +
                               "to continue using the editor.";
         //alert(saveInsructions)
         var dataRep = null;
         dataRep = document.body.all.submitData_en;
         dataRep.value = editbox_en.document.body.innerHTML;
         document.newsform.submit();
         window.location.reload();
      }
      else
      {
         editbox_en.document.execCommand('refresh', false, null);
      }
   }
}

function saveDocument()
{
   //----- Saves User Input ------
   //==============================================
   //= To change the url that the editor posts to =
   //= change the action url in the form at the   =
   //= bottom of this page.                       =
   //==============================================
   //if (editbox_en.document.body.innerHTML == "")
   //{
    //  return;
   //}
  // else
  // {
      //if (confirm("Would you like to save you entry?"))
      //{
         var dataRep = null;
         dataRep = document.body.all.submitData_en;
         dataRep.value = editbox_en.document.body.innerHTML;
		 input=document.newsform.form_submission_en.value;
         output = "";
		 for (var i = 0; i < input.length; i++) {
		 	if ((input.charCodeAt(i) == 13) && (input.charCodeAt(i + 1) == 10)) {
				i++;
				output += "";
		 	}
		else 
			output += input.charAt(i);
		}
		document.newsform.form_submission_en.value=output
        //document.newsform.submit();
      //}
      //else
     // {
      //   return;
     // }
  // }
}
function tableDialog()
{
   //----- Creates A Table Dialog And Passes Values To createTable() -----
   var rtNumRows = null;
   var rtNumCols = null;
   var rtTblAlign = null;
   var rtTblWidth = null;
   showModalDialog("table.htm",window,"status:false;dialogWidth:16em;dialogHeight:14em");
}
function createTable()
{
   //----- Creates User Defined Tables -----
   var cursor = editbox_en.document.selection.createRange();
   if (rtNumRows == "" || rtNumRows == "0")
   {
      rtNumRows = "1";
   }
   if (rtNumCols == "" || rtNumCols == "0")
   {
      rtNumCols = "1";
   }
   var rttrnum=1
   var rttdnum=1
   var rtNewTable = "<table border='1' align='" + rtTblAlign + "' cellpadding='0' cellspacing='0' width='" + rtTblWidth + "'>"
   while (rttrnum <= rtNumRows)
   {
      rttrnum=rttrnum+1
      rtNewTable = rtNewTable + "<tr>"
      while (rttdnum <= rtNumCols)
      {
         rtNewTable = rtNewTable + "<td>&nbsp;</td>"
         rttdnum=rttdnum+1
      }
      rttdnum=1
      rtNewTable = rtNewTable + "</tr>"
   }
   rtNewTable = rtNewTable + "</table>"
   cursor.pasteHTML(rtNewTable);
   editbox_it.focus();
}
function foreColor()
{
   //----- Sets Foreground Color -----
   var fColor = showModalDialog("color.htm","","dialogWidth:140px; dialogHeight:120px" );
   if (fColor != null)
   {
      editbox_en.document.execCommand("ForeColor", false, fColor);
   }
   editbox_en.focus();
}
function backColor()
{
   //----- Sets Background Color -----
   var bColor = showModalDialog("color.htm","","dialogWidth:140px; dialogHeight:120px" );
   if (bColor != null)
   {
      editbox_en.document.execCommand("BackColor", false, bColor);
   }
   editbox_en.focus();
}
function eStat(status)
{
   //----- Updates Status Bar With Information -----
   var editStat = document.getElementById("editorStatus");
   editStat.innerHTML = status;
}
function modeSelect()
{
   //----- Changes Editor Mode -----
   var HTMLtitle
   var WYSIWYGtitle
   var editorTitle
   if(viewMode == 1)
   {
      //----- Convert WYSIWYG editor to HTML -----
      iHTML = editbox_en.document.body.innerHTML; editbox_en.document.body.innerText = iHTML;
      HTMLtitle =".: Control Panel :."; editorTitle = document.getElementById("editorTitle");
      editorTitle.innerHTML = HTMLtitle; document.title = ".: Control Panel :.";
      linkCtrl.style.display = 'none';
      lineCtrl.style.display = 'none'; tableCtrl.style.display = 'none';
      //hr1Ctrl.style.display = 'none'; 
	  orderedCtrl.style.display = 'none';
      unorderedCtrl.style.display = 'none'; //hr2Ctrl.style.display = 'none';
      strikeCtrl.style.display = 'none'; subCtrl.style.display = 'none';
      superCtrl.style.display = 'none'; //hr3Ctrl.style.display = 'none';
      forecolorCtrl.style.display = 'none'; backcolorCtrl.style.display = 'none';
      //hr4Ctrl.style.display = 'none'; 
	  indentCtrl.style.display = 'none';
      outdentCtrl.style.display = 'none'; //hr5Ctrl.style.display = 'none';
      showWYSIWYGCtrl.style.display = 'inline'; showWYSIWYGCtrl2.style.display = 'none';//hr6Ctrl.style.display = 'inline';
      toolBar1.style.display = 'none'; //newCtrl.style.display = 'none';
      editbox_en.focus(); saveCtrl.style.display = 'inline'; 
      viewMode = 2;
   }
   else
   {
      //----- Convert HTML editor to WYSIWYG -----
      iText = editbox_en.document.body.innerText; editbox_en.document.body.innerHTML = iText;
      WYSIWYGtitle =".: Control Panel :." ; editorTitle = document.getElementById("editorTitle");
      editorTitle.innerHTML = WYSIWYGtitle; document.title = ".: Control Panel :."
      linkCtrl.style.display = 'inline';
      lineCtrl.style.display = 'inline'; tableCtrl.style.display = 'inline';
      //hr1Ctrl.style.display = 'inline';
	  orderedCtrl.style.display = 'inline';
      unorderedCtrl.style.display = 'inline'; //hr2Ctrl.style.display = 'inline';
      strikeCtrl.style.display = 'inline'; subCtrl.style.display = 'inline';
      superCtrl.style.display = 'inline'; //hr3Ctrl.style.display = 'inline';
      forecolorCtrl.style.display = 'inline'; backcolorCtrl.style.display = 'inline';
      //hr4Ctrl.style.display = 'inline'; 
	  indentCtrl.style.display = 'inline';
      outdentCtrl.style.display = 'inline'; //hr5Ctrl.style.display = 'inline';
      showWYSIWYGCtrl.style.display = 'none'; showWYSIWYGCtrl2.style.display = 'inline';//hr6Ctrl.style.display = 'none';
      toolBar1.style.display = 'inline'; //newCtrl.style.display = 'inline';
      editbox_en.focus(); saveCtrl.style.display = 'inline'; 
      viewMode = 1;
   }
}
</script>