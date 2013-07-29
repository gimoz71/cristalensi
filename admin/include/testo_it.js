<script language="JavaScript">
//----- Editor Initialization ------
window.onerror = handleErrors_it;
function handleErrors_it()
{
   //----- Used For Browsers That Don't Want To Behave -----
   return true;
}
var viewMode_it = 1;
function loadEditor_it()
{
  //----- Modify User Controls -----
  showWYSIWYGCtrl_it.style.display = 'none';
  ///hr6Ctrl.style.display = 'none';
  editbox_it.document.designMode="On";
  
  var datVal_it = "<%=testo_it%>"
  editbox_it.document.open();
  editbox_it.document.write(datVal_it);
  editbox_it.document.close();

}


function eButton_it(cmdButton_it, buttonval_it)
{
   //----- Controls Button Behaviors ------
   if (buttonval_it == "over")
   {
      cmdButton_it.style.backgroundColor = "threedhighlight";
      cmdButton_it.style.borderColor = "threeddarkshadow threeddarkshadow threeddarkshadow threeddarkshadow";
   }
   else if (buttonval_it == "out")
   {
      cmdButton_it.style.backgroundColor = "threedface";
      cmdButton_it.style.borderColor = "threedface";
   }
   else if (buttonval_it == "down")
   {
      cmdButton_it.style.backgroundColor = "threedlightshadow";
      cmdButton_it.style.borderColor = "threedshadow threedshadow threedshadow threedshadow";
   }
   else if (buttonval_it == "up")
   {
      cmdButton_it.style.backgroundColor = "threedhighlight";
      cmdButton_it.style.borderColor = "threedshadow threedshadow threedshadow threedshadow";
      cmdButton_it = null;
   }
   else
   {
      return;
   }
}
function newDocument_it()
{
   //----- Creates An Empty Workspace ------
   if (editbox_it.document.body.innerHTML == "")
   {
      editbox_it.document.execCommand('refresh', false, null);
   }
   else
   {
      if (confirm("Would you like to save your entry?"))
      {
         //var saveInsructions = "Click the back eButton in your browser \n" +
                               "once your input has been saved inorder \n" +
                               "to continue using the editor.";
         //alert(saveInsructions)
         var dataRep_it = null;
         dataRep_it = document.body.all.submitData_it;
         dataRep_it.value = editbox_it.document.body.innerHTML;
         document.newsform.submit();
         window.location.reload();
      }
      else
      {
         editbox_it.document.execCommand('refresh', false, null);
      }
   }
}

function saveDocument_it()
{
   //----- Saves User Input ------
   //==============================================
   //= To change the url that the editor posts to =
   //= change the action url in the form at the   =
   //= bottom of this page.                       =
   //==============================================
   //if (editbox_it.document.body.innerHTML == "")
  // {
      //return;
  // }
  // else
  // {
      //if (confirm("Would you like to save you entry?"))
      //{
         var dataRep_it = null;
         dataRep_it = document.body.all.submitData_it;
         dataRep_it.value = editbox_it.document.body.innerHTML;
		 input_it=document.newsform.form_submission_it.value;
         output_it = "";
		 for (var i = 0; i < input_it.length; i++) {
		 	if ((input_it.charCodeAt(i) == 13) && (input_it.charCodeAt(i + 1) == 10)) {
				i++;
				output_it += "";
		 	}
		else 
			output_it += input_it.charAt(i);
		}
		document.newsform.form_submission_it.value=output_it
        document.newsform.submit();
      //}
      //else
     // {
      //   return;
     // }
   //}
}
function tableDialog_it()
{
   //----- Creates A Table Dialog And Passes Values To createTable() -----
   var rtNumRows = null;
   var rtNumCols = null;
   var rtTblAlign = null;
   var rtTblWidth = null;
   showModalDialog("table.htm",window,"status:false;dialogWidth:16em;dialogHeight:14em");
}
function createTable_it()
{
   //----- Creates User Defined Tables -----
   var cursor_it = editbox_it.document.selection.createRange();
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
   cursor_it.pasteHTML(rtNewTable);
   editbox_it.focus();
}
function foreColor_it()
{
   //----- Sets Foreground Color -----
   var fColor = showModalDialog("color.htm","","dialogWidth:140px; dialogHeight:120px" );
   if (fColor != null)
   {
      editbox_it.document.execCommand("ForeColor", false, fColor);
   }
   editbox_it.focus();
}
function backColor_it()
{
   //----- Sets Background Color -----
   var bColor = showModalDialog("color.htm","","dialogWidth:140px; dialogHeight:120px" );
   if (bColor != null)
   {
      editbox_it.document.execCommand("BackColor", false, bColor);
   }
   editbox_it.focus();
}
function eStat_it(status_it)
{
   //----- Updates Status Bar With Information -----
   var editStat_it = document.getElementById("editorStatus_it");
   editStat_it.innerHTML = status_it;
}
function modeSelect_it()
{
   //----- Changes Editor Mode -----
   var HTMLtitle_it
   var WYSIWYGtitle_it
   var editorTitle_it
   if(viewMode_it == 1)
   {
      //----- Convert WYSIWYG editor to HTML -----
      iHTML_it = editbox_it.document.body.innerHTML; editbox_it.document.body.innerText = iHTML_it;
      HTMLtitle_it =".: Control Panel :."; editorTitle_it = document.getElementById("editorTitle");
      editorTitle_it.innerHTML = HTMLtitle_it; document.title = ".: Control Panel :.";
      linkCtrl_it.style.display = 'none';
      lineCtrl_it.style.display = 'none'; tableCtrl_it.style.display = 'none';
      //hr1Ctrl.style.display = 'none'; 
	  orderedCtrl_it.style.display = 'none';
      unorderedCtrl_it.style.display = 'none'; //hr2Ctrl.style.display = 'none';
      strikeCtrl_it.style.display = 'none'; subCtrl_it.style.display = 'none';
      superCtrl_it.style.display = 'none'; //hr3Ctrl.style.display = 'none';
      forecolorCtrl_it.style.display = 'none'; backcolorCtrl_it.style.display = 'none';
      //hr4Ctrl.style.display = 'none'; 
	  indentCtrl_it.style.display = 'none';
      outdentCtrl_it.style.display = 'none'; //hr5Ctrl.style.display = 'none';
      showWYSIWYGCtrl_it.style.display = 'inline'; showWYSIWYGCtrl2_it.style.display = 'none';//hr6Ctrl.style.display = 'inline';
      toolBar1_it.style.display = 'none'; //newCtrl.style.display = 'none';
      editbox_it.focus(); //saveCtrl_it.style.display = 'inline'; 
      viewMode_it = 2;
   }
   else
   {
      //----- Convert HTML editor to WYSIWYG -----
      iText_it = editbox_it.document.body.innerText; editbox_it.document.body.innerHTML = iText_it;
      WYSIWYGtitle_it =".: Control Panel :." ; editorTitle_it = document.getElementById("editorTitle");
      editorTitle_it.innerHTML = WYSIWYGtitle_it; document.title = ".: Control Panel :."
      linkCtrl_it.style.display = 'inline';
      lineCtrl_it.style.display = 'inline'; tableCtrl_it.style.display = 'inline';
      //hr1Ctrl.style.display = 'inline';
	  orderedCtrl_it.style.display = 'inline';
      unorderedCtrl_it.style.display = 'inline'; //hr2Ctrl.style.display = 'inline';
      strikeCtrl_it.style.display = 'inline'; subCtrl_it.style.display = 'inline';
      superCtrl_it.style.display = 'inline'; //hr3Ctrl.style.display = 'inline';
      forecolorCtrl_it.style.display = 'inline'; backcolorCtrl_it.style.display = 'inline';
      //hr4Ctrl.style.display = 'inline'; 
	  indentCtrl_it.style.display = 'inline';
      outdentCtrl_it.style.display = 'inline'; //hr5Ctrl.style.display = 'inline';
      showWYSIWYGCtrl_it.style.display = 'none'; showWYSIWYGCtrl2_it.style.display = 'inline';//hr6Ctrl.style.display = 'none';
      toolBar1_it.style.display = 'inline'; //newCtrl.style.display = 'inline';
      editbox_it.focus(); //saveCtrl_it.style.display = 'inline'; 
      viewMode_it = 1;
   }
}
</script>