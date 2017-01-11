#include <AD.au3>
#include <array.au3>
#include <IE.au3>
#include <Constants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <GUIConstantsEx.au3>
#include <guiconstants.au3>
#include <GuiComboBox.au3>
#include <WindowsConstants.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include '_SelfUpdate.au3'
#include <Misc.au3>

If _Singleton("MyScriptName", 1) = 0 Then
	; If successful, running our script a second time should cause us to fall through here
	MsgBox($MB_ICONERROR, "User Generated Error Message", "Error: This script is already running!")
	Exit
EndIf

Global $oCOMErrorHandler = ObjEvent("AutoIt.Error", _User_ErrFunc)
_IEErrorHandlerRegister(_User_ErrFunc)

#Region ;Variable declaration
Global $ClippArray[5][2]
$ClippArray[0][0] = "F1"
$ClippArray[1][0] = "F2"
$ClippArray[2][0] = "F3"
$ClippArray[3][0] = "F4"
$ClippArray[4][0] = "F5"
Global $sToggle = "Provisioning"
Global $start = ""
Global Const $jsMainApp = 'var MainAppWin="undefined"!=typeof g_form?window:document.getElementById("gsft_main").contentWindow;'
Global $sCC, $sBCC = ""
Global Const $iButtonwidth = 120
Global Const $iButtonheight = 26
Global Const $iButtonofffset = $iButtonheight + 1
Global Const $iHeight = 752
Global $hGui, $hCombo = ""
Global $button1, $button2, $button3, $button4, $button5, $button6, $button7 = ""
Global $button8, $button9, $button10, $button11, $button12, $button13, $button14 = ""
Global $button15, $button16, $button17, $button18, $button19, $button20, $button21 = ""
Global $button22, $button23, $button24, $button25, $button26, $button27
#EndRegion ;Variable declaration

HotKeySetter()
DrawGUI()
CheckIni()

While 1
	Local $oIE, $string, $to, $address = ""
	GuiPosition()
	Local $msg = GUIGetMsg()
	If $sToggle = "Provisioning" Then ;Tasks
		Switch $msg
			Case -3
				Update()
				Exit
			Case $button1 ; Email
				Email()
			Case $button2 ;NTR
				NTR()
			Case $button3 ;CNTR
				CNTR()
			Case $button4 ;Info
				NeedInfo()
			Case $button5 ;AD Groups
				$oIE = GetIe()
				CheckAD($oIE)
			Case $button6 ;Approver DB
				$oIE = GetIe()
				AppDB($oIE)
			Case $button7 ; Pull Approver
				$oIE = GetIe()
				Pullapprover($oIE)
				$string = 'document.getElementById("sysverb_insert_and_stay").click()'
				MainApp($oIE, $string)
			Case $button8 ;CTASK
				Ctask()
			Case $button9 ;Approval
				$string = $jsMainApp & 'MainAppWin.document.getElementById("x_human_access_task.sysapproval_approver.sysapproval_choice_actions").scrollIntoView(),MainAppWin.document.getElementById("allcheck_x_human_access_task.sysapproval_approver.sysapproval").checked=!0,MainAppWin.document.querySelector(".list2_body .input-group-checkbox .checkbox").checked=!0;var grabmenu=MainAppWin.document.querySelector(".list_action_option");grabmenu.selectedIndex=4,grabmenu.onchange();'
				MainApp("", $string)
			Case $button10 ;Clone
				CloneTask()
			Case $button11 ;Populate task
				$oIE = GetIe()
				PopTask($oIE)
			Case $button12 ;ID
				$string = $jsMainApp & 'desc=MainAppWin.g_form.getValue("description"),grp=desc.split("--- Original request below ---")[0],grp=grp.match(/ {1,}- ([A-Z0-9_\- ~&–.]{7,})\w| {1,}- (g|G)_([A-Z0-9_\- ~&–.]{1,5})\w/gi).join(";"),grp=grp.replace(/ {1,}- /g,"");var ids=desc.match(/( {1,}- \w{3}\d{4})/g).join(";");ids=ids.replace(/\s-\s/g,"").replace(/;;/g,";"),MainAppWin.g_form.addInfoMessage(ids+"<br><br>(|(sAMAccountName="+ids.replace(/;/g,")(sAMAccountName=")+"))<br><br>"+grp),MainAppWin.g_form.setValue("work_notes","Added "+ids+" to "+grp),MainAppWin.g_form.setValue("close_notes","The receiver(s) will need to restart his/her workstation(s) for the changes to take effect.");'
				MainApp("", $string)
				WinActivate("Find Users, Contacts, and Groups")
				WinActivate("Find Custom Search")
				WinActivate("[REGEXPTITLE:(Vintela Active Directory Users and Computers.*)]")
			Case $button13
				CheckSQL()
			Case $button14
				CheckReceivers()
			Case $button15 ;Create DL
				DL()
			Case $button16 ;Notes Prep
				$string = $jsMainApp & 'if("x_human_access_task"==MainAppWin.g_form.tableName){var n=document.getElementById("sys_readonly.x_human_access_task.number").value,a=[],t=new MainAppWin.GlideRecord("x_human_access_task");if(t.addQuery("sys_id",MainAppWin.g_form.getUniqueValue()),t.query(),t.next()){var s=new MainAppWin.GlideRecord("x_human_access_request");if(s.addQuery("sys_id",t.parent),s.query(),s.next()){var r=new MainAppWin.GlideRecord("sys_user"),u="^sys_idIN"+s.getValue("opened_for");for(""!==s.getValue("receivers")&&(u+=","+s.getValue("receivers")),u+="^EQ",r.encodedQuery=u,r.query();r.next();)a.push(r.getValue("user_name")+";"+(r.getValue("u_nickname")||r.getValue("first_name"))+";"+r.getValue("last_name")+";"+r.getValue("email")+";"+n)}var d="";d=a.length>0?a.map(function(e){return""+e}).join("<br>")+"<br>":"",MainAppWin.g_form.addInfoMessage(d)}}else alert("This can only be used on ART tickets");'
				MainApp("", $string)
			Case $button17 ;Format SQL
				FormatSQL()
			Case $button18 ;Format AD
				FormatAD()
			Case $button19
				$oIE = GetIe()
				Local $results = ReceiverList($oIE, False)
				SetClip($results)
			Case $button20 ;Bulk IIQ
				$string = 'function doXHR(e){var e=e||{};if(e.url=e.url||void 0,e.method=e.method||"GET",e.async=e.async||!1,"undefined"!=typeof e.url){var t=new XMLHttpRequest;return t.open(e.method,e.url,e.async),t.setRequestHeader("Content-Type","application/json"),t.send(),t.responseText}}function bulkAdd(e){var e=e||{};e.usernames=e.usernames||[],e.delay=e.delay||400,0==e.usernames.length&&alert("Error: bulkAdd() - No usernames specified");var t=setInterval(function(){if(0==e.usernames.length)return void clearInterval(t);for(var r=e.usernames.shift().trim(),n=JSON.parse(doXHR({url:"/include/identityQuery.json?query="+r+"&limit=10&start=0"})),s=n.identities,a=0;a<s.length;a++){var i=s[a];SailPoint.LCM.ChooseIdentities.addIdentities(i.id)}},e.delay)}try{var usernames=prompt("Enter a comma- or semicolon-separated list of usernames to add");usernames&&bulkAdd({usernames:usernames.split(/[,;]/g).filter(function(e){return""!==e}).map(function(e){return e.trim()}),delay:100})}catch(ex){console.log(ex.stack),alert(ex.stack)}'
				MainApp("", $string)
			Case $button21 ;Bulk IIQ
				SetRolesIIQ()
			Case $button22 ;Generate Password
				Password()
			Case $button23
				CheckEmail()
			Case $button24
				ManagerLookUp()
			Case $button25
				FolderResearch()
			Case $button26

			Case $button27
				FinalizeTask()
			Case $hCombo
				Flip()
		EndSwitch
	ElseIf $sToggle = "Research" Then ;Research
		Switch $msg
			Case -3
				Update()
				Exit
			Case $button1 ; Email
				Email()
			Case $button2 ;NTR
				NTR()
			Case $button3 ;CNTR
				CNTR()
			Case $button4 ;Info
				NeedInfo()
			Case $button5 ; DNC
				$string = $jsMainApp & 'var c=["d6c7bc6d4f2cd2403cda6cd18110c74a","d842f4e14f2cd2403cda6cd18110c7bb","d4c3f4a54f2cd2403cda6cd18110c720","c240b8ed4fe8d2403cda6cd18110c790","dfc0b8214f2cd2403cda6cd18110c71b","bad9f4614f6cd2403cda6cd18110c7a7","9fd9f4614f6cd2403cda6cd18110c7f3","6b507ced4fe8d2403cda6cd18110c745","a7d938614f6cd2403cda6cd18110c700","124938214f6cd2403cda6cd18110c725","7353f8654f2cd2403cda6cd18110c713","8171bc614f2cd2403cda6cd18110c797","86c078214f2cd2403cda6cd18110c791","269a7ca14f6cd2403cda6cd18110c708","aa9a7ca14f6cd2403cda6cd18110c707","d484bce54f2cd2403cda6cd18110c7aa","bd8130a14f2cd2403cda6cd18110c7fe","2d9a3ca14f6cd2403cda6cd18110c790","b49a3ca14f6cd2403cda6cd18110c746","478170a14f2cd2403cda6cd18110c78f","0f633c654f2cd2403cda6cd18110c7aa","c12670e94f2cd2403cda6cd18110c7c8"];if(d=[MainAppWin.g_form.getValue("opened_for")].concat(MainAppWin.g_form.getValue("receivers").split(",")).filter(function(a){return""!==a}),a=[],d.length>0){var s={},n={},f=new MainAppWin.GlideRecord("sys_user");for(f.encodedQuery="sys_class_name=sys_user^sys_idIN"+d.join(",")+"^EQ",f.query();f.next();)s[f.getValue("sys_id")]={sys_id:f.getValue("sys_id"),name:(f.getValue("u_nickname")||f.getValue("first_name"))+" "+f.getValue("last_name"),tsoid:f.getValue("user_name"),manager:f.getValue("manager")},-1==a.indexOf(f.getValue("manager"))&&a.push(f.getValue("manager"));for(f=new MainAppWin.GlideRecord("sys_user"),f.encodedQuery="sys_class_name=sys_user^sys_idIN"+a.join(",")+"^EQ",f.query();f.next();)n[f.getValue("sys_id")]={sys_id:f.getValue("sys_id"),name:(f.getValue("u_nickname")||f.getValue("first_name"))+" "+f.getValue("last_name"),tsoid:f.getValue("user_name"),manager:f.getValue("manager")};for(var r=[],t=!1,g=0;g<a.length;g++){var u=n[a[g]],i=c.indexOf(u.sys_id)>-1;i&&(t=!0),r.push((i?''< span style = "background: #aa0000; color: #ffffff; padding: 2px;" >'':"")+u.name+" ("+u.tsoid+")"+(i?"</span>":""));for(var o=[],l=0;l<d.length;l++){var _=s[d[l]];_.manager==u.sys_id&&o.push("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+_.name+" ("+_.tsoid+")")}r=r.concat(o)}MainAppWin.g_form.addInfoMessage(r.join("<br>"))}else MainAppWin.g_form.addErrorMessage("Error: No receivers were listed");'
				MainApp("", $string)
			Case $button6 ;Format AD
				FormatAD()
			Case $button7 ;AD Groups
				CheckAD()
			Case $button8 ;Format SQL
				FormatSQL()
			Case $button9 ;Format Unix
				$string = $jsMainApp & 'MainAppWin.g_form.setValue("short_description","Grant access - Unix");var text=MainAppWin.g_form.getValue("description");text=text.replace(/\n\n/g,"\n").replace(/\? Yes or No\?/g,":").replace(/If “Yes”, |If “No”, | \(Mark One\)/g,"").replace(/IF desired group is not listed, please specify:/g,"Specify group:").replace(/Develop:\s{0,}(X|x)/g,"Develop ").replace(/DBA:\s{0,}X/g,"DBA ").replace(/SYSADMIN:\s{0,}X/g,"SYSADMIN").replace(/Develop:/g,"").replace(/DBA:/g,"").replace(/SYSADMIN:/g,"").replace(/,\s{0,}/g,"\n                                ").replace(/[(]{0,1}[0-9]{3}[)]{0,1}[-\s{1,}\.]{0,1}[0-9]{3}[-\s{1,}\.]{0,1}[0-9]{4}/g,"").replace(/Phone Number:\t{0,}\n/g,"").replace(/Unix Server\s{1,}:/g,"Unix Server:").replace(/Humana Employee:.*\n/g,"").replace(/Vendor Name:.*\n/g,"").replace(/5-Digit Facility ID:.*\n/g,""),MainAppWin.g_form.setValue("description",text);'
				$sToggle = "Provisioning"
				MainApp("", $string)
				$sToggle = "Research"
				Flip()
			Case $button10 ;Contractor Email
				$oIE = CopyReceivers()
				$string = $jsMainApp & 'MainAppWin.g_form.setValue("short_description","Contractor Email");var groups=MainAppWin.g_form.getValue("description");MainAppWin.g_form.setValue("description","ID: "+"Please setup an email address for the contractor(s)\n"+MainAppWin.g_form.getValue("description"));MainAppWin.g_form.setValue("cmdb_ci","23f3922f139152007c1331a63244b064")'
				MainApp($oIE, $string)
			Case $button11 ;CTASK
				Ctask()
			Case $button12 ;Copy
				$string = $jsMainApp & 'if("x_human_access_request"==MainAppWin.g_form.tableName){var gr=new MainAppWin.GlideRecord("x_human_access_task");if(gr.addQuery("active",!0),gr.addQuery("parent",MainAppWin.g_form.getUniqueValue()),gr.query(),0==gr.rows.length)alert("No tasks");else if(gr.rows.length>1){MainAppWin.g_form.setValue("description","");for(var i=0;gr.next();)i++,MainAppWin.g_form.setValue("short_description",gr.short_description),MainAppWin.g_form.setValue("description",(MainAppWin.g_form.getValue("description")+"\n\n"+i+". "+gr.description).trim())}else gr.next(),MainAppWin.g_form.setValue("short_description",gr.short_description),MainAppWin.g_form.setValue("description",gr.description)}else if("x_human_access_task"==MainAppWin.g_form.tableName){var gr=new MainAppWin.GlideRecord("x_human_access_request");gr.addQuery("sys_id",MainAppWin.g_form.getValue("parent")),gr.query(),gr.next(),MainAppWin.g_form.setValue("priority",gr.getValue("priority")),MainAppWin.g_form.setValue("short_description",gr.getValue("short_description")),MainAppWin.g_form.setValue("description",gr.getValue("description")),MainAppWin.g_form.setMandatory("cmdb_ci",!0);var services={"8d2e08204f2e96006d3b97dd0210c762":[/(Domain group membership.*RX1AD)/i],"6ff3922f139152007c1331a63244b063":[/(Domain group creation)/i,/(Domain group membership)/i,/Grant access - AD/i,/Requesting Security Access for Local or Development Administrative Rights./i],"23f3922f139152007c1331a63244b064":[/(Email account)/i],a3f3922f139152007c1331a63244b064:[/(Distribution group)/i],"67f3922f139152007c1331a63244b065":[/(EDW access)/i,/(Request: EDW)/i],"67f3922f139152007c1331a63244b064":[/(GOOD - Access\\Software Request - Security Setup Request.)/i],"5dfdc4204f2e96006d3b97dd0210c73d":[/(Lync Group Chat)/i],a7f3922f139152007c1331a63244b065:[/(Netezza access)/i],"27f3922f139152007c1331a63244b065":[/(Oracle Financials Security Request)/i],"6bf3922f139152007c1331a63244b064":[/(Shared mailbox)/i],e7f3922f139152007c1331a63244b065:[/Grant access - SQL/i],"2bf3922f139152007c1331a63244b065":[/Grant access - Unix/i],"59c06dd4137ce60002bb3d576144b082":[/CompBenefits AS400/i],"23f3922f139152007c1331a63244b064":[/Contractor Email/i],aa5e99d4137ce60002bb3d576144b0ee:[/QicLink/i],eff3922f139152007c1331a63244b064:[/Request: FIMMAS Security.*/i]};e:for(var key in services)for(var vals=services[key],l=vals.length,i=0;l>i;i++)if(val=vals[i],MainAppWin.g_form.getValue("short_description").match(val)){MainAppWin.g_form.setValue("cmdb_ci",key),MainAppWin.g_form.getElement("sys_readonly.x_human_access_task.number").focus();break e}}'
				MainApp("", $string)
			Case $button13 ;DDA
				DDA()
			Case $button14 ; Copy Receivers
				CopyReceivers()
			Case $button15 ; Pop Tasks
				$string = $jsMainApp & 'MainAppWin.g_form.setValue("cmdb_ci","6ff3922f139152007c1331a63244b063"),MainAppWin.g_form.getElement("sys_readonly.x_human_access_task.number").focus(),MainAppWin.g_form.setValue("assigned_to"," 1ee87ced4f2cd2403cda6cd18110c78a");'
				MainApp("", $string)
			Case $button16 ; Clone AR
				CloneAR()
			Case $button17 ; CheckReceivers
				CheckReceivers()
			Case $button18 ; Chameleon & ECW Template
				ChameleonECW()
			Case $button19 ; Shrink
				Replace("/\n\n/g,'\n'")
			Case $button20 ; New Lines
				Replace("/,|;/g,'\n'")
			Case $button21 ; Add Tabs
				Replace("/\t/g,'\t\t\t\t\t\t'")
			Case $button22 ; Approver DB
				$oIE = GetIe()
				AppDB($oIE)
			Case $button23
				ManagerLookUp()
			Case $button24
				CheckEmail()
			Case $button25
				FolderResearch()
			Case $button26
				$string = $jsMainApp & 'gr=new MainAppWin.GlideRecord("sys_user"),gr.addQuery("sys_id",MainAppWin.g_form.getValue("opened_by")),gr.query(),gr.next();var firstN=gr.getValue("first_name");MainAppWin.g_form.setValue("work_notes","I am cancelling the request. I have emailed instruction to the requestor on how to proceed."),MainAppWin.g_form.setValue("substate","DNK Communication");'
				MainApp("", $string)
			Case $button27

			Case $hCombo
				Flip()
		EndSwitch
	ElseIf $sToggle = "Comments" Then ;Comments
		Switch $msg
			Case -3
				Update()
				Exit
			Case $button1
				$string = 'I created the chat room, control group and added the owner.'
				GetSetValues("", "work_notes", $string)
			Case $button2
				$string = 'I submitted the IIQ request for email access.I verified the mailbox created properly'
				GetSetValues("", "work_notes", $string)
			Case $button3
				$string = 'I submitted the IIQ request for email access.I added the SMTP and/or SIP manually.'
				GetSetValues("", "work_notes", $string)
			Case $button4
				$string = 'I created AWCP XXXX for this request.'
				GetSetValues("", "work_notes", $string)
			Case $button5
				$string = 'I created the DL and added the owner(s).'
				GetSetValues("", "work_notes", $string)
			Case $button6
				$string = 'Receiver(s) already have the requested access. No action taken.'
				GetSetValues("", "work_notes", $string)
			Case $button7
				$string = 'Complete - the notes ID file is created. The password was sent in a separate email per policy.'
				GetSetValues("", "work_notes", $string)
			Case $button8
				$string = 'I added the receiver(s) as an owner for the DL.'
				GetSetValues("", "work_notes", $string)
			Case $button9
				$string = 'The requested SQL access has been granted.'
				GetSetValues("", "work_notes", $string)
			Case $button10
				$string = 'I have created the Shared Mailbox. Added the owner(s), and completed the SharePoint form.'
				GetSetValues("", "work_notes", $string)
			Case $button11
				$string = 'I granted the requested access.'
				GetSetValues("", "work_notes", $string)
			Case $button12
				$string = 'Receiver(s) is a member of the below group(s). I am removing these from the request.' & @LF & '- Groups:' & @LF
				GetSetValues("", "work_notes", $string)
			Case $button13
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button14
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button15
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button16
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button17
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button18
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button19
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button20
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button21
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button22
				$string = ''
				GetSetValues("", "work_notes", $string)
			Case $button23
				$string = ""
				MainApp("", $string)
			Case $button24
				$string = ""
				MainApp("", $string)
			Case $button25

			Case $button26

			Case $button27

			Case $hCombo
				Flip()
		EndSwitch
	ElseIf $sToggle = "Email" Then ;Email
		If Not WinExists("Compose Email - ServiceNow") Then
			_GUICtrlComboBox_SetCurSel($hCombo, $start)
			Flip()
		EndIf
		Switch $msg
			Case -3
				Update()
				Exit
			Case $button1
				$string = '<p>Additional information is required to process this request.</p><p>[input information needed]</p><p>This access request cannot be completed until this information is received in the Security Administration mailbox.<p>'
				MainApp("", $string)
			Case $button2
				$string = '<p>Please be advised, this is our second attempt to obtain information required to complete this request.</p><p>[input information needed]</p><p>This access request cannot be completed until this information is received in the Security Administration mailbox.</p>'
				MainApp("", $string)
			Case $button3
				$string = '<p>Please be advised, this is our third and final attempt to obtain information required to complete this request.</p><p>[input information needed]</p><p>If we do not receive the requested information by this request will be closed and a new request will need to be opened.</p>'
				MainApp("", $string)
			Case $button4
				$string = '<p>Please be advised, this request has been canceled as there have been three attempts to obtain information required to complete this request, but no reply was received. If this access is still needed, a new request will need to be submitted.</p>'
				MainApp("", $string)
			Case $button5 ;EDW
				$string = '<p>Please assist with identifying which roles are required for the below request:</p><p>[enter access request details here]</p>'
				$to = "Ed or Raghu"
				$address = "emiles@humana.com;rram@humana.com"
				MainApp("", $string, $to, $address)
			Case $button6
				$string = '<p>Security Administration does not provision access to the department folders on the Q: drive.  This is managed by the assigned DDAs.  You can use the link below to determine who the DDAs are.  The DDAs are listed by their group associations.  These groups are named like so: G_D###_F#####_DDA.  You need to determine the group that has your DDAs.  Thus look for the group with the D### that matches the department folder you are accessing and matches the F##### of the facility folder you are accessing.  <strong>Q:\\DDAs\\Windows DDA\''s.htm</strong> NOTE: The "AllDept" folders can be managed by any DDA in a department.<br><br>DDA(s) for this folder:</p>'
				MainApp("", $string)
			Case $button7
				$string = '<p>Please go to go/goodpin to reset your PIN. If you have any issues please contact CSS @ 888-224-2700 for further assistance.</p>'
				MainApp("", $string)
			Case $button8
				$string = '<p>Please provide the full path for the folder(s) you need access to or a screen shot of the folder(s).</p>'
				MainApp("", $string)
			Case $button9
				$string = '<p>I am cancelling this ticket as is it a duplicate to [ticket].</p>'
				MainApp("", $string)
			Case $button10
				$string = '<p>We cannot grant access to this level of the [letter] drive. Please specify the specific folder(s) required.</p>'
				MainApp("", $string)
			Case $button11
				$string = '<p>Does [TSO] - [name] have an approved SRC to have local admin access?></p>'
				$to = 'SRC'
				$address = 'src@humana.com'
				$oIE = GetIe()
				$oIE.document.getElementById('MsgCcUI').innerHTML = $oIE.document.getElementById('MsgToUI').innerHTML
				MainApp("", $string, $to, $address)
			Case $button12
				_AD_Open()
				Local $displayname = _AD_GetObjectAttribute(@UserName, "displayname")
				_AD_Close()
				$string = '<p>We have a request to delete the below group(s). This was submitted by one of the current owners. Do we have your approval to process?<br><br><strong>Group(s):</strong><br><br><br>' & $displayname & '</p>'
				$to = "Greg"
				$address = 'gfarr@humana.com'
				MainApp("", $string, $to, $address)
			Case $button13
				$string = '<p>XXX is a member of the below group(s). I am removing these from the request.<br><br>- Groups:<br></p>'
				MainApp("", $string)
			Case $button14
				$string = '<p>A SharePoint task has been assigned to you for the certification ##### for ####. This request cannot be processed without the previously stated certification\''s approval. Your approval must be captured by "approving" through SharePoint, not replying to this email. An email from SharePoint was sent to you regarding this the same day the certification was created. If you did not receive an email, let me know and I will force the system to send another.</p>'
				MainApp("", $string)
			Case $button15
				$to = "Karen or Pamela"
				$address = 'kash@humana.com; ppoole@humana.com'
				$string = '<p>Do you approve of the below request?<br><br><strong>Request Details:</strong><br><br></p>'
				MainApp("", $string, $to, $address)
			Case $button16
				$string = '<p>We do not have owner information for the below item(s). Do you know who should be the owner?<br><br><br></p>'
				MainApp("", $string)
			Case $button17
				$string = '<p><br></p>'
				$to = 'BI Tool team'
				$address = 'BI_Tool@humana.com'
				$oIE = GetIe()
				$oIE.document.getElementById('MsgCcUI').innerHTML = $oIE.document.getElementById('MsgToUI').innerHTML
				MainApp("", $string, $to, $address)
			Case $button18
				$string = '<p></p>'
				MainApp("", $string)
			Case $button19
				$string = '<p></p>'
				MainApp("", $string)
			Case $button20
				$string = '<p></p>'
				MainApp("", $string)
			Case $button21
				$string = '<p></p>'
				MainApp("", $string)
			Case $button22
				$string = '<p></p>'
				MainApp("", $string)
			Case $button23
			Case $button24
			Case $button25
			Case $button26 ; VP Approval
				$string = "<p>A request has been submitted to delete a system account(s) on an application supported in your organization.  To ensure proper due diligence was completed prior to account removal, " & _
						"your approval or denial is required to proceed with this account deletion request.  We have already received approval to delete the account(s) from the account(s) owner(s).  " & _
						"This final approval step has been put in place to ensure system reliability and maximize the customer experience.</p>" & _
						"<p><strong>Account(s) Name:</strong><br><br></p>" & _
						"<p><strong>Account(s) Owners:</strong><br><br></p>" & _
						"<p><strong>Account(s) Environment:</strong><br><br></p>" & _
						'<p style="color:#5C9A1B;">What do you need from me?</p>' & _
						"<p>This request cannot be completed until your approval or denial is received. <strong>If changes are required for this request, please reply to this email and provide the necessary information.</strong></p>" & _
						'<p style="color:#5C9A1B;">What do I do if I have questions?</p>' & _
						"<p>If you have any questions or need additional guidance, please reply to this email.</p>"
				MainApp("", $string)
			Case $button27 ; Add Requestor's name
				$string = "<br><br>"
				MainApp("", $string)
			Case $hCombo
				Flip()
		EndSwitch
	EndIf
WEnd

Func AddApprover($flag, $results, $desc) ; Adds approver email addresses after each item found
	Local $replaceString = ""
	For $i = 1 To UBound($results) - 1
		If $flag = "Unix" Then
			$replaceString = StringReplace($results[$i][0], "AIX Server/PROD/MACHINE/", "")
			$replaceString = StringRegExpReplace($replaceString, "\.[A-Za-z0-9_\-]{4,}\.com", "")
		ElseIf $flag = "AD" Then
			$replaceString = StringReplace($results[$i][0], "Active Directory/PROD/GROUP/humad.com\", "")
		ElseIf $flag = "SQL" Then
			$replaceString = StringReplace($results[$i][0], "/QA/", "/")
			$replaceString = StringReplace($replaceString, "/INT/", "/")
			$replaceString = StringReplace($replaceString, "/Prod/", "/")
			$replaceString = StringReplace($replaceString, "/STAGE/", "/")
			$replaceString = StringReplace($replaceString, "/UNKNOWN/", "/")
			$replaceString = StringReplace($replaceString, "/DEV/", "/")
			$replaceString = StringReplace($replaceString, "/CORRECTED/", "/")
			$replaceString = StringReplace($replaceString, "/", @TAB)
		ElseIf $flag = "SRV" Then
			$replaceString = StringReplace($results[$i][0], "Windows Local Servers/Server/", "")
		EndIf
		If $results[$i][7] = "Approved" Then
			$desc = StringReplace($desc, $replaceString, $replaceString & " (" & $results[$i][8] & ")" & $results[$i][7], 1)
		ElseIf $results[$i][6] = "" Then
			$desc = StringReplace($desc, $replaceString, $replaceString & " (" & $results[$i][3] & ")" & $results[$i][7], 1)
		Else
			$desc = StringReplace($desc, $replaceString, $replaceString & " (" & $results[$i][3] & ")" & $results[$i][7] & "(" & $results[$i][6] & ")", 1)
		EndIf
	Next

	$desc = StringReplace($desc, "; ; )", ")")
	$desc = StringReplace($desc, "; )", ")")
	Return $desc
EndFunc   ;==>AddApprover

Func AppDB($oIE) ;Main approver DB logic
	Local $desc, $flag, $lookup, $systemMSG, $desc2, $descClean, $array, $group, $results = ""
	$desc = GetSetValues("", "description")
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf

	If StringInStr($desc, "Unix Server:") <> 0 Then
		$flag = "Unix"
	ElseIf StringInStr($desc, "Grant the following domain group membership:") <> 0 Then
		$flag = "AD"
	ElseIf StringInStr($desc, "SQL Server") <> 0 Then
		$flag = "SQL"
	ElseIf StringInStr($desc, "Server") <> 0 Then
		$flag = "SRV"
	Else
		MsgBox(0, "", "Failed to determine ticket type for search.")
		Return
	EndIf

	If $flag = "Unix" Then
		$desc2 = StringMid($desc, 1, StringInStr($desc, "Unix Server:") - 1)
		$descClean = StringReplace($desc, $desc2, "")
		$desc2 = StringMid($descClean, StringInStr($descClean, "Group:"), StringLen($descClean))
		$descClean = StringReplace($descClean, $desc2, "")
		$descClean = StringReplace($descClean, "Unix Server:", "")
		$array = StringSplit($descClean, @LF)
		For $i = 1 To UBound($array) - 1
			$array[$i] = StringStripWS($array[$i], 3)
			If StringLen($array[$i]) > 3 Then
				If $lookup <> "" Then
					$lookup = $lookup & " or app.[Element Path] like  'AIX Server/PROD/MACHINE/" & $array[$i] & "%'"
				Else
					$lookup = "app.[Element Path] like  'AIX Server/PROD/MACHINE/" & $array[$i] & "%'"
				EndIf
			EndIf
		Next
		$systemMSG = "AIX Server"
	ElseIf $flag = "AD" Then
		$desc2 = StringMid($desc, StringInStr($desc, "- Groups:"), StringLen($desc))
		$group = StringRegExp($desc2, "(?<= - )[A-Za-z0-9_\- ~&–]{4,}", 3)

		For $i = 0 To UBound($group) - 1
			If $i <> 0 Then
				$lookup = $lookup & ","
			EndIf
			$lookup = $lookup & "'Active Directory/PROD/GROUP/humad.com\" & $group[$i] & "'"
		Next
		$lookup = "app.[Element Path] in (" & $lookup & ")"
		$systemMSG = "Active Directory"
	ElseIf $flag = "SQL" Then
		$array = StringRegExp($desc, "SQL Server.*", 3)
		For $i = 0 To UBound($array) - 1
			$array[$i] = StringReplace($array[$i], @TAB, "%")
			$array[$i] = StringReplace($array[$i], "/", "/%", 1)
			$array[$i] = StringReplace($array[$i], "% ", "%")
			If StringLen($array[$i]) > 3 Then
				If $lookup <> "" Then
					$lookup = $lookup & " or app.[Element Path] like '" & $array[$i] & "'"
				Else
					$lookup = "app.[Element Path] like '" & $array[$i] & "'"
				EndIf
			EndIf
		Next
		$systemMSG = "SQL Server"
	ElseIf $flag = "SRV" Then
		$group = StringRegExp($desc, "[A-Za-z0-9]{10,}", 3)

		For $i = 0 To UBound($group) - 1
			If $i <> 0 Then
				$lookup = $lookup & ","
			EndIf
			$lookup = $lookup & "'Windows Local Servers/Server/" & $group[$i] & "'"
		Next
		$lookup = "app.[Element Path] in (" & $lookup & ")"
		$systemMSG = "Windows Local Servers"
	EndIf

	Local $query = "select " & _
			"app.[Element Path], " & _
			"app.[Description], " & _
			"app.[Approval Step Number], " & _
			"app.[Required number of Approvers (per step)], " & _
			"app.[Business Description], " & _
			"prim.[EmailAddress] as Pri, " & _
			"bup.[EmailAddress] as Bup, " & _
			"ste.[EmailAddress] as Ste, " & _
			"app.[Primary HISL] " & _
			"from tbl_Approver as app " & _
			"left join " & _
			"tbl_PersonnelActive as prim " & _
			"on app.[Primary HISL] = prim.AIN " & _
			"left join " & _
			"tbl_PersonnelActive as bup " & _
			"on app.[Backup HiSL] = bup.AIN " & _
			"left join " & _
			"tbl_PersonnelActive as ste " & _
			"on app.[Steward HiSL] = ste.AIN " & _
			"WHERE " & $lookup & " order by app.[Element Path], app.[Approval Step Number]"

	$array = SQLConnection($query)

	$query = "select [Instructions] from dbo.tbl_System where [system] = '" & $systemMSG & "'"

	$systemMSG = SQLConnection($query)

	$results = FlattenResults($array, $oIE, True)

	If IsArray($results) Then
		#Region Display System Message if one exists
		_ArrayDisplay($results, "Results", "|0:7", 64, Default, Default, Default, 0xe6e8ea)
		If $systemMSG[0][0] <> Null Then MsgBox(0, "System Message", $systemMSG[0][0])
		#EndRegion Display System Message if one exists

		If $sToggle = "Provisioning" Then
			$desc = AddApprover($flag, $results, $desc)
		EndIf

		GetSetValues($oIE, "description", $desc)
		If $desc = False Then
			MsgBox(0, "", "This button can only be used on an AR0 or ART.")
			Return
		EndIf
	EndIf
EndFunc   ;==>AppDB

Func BuildCommentsButtons()
	$button1 = GUICtrlCreateButton("Lync Chat", 1, $iButtonofffset * 0, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Puts a note in the work log for creating a Lynch chat room.", "Button Description", 1, 1)
	$button2 = GUICtrlCreateButton("IIQ Email Created", 1, $iButtonofffset * 1, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button3 = GUICtrlCreateButton("IIQ Email Failed", 1, $iButtonofffset * 2, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button4 = GUICtrlCreateButton("AWCP Note", 1, $iButtonofffset * 3, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button5 = GUICtrlCreateButton("DL Create", 1, $iButtonofffset * 4, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button6 = GUICtrlCreateButton("Access exists", 1, $iButtonofffset * 5, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button7 = GUICtrlCreateButton("Lotus Notes", 1, $iButtonofffset * 6, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button8 = GUICtrlCreateButton("DL Owner", 1, $iButtonofffset * 7, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button9 = GUICtrlCreateButton("SQL Access", 1, $iButtonofffset * 8, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button10 = GUICtrlCreateButton("Shared Mailbox", 1, $iButtonofffset * 9, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button11 = GUICtrlCreateButton("Generic completion", 1, $iButtonofffset * 10, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button12 = GUICtrlCreateButton("Partial Access AD", 1, $iButtonofffset * 11, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button13 = GUICtrlCreateButton("", 1, $iButtonofffset * 12, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button14 = GUICtrlCreateButton("", 1, $iButtonofffset * 13, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button15 = GUICtrlCreateButton("", 1, $iButtonofffset * 14, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button16 = GUICtrlCreateButton("", 1, $iButtonofffset * 15, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button17 = GUICtrlCreateButton("", 1, $iButtonofffset * 16, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button18 = GUICtrlCreateButton("", 1, $iButtonofffset * 17, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button19 = GUICtrlCreateButton("", 1, $iButtonofffset * 18, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button20 = GUICtrlCreateButton("", 1, $iButtonofffset * 19, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button21 = GUICtrlCreateButton("", 1, $iButtonofffset * 20, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button22 = GUICtrlCreateButton("", 1, $iButtonofffset * 21, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button23 = GUICtrlCreateButton("", 1, $iButtonofffset * 22, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button24 = GUICtrlCreateButton("", 1, $iButtonofffset * 23, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button25 = GUICtrlCreateButton("", 1, $iButtonofffset * 24, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button26 = GUICtrlCreateButton("", 1, $iButtonofffset * 25, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button27 = GUICtrlCreateButton("", 1, $iButtonofffset * 26, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
EndFunc   ;==>BuildCommentsButtons

Func BuildEmailButtons()
	$button1 = GUICtrlCreateButton("Need Info #1", 1, $iButtonofffset * 0, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button2 = GUICtrlCreateButton("Need Info #2", 1, $iButtonofffset * 1, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button3 = GUICtrlCreateButton("Need Info #3", 1, $iButtonofffset * 2, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button4 = GUICtrlCreateButton("Cancel No Reply", 1, $iButtonofffset * 3, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button5 = GUICtrlCreateButton("EDW Roles", 1, $iButtonofffset * 4, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button6 = GUICtrlCreateButton("DDA", 1, $iButtonofffset * 5, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button7 = GUICtrlCreateButton("Good Pin", 1, $iButtonofffset * 6, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button8 = GUICtrlCreateButton("Provide folder", 1, $iButtonofffset * 7, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button9 = GUICtrlCreateButton("Duplicate", 1, $iButtonofffset * 8, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button10 = GUICtrlCreateButton("Entire Drive", 1, $iButtonofffset * 9, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button11 = GUICtrlCreateButton("SRC Approval", 1, $iButtonofffset * 10, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button12 = GUICtrlCreateButton("Rod/Greg Approval", 1, $iButtonofffset * 11, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button13 = GUICtrlCreateButton("AD in GRPS", 1, $iButtonofffset * 12, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button14 = GUICtrlCreateButton("SharePoint Certification", 1, $iButtonofffset * 13, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button15 = GUICtrlCreateButton("HR Approval", 1, $iButtonofffset * 14, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button16 = GUICtrlCreateButton("Owner?", 1, $iButtonofffset * 15, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button17 = GUICtrlCreateButton("BI Tool", 1, $iButtonofffset * 16, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button18 = GUICtrlCreateButton("", 1, $iButtonofffset * 17, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button19 = GUICtrlCreateButton("", 1, $iButtonofffset * 18, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button20 = GUICtrlCreateButton("", 1, $iButtonofffset * 19, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button21 = GUICtrlCreateButton("", 1, $iButtonofffset * 20, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button22 = GUICtrlCreateButton("", 1, $iButtonofffset * 21, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button23 = GUICtrlCreateButton("", 1, $iButtonofffset * 22, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button24 = GUICtrlCreateButton("", 1, $iButtonofffset * 23, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button25 = GUICtrlCreateButton("", 1, $iButtonofffset * 24, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button26 = GUICtrlCreateButton("VP Approval", 1, $iButtonofffset * 25, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	$button27 = GUICtrlCreateButton("Add Requestor's name", 1, $iButtonofffset * 26, $iButtonwidth, $iButtonheight, $BS_NOTIFY)

EndFunc   ;==>BuildEmailButtons

Func BuildTaskButtons()
	$button1 = GUICtrlCreateButton("Email", 1, $iButtonofffset * 0, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Opens the email client. The buttons will change to provide email templates.", "Button Description", 1, 1)
	$button2 = GUICtrlCreateButton("Need to review", 1, $iButtonofffset * 1, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Adds work note 'Sent follow up email.' repeating clicking the button will update to 2nd and 3rd follow ups. ", "Button Description", 1, 1)
	$button3 = GUICtrlCreateButton("Clear NTR", 1, $iButtonofffset * 2, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Adds work note 'I am clearing the NTR as no response information required back to end user.'", "Button Description", 1, 1)
	$button4 = GUICtrlCreateButton("Info", 1, $iButtonofffset * 3, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Adds [info] to the beginning of short description and adds 'Sent for more information' in the work logs.", "Button Description", 1, 1)
	$button5 = GUICtrlCreateButton("AD Groups", 1, $iButtonofffset * 4, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Checks if the receivers are members of the listed groups. If they are in the group A (+) will be listed under their TSO. Then it will launch the Approver DB button.", "Button Description", 1, 1)
	$button6 = GUICtrlCreateButton("Approver DB", 1, $iButtonofffset * 5, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Looks up approvers after the ticket is formatted for AD group.", "Button Description", 1, 1)
	$button7 = GUICtrlCreateButton("Pull Approver", 1, $iButtonofffset * 6, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Pulls the approver after formatted for AD groups and approver look up. If multiple approvers are on the task it will order them.", "Button Description", 1, 1)
	$button8 = GUICtrlCreateButton("CTASK", 1, $iButtonofffset * 7, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "On ARO this will create a ART. On a ART it will create an approval task. On the approval task it will provide a box to enter the email address of the approver.", "Button Description", 1, 1)
	$button9 = GUICtrlCreateButton("Submit approval", 1, $iButtonofffset * 8, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "This will submit for approval if you have an approval task.", "Button Description", 1, 1)
	$button10 = GUICtrlCreateButton("Clone Task", 1, $iButtonofffset * 9, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "This will open a new tab, then clone that ART, and lastly it will populate the Service, Assignment group, and Assigned to.", "Button Description", 1, 1)
	$button11 = GUICtrlCreateButton("Populate Task", 1, $iButtonofffset * 10, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "This will populate the Service, Assignment group, and Assigned to.", "Button Description", 1, 1)
	$button12 = GUICtrlCreateButton("ID", 1, $iButtonofffset * 11, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "This will show a notice box at the top of the ticket. It will list the ID and groups to easily paste in AD search.", "Button Description", 1, 1)
	$button13 = GUICtrlCreateButton("Check SQL access", 1, $iButtonofffset * 12, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Only use on SQL formatted tickets. This will check the users to see if they already have access.", "Button Description", 1, 1)
	$button14 = GUICtrlCreateButton("Check Receivers", 1, $iButtonofffset * 13, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Checks all added receivers against TSOIDs listed in the description. It will show any differences.", "Button Description", 1, 1)
	$button15 = GUICtrlCreateButton("Create DL", 1, $iButtonofffset * 14, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Helps to automate the DL creation process.", "Button Description", 1, 1)
	$button16 = GUICtrlCreateButton("Notes Prep", 1, $iButtonofffset * 15, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Pulls the needed information to use with the Notes automation tools.", "Button Description", 1, 1)
	$button17 = GUICtrlCreateButton("Format SQL", 1, $iButtonofffset * 16, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "After pasting in the Word document, to the description this button will make the format presentable.", "Button Description", 1, 1)
	$button18 = GUICtrlCreateButton("Format AD", 1, $iButtonofffset * 17, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Formats the ticket for AD group(s).", "Button Description", 1, 1)
	$button19 = GUICtrlCreateButton("Receivers to clipboard", 1, $iButtonofffset * 18, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Pulls the reivers to clipboard. This will be useful when doing bulk IIQ requests.", "Button Description", 1, 1)
	$button20 = GUICtrlCreateButton("Bulk IIQ", 1, $iButtonofffset * 19, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button21 = GUICtrlCreateButton("IIQ Profile", 1, $iButtonofffset * 20, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button22 = GUICtrlCreateButton("Generate Password", 1, $iButtonofffset * 21, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Generates a password and adds it to the clipboard manager.", "Button Description", 1, 1)
	$button23 = GUICtrlCreateButton("Check Email", 1, $iButtonofffset * 22, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Pulls all of the receivers and then checks their SIP and SMTP address. It will display the results. This will primarily be used after requesting contractor email account(s).", "Button Description", 1, 1)
	$button24 = GUICtrlCreateButton("Manager Lookup", 1, $iButtonofffset * 23, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Pulls all of the receivers and then checks their manager and next level manager.", "Button Description", 1, 1)
	$button25 = GUICtrlCreateButton("Folder Research", 1, $iButtonofffset * 24, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Prompts for folder path. Returns the ACL and then you can highlight a single row and click the run button to do an ApproverDB search.", "Button Description", 1, 1)
	$button26 = GUICtrlCreateButton("", 1, $iButtonofffset * 25, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
	$button27 = GUICtrlCreateButton("Finalize Ticket", 1, $iButtonofffset * 26, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Only for use with AD currently - Combines the button Pull Approver, CTASK, and Submit Approval", "Button Description", 1, 1)
EndFunc   ;==>BuildTaskButtons

Func BuildResearchButtons()
	$button1 = GUICtrlCreateButton("Email", 1, $iButtonofffset * 0, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Opens the email client. The buttons will change to provide email templates.", "Button Description", 1, 1)
	$button2 = GUICtrlCreateButton("Need to review", 1, $iButtonofffset * 1, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Adds work note 'Sent follow up email.' repeating clicking the button will update to 2nd and 3rd follow ups. ", "Button Description", 1, 1)
	$button3 = GUICtrlCreateButton("Clear NTR", 1, $iButtonofffset * 2, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Adds work note 'I am clearing the NTR as no response information required back to end user.'", "Button Description", 1, 1)
	$button4 = GUICtrlCreateButton("Info", 1, $iButtonofffset * 3, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Adds [info] to the beginning of short description and adds 'Sent for more information' in the work logs.", "Button Description", 1, 1)
	$button5 = GUICtrlCreateButton("DNC Check", 1, $iButtonofffset * 4, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Checks to see if any of the receivers or their managers are on the DNC.", "Button Description", 1, 1)
	$button6 = GUICtrlCreateButton("Format AD", 1, $iButtonofffset * 5, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Formats the ticket for AD group(s).", "Button Description", 1, 1)
	$button7 = GUICtrlCreateButton("AD Groups", 1, $iButtonofffset * 6, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Checks if the receivers are members of the listed groups. If they are in the group A (+) will be listed under their TSO. Then it will launch the Approver DB button.", "Button Description", 1, 1)
	$button8 = GUICtrlCreateButton("Format SQL", 1, $iButtonofffset * 7, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "After pasting in the Word document, to the description this button will make the format presentable.", "Button Description", 1, 1)
	$button9 = GUICtrlCreateButton("Format Unix", 1, $iButtonofffset * 8, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "After pasting in the Word document, to the description this button will make the format presentable.", "Button Description", 1, 1)
	$button10 = GUICtrlCreateButton("Contractor Email", 1, $iButtonofffset * 9, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Formats the ticket for contractor email", "Button Description", 1, 1)
	$button11 = GUICtrlCreateButton("CTASK", 1, $iButtonofffset * 10, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "On ARO this will create a ART. On a ART it will create an approval task. On the approval task it will provide a box to enter the email address of the approver.", "Button Description", 1, 1)
	$button12 = GUICtrlCreateButton("Copy", 1, $iButtonofffset * 11, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Copies the description and short description from the AR to the ART.", "Button Description", 1, 1)
	$button13 = GUICtrlCreateButton("DDA", 1, $iButtonofffset * 12, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "If you have a DDA folder or group it will provide the DDA members.", "Button Description", 1, 1)
	$button14 = GUICtrlCreateButton("Copy receivers", 1, $iButtonofffset * 13, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Copies the receivers from the ARO and puts in the ART description.", "Button Description", 1, 1)
	$button15 = GUICtrlCreateButton("Populate Task", 1, $iButtonofffset * 14, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "This will populate the Service and Assignment group for LAN.", "Button Description", 1, 1)
	$button16 = GUICtrlCreateButton("Clone AR", 1, $iButtonofffset * 15, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "This will open a new tab, then clone that AR, and lastly it will populate the Description, Short Description, Requestor, Service, Assignment group, and Assigned to.", "Button Description", 1, 1)
	$button17 = GUICtrlCreateButton("Check Receivers", 1, $iButtonofffset * 16, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Checks all added receivers against TSOIDs listed in the description. It will show any differences.", "Button Description", 1, 1)
	$button18 = GUICtrlCreateButton("Chameleon & ECW", 1, $iButtonofffset * 17, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Template for Chameleon requests.", "Button Description", 1, 1)
	$button19 = GUICtrlCreateButton("Shrink", 1, $iButtonofffset * 18, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Changes double new lines to single. This helps to make the text more readable.", "Button Description", 1, 1)
	$button20 = GUICtrlCreateButton("New Line", 1, $iButtonofffset * 19, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Convert , or ; to new line to break apart chained items.", "Button Description", 1, 1)
	$button21 = GUICtrlCreateButton("Add Tabs", 1, $iButtonofffset * 20, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Add tabs to make the form more readable.", "Button Description", 1, 1)
	$button22 = GUICtrlCreateButton("Approver DB", 1, $iButtonofffset * 21, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Looks up approvers after the ticket is formatted for AD group.", "Button Description", 1, 1)
	$button23 = GUICtrlCreateButton("Manager Lookup", 1, $iButtonofffset * 22, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Pulls all of the receivers and then checks their manager and next level manager.", "Button Description", 1, 1)
	$button24 = GUICtrlCreateButton("Check Email", 1, $iButtonofffset * 23, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Pulls all of the receivers and then checks their SIP and SMTP address. It will display the results. This will primarily be used after requesting contractor email account(s).", "Button Description", 1, 1)
	$button25 = GUICtrlCreateButton("Folder Research", 1, $iButtonofffset * 24, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Prompts for folder path. Returns the ACL and then you can highlight a single row and click the run button to do an ApproverDB search.", "Button Description", 1, 1)
	$button26 = GUICtrlCreateButton("Cancel AR", 1, $iButtonofffset * 25, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "Adds work notes and sets the substate to DNK Communication.", "Button Description", 1, 1)
	$button27 = GUICtrlCreateButton("", 1, $iButtonofffset * 26, $iButtonwidth, $iButtonheight, $BS_NOTIFY)
	GUICtrlSetTip(-1, "", "Button Description", 1, 1)
EndFunc   ;==>BuildResearchButtons

Func ChameleonECW() ;ChameleonECW Template
	$oIE = GetIe()
	$desc = GetSetValues($oIE, "description")
	$string = ' --- Chameleon request --- ' & @LF & _
			'Application: Chameleon ' & @LF & _
			'Entity: ' & @LF & _
			'Environment: ' & @LF & _
			'Provider: No\Yes.  Details of provider detailed in the attached request form. ' & @LF & _
			'Role for CAC: ' & @LF & _
			'Role for CAC Care: ' & @LF & _
			'Role for CNU: ' & @LF & _
			'Role for PIPC: ' & @LF & _
			'Facility: ' & @LF & _
			'Access to edit Enrollments: ' & @LF & _
			'Required Group Memberships: ' & @LF & _
			@LF & _
			' --- eCW request --- ' & @LF & _
			'Application: eCW ' & @LF & _
			'Entity: ' & @LF & _
			'Environment: ' & @LF & _
			'Provider: No\Yes.  Details of provider detailed in the attached request form. ' & @LF & _
			'Role for CAC: ' & @LF & _
			'Role for CAC Care: ' & @LF & _
			'Role for CNU: ' & @LF & _
			'Role for PIPC: ' & @LF & _
			'Role for Metcare: ' & @LF & _
			'Facility: ' & @LF & _
			'License Healthcare Professional or Credentialed Medical Assistant: ' & @LF & _
			'Patient Documents Folder access: ' & @LF & _
			'eCW Fax Inbox Access: ' & @LF & _
			'If yes, what groups are required:' & @LF & _
			@LF & _
			'Care Delivery IT Support Statement:' & @LF & _
			@LF & _
			' --- Original request below --- ' & @LF & _
			@LF & _
			$desc

	GetSetValues($oIE, "description", $string)
EndFunc   ;==>ChameleonECW

Func CheckAD($oIE = "") ;Verifies AD groups are valid and checks to see if receivers are members
	If $oIE = "" Then
		$oIE = GetIe()
	EndIf
	Local $ids, $desc, $desc2, $group, $lookup, $aGroups = ""
	Local $desc = GetSetValues($oIE, "description")
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf
	If StringInStr($desc, "Grant the following domain group membership:") <> 0 Then
		$ids = StringRegExp($desc, "(?: - )(\w{3}\d{4}[xXsSaA]?)", 3)
		$desc = StringMid($desc, StringInStr($desc, "- Groups:"), StringLen($desc))
		$desc2 = StringMid($desc, StringInStr($desc, "- Groups:"), StringLen($desc))
		Local $checkForOriginal = StringInStr($desc2, "--- Original request below ---")
		If $checkForOriginal <> 0 Then
			$desc2 = StringMid($desc2, 1, $checkForOriginal - 1)
		EndIf
		$group = StringRegExp($desc2, "(?<= - )[A-Za-z0-9_\- ~&–]{4,}", 3)

		Local $lookup[UBound($group) + 1][UBound($ids) + 2]

		For $i = 1 To UBound($group)
			$lookup[$i][0] = StringStripWS($group[$i - 1], 2)
		Next

		For $i = 1 To UBound($ids)
			$lookup[0][$i + 1] = $ids[$i - 1]
		Next

		$lookup[0][1] = "Exists"

		_AD_Open()
		If @error Then Exit MsgBox(16, "Error", "Function _AD_Open encountered a problem. @error = " & @error & ", @extended = " & @extended)

		; Get an array of group names (FQDN) that the user is immediately a member of with element 0 containing the number of groups.
		For $i = 0 To UBound($ids) - 1
			$aGroups = _AD_GetUserGroups($ids[$i])
			For $y = 1 To UBound($aGroups) - 1
				For $x = 1 To UBound($lookup) - 1
					If $y = 1 Then
						If _AD_ObjectExists($lookup[$x][0], "cn") = 1 Then
							$lookup[$x][1] = "+"
						Else
							$lookup[$x][1] = "X"
						EndIf
					EndIf
					If StringInStr($aGroups[$y], "CN=" & StringStripWS($lookup[$x][0], 2) & ",") <> 0 Then
						$lookup[$x][$i + 2] = "+"
					EndIf
				Next
			Next
		Next

		_ArrayDisplay($lookup, "AD Groups", Default, 64, Default, Default, Default, 0xe6e8ea)

		; Close Connection to the Active Directory
		_AD_Close()
	Else
		MsgBox(0, "", "This does not appear to be a properly formatted AD ticket.")
		Return
	EndIf
	If $sToggle = "Provisioning" Then
		AppDB($oIE)
	EndIf
EndFunc   ;==>CheckAD

Func CheckEmail() ;Queries to verify if all receivers email accounts are created correctly
	$oIE = GetIe()
	Local $array = ReceiverList($oIE)

	Local $result[UBound($array)][4]
	$result[0][0] = "TSOID"
	$result[0][1] = "SIP"
	$result[0][2] = "SMTP"
	$result[0][3] = "SMTP Exch"
	_AD_Open()
	For $i = 1 To UBound($array) - 1
		$sResult = _AD_GetObjectAttribute($array[$i], "ProxyAddresses")
		$result[$i][0] = $array[$i]
		If IsArray($sResult) Then
			For $x = 0 To UBound($sResult) - 1
				If StringInStr($sResult[$x], "sip") <> 0 Then
					$result[$i][1] = StringReplace($sResult[$x], "sip:", "")
				ElseIf StringInStr($sResult[$x], "exch") <> 0 Then
					$result[$i][2] = StringReplace($sResult[$x], "smtp:", "")
				ElseIf StringInStr($sResult[$x], "smtp") <> 0 Then
					$result[$i][3] = StringReplace($sResult[$x], "smtp:", "")
				EndIf
			Next
		EndIf
		If @error = 2 Then MsgBox(16, "Error", "Attribute 'mail' does not exist for user " & @UserName)
	Next
	_AD_Close()

	_ArrayDisplay($result, "Email Addresses", Default, 64, Default, Default, Default, 0xe6e8ea)
EndFunc   ;==>CheckEmail

Func CheckIni() ;Pulls settings if ini file exists
	Local Const $sFileIni = @ScriptDir & "\toolbar.ini"
	If FileExists($sFileIni) Then
		$sCC = IniRead($sFileIni, "General", "CC", "")
		$sBCC = IniRead($sFileIni, "General", "BCC", "")
		Local $sMODE = IniRead($sFileIni, "General", "MODE", "")
		If $sMODE = "2" Then
			$sToggle = $sMODE
			_GUICtrlComboBox_SetCurSel($hCombo, 3) ; Set to research
			$start = _GUICtrlComboBox_GetCurSel($hCombo)
			Flip()
		EndIf
	EndIf
EndFunc   ;==>CheckIni

Func CheckReceivers() ;Check if any receivers are missing from the description or that exist in the description but are not added as a receiver
	Local $results[0][2]
	$oIE = GetIe()
	$desc = GetSetValues($oIE, "description")
	$ids = StringRegExp($desc, "[A-Za-z]{3}\d{4}[sSaAtT]{0,1}", 3)
	Local $array = ReceiverList($oIE)
	$counter = 0
	For $i = 0 To UBound($ids) - 1
		$iIndex = _ArraySearch($array, $ids[$i])
		If $iIndex = -1 Then
			ReDim $results[$counter + 1][2]

			$results[$counter][1] = "This receiver has not been added yet."
			$results[$counter][0] = $ids[$i]

			$counter += 1

		EndIf
	Next

	For $i = 1 To UBound($array) - 1
		$iIndex = _ArraySearch($ids, $array[$i])
		If $iIndex = -1 Then
			ReDim $results[$counter + 1][2]

			$results[$counter][1] = "This receiver is not listed on the description."
			$results[$counter][0] = $array[$i]

			$counter += 1

		EndIf
	Next

	If UBound($results) = 0 Then
		MsgBox(0, "", "All receivers are accounted for.")
	Else
		_ArrayDisplay($results, "Receiver Discrepancies", Default, 64, Default, Default, Default, 0xe6e8ea)
	EndIf
EndFunc   ;==>CheckReceivers

Func CheckSQL() ;Pull ticket info to feed into CheckSQLDB
	Local $oIE = GetIe()

	Local $id, $config, $srv, $db, $desc = ""
	$desc = GetSetValues($oIE, "description")
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf

	If StringInStr($desc, "SQL Server") <> 0 Then
		$array = StringSplit($desc, @LF & @LF, 1)

		For $x = 1 To UBound($array) - 1
			$id = StringRegExp($desc, "(?<=User Name:       ).*", 3)
			$id = StringRegExp($id[0], "(?<=\().*(?=\))", 3)

			$config = StringRegExp($desc, "SQL Server.*", 3)
			For $i = 0 To UBound($config) - 1
				$srv = StringRegExp($config[$i], "(?<=SQL Server               ).*(?=    Database)", 3)
				_ArrayDisplay($srv, "$srv")
				$db = StringRegExp($config[$i], "(?<=Database  ).*(?=\()", 3)
				If Not IsArray($db) Then
					$db = StringRegExp($config[$i], "(?<=Database               ).*", 3)
				EndIf
			Next
			CheckSQLDB($srv[0], $db[0], $id[0])
		Next
	Else
		MsgBox(0, "", "This does not appear to be a correctly formatted SQL ticket.")
	EndIf
EndFunc   ;==>CheckSQL

Func CheckSQLDB($arg1, $arg2, $arg3) ;Attempts to connect to a SQL server to see if access already exists
	Local $conn = ObjCreate("ADODB.Connection")
	Local $RS = ObjCreate("ADODB.Recordset")
	Local $DSN = "Provider=SQLOLEDB;Data Source=" & $arg1 & ";Initial Catalog=" & $arg2 & ";Integrated Security=SSPI;"
	$conn.Open($DSN)

	Local $query = "SELECT " & _
			"             susers.[name] AS LogInAtServerLevel, " & _
			"             users.[name] AS UserAtDBLevel, " & _
			"             DB_NAME() AS [Database], " & _
			"             roles.name AS DatabaseRoleMembership " & _
			"from sys.database_principals users " & _
			"             inner join sys.database_role_members link " & _
			"             on link.member_principal_id = users.principal_id " & _
			"             inner join sys.database_principals roles " & _
			"             on roles.principal_id = link.role_principal_id " & _
			"             inner join sys.server_principals susers " & _
			" on susers.sid = users.sid " & _
			"             where susers.[name] like '%" & $arg3 & "' or " & _
			"             users.[name] like '%" & $arg3 & "' "

	$RS.open($query, $conn)
	If $RS.state = 0 Then
		$conn.close
		$RS = ""
		$conn = ""
		$DSN = ""
		MsgBox(0, "", "The server or database name is invalid, or unable to make a connection.")
		Return
	EndIf
	If $RS.EOF Then
		MsgBox(0, "", "No access found for this server, DB, user configuration." & @LF & $arg1 & " - " & $arg2 & " - " & $arg3)
		$array = ""
	Else
		$array = $RS.GetRows()
	EndIf
	$conn.close
	$RS = ""
	$conn = ""
	$DSN = ""
EndFunc   ;==>CheckSQLDB

Func CloneAR() ;Creates a clone AR
	Local $oIE = GetIe()
	Local $desc = GetSetValues($oIE, "description")
	Local $sdesc = GetSetValues($oIE, "short_description")
	Local $open_by = GetSetValues($oIE, "opened_by")
	Local $assign_grp = GetSetValues($oIE, "assignment_group")
	Local $assign_to = GetSetValues($oIE, "assigned_to")
	NewTab($oIE)
	$oIE = IEWait("")

	Local $string = "GlideList2.get('x_human_access_request.x_human_access_request.copied_from').action('7b37cc370a0a0b34005bd7d7c7255583', 'sysverb_new');"
	$oIE.document.parentwindow.execScript($string)

	Sleep(2000)
	$oIE = IEWait("")

	GetSetValues($oIE, "description", $desc)
	GetSetValues($oIE, "short_description", $sdesc)
	GetSetValues($oIE, "opened_by", $open_by)
	GetSetValues($oIE, "assignment_group", $assign_grp)
	GetSetValues($oIE, "assigned_to", $assign_to)
EndFunc   ;==>CloneAR

Func CloneTask($arg = "", $arg2 = "") ;Creates a clone task
	$oIE = NewTab($arg, $arg2)
	$string = $jsMainApp & 'MainAppWin.document.getElementById("73b36716ef13200035c61ab995c0fbe8").click();'
	$oIE.document.parentwindow.execScript($string)
	$oIE = IEWait($oIE)
	PopTask($oIE)
	Return $oIE
EndFunc   ;==>CloneTask

Func CNTR() ;Note to clear NTR
	Local $string = 'I am clearing the NTR as no response information required back to end user.'
	GetSetValues("", "work_notes", $string)
EndFunc   ;==>CNTR

Func CopyReceivers($oIE = "") ;Copies receivers to the description
	If $oIE = "" Then
		$oIE = GetIe()
	EndIf
	$string = $jsMainApp & 'if("x_human_access_task"==MainAppWin.g_form.tableName){var n=[],a=new MainAppWin.GlideRecord("x_human_access_task");if(a.addQuery("sys_id",MainAppWin.g_form.getUniqueValue()),a.query(),a.next()){var r=new MainAppWin.GlideRecord("x_human_access_request");if(r.addQuery("sys_id",a.parent),r.query(),r.next()){var t=new MainAppWin.GlideRecord("sys_user"),s="^sys_idIN"+r.getValue("opened_for");for(""!==r.getValue("receivers")&&(s+=","+r.getValue("receivers")),s+="^EQ",t.encodedQuery=s,t.query();t.next();)n.push(t.getValue("user_name")+" ("+(t.getValue("u_nickname")||t.getValue("first_name"))+" "+t.getValue("last_name")+")")}n.length>0?MainAppWin.g_form.setValue("description","Receivers:\n"+n.map(function(e){return"- "+e}).join("\n")+"\n\n"+MainAppWin.g_form.getValue("description")):MainAppWin.g_form.setValue("description","Receivers:\n- None\n\n"+MainAppWin.g_form.getValue("description"))}}if("x_human_access_request"==MainAppWin.g_form.tableName){var n=[],a=new MainAppWin.GlideRecord("x_human_access_request"),r=new MainAppWin.GlideRecord("x_human_access_request");if(r.addQuery("sys_id",MainAppWin.g_form.getUniqueValue()),r.query(),r.next()){var t=new MainAppWin.GlideRecord("sys_user"),s="^sys_idIN"+r.getValue("opened_for");for(""!==r.getValue("receivers")&&(s+=","+r.getValue("receivers")),s+="^EQ",t.encodedQuery=s,t.query();t.next();)n.push(t.getValue("user_name")+" ("+(t.getValue("u_nickname")||t.getValue("first_name"))+" "+t.getValue("last_name")+")")}n.length>0?MainAppWin.g_form.setValue("description","Receivers:\n"+n.map(function(e){return"- "+e}).join("\n")+"\n\n"+MainAppWin.g_form.getValue("description")):MainAppWin.g_form.setValue("description","Receivers:\n- None\n\n"+MainAppWin.g_form.getValue("description"))}else alert("This can only be used on ART tickets");'
	MainApp($oIE, $string)
	Return $oIE
EndFunc   ;==>CopyReceivers

Func Ctask() ;On ARO this will create a ART. On a ART it will create an approval task. On the approval task it will provide a box to enter the email address of the approver.
	$string = '!function(){var e="undefined"!=typeof g_form?window:void 0!=document.getElementById("gsft_main")?document.getElementById("gsft_main").contentWindow:window;if("humana.service-now.com"==e.document.location.hostname)if(void 0!=e.g_form)if("x_human_access_request"==e.g_form.tableName){var t=new e.GlideRecord("x_human_access_task");t.addQuery("active",!0),t.addQuery("parent",e.g_form.getUniqueValue()),t.query(),0==t.rows.length?e.GlideList2.get("x_human_access_request.x_human_access_task.parent").action("be9263801b33200050fdfbcd2c0713e0","sysverb_new")||document.querySelector("[gsft_action_name=add_task]").click():confirm("Task already exists. Continue creating additional task?")&&(e.GlideList2.get("x_human_access_request.x_human_access_task.parent").action("be9263801b33200050fdfbcd2c0713e0","sysverb_new")||document.querySelector("[gsft_action_name=add_task]").click())}else"x_human_access_task"==e.g_form.tableName?e.GlideList2.get("x_human_access_task.sysapproval_approver.sysapproval").action("7ca0a8d60a0a0b340080d6f48c880640","sysverb_new"):alert("This can only be used on AR or ARTASK tickets, or the approver editor page");else if(e.document.location.pathname.match(/sys_m2m_template\.do/)){var a=prompt("Input the email address");if(a){var e="undefined"!=typeof g_form?window:void 0!=document.getElementById("gsft_main")?document.getElementById("gsft_main").contentWindow:window,t=new e.GlideRecord("sys_user");if(t.addQuery("email",a),t.query(),t.next()){var n=(t.getValue("u_nickname")||t.getValue("first_name"))+" "+t.getValue("last_name"),s=t.getValue("sys_id"),o=e.document.getElementById("select_1"),c=e.document.createElement("option");c.setAttribute("value",s),c.innerHTML=n,o.insert(c),document.getElementById("sysverb_save").click()}else e.alert("No employees found (email: "+a+")")}}else alert("This can only be used on AR or ARTASK tickets, or the approver editor page");else alert("This can only be used at humana.service-now.com")}();'
	MainApp($oIE, $string)
EndFunc   ;==>Ctask

Func DB($results, $aSelected) ;Builds query for folder approver lookup
	Local $search = StringReplace($results[$aSelected[1]][1], "\", ".com\")
	Local $lookup = "app.[Element Path] in ('Active Directory/PROD/GROUP/" & $search & "')"

	Local $query = "select " & _
			"app.[Element Path], " & _
			"app.[Description], " & _
			"app.[Approval Step Number], " & _
			"app.[Required number of Approvers (per step)], " & _
			"app.[Business Description], " & _
			"prim.[EmailAddress] as Pri, " & _
			"bup.[EmailAddress] as Bup, " & _
			"ste.[EmailAddress] as Ste, " & _
			"app.[Primary HISL] " & _
			"from tbl_Approver as app " & _
			"left join " & _
			"tbl_PersonnelActive as prim " & _
			"on app.[Primary HISL] = prim.AIN " & _
			"left join " & _
			"tbl_PersonnelActive as bup " & _
			"on app.[Backup HiSL] = bup.AIN " & _
			"left join " & _
			"tbl_PersonnelActive as ste " & _
			"on app.[Steward HiSL] = ste.AIN " & _
			"WHERE " & $lookup & " order by app.[Element Path], app.[Approval Step Number]"

	Local $array = SQLConnection($query)

	Local $appResults = FlattenResults($array)

	If IsArray($appResults) Then
		_ArrayDisplay($appResults, "Results", "|0:6", 64, Default, Default, Default, 0xe6e8ea)
	EndIf
EndFunc   ;==>DB

Func DDA() ;Pulls department and facility information if it exists in the description and then return the DDA members list
	Local $oIE = GetIe()
	Local $desc = GetSetValues($oIE, "description")
	Local $sName, $cleanName, $aMembers, $count = ""
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf
	$sName = StringRegExp($desc, "\D\d{3}[\\|_]\D\d{5}", 3)
	If IsArray($sName) Then
		_AD_Open()
		For $x = 0 To UBound($sName) - 1
			If $x <> 0 And $sName[$x - 1] = $sName[$x] Then
				ContinueLoop
			EndIf
			$cleanName = "G_" & StringReplace($sName[$x], "\", "_") & "_DDA"
			$aMembers = _AD_GetGroupMembers($cleanName)
			For $i = 1 To UBound($aMembers) - 1
				$count = StringInStr($aMembers[$i], ",")
				$aMembers[$i] = StringMid($aMembers[$i], 4, $count - 4)
			Next
			_ArrayDisplay($aMembers, $cleanName, Default, Default, Default, Default, Default, 0xe6e8ea)
		Next
		_AD_Close()
	Else
		MsgBox(0, "", "A DDA group or folder was not found.")
	EndIf
EndFunc   ;==>DDA

Func DL() ;Script to automate creating a DL
	Local $dl = InputBox("Input", "Please enter the DL name.")

	If @error = 1 Then
		Return
	EndIf
	$dl = StringStripWS($dl, 2)
	Local $dl2 = StringMid(StringStripWS($dl, 8), 1, 20)

	If Not WinExists("Exchange Management Console") Then
		MsgBox(0, "", "Please open Exchange Management Console")
		Return
	EndIf

	If WinExists("Find - Entire Forest") Then
		WinClose("Find - Entire Forest")
	EndIf

	Local $hWnd = WinGetHandle("Exchange Management Console", "")
	WinActivate($hWnd)
	Sleep(100)
	ControlSend($hWnd, "", "ToolbarWindow322", "!a")
	Sleep(100)
	ControlSend($hWnd, "", "ToolbarWindow322", "g")
	WinWaitActive("New Distribution Group", "", 10)

	Local $hWnd2 = WinGetHandle("New Distribution Group", "")
	ControlClick("New Distribution Group", "", "[NAME:next]")
	Sleep(200)
	ControlSend("New Distribution Group", "", "", "!o")
	ControlSend("New Distribution Group", "", "", "!r")
	WinWaitActive("Select Organizational Unit")

	Opt("SendKeyDelay", 200)
	Sleep(600)
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "h")
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "{RIGHT down}")
	Sleep(200)
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "exc")
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "{RIGHT down}")
	Sleep(400)
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "di")
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "{RIGHT down}")
	Sleep(400)
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "a")
	ControlSend("Select Organizational Unit", "", "[NAME:resultTreeView]", "{ENTER}")
	Sleep(400)
	Opt("SendKeyDelay", 5)
	ControlFocus("New Distribution Group", "", "[NAME:exchangeTextBoxGroupName]")
	ControlSend("New Distribution Group", "", "[NAME:exchangeTextBoxGroupName]", $dl)
	Sleep(200)
	ControlFocus("New Distribution Group", "", "[NAME:exchangeTextBoxAlias]")
	ControlSend("New Distribution Group", "", "[NAME:exchangeTextBoxAlias]", $dl2)
EndFunc   ;==>DL

Func DrawGUI() ;Creates the toolbar GUI
	Local $iGuiX, $iGuiY = ""
	Local Const $sFileIni = @ScriptDir & "\toolbar.ini"
	If FileExists($sFileIni) Then
		$iGuiX = IniRead($sFileIni, "General", "GuiX", "")
		$iGuiY = IniRead($sFileIni, "General", "GuiY", "")
	EndIf

	Local $iGuiWidth = 122
	Local $iGuiWidthFrame = $iGuiWidth + 10
	If $iGuiX = "" Then
		$iGuiY = 234
		If @DesktopWidth > 1920 Then
			$iGuiX = (@DesktopWidth / 2) - $iGuiWidthFrame
		Else
			$iGuiX = @DesktopWidth - $iGuiWidthFrame
		EndIf
	EndIf

	$hGui = GUICreate("ToolBar", $iGuiWidth, $iHeight, $iGuiX, $iGuiY, -1, $WS_EX_TOPMOST)
	GUISetBkColor(0x0A4F88)
	GUISetIcon(@ScriptDir & "\Tool-kit.ico")
	BuildTaskButtons()
	$hCombo = GUICtrlCreateCombo("Provisioning", 2, 729, 118, 30, $CBS_DROPDOWNLIST)
	GUICtrlSetData($hCombo, "Comments|Email|Research")
	GUISetState(@SW_SHOW)
EndFunc   ;==>DrawGUI

Func Email() ;Opens the ServiceNow email client
	Local $oIE = GetIe()
	$oIE.Document.parentWindow.execScript("document.body.oAutoIt = eval;")
	Local $eval = Execute("$oIE.Document.body.oAutoIt")
	Local $verify = $eval($jsMainApp)
	Local $result = $eval("MainAppWin.g_form.getValue('number') + ' - ' + MainAppWin.g_form.getValue('short_description');")
	$oIE.document.parentwindow.execScript('var MainAppWin = "undefined" != typeof g_form ? window : document.getElementById("gsft_main").contentWindow;MainAppWin.document.getElementById("email_client_open").click()')
	WinWaitActive("Compose Email - ServiceNow")
	Do
		Sleep(200)
		$oIE = GetIe()
	Until IsObj($oIE.document.getElementById('subject'))
	$oIE.document.getElementById('subject').value = $result
	$start = _GUICtrlComboBox_GetCurSel($hCombo)
	_GUICtrlComboBox_SetCurSel($hCombo, 2) ; Set to email
	Flip()
EndFunc   ;==>Email

Func FinalizeTask($oIE = "") ;Combo of pull appreover, Ctask, and submit approval buttons
	If $oIE = "" Then
		$oIE = GetIe()
	EndIf
	Pullapprover($oIE)
	Local $string = 'document.getElementById("sysverb_insert_and_stay").click()'
	MainApp($oIE, $string)
	IEWait($oIE)
	$string = '!function(){var e="undefined"!=typeof g_form?window:void 0!=document.getElementById("gsft_main")?document.getElementById("gsft_main").contentWindow:window;if("humana.service-now.com"==e.document.location.hostname||"humanaqa.service-now.com"==e.document.location.hostname)if(void 0!=e.g_form)if("x_human_access_request"==e.g_form.tableName){var t=new e.GlideRecord("x_human_access_task");t.addQuery("active",!0),t.addQuery("parent",e.g_form.getUniqueValue()),t.query(),0==t.rows.length?e.GlideList2.get("x_human_access_request.x_human_access_task.parent").action("be9263801b33200050fdfbcd2c0713e0","sysverb_new")||document.querySelector("[gsft_action_name=add_task]").click():confirm("Task already exists. Continue creating additional task?")&&(e.GlideList2.get("x_human_access_request.x_human_access_task.parent").action("be9263801b33200050fdfbcd2c0713e0","sysverb_new")||document.querySelector("[gsft_action_name=add_task]").click())}else"x_human_access_task"==e.g_form.tableName?e.GlideList2.get("x_human_access_task.sysapproval_approver.sysapproval").action("7ca0a8d60a0a0b340080d6f48c880640","sysverb_new"):alert("This can only be used on AR or ARTASK tickets, or the approver editor page");else if(e.document.location.pathname.match(/sys_m2m_template\.do/)){var a=prompt("Input the email address");if(a){var e="undefined"!=typeof g_form?window:void 0!=document.getElementById("gsft_main")?document.getElementById("gsft_main").contentWindow:window,t=new e.GlideRecord("sys_user");if(t.addQuery("email",a),t.query(),t.next()){var n=(t.getValue("u_nickname")||t.getValue("first_name"))+" "+t.getValue("last_name"),s=t.getValue("sys_id"),o=e.document.getElementById("select_1"),c=e.document.createElement("option");c.setAttribute("value",s),c.innerHTML=n,o.insert(c),e.document.getElementById("sysverb_save").click()}else e.alert("No employees found (email: "+a+")")}}else alert("This can only be used on AR or ARTASK tickets, or the approver editor page");else alert("This can only be used at humana.service-now.com")}();'
	MainApp("", $string)
	IEWait($oIE)
	$string = 'var a="' & $ClippArray[0][1] & '";if(a){var e="undefined"!=typeof g_form?window:void 0!=document.getElementById("gsft_main")?document.getElementById("gsft_main").contentWindow:window,t=new e.GlideRecord("sys_user");if(t.addQuery("email",a),t.query(),t.next()){var n=(t.getValue("u_nickname")||t.getValue("first_name"))+" "+t.getValue("last_name"),s=t.getValue("sys_id"),o=e.document.getElementById("select_1"),c=e.document.createElement("option");c.setAttribute("value",s),c.innerHTML=n,o.insert(c),e.document.getElementById("sysverb_save").click()}else e.alert("No employees found (email: "+a+")")}'
	MainApp("", $string)
	IEWait($oIE)
	$string = $jsMainApp & 'MainAppWin.document.getElementById("x_human_access_task.sysapproval_approver.sysapproval_choice_actions").scrollIntoView(),MainAppWin.document.getElementById("allcheck_x_human_access_task.sysapproval_approver.sysapproval").checked=!0,MainAppWin.document.querySelector(".list2_body .input-group-checkbox .checkbox").checked=!0;var grabmenu=MainAppWin.document.querySelector(".list_action_option");grabmenu.selectedIndex=4,grabmenu.onchange();'
	MainApp("", $string)
	IEWait($oIE)
EndFunc   ;==>FinalizeTask

Func FlattenResults($array, $oIE = "", $arg = False) ;Formats the results from approver DB pulls
	If $arg = True Then
		#Region Requestor lookup
		$oIE.Document.parentWindow.execScript("document.body.oAutoIt = eval;")
		$eval = Execute("$oIE.Document.body.oAutoIt")
		$eval('var requestor;var MainAppWin="undefined"!=typeof g_form?window:document.getElementById("gsft_main").contentWindow,gr=new MainAppWin.GlideRecord("x_human_access_request");gr.addQuery("sys_id",MainAppWin.g_form.getValue("parent")),gr.query(function(e){if(e.next()){var n=new MainAppWin.GlideRecord("sys_user");n.addQuery("sys_id",e.getValue("opened_by")),n.query(function(e){e.next(); requestor = e.email})}});')
		Local $iCounter = 0
		Do
			$requestor = $eval('requestor')
			Sleep(200)
			$iCounter += 1
			If $iCounter = 30 Then
				$requestor = " "
			EndIf
		Until $requestor <> ""
		#EndRegion Requestor lookup
	Else
		$requestor = "Empty"
	EndIf

	If IsArray($array) Then
		Local $results[1][9]
		$results[0][0] = "Group"
		$results[0][1] = "Description"
		$results[0][2] = "Business Description"
		$results[0][3] = "Primary"
		$results[0][4] = "BackUp"
		$results[0][5] = "Steward"
		$results[0][6] = "Seconday"
		$results[0][7] = "Approved"
		$results[0][8] = "Approver"

		$row = UBound($results) - 1

		#Region Flatten results
		For $i = 0 To UBound($array) - 1
			If $results[$row][0] <> $array[$i][0] Then
				ReDim $results[UBound($results) + 1][9]
			EndIf
			$results[$row][8] = $requestor
			$row = UBound($results) - 1
			If $array[$i][2] <> 2 Then
				$results[$row][0] = $array[$i][0]
				If $array[$i][1] <> Null And $results[$row][1] <> Null And $array[$i][1] <> "" And $results[$row][1] <> "" Then
					$results[$row][1] = $results[$row][1] & "; " & $array[$i][1]
				ElseIf $array[$i][1] <> Null And $array[$i][1] <> "" Then
					$results[$row][1] = $array[$i][1]
				EndIf
				If $array[$i][2] = 1 Then
					If $array[$i][4] <> Null And $results[$row][2] <> Null And $array[$i][4] <> "" And $results[$row][2] <> "" Then
						$results[$row][2] = $results[$row][2] & "; " & $array[$i][4]
					ElseIf $array[$i][4] <> Null And $array[$i][4] <> "" Then
						$results[$row][2] = $array[$i][4]
					EndIf
					If $array[$i][5] <> Null And $results[$row][3] <> Null And $array[$i][5] <> "" And $results[$row][3] <> "" Then
						$results[$row][3] = $results[$row][3] & "; " & $array[$i][5]
					ElseIf $array[$i][5] <> Null And $array[$i][5] <> "" Then
						$results[$row][3] = $array[$i][5]
					EndIf
					If $array[$i][6] <> Null And $results[$row][4] <> Null And $array[$i][6] <> "" And $results[$row][4] <> "" Then
						$results[$row][4] = $results[$row][4] & "; " & $array[$i][6]
					ElseIf $array[$i][6] <> Null And $array[$i][6] <> "" Then
						$results[$row][4] = $array[$i][6]
					EndIf
					If $array[$i][7] <> Null And $results[$row][5] <> Null And $array[$i][7] <> "" And $results[$row][5] <> "" Then
						$results[$row][5] = $results[$row][5] & "; " & $array[$i][7]
					ElseIf $array[$i][7] <> Null And $array[$i][7] <> "" Then
						$results[$row][5] = $array[$i][7]
					EndIf
				EndIf
				If $arg = True Then
					If ($requestor = $array[$i][5] Or $requestor = $array[$i][6] Or $requestor = $array[$i][7]) And $requestor <> Null Then
						$results[$row][7] = "Approved"
					EndIf
				EndIf
				If $array[$i][8] = 0 Then
					$results[$row][3] = "Manager Only"
				EndIf
			EndIf
			If $array[$i][2] = 2 Then
				If $array[$i][6] <> Null And $results[$row][6] <> Null And $array[$i][6] <> "" And $results[$row][6] <> "" Then
					$results[$row][6] = $results[$row][6] & "; " & $array[$i][5]
				Else
					$results[$row][6] = $array[$i][5]
				EndIf
			EndIf
		Next
		#EndRegion Flatten results

		Return $results
	EndIf
EndFunc   ;==>FlattenResults

Func FormatAD() ;AD template
	Local $string = $jsMainApp & 'if(r=MainAppWin.g_form.getValue("short_description"),n="Domain group membership",null!==r.match("Urgent Request: Virtual Training Room Access")&&(n=r),"x_human_access_request"==MainAppWin.g_form.getTableName()||"x_human_access_task"==MainAppWin.g_form.getTableName()){var t="x_human_access_request"==MainAppWin.g_form.getTableName()?MainAppWin.g_form.getValue("receiver_ids").split(/,/):MainAppWin.g_form.getValue("x_human_access_task.parent.ref_x_human_access_request.receiver_ids").split(/,/),s=[],a=MainAppWin.g_form.getValue("description").match(/(G_.*)/gi);if(0===t.length)MainAppWin.g_form.addErrorMessage("Error: No receivers were listed");else{var o=new MainAppWin.GlideRecord("sys_user");for(o.addQuery("user_name","IN",t.join(",")),o.query();o.next();)s.push(o.getValue("user_name")+" ("+(""!==o.getValue("u_nickname")?o.getValue("u_nickname"):o.getValue("first_name"))+" "+o.getValue("last_name")+")")}0===a.length?(a.push("G_TS_Remote_Desktop_Connection"),MainAppWin.g_form.addErrorMessage("Error: Unable to find groups (G_*)")):a.map(function(e){return e.trim()}),MainAppWin.g_form.setValue("short_description",n),MainAppWin.g_form.setValue("description","Grant the following domain group membership:\n- User IDs:\n  - "+s.join("\n  - ")+"\n- Groups:\n  - "+a.join("\n  - ")+"\n\n--- Original request below ---\n\n"+MainAppWin.g_form.getValue("description"))}else alert("This can only be used on AR or ARTASK tickets");'
	MainApp("", $string)
EndFunc   ;==>FormatAD

Func FormatSQL() ;SQL template
	Local $array, $serv, $db, $user, $id = ""
	Local $string = $jsMainApp & 'MainAppWin.g_form.setValue("short_description","Grant access - SQL");var text=MainAppWin.g_form.getValue("description"),nth=0;text=text.replace(/(?:\r\n|\r|\n)/g,"; "),text=text.replace(/Database Name:/g,"\nDatabase:").replace(/User Name:/g,"\n\nUser Name:").replace(/SQL Server Name:/g,"\nSQL Server:").replace(/User ID:/g,"\nUser ID:").replace(/Telephone #:/g,"\nTelephone #:").replace(/Access Needed \/ Role:/g,"\nRole:").replace(/Reason Access is Needed:/g,"\nReason:").replace(/Name of Special Role\(s\):/g,"\nSpecial Role(s):").replace(/Special Requirements:; \(default is HUMAD or HMHSChamp\)/g,"\nSpecial Requirements:").replace(/\*Special application\/system  Account ID /g,"\nSpecial application or ID: ").replace(/\*Account Owner ID /g,"\nAccount Owner ID: ").replace(/Active Directory; Group Name:/g,"\nGroup Name:").replace(/;\s{1,};/g,"").replace(/1\s{1,}DataReader\; {1,}/g,"DataReader;  ").replace(/0\s{1,}DataReader\; {1,}/g,"").replace(/1\s{1,}DataWriter\; {1,}/g,"DataWriter; ").replace(/0\s{1,}DataWriter\; {1,}/g,"").replace(/1\s{1,}SP_Execute\; {1,}/g,"SP_Execute; ").replace(/0\s{1,}SP_Execute\; {1,}/g,"").replace(/1\s{1,}Other/g,"Other").replace(/0\s{1,}Other {1,}/g,"").replace(/0\s{1,}New\; {1,}/g,"").replace(/1\s{1,}New\; {1,}/g,"New; ").replace(/0\s{1,}Existing/g,"").replace(/1\s{1,}Existing/g,"Existing; ").replace(/[(]{0,1}[0-9]{3}[)]{0,1}[-\s{1,}\.]{0,1}[0-9]{3}[-\s{1,}\.]{0,1}[0-9]{4}/g,"").replace(/Telephone #:\s{0,}; {1,}\n/g,"").replace(/Special Role\(s\): {1,}; {1,}\n/g,"").replace(/Special Requirements: {1,}\n/g,"").replace(/\(required for special accounts\)\:\s{1,}/g,"\n").replace(/;\s{1,}Reason:/g,"\nReason:").replace(/\(if applicable\)\: {1,}/g,"\n").replace(/^;\s{1,}\n\n/g,"").replace(/;\s\n/g,"\n").replace(/Telephone #:\s{1,}\n/g,"\n").replace(/:\s{1,}\n/g,":        \n").replace(/.*\:        \n/g,"").replace(/; {1,}\n/g,"").replace(/User Name\:/g,function(e,a,r){return nth++,nth>1?"\n\nUser Name:":e}),MainAppWin.g_form.setValue("description",text);'
	Local $oIE = MainApp("", $string)
	Local $desc = GetSetValues($oIE, "description")
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf

	$array = StringSplit($desc, @LF & @LF, 1)

	For $x = 1 To UBound($array) - 1
		$serv = StringRegExp($array[$x], "SQL Server:.*", 3)
		$db = StringRegExp($array[$x], "Database:.*", 3)
		$user = StringRegExp($array[$x], "User Name: .*", 3)
		$id = StringRegExp($array[$x], "User ID:.*", 3)

		For $i = 0 To UBound($serv) - 1
			$array[$x] = StringReplace($array[$x], @LF & $serv[$i], "")
			$array[$x] = StringReplace($array[$x], $db[$i], StringStripWS(StringReplace($serv[$i], ":" & @TAB, @TAB), 3) & @TAB & StringStripWS(StringReplace($db[$i], ":" & @TAB, @TAB), 3))

			$array[$x] = StringReplace($array[$x], @LF & $id[$i], "")
			$array[$x] = StringReplace($array[$x], $user[$i], $user[$i] & " (" & StringReplace($id[$i], "User ID:           ", "") & ")")
		Next
		Do
			$id = StringRegExp($array[$x], "(?<=\()[A-Za-z0-9_\- ~&–.]{7,}; ", 3)
			If IsArray($id) Then
				$array[$x] = StringReplace($array[$x], $id[0], "")
				$id[0] = StringReplace($id[0], "; ", "")
				$array[$x] = StringReplace($array[$x], "; ", " (" & $id[0] & "), ", 1)
			Else
				ContinueLoop
			EndIf

		Until Not IsArray($id)
	Next

	$desc = _ArrayToString($array, @LF & @LF, 1)

	GetSetValues($oIE, "description", $desc)
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf
EndFunc   ;==>FormatSQL

Func Flip() ;Changes the mode of the toolbar
	$sToggle = GUICtrlRead($hCombo)
	If $sToggle = "Provisioning" Then
		WipeButtons()
		BuildTaskButtons()
	ElseIf $sToggle = "Research" Then
		WipeButtons()
		BuildResearchButtons()
	ElseIf $sToggle = "Email" Then
		WipeButtons()
		BuildEmailButtons()
	ElseIf $sToggle = "Comments" Then
		WipeButtons()
		BuildCommentsButtons()
	EndIf
EndFunc   ;==>Flip

Func FolderResearch() ;Cleans and displays the results of GetAclByCacls
	Local $lookup = InputBox("Input required", "Enter folder path you would like to research.")

	Local $aclResults = GetAclByCacls($lookup)
	Local $aclArray = StringSplit($aclResults, @CRLF)

	Local $results[UBound($aclArray)][5]

	For $i = 0 To UBound($aclArray) - 1
		If $i = 0 Then
			$results[$i][0] = "Folder"
			$results[$i][1] = "Group"
			$results[$i][2] = "Inheritting to"
			$results[$i][3] = "Permissions"
			$results[$i][4] = "Inheritting from"
			ContinueLoop
		EndIf
		If StringInStr($aclArray[$i], "  ") <> 0 Then
			$results[$i][1] = StringStripWS($aclArray[$i], 3)
		ElseIf StringInStr($aclArray[$i], $lookup) <> 0 Then
			$results[$i][1] = StringReplace($aclArray[$i], $lookup & " ", "")
			$results[$i][0] = $lookup
		Else
			$results[$i][0] = $aclArray[$i]
		EndIf

		#Region replace Icacls shorthand with definition
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(RX\)|\(RX,", "(Read & Execute)")
		If StringInStr($results[$i][1], "(Read & Execute)") <> 0 Then
			$results[$i][3] &= "(Read & Execute)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Read & Execute\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(RA\)|\(RA,", "(Read attributes)")
		If StringInStr($results[$i][1], "(Read attributes)") <> 0 Then
			$results[$i][3] &= "(Read attributes)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Read attributes\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(WA\)|\(WA,", "(Write attributes)")
		If StringInStr($results[$i][1], "(Write attributes)") <> 0 Then
			$results[$i][3] &= "(Write attributes)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Write attributes\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(R\)|\(R", "(Read)")
		If StringInStr($results[$i][1], "(Read)") <> 0 Then
			$results[$i][3] &= "(Read)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Read\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(W\)|\(W,", "(Write-only access)")
		If StringInStr($results[$i][1], "(Write-only access)") <> 0 Then
			$results[$i][3] &= "(Write-only access)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Write\-only access\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(F\)|\(F", "(Full Control)")
		If StringInStr($results[$i][1], "(Full Control)") <> 0 Then
			$results[$i][3] &= "(Full Control)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Full Control\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(M\)|\(M,", "(Modify)")
		If StringInStr($results[$i][1], "(Modify)") <> 0 Then
			$results[$i][3] &= "(Modify)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Modify\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(OI\)\(CI\)\(IO\)", "(Subfolders and files only)")
		If StringInStr($results[$i][1], "(Subfolders and files only)") <> 0 Then
			$results[$i][2] &= "(Subfolders and files only)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Subfolders and files only\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(OI\)\(CI\)", "(This folder, subfolders, and files.)")
		If StringInStr($results[$i][1], "(This folder, subfolders, and files.)") <> 0 Then
			$results[$i][2] &= "(This folder, subfolders, and files.)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(This folder, subfolders, and files\.\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(CI\)", "(This folder and subfolders)")
		If StringInStr($results[$i][1], "(This folder and subfolders)") <> 0 Then
			$results[$i][2] &= "(This folder and subfolders)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(This folder and subfolders\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(I\)", "(Inherited from parent)")
		If StringInStr($results[$i][1], "(Inherited from parent)") <> 0 Then
			$results[$i][4] &= "(Inherited from parent)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Inherited from parent\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "DC\)|DC,", "(Delete child)")
		If StringInStr($results[$i][1], "(Delete child)") <> 0 Then
			$results[$i][3] &= "(Delete child)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Delete child\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "WDAC\)|WDAC", "(Write DAC)")
		If StringInStr($results[$i][1], "(Write DAC)") <> 0 Then
			$results[$i][3] &= "(Write DAC)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Write DAC\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "WO\)|WO,", "(Write Owner)")
		If StringInStr($results[$i][1], "(Write Owner)") <> 0 Then
			$results[$i][3] &= "(Write Owner)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Write Owner\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "DE\)|DE,", "(Delete)")
		If StringInStr($results[$i][1], "(Delete)") <> 0 Then
			$results[$i][3] &= "(Delete)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Delete\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "RC\)|RC,", "(Read control)")
		If StringInStr($results[$i][1], "(Read control)") <> 0 Then
			$results[$i][3] &= "(Read control)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Read control\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "AS\)|AS,", "(Access system security)")
		If StringInStr($results[$i][1], "(Access system security)") <> 0 Then
			$results[$i][3] &= "(Access system security)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Access system security\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "MA\)|MA,", "(Maxium allowed)")
		If StringInStr($results[$i][1], "(Maxium allowed)") <> 0 Then
			$results[$i][3] &= "(Maxium allowed)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Maxium allowed\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "GR\)|GR,", "(Generic read)")
		If StringInStr($results[$i][1], "(Generic read)") <> 0 Then
			$results[$i][3] &= "(Generic read)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Generic read\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "GW\)|GW,", "(Generic write)")
		If StringInStr($results[$i][1], "(Generic write)") <> 0 Then
			$results[$i][3] &= "(Generic write)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Generic write\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "GE\)|GE,", "(Generic execute)")
		If StringInStr($results[$i][1], "(Generic execute)") <> 0 Then
			$results[$i][3] &= "(Generic execute)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Generic execute\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "GA\)|GA,", "(Generic all)")
		If StringInStr($results[$i][1], "(Generic all)") <> 0 Then
			$results[$i][3] &= "(Generic all)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Generic all\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "RD\)|RD,", "(Read data/list directory)")
		If StringInStr($results[$i][1], "(Read data/list directory)") <> 0 Then
			$results[$i][3] &= "(Read data/list directory)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Read data/list directory\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "WD\)|WD,", "(Write data/add file)")
		If StringInStr($results[$i][1], "(Write data/add file)") <> 0 Then
			$results[$i][3] &= "(Write data/add file)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Write data/add file\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "AD\)|AD,", "(Append data/add subdirectory)")
		If StringInStr($results[$i][1], "(Append data/add subdirectory)") <> 0 Then
			$results[$i][3] &= "(Append data/add subdirectory)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Append data/add subdirectory\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "REA\)|REA,", "(Read extended atributes)")
		If StringInStr($results[$i][1], "(Read extended atributes)") <> 0 Then
			$results[$i][3] &= "(Read extended atributes)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Read extended atributes\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "WEA\)|WEA,", "(Write extended atributes)")
		If StringInStr($results[$i][1], "(Write extended atributes)") <> 0 Then
			$results[$i][3] &= "(Write extended atributes)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Write extended atributes\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "X\)|X,", "(Execute/traverse)")
		If StringInStr($results[$i][1], "(Execute/traverse)") <> 0 Then
			$results[$i][3] &= "(Execute/traverse)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Execute/traverse\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "\(D\)", "(Delete)")
		If StringInStr($results[$i][1], "(Delete)") <> 0 Then
			$results[$i][3] &= "(Delete)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Delete\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], "S\)", "(Synchronize)")
		If StringInStr($results[$i][1], "(Synchronize)") <> 0 Then
			$results[$i][3] &= "(Synchronize)"
			$results[$i][1] = StringRegExpReplace($results[$i][1], "\(Synchronize\)", "")
		EndIf
		;##########################################################################################
		$results[$i][1] = StringRegExpReplace($results[$i][1], ":", "")
		$results[$i][1] = StringRegExpReplace($results[$i][1], ",", "")
		#EndRegion replace Icacls shorthand with definition

		If $results[$i][1] = "" And $results[$i][3] <> "" Then
			$results[$i][1] = $results[$i - 1][1] & " (Special)"
		EndIf
	Next

	ReDim $results[UBound($results) - 4][5]

	$hUserFunction = DB
	_ArrayDisplay($results, "ACL Results", Default, 64, Default, Default, Default, 0xe6e8ea, $hUserFunction)
EndFunc   ;==>FolderResearch

Func GetAclByCacls($arg) ;Runs ICACLS.exe
	Local $CaclsPath = @SystemDir & "\ICACLS.exe"
	Local $Run = FileGetShortName($CaclsPath) & ' "' & $arg & '" '
	Local $Pid = Run($Run, '', @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	ProcessWaitClose($Pid)
	Local $sOutput = StdoutRead($Pid)
	Return $sOutput
EndFunc   ;==>GetAclByCacls

Func GetIe() ;Gets the open IE session
	Local $counter = 0
	Local $oGIE, $sString, $hControl, $sText, $aWinlist = ""

	Do
		If $counter = 30 Then
			MsgBox(0, "", "Unable to get control of IE. Please re-open the browser and re-launch the application.")
			Exit
		EndIf

		$hWnd = WinGetHandle("[CLASS:IEFrame]", "")
;~              ConsoleWrite("GetIE: " & $hWnd & @CRLF)
		WinActivate($hWnd)
		$sText = WinGetTitle("[ACTIVE]")
		If $sText = "Blank Page - Internet Explorer" Then
			Sleep(400)
			ContinueLoop
		EndIf

		$counter += 1

		$hWnd = WinGetHandle("[CLASS:IEFrame]", "")
		$hControl = ControlGetHandle($hWnd, "", "")
		$oGIE = __IEControlGetObjFromHWND($hControl)
		If IsObj($oGIE) Then
			Return $oGIE
		EndIf

		If Not IsObj($oGIE) Then Sleep(25)
	Until IsObj($oGIE)
;~    ConsoleWrite("GetIE: " & $oGIE.document.title & @CRLF)
;~    ConsoleWrite("GetIE: " & String(ObjName($oGIE)) & @CRLF)
	Return $oGIE
EndFunc   ;==>GetIe

Func GetSetValues($oIE, $arg1, $arg2 = "") ;Get or set ServiceNow attribute
	If Not IsObj($oIE) Then $oIE = GetIe()
	$oIE.Document.parentWindow.execScript("document.body.oAutoIt = eval;")
	Local $eval = Execute("$oIE.Document.body.oAutoIt")
	Local $verify = $eval($jsMainApp)
	If $arg2 <> "" Then
		$arg2 = StringReplace($arg2, '\', '\\')
		$arg2 = StringReplace($arg2, @LF, "\n")
		$arg2 = StringReplace($arg2, @CR, "\n")
		$arg2 = StringReplace($arg2, @CRLF, "\n")
		$arg2 = StringReplace($arg2, @TAB, "\t")
		$arg2 = StringReplace($arg2, '"', '\"')
		$eval('MainAppWin.g_form.setValue("' & $arg1 & '", "' & $arg2 & '");')
	EndIf
	Local $result = $eval("MainAppWin.g_form.getValue('" & $arg1 & "');")
	If $result = Null Then Return False
	$oIE.Document.parentWindow.execScript("document.body.oAutoIt = null;")
	If $sToggle = "Email" Or $sToggle = "Comments" Then
		_GUICtrlComboBox_SetCurSel($hCombo, $start)
		Flip()
	EndIf
	Return $result
EndFunc   ;==>GetSetValues

Func GuiPosition() ;Resizes the toolbar based on mouse location
	Local $aPos = WinGetPos("ToolBar")
	Local $aMPos = MouseGetPos()
	Local $overlap = 8
	Local $aClientSize = WinGetClientSize($hGui)
	If ($aMPos[0] >= $aPos[0] - $overlap And $aMPos[0] <= $aPos[0] + $aPos[2] + $overlap And $aMPos[1] >= $aPos[1] - $overlap And $aMPos[1] <= $aPos[1] + $aPos[3] + 50) Or GUICtrlRead($hCombo) = "Comments" Or GUICtrlRead($hCombo) = "Email" Then
		If $aClientSize[1] <> $iHeight And $aClientSize[1] <> ($iHeight + 3) Then
			WinMove($hGui, "", $aPos[0], $aPos[1], 128, ($iHeight + 28))
			WinSetTrans($hGui, "", 255)
		EndIf
	Else
		If $aClientSize[1] <> 0 And $aClientSize[1] <> 1 Then
			WinMove($hGui, "", $aPos[0], $aPos[1], 128, 26)
			WinSetTrans($hGui, "", 160)
		EndIf
	EndIf
EndFunc   ;==>GuiPosition

Func HotKeyPressed() ;Run function connected to the hotkey
	Switch @HotKeyPressed ; The last hotkey pressed.
		Case "^c" ; String is the Ctrl C hotkey.
			HotKeySet("^c")
			Send("^c")
			HotKeySet("^c", "HotKeyPressed")
			SetClip(StringStripWS(ClipGet(), 2))
		Case "^{f1}"
			ClipPut($ClippArray[0][1])
			Send("^v")
		Case "^{f2}"
			ClipPut($ClippArray[1][1])
			Send("^v")
		Case "^{f3}"
			ClipPut($ClippArray[2][1])
			Send("^v")
		Case "^{f4}"
			ClipPut($ClippArray[3][1])
			Send("^v")
		Case "^{f5}"
			ClipPut($ClippArray[4][1])
			Send("^v")
		Case "^{f6}"
			Send("humad\")
		Case "^{f8}"
			FinalizeTask()
		Case "^{f9}"
			SpawnTasks()
		Case "^{f11}"
			SetLocation()
		Case "^{f12}"
			_ArrayDisplay($ClippArray, "Current Clipboard", Default, 32 + 64, Default, "Key|Saved Item", Default, 0xe4f1fe)
		Case "!z"
			WinActivate("ToolBar")
			$start = _GUICtrlComboBox_GetCurSel($hCombo)
			_GUICtrlComboBox_SetCurSel($hCombo, 1) ; Set to email
			Flip()
	EndSwitch
EndFunc   ;==>HotKeyPressed

Func HotKeySetter() ;Sets hot keys and Function called when clicked.
	HotKeySet("^c", "HotKeyPressed")
	HotKeySet("^{f1}", "HotKeyPressed")
	HotKeySet("^{f2}", "HotKeyPressed")
	HotKeySet("^{f3}", "HotKeyPressed")
	HotKeySet("^{f4}", "HotKeyPressed")
	HotKeySet("^{f5}", "HotKeyPressed")
	HotKeySet("^{f6}", "HotKeyPressed")
	HotKeySet("^{f8}", "HotKeyPressed")
	HotKeySet("^{f9}", "HotKeyPressed")
	HotKeySet("^{f11}", "HotKeyPressed")
	HotKeySet("^{f12}", "HotKeyPressed")
	HotKeySet("!z", "HotKeyPressed")
EndFunc   ;==>HotKeySetter

Func IEState($arg) ;A second check for IEWait when object type is HTMLWindow2
	Local $arg2 = $arg.document
	Local $oTemp = $arg2.parentWindow
	While Not (String($oTemp.top.document.readyState) = "complete" Or $arg2.top.document.readyState = 4)
		Sleep(200)
		$arg = GetIe()
		$arg2 = $arg.document
		$oTemp = $arg2.parentWindow
	WEnd
	Return $arg
EndFunc   ;==>IEState

Func IEWait($arg) ;Holds a loop until the broswer state is complete
;~    ConsoleWrite("IEWait: " & String(ObjName($arg)) & @CRLF)
	Local $counter = 0
	Do
		If Not IsObj($arg) Then
			Sleep(100)
			$arg = GetIe()
		ElseIf $counter = 20 Then
			$arg = GetIe()
			$counter = 0
		Else
			If String(ObjName($arg)) = "HTMLWindow2" Then
				$arg = IEState($arg)
				Return $arg
			EndIf
;~                             ConsoleWrite("IEWait: " & _IEPropertyGet($arg, "busy") & @CRLF)
;~                             ConsoleWrite("IEWait: " & _IEPropertyGet($arg, "readystate") & @CRLF)
			$counter += 1
			Sleep(200)
		EndIf
	Until _IEPropertyGet($arg, "readystate") = 4 And _IEPropertyGet($arg, "busy") = False
	Return $arg
EndFunc   ;==>IEWait

Func MainApp($oIE, $arg1, $arg2 = "", $arg3 = "")
	If $oIE = "" Then $oIE = GetIe()
	If $sToggle = "Provisioning" Or $sToggle = "Research" Then
		$oIE.document.parentwindow.execScript($arg1)
	ElseIf $sToggle = "Comments" Then
		Local $script = $jsMainApp & 'MainAppWin.g_form.setValue("work_notes","' & $arg1 & '");'
		$oIE.document.parentwindow.execScript($script)
		_GUICtrlComboBox_SetCurSel($hCombo, $start)
		Flip()
	ElseIf $sToggle = "Email" Then
		$oIE.document.parentwindow.execScript('document.getElementById("message.text_ifr").contentWindow.document.getElementById("tinymce").getElementsByTagName("p")[0].getElementsByTagName("br")[0].outerHTML += "<br><br>"')
		$oIE.document.parentwindow.execScript('document.getElementById("message.text_ifr").contentWindow.document.getElementById("tinymce").getElementsByTagName("br")[2].outerHTML = ' & "'" & $arg1 & "'")
		If $arg2 <> '' Then
			$oIE.document.parentwindow.execScript('document.getElementById("message.text_ifr").contentWindow.document.getElementById("tinymce").getElementsByTagName("br")[0].outerHTML = ' & "'" & $arg2 & ",'")
		Else
			Local $requestor = $oIE.document.querySelector('#MsgToUI span').innerHTML
			$requestor = StringMid($requestor, 1, StringInStr($requestor, ";") - 1)
			Local $aObjects = NameLookUp($requestor)
			If IsArray($aObjects) Then
				$oIE.document.parentwindow.execScript('document.getElementById("message.text_ifr").contentWindow.document.getElementById("tinymce").getElementsByTagName("br")[0].outerHTML = ' & "'" & $aObjects[0][0] & ",'")
			Else
				MsgBox(0, "", "Unable to locate this user's profile. Their name cannot be added.")
			EndIf
		EndIf
		If $arg3 <> '' Then
			$oIE.document.parentwindow.execScript('document.getElementById("MsgToUI").getElementsByTagName("span")[0].remove()')
			Local $names = StringSplit($arg3, ";")
			For $name = 1 To UBound($names) - 1
				$oIE.document.parentwindow.execScript('var spn=document.createElement("span");spn.setAttribute("class","address"),spn.setAttribute("onclick","addressOnClick(event, this)"),spn.setAttribute("value","' & $names[$name] & '"),spn.innerHTML=''' & $names[$name] & ';<span class="action-delete">x</span>'';var test=document.getElementById("MsgToUI");test.appendChild(spn);')
			Next
		EndIf
		If $sCC <> "" Then
			$string = 'var spn=document.createElement("span");spn.setAttribute("class","address"),spn.setAttribute("onclick","addressOnClick(event, this)"),spn.setAttribute("value","' & $sCC & '"),spn.innerHTML=''' & $sCC & ';<span class="action-delete">x</span>'';document.getElementById("MsgCcUI").appendChild(spn);'
			$oIE.document.parentwindow.execScript($string)
		EndIf
		If $sBCC <> "" Then
			$string = 'var spn=document.createElement("span");spn.setAttribute("class","address"),spn.setAttribute("onclick","addressOnClick(event, this)"),spn.setAttribute("value","' & $sBCC & '"),spn.innerHTML=''' & $sBCC & ';<span class="action-delete">x</span>'';document.getElementById("MsgBccUI").appendChild(spn);'
			$oIE.document.parentwindow.execScript($string)
		EndIf

		_GUICtrlComboBox_SetCurSel($hCombo, $start)
		Flip()
	EndIf
	Return $oIE
EndFunc   ;==>MainApp

Func ManagerLookUp() ;Looks up 1st manager and 2nd manager for all receivers
	Local $oIE = GetIe()
	Local $results = ReceiverList($oIE, False)
	$results = StringReplace($results, ";", "','")

	Local $query = "select 'TSOID' as TSOID " & _
			",'LookUpDesc' as LookUpDesc " & _
			",'NickName' as NickName " & _
			",'FullName' as FullName " & _
			",'LastName' as LastName " & _
			",'FirstName' as FirstName " & _
			",'EmailAddress' as EmailAddress " & _
			",'SupervisorTSOID' as SupervisorTSOID " & _
			",'SupervisorFullName' as SupervisorFullName " & _
			",'SupervisorEmailAddress' as SupervisorEmailAddress " & _
			",'NextSupervisorTSOID' as NextSupervisorTSOID " & _
			",'NextSupervisorSupervisorFullName' as NextSupervisorSupervisorFullName " & _
			",'NextSupervisorEmailAddress' as NextSupervisorEmailAddress " & _
			"union all  " & _
			"( " & _
			"SELECT assoc.[TSOID] " & _
			",assoc.[LookUpDesc] " & _
			",assoc.[NickName] " & _
			",assoc.[FullName] " & _
			",assoc.[LastName] " & _
			",assoc.[FirstName] " & _
			",assoc.[EmailAddress] " & _
			",manager.[TSOID] as SupervisorTSOID " & _
			",assoc.[SupervisorFullName] " & _
			",manager.[EmailAddress] as SupervisorEmailAddress " & _
			",nextmanager.[TSOID] as NextSupervisorTSOID " & _
			",manager.[SupervisorFullName] as NextSupervisorSupervisorFullName " & _
			",nextmanager.[EmailAddress] as NextSupervisorEmailAddress " & _
			"FROM [AccessApprovers].[dbo].[tbl_PersonnelActive] as assoc " & _
			"join [AccessApprovers].[dbo].[tbl_PersonnelActive] as manager " & _
			"on assoc.[Supervisornumber] = manager.[AIN] " & _
			"join [AccessApprovers].[dbo].[tbl_PersonnelActive] as nextmanager " & _
			"on manager.[Supervisornumber] = nextmanager.[AIN] " & _
			"where assoc.TSOID in ('" & $results & "') " & _
			")"
	Local $array = SQLConnection($query)
	If IsArray($array) Then
		_ArrayDisplay($array, "Manager Lookup", Default, 64, Default, Default, Default, 0xe6e8ea)
	EndIf
EndFunc   ;==>ManagerLookUp

Func NameLookUp($arg) ;Find name from approver DB based on email address
	Local $query = "select case when NickName = '' Then " & _
			"FirstName " & _
			"when NickName is null Then " & _
			"FirstName " & _
			"else " & _
			"NickName " & _
			"end as Name " & _
			"from [AccessApprovers].[dbo].[tbl_PersonnelAll] where EmailAddress = '" & $arg & "' "

	Local $array = SQLConnection($query)

	Return $array
EndFunc   ;==>NameLookUp

Func NTR() ;Updates work notes
	Local $string = $jsMainApp & 'work_log=MainAppWin.g_form.getValue("work_notes");""===work_log?MainAppWin.g_form.setValue("work_notes","Sent follow up email."):"Sent follow up email."===work_log?MainAppWin.g_form.setValue("work_notes","Sent 2nd follow up email."):"Sent 2nd follow up email."===work_log?MainAppWin.g_form.setValue("work_notes","Sent 3rd follow up email."):MainAppWin.g_form.setValue("work_notes","Sent follow up email.");'
	MainApp("", $string)
EndFunc   ;==>NTR

Func NeedInfo() ;Updates short description and work notes
	Local $string = $jsMainApp & 'MainAppWin.g_form.setValue("short_description","[info] "+MainAppWin.g_form.getValue("short_description")),MainAppWin.g_form.setValue("work_notes","Sent for more info");'
	MainApp("", $string)
EndFunc   ;==>NeedInfo

Func NewTab($oNTIE = "", $number = "") ;Creates a new tab that is a copy of the active ticket or the ticket provided
	Local Const $navOpenInNewTab = 0x0800
	Local $curURL = ""
	If $oNTIE = "" Then
		$oNTIE = GetIe()
	EndIf
	$curURL = _IEPropertyGet($oNTIE, "locationurl")
	If StringInStr($curURL, "humanaqa") <> 0 Then
		Local $domain = "humanaqa"
	Else
		Local $domain = "humana"
	EndIf

	If $number = "" Then
		$number = GetSetValues($oNTIE, "number")
	EndIf
	If StringInStr($number, "ART") <> 0 Then
		Local $url = "https://" & $domain & ".service-now.com/x_human_access_task.do?sysparm_query=number=" & $number
	ElseIf StringInStr($number, "AR") <> 0 Then
		Local $url = "https://" & $domain & ".service-now.com/x_human_access_request.do?sysparm_query=number=" & $number
	Else
		Return False
	EndIf
	$oNTIE = _IEAttach($curURL, "url")
	$oNTIE.Navigate2($url, $navOpenInNewTab)
	$oNTIE = GetIe()

	$oNTIE = _IEAttach($curURL, "url", 2)
	$oNTIE = IEState($oNTIE)

	Return $oNTIE
EndFunc   ;==>NewTab

Func Password() ;Password generator
	Local $string = "abcdefghijklmnopqrstuvwxyz"
	Local $number = "0123456789"
	Local $character = "$#"
	Local $password = ""
	Local $flag, $random, $password = ""

	For $x = 0 To 9
		If $x = 0 Or $x = 9 Then
			$flag = Random(1, 2, 1)
		Else
			$flag = Random(1, 4, 1)
		EndIf

		If $flag = 1 Then
			$random = Random(1, 26, 1)
			$password = $password & StringMid($string, $random, 1)
		ElseIf $flag = 2 Then
			$random = Random(1, 26, 1)
			$password = $password & StringUpper(StringMid($string, $random, 1))
		ElseIf $flag = 3 Then
			$random = Random(1, 10, 1)
			$password = $password & StringMid($number, $random, 1)
		ElseIf $flag = 4 Then
			$random = Random(1, 2, 1)
			$password = $password & StringMid($character, $random, 1)
		EndIf
	Next
	ClipPut($password)
	SetClip($password)
EndFunc   ;==>Password

Func PopTask($oIE) ;Sets service, assignment group, and assigned to
	Local $string = $jsMainApp & 'var mytemp;mytemp,clone=MainAppWin.document.parentWindow.g_form.getValue("cloned_from_label"),grIncident=new MainAppWin.GlideRecord("x_human_access_task");grIncident.addQuery("number",clone),grIncident.query(),grIncident.next(),mytemp=grIncident.cmdb_ci,undefined===mytemp&&(mytemp="6ff3922f139152007c1331a63244b063"),MainAppWin.document.parentWindow.g_form.setValue("cmdb_ci",mytemp);'
	$oIE.document.parentwindow.execScript($string)

	#Region wait until Assignment Group is populated
	Local $counter = 0
	$oIE.Document.parentWindow.execScript("document.body.oAutoIt = eval;")
	Local $eval = Execute("$oIE.Document.body.oAutoIt")
	$eval('var assignment_group;var MainAppWin="undefined"!=typeof g_form?window:document.getElementById("gsft_main").contentWindow;')
	Do
		$eval('assignment_group=MainAppWin.g_form.getElement("assignment_group").value')
		$assignment_group = $eval('assignment_group')
		$counter += 1
		Sleep(200)
	Until $assignment_group <> "" Or $counter = 20
	#EndRegion wait until Assignment Group is populated
	If $counter = 20 Then
		MsgBox(0, "", "Not able to grab control of the IE session. Please reload the page and try again.")
		Return
	EndIf
	Sleep(200)
	$string = $jsMainApp & 'MainAppWin.g_form.setValue("assigned_to",MainAppWin.g_user.userID),MainAppWin.g_form.getElement("sys_readonly.x_human_access_task.number").focus();'
	$oIE.document.parentwindow.execScript($string)
EndFunc   ;==>PopTask

Func Pullapprover($oIE)
	Local $desc = GetSetValues($oIE, "description")
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf
	Local $aName = StringRegExp($desc, "((?<=\()[A-Za-z0-9._%+-]{2,}@[A-Za-z0-9._%+-]{4,}|Manager Only)", 3)

	If IsArray($aName) Then
		If UBound($aName) = 1 Then
			SetClip($aName[0])
		Else ;Orders approvers
			$aName = _ArrayUnique($aName)
			Local $partialDescIds = StringMid($desc, 1, StringInStr($desc, "- Groups:") + 8)
			If $partialDescIds = "" Then
				SetClip($aName[1])
				Return
			EndIf
			Local $partialDescGrps = StringReplace($desc, $partialDescIds, "")

			Local $checkForOriginal = StringInStr($desc, "--- Original request below ---")
			If $checkForOriginal <> 0 Then
				$checkForOriginal = StringMid($desc, $checkForOriginal, StringLen($desc))
				$partialDescGrps = StringReplace($partialDescGrps, $checkForOriginal, "")
			Else
				$checkForOriginal = ""
			EndIf

			Local $groups = StringRegExp($partialDescGrps, "  -.*", 3)
			Local $process = True
			Local $newDesc = $partialDescIds
			For $i = 1 To UBound($aName) - 1
				For $x = 0 To UBound($groups) - 1
					If StringInStr($groups[$x], "Approved") <> 0 And $aName[$i] <> "Manager Only" Then
						ContinueLoop
					ElseIf StringInStr($groups[$x], $aName[$i]) <> 0 Or (StringInStr($groups[$x], "Approved") <> 0 And $aName[$i] = "Manager Only") Then
						If $process = True Then
							SetClip($aName[$i])
							$process = False
						EndIf
						$newDesc &= @CR & $groups[$x]
					ElseIf StringInStr($groups[$x], "(") = 0 Or StringInStr($groups[$x], "()") <> 0 Then
						GetSetValues($oIE, "description", $desc)
						MsgBox(0, "", "Please review some groups do not have an approver listed.")
						Return
					EndIf
				Next
			Next
			GetSetValues($oIE, "description", $newDesc & @CRLF & $checkForOriginal)
		EndIf
	Else
		MsgBox(0, "", "This ticket does not appear to have an approver listed. Please run the approver lookup first.")
	EndIf
EndFunc   ;==>Pullapprover

Func ReceiverList($oIE, $arg = True) ;Looks up all receivers TSOIDs and returns them as ; seperated string
	Local $desc = GetSetValues($oIE, "description")
	Local $ids = StringRegExp($desc, "[A-Za-z]{3}\d{4}[sSaAtT]{0,1}", 3)
	$oIE.Document.parentWindow.execScript("document.body.oAutoIt = eval;")
	Local $eval = Execute("$oIE.Document.body.oAutoIt")
	$eval($jsMainApp & 'if("x_human_access_task"==MainAppWin.g_form.tableName){var n=[],a=new MainAppWin.GlideRecord("x_human_access_task");if(a.addQuery("sys_id",MainAppWin.g_form.getUniqueValue()),a.query(),a.next()){var r=new MainAppWin.GlideRecord("x_human_access_request");if(r.addQuery("sys_id",a.parent),r.query(),r.next()){var t=new MainAppWin.GlideRecord("sys_user"),s="^sys_idIN"+r.getValue("opened_for");for(""!==r.getValue("receivers")&&(s+=","+r.getValue("receivers")),s+="^EQ",t.encodedQuery=s,t.query();t.next();)n.push(t.getValue("user_name"))}}}if("x_human_access_request"==MainAppWin.g_form.tableName){var n=[],a=new MainAppWin.GlideRecord("x_human_access_request");if(a.addQuery("sys_id",MainAppWin.g_form.getUniqueValue()),a.query(),a.next()){var r=new MainAppWin.GlideRecord("x_human_access_request");if(r.addQuery("sys_id",a.sys_id),r.query(),r.next()){var t=new MainAppWin.GlideRecord("sys_user"),s="^sys_idIN"+r.getValue("opened_for");for(""!==r.getValue("receivers")&&(s+=","+r.getValue("receivers")),s+="^EQ",t.encodedQuery=s,t.query();t.next();)n.push(t.getValue("user_name"))}}}')
	Local $string = $eval("n.join(';')")
	If $arg = True Then
		Local $array = StringSplit($string, ";")
		Return $array
	Else
		Return $string
	EndIf
EndFunc   ;==>ReceiverList

Func Replace($arg) ;A Regex statement to find and replace items in the description
	Local $string = $jsMainApp & "var text = MainAppWin.g_form.getValue('description');text = text.replace(" & $arg & ");MainAppWin.g_form.setValue('description',text);"
	MainApp("", $string)
EndFunc   ;==>Replace

Func SetClip($arg) ;Shuffles clipboard as new items are copied
	$ClippArray[4][1] = $ClippArray[3][1]
	$ClippArray[3][1] = $ClippArray[2][1]
	$ClippArray[2][1] = $ClippArray[1][1]
	$ClippArray[1][1] = $ClippArray[0][1]
	$ClippArray[0][1] = $arg
	ClipPut($arg)
EndFunc   ;==>SetClip

Func SetLocation()
	Local $aClientSize = WinGetPos($hGui)
	Local Const $sFileIni = @ScriptDir & "\toolbar.ini"
	IniWrite($sFileIni, "General", "GuiX", $aClientSize[0])
	IniWrite($sFileIni, "General", "GuiY", $aClientSize[1])
EndFunc   ;==>SetLocation

Func SetRolesIIQ() ;Used by customer service when submitting RBAC requests
	$oIE = GetIe()
	$body = $oIE.document.body.innerHTML

	$array = StringRegExp($body, '(?<=\<div class\="x-trigger-index-0 x-form-trigger x-form-arrow-trigger x-form-trigger-last x-unselectable" id\=")ext-gen\d{4}(?=" role="button"\>\<\/div\>)', 3)
	$array2 = StringRegExp($body, 'boundlist-\d{4}-listEl', 3)
	For $i = 3 To UBound($array) - 1
		$oIE.document.getElementById($array[$i]).click()
		Sleep(400)
		$oIE.document.getElementById($array[$i]).click()
		Sleep(400)
		$body = $oIE.document.body.innerHTML
		$array2 = StringRegExp($body, 'boundlist-\d{4}-listEl', 3)
		$oIE.document.getElementById($array[$i]).click()
		Sleep(400)
		$oIE.Document.parentWindow.execScript("document.getElementById('" & $array2[UBound($array2) - 1] & "').children[0].getElementsByTagName('LI')[1].click()")
		$i += 1
	Next
EndFunc   ;==>SetRolesIIQ

Func SpawnTasks() ;Creates all needed clone tasks for an AD ticket
	Local $iReply = MsgBox(4, "Alert", "You are about to clone tasks for each approver email address on this ticket. Do you want to proceed?", $MB_TOPMOST)
	If $iReply = 7 Then
		Return
	EndIf
	$oIE = GetIe()
	Local $desc = GetSetValues($oIE, "description")
	If $desc = False Then
		MsgBox(0, "", "This button can only be used on an AR0 or ART.")
		Return
	EndIf
	Local $aName = StringRegExp($desc, "((?<=\()[A-Za-z0-9._%+-]{2,}@[A-Za-z0-9._%+-]{4,}|Manager Only)", 3)

	If IsArray($aName) Then
		If UBound($aName) = 1 Then
			SetClip($aName[0])
		Else ;Orders approvers
			$aName = _ArrayUnique($aName)
			Local $partialDescIds = StringMid($desc, 1, StringInStr($desc, "- Groups:") + 8)
			Local $partialDescGrps = StringReplace($desc, $partialDescIds, "")

			Local $checkForOriginal = StringInStr($desc, "--- Original request below ---")
			If $checkForOriginal <> 0 Then
				$checkForOriginal = StringMid($desc, $checkForOriginal, StringLen($desc))
				$partialDescGrps = StringReplace($partialDescGrps, $checkForOriginal, "")
			Else
				$checkForOriginal = ""
			EndIf

			Local $groups = StringRegExp($partialDescGrps, "  -.*", 3)
			Local $process = True
			Local $number = GetSetValues($oIE, "number")
			Local $firstPass = True
			For $i = 1 To UBound($aName) - 1
				Local $newDesc = $partialDescIds
				Local $adddescription = False
				For $x = 0 To UBound($groups) - 1
					If StringInStr($groups[$x], "Approved") <> 0 And $aName[$i] <> "Manager Only" Then
						ContinueLoop
					ElseIf StringInStr($groups[$x], $aName[$i]) <> 0 Or (StringInStr($groups[$x], "Approved") <> 0 And $aName[$i] = "Manager Only") Then
						If $process = True Then
							SetClip($aName[$i])
							$process = False
						EndIf
						$newDesc &= @CR & $groups[$x]
						$adddescription = True
					ElseIf StringInStr($groups[$x], "(") = 0 Or StringInStr($groups[$x], "()") <> 0 Then
						GetSetValues($oIE, "description", $desc)
						MsgBox(0, "", "Please review some groups do not have an approver listed.")
						Return
					EndIf
				Next
				If $adddescription = True Then
					If $firstPass = False Then
						$oIE = CloneTask($oIE, $number)
					EndIf
					GetSetValues($oIE, "description", $newDesc & @CRLF & $checkForOriginal)
					If $aName[$i] <> "Manager Only" And StringInStr($newDesc, "Approved") = 0 Then
						FinalizeTask($oIE)
					EndIf
					$firstPass = False
;~                                               MsgBox(0,"","Verify the task is correct")
				EndIf
			Next
		EndIf
	Else
		MsgBox(0, "", "This ticket does not appear to have an approver listed. Please run the approver lookup first.")
	EndIf
EndFunc   ;==>SpawnTasks

Func SQLConnection($query, $table = "Table") ;Opens SQL connection runs query and returns results
	Local $conn = ObjCreate("ADODB.Connection")
	Local $RS = ObjCreate("ADODB.Recordset")
	Local $DSN = "Provider=SQLOLEDB;Data Source=Server;Initial Catalog=" & $table & ";Integrated Security=SSPI;"
	$conn.Open($DSN)

	$RS.open($query, $conn)
	If $RS.EOF Then
		MsgBox(0, "", "No record(s) found for these group(s).")
		Return
	Else
		$array = $RS.GetRows()
	EndIf

	$conn.close
	$RS = ""
	$conn = ""
	$DSN = ""
	Return $array
EndFunc   ;==>SQLConnection

Func Update() ;Script to self update from network
;~    Local $source = FileGetTime (@ScriptDir & '\Toolbar.exe')
;~    Local $update = FileGetTime ('V:\Security\ITSecAdm_Users\jboley\Toolbar.exe')
;~    $source = _ArrayToString($source)
;~    $update = _ArrayToString($update)

;~    If $source <> $update Then
	_SelfUpdate('V:\Security\ITSecAdm_Users\jboley\Toolbar.exe', False, 0, False, False)
EndFunc   ;==>Update

Func WipeButtons() ;Deletes all button
	GUICtrlDelete($button1)
	GUICtrlDelete($button2)
	GUICtrlDelete($button3)
	GUICtrlDelete($button4)
	GUICtrlDelete($button5)
	GUICtrlDelete($button6)
	GUICtrlDelete($button7)
	GUICtrlDelete($button8)
	GUICtrlDelete($button9)
	GUICtrlDelete($button10)
	GUICtrlDelete($button11)
	GUICtrlDelete($button12)
	GUICtrlDelete($button13)
	GUICtrlDelete($button14)
	GUICtrlDelete($button15)
	GUICtrlDelete($button16)
	GUICtrlDelete($button17)
	GUICtrlDelete($button18)
	GUICtrlDelete($button19)
	GUICtrlDelete($button20)
	GUICtrlDelete($button21)
	GUICtrlDelete($button22)
	GUICtrlDelete($button23)
	GUICtrlDelete($button24)
	GUICtrlDelete($button25)
	GUICtrlDelete($button26)
	GUICtrlDelete($button27)
EndFunc   ;==>WipeButtons

Func _User_ErrFunc($oError) ;Error Handler
	; User's COM error function.
	; After SetUp with ObjEvent("AutoIt.Error", ....) will be called if COM error occurs
	; Do anything here.
	ConsoleWrite(@CRLF & @ScriptFullPath & " (" & $oError.scriptline & ") : ==> COM Error intercepted !" & @CRLF & _
			@TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oError.number) & @CRLF & _
			@TAB & "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
			@TAB & "err.description is: " & @TAB & $oError.description & @CRLF & _
			@TAB & "err.source is: " & @TAB & @TAB & $oError.source & @CRLF & _
			@TAB & "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
			@TAB & "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
			@TAB & "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
			@TAB & "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
			@TAB & "err.retcode is: " & @TAB & "0x" & Hex($oError.retcode) & @CRLF & @CRLF)
EndFunc   ;==>_User_ErrFunc
