/**
 * A Blessing to Share — persistence + Gmail backend
 * Uses Google Sheets for creation metadata/state and Google Drive for uploaded photos.
 * Deploy as Web App: Execute as Me, Who has access: Anyone.
 */
var ADMIN_PASSWORD_HASH = 'fd7adfefd51bfceb2db1e9a34d1baa2857188ae937975c0ad891947051cc8b05';
var SHEET_NAME = 'Creations';
var DRIVE_FOLDER_NAME = 'A Blessing to Share Photos';

function authorizeServices(){
  DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  ss_();
  GmailApp.getAliases();
  Logger.log('Authorization successful: Drive, Sheets, and Gmail are accessible.');
}

function authorizeDrive(){ authorizeServices(); }

function doGet(e){
  var p=e&&e.parameter?e.parameter:{};
  var action=String(p.action||'');
  try{
    if(action==='getHistory') return jsonp_(historyForUser_(p.userId,p.userKey),p.callback);
    if(action==='getPublic') return jsonp_(publicCreation_(p.creationId),p.callback);
    if(action==='adminList') return jsonp_(adminList_(p.adminToken),p.callback);
    if(action==='adminLogin') return jsonp_(adminLoginData_(p),p.callback);
    return output_('A Blessing to Share backend is running.');
  }catch(err){return jsonp_({ok:false,message:err.message},p.callback)}
}
function doPost(e){
  try{
    var p=e&&e.parameter?e.parameter:{};
    var action=String(p.action||'sendEmail');
    if(action==='save') return saveCreation_(p);
    if(action==='delete') return deleteCreation_(p);
    if(action==='adminDelete') return adminDelete_(p);
    if(action==='adminLogin') return htmlPostMessage_(adminLoginData_(p));
    if(action==='sendEmail') return sendEmail_(p);
    return output_('ERROR: Unknown action.');
  }catch(err){return output_('ERROR: '+err.message)}
}
function ss_(){var props=PropertiesService.getScriptProperties();var id=props.getProperty('CREATIONS_SHEET_ID');var ss=id?SpreadsheetApp.openById(id):SpreadsheetApp.create('A Blessing to Share — Creations');if(!id){props.setProperty('CREATIONS_SHEET_ID',ss.getId())}var sh=ss.getSheetByName(SHEET_NAME)||ss.insertSheet(SHEET_NAME);if(sh.getLastRow()===0)sh.appendRow(['id','userId','userKeyHash','cardType','stateJson','createdAt','updatedAt']);return sh}
function folder_(){var props=PropertiesService.getScriptProperties(),id=props.getProperty('PHOTO_FOLDER_ID');if(id)return DriveApp.getFolderById(id);var it=DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);var f=it.hasNext()?it.next():DriveApp.createFolder(DRIVE_FOLDER_NAME);props.setProperty('PHOTO_FOLDER_ID',f.getId());return f}
function hash_(s){var b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8);return b.map(function(x){var v=(x<0?x+256:x).toString(16);return v.length===1?'0'+v:v}).join('')}
function authUser_(p){if(!p.userId||!p.userKey)throw new Error('User identity missing.');return {id:String(p.userId),keyHash:hash_(p.userKey)}}
function processPhotos_(photos,oldPhotos){
  photos=Array.isArray(photos)?photos.slice(0,3):[];oldPhotos=oldPhotos||[];var folder=folder_();var oldIds={};oldPhotos.forEach(function(x){if(x.fileId)oldIds[x.fileId]=true});
  return photos.map(function(p){if(p.fileId&&p.url)return {id:p.id,fileId:p.fileId,url:p.url,x:Number(p.x)||0,y:Number(p.y)||0,width:Number(p.width)||90,height:Number(p.height)||90,rotation:Number(p.rotation)||0};if(!p.dataUrl)throw new Error('Photo data is missing.');var m=String(p.dataUrl).match(/^data:image\/(jpeg|png|webp);base64,(.+)$/);if(!m)throw new Error('Unsupported photo data.');var blob=Utilities.newBlob(Utilities.base64Decode(m[2]),'image/jpeg','blessing-'+p.id+'.jpg');var file=folder.createFile(blob);file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return {id:p.id,fileId:file.getId(),url:'https://drive.google.com/uc?export=view&id='+file.getId(),x:Number(p.x)||0,y:Number(p.y)||0,width:Number(p.width)||90,height:Number(p.height)||90,rotation:Number(p.rotation)||0}});
}
function saveCreation_(p){var u=authUser_(p),sh=ss_(),id=String(p.creationId||Utilities.getUuid()),state=JSON.parse(String(p.state||'{}')),rows=sh.getDataRange().getValues(),found=-1,old={};for(var i=1;i<rows.length;i++){if(String(rows[i][0])===id){found=i+1;old=JSON.parse(rows[i][4]||'{}');if(String(rows[i][2])!==u.keyHash||String(rows[i][1])!==u.id)throw new Error('You do not own this creation.');break}}
  state.photos=processPhotos_(state.photos,old.photos);state.creationId=id;var now=new Date().toISOString();if(found<0)sh.appendRow([id,u.id,u.keyHash,String(p.cardType||'Blessing Card'),JSON.stringify(state),now,now]);else sh.getRange(found,4,1,4).setValues([[String(p.cardType||'Blessing Card'),JSON.stringify(state),rows[found-1][5],now]]);return output_('OK')}
function deleteCreation_(p){var u=authUser_(p),sh=ss_(),rows=sh.getDataRange().getValues();for(var i=1;i<rows.length;i++){if(String(rows[i][0])===String(p.creationId)){if(String(rows[i][1])!==u.id||String(rows[i][2])!==u.keyHash)throw new Error('You do not own this creation.');sh.deleteRow(i+1);return output_('OK')}}return output_('OK')}
function historyForUser_(uid,key){if(!uid||!key)throw new Error('User identity missing.');var sh=ss_(),rows=sh.getDataRange().getValues(),h=hash_(key),items=[];for(var i=1;i<rows.length;i++){if(String(rows[i][1])===String(uid)&&String(rows[i][2])===h)items.push(row_(rows[i]))}items.reverse();return {ok:true,items:items}}
function publicCreation_(id){var sh=ss_(),rows=sh.getDataRange().getValues();for(var i=1;i<rows.length;i++)if(String(rows[i][0])===String(id)){return {ok:true,state:JSON.parse(rows[i][4]),id:rows[i][0]}}return {ok:false,message:'Creation not found.'}}
function row_(r){return {id:r[0],userId:r[1],cardType:r[3],state:JSON.parse(r[4]||'{}'),createdAt:r[5],updatedAt:r[6]}}
function adminLoginData_(p){var ok=hash_(String(p.password||''))===ADMIN_PASSWORD_HASH;if(!ok)return {type:'ADMIN_LOGIN',ok:false,message:'Invalid administrator password.'};var token=Utilities.getUuid()+'-'+Utilities.getUuid();CacheService.getScriptCache().put('admin:'+token,'1',21600);return {type:'ADMIN_LOGIN',ok:true,token:token}}
function adminLogin_(p){return htmlPostMessage_(adminLoginData_(p))}
function adminList_(token){if(!CacheService.getScriptCache().get('admin:'+String(token||'')))throw new Error('Administrator session expired.');var sh=ss_(),rows=sh.getDataRange().getValues(),items=[];for(var i=1;i<rows.length;i++)items.push(row_(rows[i]));items.reverse();return {ok:true,items:items}}
function adminDelete_(p){if(!CacheService.getScriptCache().get('admin:'+String(p.adminToken||'')))throw new Error('Administrator session expired.');var sh=ss_(),rows=sh.getDataRange().getValues();for(var i=1;i<rows.length;i++)if(String(rows[i][0])===String(p.creationId)){sh.deleteRow(i+1);return output_('OK')}return output_('OK')}
function sendEmail_(p){var email=String(p.email||'').trim(),to=String(p.to||'you').trim(),from=String(p.from||'Young Servants of the Lord').trim(),subject=String(p.subject||'A blessing for you').trim(),message=String(p.message||'').trim(),cardLink=String(p.cardLink||'').trim();if(!/^.+@gmail\.com$/i.test(email))return output_('Please provide a valid Gmail address.');if(!cardLink)return output_('The card link is missing.');var html='<div style="font-family:Arial,sans-serif;line-height:1.6;color:#2a0a10;max-width:620px"><h2 style="color:#6e2530">A blessing for '+escapeHtml_(to)+'</h2><p>'+escapeHtml_(message).replace(/\n/g,'<br>')+'</p><p style="margin:24px 0"><a href="'+escapeHtml_(cardLink)+'" style="display:inline-block;background:#c99a3d;color:#241012;text-decoration:none;padding:12px 20px;border-radius:5px;font-weight:bold">Open Your Blessing Card</a></p><p style="font-size:13px;color:#6e2530">— '+escapeHtml_(from)+'</p></div>';GmailApp.sendEmail(email,subject,message,{htmlBody:html,name:'Young Servants of the Lord'});return htmlPostMessage_({type:'EMAIL_RESULT',ok:true,message:'Blessing sent.'})}
function htmlPostMessage_(obj){var j=JSON.stringify(obj).replace(/<\/script/gi,'<\\/script');return HtmlService.createHtmlOutput('<script>window.top.postMessage('+j+',"*");</script>')}
function jsonp_(obj,cb){var j=JSON.stringify(obj).replace(/<\/script/gi,'<\\/script');if(cb&&/^[A-Za-z_$][\w$]*$/.test(cb))return ContentService.createTextOutput(cb+'('+j+')').setMimeType(ContentService.MimeType.JAVASCRIPT);return output_(j)}
function output_(text){return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.TEXT)}
function escapeHtml_(text){return String(text).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;')}
