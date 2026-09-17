/**
 * A Blessing to Share
 * Persistence + Gmail + Google Sheets + Google Drive
 *
 * CHANGE FROM YOUR VERSION:
 *   sendEmail_() used to require the recipient address to end in
 *   "@gmail.com". That's why "Send Blessing" was failing for most
 *   people, since a friend's address is rarely a Gmail address.
 *   It now accepts any normally-formatted email address.
 */
 
var ADMIN_PASSWORD_HASH =
  'fd7adfefd51bfceb2db1e9a34d1baa285718ae937975c0ad891947051cc8b05';
 
var SHEET_NAME = 'Creations';
 
var DRIVE_FOLDER_NAME =
  'A Blessing to Share Photos';
 
 
/* ============================================================
   AUTHORIZATION
   ============================================================ */
 
function authorizeServices() {
 
  var folders =
    DriveApp.getFoldersByName(
      DRIVE_FOLDER_NAME
    );
 
  Logger.log(
    'Drive authorization successful.'
  );
 
  Logger.log(
    'Folder exists: ' +
    folders.hasNext()
  );
 
  var ss = ss_();
 
  Logger.log(
    'Sheets authorization successful.'
  );
 
  Logger.log(
    'Spreadsheet: ' +
    ss.getName()
  );
 
  var aliases =
    GmailApp.getAliases();
 
  Logger.log(
    'Gmail authorization successful.'
  );
 
  Logger.log(
    'Gmail aliases: ' +
    aliases.length
  );
 
  Logger.log(
    '========================================'
  );
 
  Logger.log(
    'A BLESSING TO SHARE authorization complete.'
  );
 
  Logger.log(
    'Drive: OK'
  );
 
  Logger.log(
    'Sheets: OK'
  );
 
  Logger.log(
    'Gmail: OK'
  );
 
  Logger.log(
    '========================================'
  );
}
 
 
function authorizeDrive() {
  authorizeServices();
}
 
 
/* ============================================================
   GET
   ============================================================ */
 
function doGet(e) {
 
  var p =
    e && e.parameter
      ? e.parameter
      : {};
 
  var action =
    String(p.action || '');
 
  try {
 
    if(action === 'getHistory') {
 
      return jsonp_(
        historyForUser_(
          p.userId,
          p.userKey
        ),
        p.callback
      );
 
    }
 
    if(action === 'getPublic') {
 
      return jsonp_(
        publicCreation_(
          p.creationId
        ),
        p.callback
      );
 
    }
 
    if(action === 'adminList') {
 
      return jsonp_(
        adminList_(
          p.adminToken
        ),
        p.callback
      );
 
    }
 
    if(action === 'adminLogin') {
 
      return jsonp_(
        adminLoginData_(p),
        p.callback
      );
 
    }
 
    return output_(
      'A Blessing to Share backend is running.'
    );
 
  } catch(err) {
 
    return jsonp_(
      {
        ok:false,
        message:err.message
      },
      p.callback
    );
 
  }
 
}
 
 
/* ============================================================
   POST
   ============================================================ */
 
function doPost(e) {
 
  try {
 
    var p =
      e && e.parameter
        ? e.parameter
        : {};
 
    var action =
      String(
        p.action || 'sendEmail'
      );
 
    if(action === 'save') {
      return saveCreation_(p);
    }
 
    if(action === 'delete') {
      return deleteCreation_(p);
    }
 
    if(action === 'adminDelete') {
      return adminDelete_(p);
    }
 
    if(action === 'adminLogin') {
      return htmlPostMessage_(
        adminLoginData_(p)
      );
    }
 
    if(action === 'sendEmail') {
      return sendEmail_(p);
    }
 
    return output_(
      'ERROR: Unknown action.'
    );
 
  } catch(err) {
 
    return htmlPostMessage_({
      type:'SAVE_RESULT',
      ok:false,
      message:err.message
    });
 
  }
 
}
 
 
/* ============================================================
   SHEETS
   ============================================================ */
 
function ss_() {
 
  var props =
    PropertiesService
      .getScriptProperties();
 
  var id =
    props.getProperty(
      'CREATIONS_SHEET_ID'
    );
 
  var ss =
    id
      ? SpreadsheetApp.openById(id)
      : SpreadsheetApp.create(
          'A Blessing to Share — Creations'
        );
 
  if(!id) {
 
    props.setProperty(
      'CREATIONS_SHEET_ID',
      ss.getId()
    );
 
  }
 
  var sh =
    ss.getSheetByName(
      SHEET_NAME
    );
 
  if(!sh) {
 
    sh =
      ss.insertSheet(
        SHEET_NAME
      );
 
  }
 
  if(sh.getLastRow() === 0) {
 
    sh.appendRow([
      'id',
      'userId',
      'userKeyHash',
      'cardType',
      'stateJson',
      'createdAt',
      'updatedAt'
    ]);
 
  }
 
  return sh;
 
}
 
 
/* ============================================================
   DRIVE
   ============================================================ */
 
function folder_() {
 
  var props =
    PropertiesService
      .getScriptProperties();
 
  var id =
    props.getProperty(
      'PHOTO_FOLDER_ID'
    );
 
  if(!id) {
 
    throw new Error(
      'PHOTO_FOLDER_ID is not configured.'
    );
 
  }
 
  return DriveApp.getFolderById(id);
 
}
 
 
/* ============================================================
   HELPERS
   ============================================================ */
 
function hash_(s) {
 
  var b =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(s),
      Utilities.Charset.UTF_8
    );
 
  return b.map(function(x) {
 
    var v =
      (
        x < 0
          ? x + 256
          : x
      ).toString(16);
 
    return v.length === 1
      ? '0' + v
      : v;
 
  }).join('');
 
}
 
 
/* ============================================================
   SHORT CREATION ID
   ============================================================ */
 
function shortId_() {
 
  var chars =
    'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789';
 
  var result = '';
 
  for(var i=0;i<8;i++) {
 
    result +=
      chars.charAt(
        Math.floor(
          Math.random() *
          chars.length
        )
      );
 
  }
 
  return result;
 
}
 
 
function uniqueCreationId_() {
 
  var sh = ss_();
 
  var rows =
    sh.getDataRange()
      .getValues();
 
  for(
    var attempt=0;
    attempt<20;
    attempt++
  ) {
 
    var id =
      shortId_();
 
    var exists=false;
 
    for(
      var i=1;
      i<rows.length;
      i++
    ) {
 
      if(
        String(rows[i][0]) === id
      ) {
 
        exists=true;
        break;
 
      }
 
    }
 
    if(!exists) {
      return id;
    }
 
  }
 
  throw new Error(
    'Could not create a unique card ID.'
  );
 
}
 
 
/* ============================================================
   USER AUTH
   ============================================================ */
 
function authUser_(p) {
 
  if(!p.userId || !p.userKey) {
 
    return null;
 
  }
 
  return {
    id:String(p.userId),
    keyHash:hash_(p.userKey)
  };
 
}
 
 
/* ============================================================
   PHOTOS
   ============================================================ */
 
function processPhotos_(
  photos,
  oldPhotos
) {
 
  photos =
    Array.isArray(photos)
      ? photos.slice(0,3)
      : [];
 
  oldPhotos =
    Array.isArray(oldPhotos)
      ? oldPhotos
      : [];
 
  var folder =
    folder_();
 
  return photos.map(function(p) {
 
    /*
     * Existing Drive photo.
     */
 
    if(
      p.fileId &&
      p.url
    ) {
 
      return {
        id:p.id,
        fileId:p.fileId,
        url:p.url,
        name:p.name || '',
        x:Number(p.x) || 50,
        y:Number(p.y) || 50,
        width:Number(p.width) || 180,
        height:Number(p.height) || 180,
        scale:Number(p.scale) || 1,
        rotation:Number(p.rotation) || 0,
        ratio:Number(p.ratio) || 1
      };
 
    }
 
 
    /*
     * If the photo has no data URL,
     * try to preserve an existing photo
     * from the previous saved state.
     */
 
    var previous = null;
 
    for(
      var i=0;
      i<oldPhotos.length;
      i++
    ) {
 
      if(
        oldPhotos[i] &&
        oldPhotos[i].id === p.id
      ) {
 
        previous =
          oldPhotos[i];
 
        break;
 
      }
 
    }
 
    if(
      !p.dataUrl &&
      previous &&
      previous.url
    ) {
 
      return {
        id:p.id,
        fileId:previous.fileId,
        url:previous.url,
        name:p.name || previous.name || '',
        x:Number(p.x) || 50,
        y:Number(p.y) || 50,
        width:Number(p.width) || 180,
        height:Number(p.height) || 180,
        scale:Number(p.scale) || 1,
        rotation:Number(p.rotation) || 0,
        ratio:Number(p.ratio) || 1
      };
 
    }
 
    if(!p.dataUrl) {
 
      throw new Error(
        'Photo data is missing.'
      );
 
    }
 
    var m =
      String(p.dataUrl).match(
        /^data:image\/(jpeg|jpg|png|webp);base64,(.+)$/
      );
 
    if(!m) {
 
      throw new Error(
        'Unsupported photo data.'
      );
 
    }
 
    var mime =
      'image/jpeg';
 
    if(m[1] === 'png') {
      mime='image/png';
    }
 
    if(m[1] === 'webp') {
      mime='image/webp';
    }
 
    var blob =
      Utilities.newBlob(
        Utilities.base64Decode(m[2]),
        mime,
        'blessing-' +
        (p.id || shortId_())
      );
 
    var file =
      folder.createFile(
        blob
      );
 
    file.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
 
    return {
      id:p.id || shortId_(),
      fileId:file.getId(),
      url:
        'https://drive.google.com/uc?export=view&id=' +
        file.getId(),
      name:p.name || '',
      x:Number(p.x) || 50,
      y:Number(p.y) || 50,
      width:Number(p.width) || 180,
      height:Number(p.height) || 180,
      scale:Number(p.scale) || 1,
      rotation:Number(p.rotation) || 0,
      ratio:Number(p.ratio) || 1
    };
 
  });
 
}
 
 
/* ============================================================
   SAVE CREATION
   ============================================================ */
 
function saveCreation_(p) {
 
  var sh = ss_();
 
  var creationId =
    String(
      p.creationId ||
      uniqueCreationId_()
    );
 
  var state;
 
  try {
 
    state =
      JSON.parse(
        String(
          p.state || '{}'
        )
      );
 
  } catch(err) {
 
    throw new Error(
      'Invalid card state.'
    );
 
  }
 
  var rows =
    sh.getDataRange()
      .getValues();
 
  var found=-1;
 
  var old={};
 
  var existingUserId='';
 
  var existingKeyHash='';
 
 
  /*
   * Find existing card.
   */
 
  for(
    var i=1;
    i<rows.length;
    i++
  ) {
 
    if(
      String(rows[i][0]) === creationId
    ) {
 
      found=i+1;
 
      try {
 
        old =
          JSON.parse(
            rows[i][4] || '{}'
          );
 
      } catch(err) {
 
        old={};
 
      }
 
      existingUserId =
        String(rows[i][1] || '');
 
      existingKeyHash =
        String(rows[i][2] || '');
 
      break;
 
    }
 
  }
 
 
  /*
   * Existing cards require the same
   * edit credentials.
   */
 
  if(found >= 0) {
 
    var u =
      authUser_(p);
 
    if(!u) {
 
      throw new Error(
        'This saved card belongs to another session. Please create a new card.'
      );
 
    }
 
    if(
      existingUserId !== u.id ||
      existingKeyHash !== u.keyHash
    ) {
 
      throw new Error(
        'You do not own this creation.'
      );
 
    }
 
  }
 
 
  /*
   * New card:
   *
   * If the browser already has an identity,
   * preserve it.
   *
   * Otherwise create one.
   */
 
  var userId =
    String(
      p.userId || ''
    );
 
  var userKey =
    String(
      p.userKey || ''
    );
 
  if(!userId) {
 
    userId =
      'guest-' +
      Utilities.getUuid();
 
  }
 
  if(!userKey) {
 
    userKey =
      Utilities.getUuid() +
      Utilities.getUuid();
 
  }
 
  var userKeyHash =
    hash_(userKey);
 
 
  /*
   * Process uploaded photos.
   */
 
  state.photos =
    processPhotos_(
      state.photos,
      old.photos
    );
 
 
  state.creationId =
    creationId;
 
 
  var now =
    new Date().toISOString();
 
 
  if(found < 0) {
 
    sh.appendRow([
      creationId,
      userId,
      userKeyHash,
      String(
        p.cardType ||
        'Blessing Card'
      ),
      JSON.stringify(state),
      now,
      now
    ]);
 
  } else {
 
    sh.getRange(
      found,
      4,
      1,
      4
    ).setValues([[
      String(
        p.cardType ||
        'Blessing Card'
      ),
      JSON.stringify(state),
      rows[found-1][5],
      now
    ]]);
 
    /*
     * Preserve the existing owner.
     */
 
    userId =
      existingUserId;
 
  }
 
 
  return jsonOutput_({
    type:'SAVE_RESULT',
    ok:true,
    creationId:creationId,
    userId:userId,
    userKey:userKey
  });
 
}
 
 
/* ============================================================
   PUBLIC CARD
   ============================================================ */
 
function publicCreation_(id) {
 
  id =
    String(id || '')
      .trim();
 
  if(!id) {
 
    return {
      ok:false,
      message:'Card ID missing.'
    };
 
  }
 
  var sh=ss_();
 
  var rows=
    sh.getDataRange()
      .getValues();
 
 
  for(
    var i=1;
    i<rows.length;
    i++
  ) {
 
    if(
      String(rows[i][0]) === id
    ) {
 
      var state={};
 
      try {
 
        state =
          JSON.parse(
            rows[i][4] || '{}'
          );
 
      } catch(err) {
 
        return {
          ok:false,
          message:'Saved card data is invalid.'
        };
 
      }
 
      return {
        ok:true,
        state:state,
        id:String(rows[i][0])
      };
 
    }
 
  }
 
 
  return {
    ok:false,
    message:'Creation not found.'
  };
 
}
 
 
/* ============================================================
   DELETE
   ============================================================ */
 
function deleteCreation_(p) {
 
  var u =
    authUser_(p);
 
  if(!u) {
 
    throw new Error(
      'User identity missing.'
    );
 
  }
 
  var sh=ss_();
 
  var rows=
    sh.getDataRange()
      .getValues();
 
 
  for(
    var i=1;
    i<rows.length;
    i++
  ) {
 
    if(
      String(rows[i][0]) ===
      String(p.creationId)
    ) {
 
      if(
        String(rows[i][1]) !== u.id ||
        String(rows[i][2]) !== u.keyHash
      ) {
 
        throw new Error(
          'You do not own this creation.'
        );
 
      }
 
      sh.deleteRow(i+1);
 
      return output_('OK');
 
    }
 
  }
 
 
  return output_('OK');
 
}
 
 
/* ============================================================
   HISTORY
   ============================================================ */
 
function historyForUser_(
  uid,
  key
) {
 
  if(!uid || !key) {
 
    throw new Error(
      'User identity missing.'
    );
 
  }
 
  var sh=ss_();
 
  var rows=
    sh.getDataRange()
      .getValues();
 
  var h=
    hash_(key);
 
  var items=[];
 
 
  for(
    var i=1;
    i<rows.length;
    i++
  ) {
 
    if(
      String(rows[i][1]) === String(uid) &&
      String(rows[i][2]) === h
    ) {
 
      items.push(
        row_(rows[i])
      );
 
    }
 
  }
 
 
  items.reverse();
 
 
  return {
    ok:true,
    items:items
  };
 
}
 
 
function row_(r) {
 
  return {
    id:r[0],
    userId:r[1],
    cardType:r[3],
    state:JSON.parse(
      r[4] || '{}'
    ),
    createdAt:r[5],
    updatedAt:r[6]
  };
 
}
 
 
/* ============================================================
   ADMIN
   ============================================================ */
 
function adminLoginData_(p) {
 
  var ok =
    hash_(
      String(
        p.password || ''
      )
    ) ===
    ADMIN_PASSWORD_HASH;
 
 
  if(!ok) {
 
    return {
      type:'ADMIN_LOGIN',
      ok:false,
      message:
        'Invalid administrator password.'
    };
 
  }
 
 
  var token =
    Utilities.getUuid() +
    '-' +
    Utilities.getUuid();
 
 
  CacheService
    .getScriptCache()
    .put(
      'admin:' + token,
      '1',
      21600
    );
 
 
  return {
    type:'ADMIN_LOGIN',
    ok:true,
    token:token
  };
 
}
 
 
function adminLogin_(p) {
 
  return htmlPostMessage_(
    adminLoginData_(p)
  );
 
}
 
 
function adminList_(token) {
 
  if(
    !CacheService
      .getScriptCache()
      .get(
        'admin:' +
        String(token || '')
      )
  ) {
 
    throw new Error(
      'Administrator session expired.'
    );
 
  }
 
 
  var sh=ss_();
 
  var rows=
    sh.getDataRange()
      .getValues();
 
  var items=[];
 
 
  for(
    var i=1;
    i<rows.length;
    i++
  ) {
 
    items.push(
      row_(rows[i])
    );
 
  }
 
 
  items.reverse();
 
 
  return {
    ok:true,
    items:items
  };
 
}
 
 
function adminDelete_(p) {
 
  if(
    !CacheService
      .getScriptCache()
      .get(
        'admin:' +
        String(
          p.adminToken || ''
        )
      )
  ) {
 
    throw new Error(
      'Administrator session expired.'
    );
 
  }
 
 
  var sh=ss_();
 
  var rows=
    sh.getDataRange()
      .getValues();
 
 
  for(
    var i=1;
    i<rows.length;
    i++
  ) {
 
    if(
      String(rows[i][0]) ===
      String(p.creationId)
    ) {
 
      sh.deleteRow(i+1);
 
      return output_('OK');
 
    }
 
  }
 
 
  return output_('OK');
 
}
 
 
/* ============================================================
   GMAIL
   ============================================================ */
 
function sendEmail_(p) {
 
  var email =
    String(
      p.email || ''
    ).trim();
 
  var to =
    String(
      p.to || 'you'
    ).trim();
 
  var from =
    String(
      p.from ||
      'Young Servants of the Lord'
    ).trim();
 
  var subject =
    String(
      p.subject ||
      'A blessing for you'
    ).trim();
 
  var message =
    String(
      p.message || ''
    ).trim();
 
  var cardLink =
    String(
      p.cardLink || ''
    ).trim();
 
 
  /*
   * CHANGED: accept any properly-formatted email address
   * instead of requiring "@gmail.com". The blessing is being
   * sent TO a friend, who will rarely have a Gmail address.
   */
  if(
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(
      email
    )
  ) {
 
    return jsonOutput_({
      type:'EMAIL_RESULT',
      ok:false,
      message:
        'Please enter a valid email address.'
    });
 
  }
 
 
  if(!cardLink) {
 
    return jsonOutput_({
      type:'EMAIL_RESULT',
      ok:false,
      message:
        'The card link is missing.'
    });
 
  }
 
 
  var html =
    '<div style="' +
    'font-family:Arial,sans-serif;' +
    'line-height:1.6;' +
    'color:#2a0a10;' +
    'max-width:620px">' +
 
    '<h2 style="color:#6e2530">' +
    'A blessing for ' +
    escapeHtml_(to) +
    '</h2>' +
 
    '<p>' +
    escapeHtml_(message)
      .replace(/\n/g,'<br>') +
    '</p>' +
 
    '<p style="margin:24px 0">' +
 
    '<a href="' +
    escapeHtml_(cardLink) +
    '" ' +
 
    'style="' +
    'display:inline-block;' +
    'background:#c99a3d;' +
    'color:#241012;' +
    'text-decoration:none;' +
    'padding:12px 20px;' +
    'border-radius:5px;' +
    'font-weight:bold">' +
 
    'Open Your Blessing Card' +
 
    '</a>' +
 
    '</p>' +
 
    '<p style="font-size:13px;color:#6e2530">' +
    '— ' +
    escapeHtml_(from) +
    '</p>' +
 
    '</div>';
 
 
  GmailApp.sendEmail(
    email,
    subject,
    message ||
      'You have received a blessing card.',
    {
      htmlBody:html,
      name:'Young Servants of the Lord'
    }
  );
 
 
  return jsonOutput_({
    type:'EMAIL_RESULT',
    ok:true,
    message:
      'Blessing sent successfully.'
  });
 
}
 
 
/* ============================================================
   OUTPUT
   ============================================================ */
 
/*
 * Plain JSON response — used for the "save" and "sendEmail" actions,
 * which the website reads with fetch()+r.json(). (Other actions like
 * adminLogin keep using htmlPostMessage_ below, in case an admin
 * panel relies on the hidden-iframe + postMessage pattern.)
 */
function jsonOutput_(obj) {
 
  return ContentService
    .createTextOutput(
      JSON.stringify(obj)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
 
}
 
 
function htmlPostMessage_(obj) {
 
  var j =
    JSON.stringify(obj)
      .replace(
        /<\/script/gi,
        '<\\/script'
      );
 
 
  return HtmlService
    .createHtmlOutput(
      '<script>' +
      'window.top.postMessage(' +
      j +
      ',"*");' +
      '</script>'
    );
 
}
 
 
function jsonp_(obj,cb) {
 
  var j =
    JSON.stringify(obj)
      .replace(
        /<\/script/gi,
        '<\\/script'
      );
 
 
  if(
    cb &&
    /^[A-Za-z_$][\w$]*$/.test(cb)
  ) {
 
    return ContentService
      .createTextOutput(
        cb +
        '(' +
        j +
        ')'
      )
      .setMimeType(
        ContentService.MimeType.JAVASCRIPT
      );
 
  }
 
 
  return output_(j);
 
}
 
 
function output_(text) {
 
  return ContentService
    .createTextOutput(text)
    .setMimeType(
      ContentService.MimeType.TEXT
    );
 
}
 
 
function escapeHtml_(text) {
 
  return String(text)
    .replace(
      /&/g,
      '&amp;'
    )
    .replace(
      /</g,
      '&lt;'
    )
    .replace(
      />/g,
      '&gt;'
    )
    .replace(
      /"/g,
      '&quot;'
    )
    .replace(
      /'/g,
      '&#039;'
    );
 
}
 