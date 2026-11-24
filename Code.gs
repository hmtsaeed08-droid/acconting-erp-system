/**
 * ERP Accounting - Google Apps Script API
 * Stage 1: database schema, setup, authentication, sessions and RBAC.
 * Deploy as Web App: Execute as Me; access policy chosen for your organization.
 */
var ERP = {
  VERSION: "1.0.0-stage1",
  SESSION_HOURS: 8,
  MAX_ATTEMPTS: 5,
  LOCK_MINUTES: 15,
  SHEETS: {
    Users: ["ID","الاسم","الرقم الوظيفي","اسم المستخدم","كلمة المرور (مجزأة)","Salt","الدور","الحالة","تاريخ الإنشاء","آخر دخول","محاولات فاشلة","قفل حتى","تغيير كلمة المرور"],
    Accounts: ["رقم الحساب","اسم الحساب","نوع الحساب","الحساب الأب","الرصيد الافتتاحي","الرصيد الحالي","الحالة","مرتبط بموظف","الرقم الوظيفي للموظف"],
    JournalEntries: ["رقم القيد","التاريخ","البيان","نوع القيد","المستخدم المُدخل","حالة القيد","إجمالي المدين","إجمالي الدائن","المرجع","المرفقات"],
    JournalLines: ["رقم القيد","رقم الحساب","اسم الحساب","مدين","دائن","البيان التفصيلي","مركز التكلفة"],
    Transport: ["رقم الحركة","التاريخ","رقم الشاحنة","اسم السائق","نقطة الانطلاق","نقطة الوصول","الحمولة","الأجرة","المصروفات","الحالة","القيد المرتبط","ملاحظات"],
    SmartIncome: ["رقم العملية","التاريخ","مصدر الدخل","المبلغ","طريقة الدفع","الحساب المدين","الحساب الدائن","الوصف","القيد المرتبط","حالة الاعتماد"],
    Settings: ["المفتاح","القيمة","الوصف","آخر تعديل"],
    AuditLog: ["التاريخ","المستخدم","نوع العملية","التفاصيل","IP","المتصفح"],
    Employees: ["الرقم الوظيفي","الاسم","القسم","المنصب","الراتب","حساب الموظف المرتبط","الحالة","تاريخ التعيين"],
    SystemMeta: ["المفتاح","القيمة"]
  },
  ROLES: ["SUPER_ADMIN","ADMIN_1","ADMIN_2","VIEWER"]
};

function doGet() {
  return jsonOutput_({ok:true, service:"ERP Accounting API", version:ERP.VERSION, time:new Date().toISOString()});
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    var action = String(payload.action || "");
    if (action === "login") return jsonOutput_(login_(payload));
    if (action === "logout") return jsonOutput_(logout_(payload));
    if (action === "changePassword") return jsonOutput_(changePassword_(payload));
    if (action === "bootstrap") return jsonOutput_(bootstrap_(payload));
    if (action === "setup") return jsonOutput_(setup());
    lock.waitLock(20000);
    var user = requireSession_(payload.token);
    var result;
    if (action === "listUsers") {
      requireRole_(user, ["SUPER_ADMIN"]);
      result = {users: listUsers_()};
    } else if (action === "saveUser") {
      requireRole_(user, ["SUPER_ADMIN"]);
      result = saveUser_(payload.user, user);
    } else if (action === "disableUser") {
      requireRole_(user, ["SUPER_ADMIN"]);
      result = disableUser_(payload.userId, user);
    } else if (action === "resetPassword") {
      requireRole_(user, ["SUPER_ADMIN"]);
      result = resetPassword_(payload.userId, user);
    } else if (action === "dashboard") {
      requireRole_(user, ["SUPER_ADMIN","ADMIN_1","ADMIN_2","VIEWER"]);
      result = {user: publicUser_(user), stats:{users: listUsers_().length}};
    } else {
      throw new Error("الإجراء غير معروف");
    }
    return jsonOutput_({ok:true, data:result});
  } catch (err) {
    return jsonOutput_({ok:false, error:String(err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("افتح Apps Script من داخل Google Sheet أولاً");
  var created = [];
  Object.keys(ERP.SHEETS).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); created.push(name); }
    var headers = ERP.SHEETS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#17365D").setFontColor("#FFFFFF");
    }
  });
  seedSettings_();
  var users = sheet_("Users");
  if (users.getLastRow() < 2) {
    var tempPassword = randomPassword_();
    var salt = Utilities.getUuid();
    users.appendRow([Utilities.getUuid(),"مدير النظام","0001","admin",hashPassword_(tempPassword,salt),salt,"SUPER_ADMIN","نشط",new Date(),"",0,"",true]);
    audit_("SYSTEM","SETUP","إنشاء قاعدة البيانات والمدير الأول");
    return {ok:true, createdSheets:created, adminUsername:"admin", temporaryPassword:tempPassword, message:"احفظ كلمة المرور المؤقتة الآن؛ لن تظهر مرة أخرى."};
  }
  return {ok:true, createdSheets:created, message:"تم فحص وإنشاء هيكل الأوراق. لم يتم تغيير المستخدمين."};
}

function login_(p) {
  var username = String(p.username || "").trim().toLowerCase();
  var password = String(p.password || "");
  if (!username || !password) throw new Error("أدخل اسم المستخدم وكلمة المرور");
  var sh = sheet_("Users"), rows = sh.getDataRange().getValues();
  for (var i=1;i<rows.length;i++) {
    var r=rows[i];
    if (String(r[3]).toLowerCase() === username || String(r[2]) === username) {
      if (r[7] !== "نشط") throw new Error("الحساب موقوف");
      if (r[11] && new Date(r[11]).getTime() > Date.now()) throw new Error("الحساب مقفل مؤقتاً؛ حاول لاحقاً");
      if (hashPassword_(password,String(r[5])) !== String(r[4])) {
        var attempts = Number(r[10] || 0)+1;
        sh.getRange(i+1,11).setValue(attempts);
        if (attempts >= ERP.MAX_ATTEMPTS) {
          sh.getRange(i+1,12).setValue(new Date(Date.now()+ERP.LOCK_MINUTES*60000));
          sh.getRange(i+1,11).setValue(0);
        }
        audit_(username,"LOGIN_FAILED","محاولة دخول فاشلة");
        throw new Error("بيانات الدخول غير صحيحة");
      }
      sh.getRange(i+1,11,1,3).setValues([[0,"",new Date()]]);
      var token = Utilities.getUuid()+Utilities.getUuid();
      var expires = new Date(Date.now()+ERP.SESSION_HOURS*3600000);
      var sess = PropertiesService.getScriptProperties();
      sess.setProperty("ERP_SESSION_"+token, JSON.stringify({userId:r[0], username:r[3], role:r[6], expires:expires.toISOString(), lastSeen:Date.now()}));
      audit_(username,"LOGIN","تسجيل دخول ناجح");
      return {ok:true, token:token, expiresAt:expires.toISOString(), mustChangePassword:Boolean(r[12]), user:{id:r[0],name:r[1],employeeNo:r[2],username:r[3],role:r[6]}};
    }
  }
  audit_(username,"LOGIN_FAILED","اسم مستخدم غير موجود");
  throw new Error("بيانات الدخول غير صحيحة");
}

function requireSession_(token) {
  if (!token) throw new Error("يلزم تسجيل الدخول");
  var props=PropertiesService.getScriptProperties(), key="ERP_SESSION_"+token, raw=props.getProperty(key);
  if (!raw) throw new Error("الجلسة غير صالحة؛ سجل الدخول مجدداً");
  var s=JSON.parse(raw);
  if (new Date(s.expires).getTime()<Date.now() || Date.now()-Number(s.lastSeen)>ERP.SESSION_HOURS*3600000) {
    props.deleteProperty(key); throw new Error("انتهت الجلسة؛ سجل الدخول مجدداً");
  }
  s.lastSeen=Date.now(); props.setProperty(key,JSON.stringify(s));
  var u=findUserById_(s.userId);
  if (!u || u.status!=="نشط") { props.deleteProperty(key); throw new Error("الحساب غير نشط"); }
  return u;
}
function logout_(p) {
  var token=String(p.token||"");
  if (token) PropertiesService.getScriptProperties().deleteProperty("ERP_SESSION_"+token);
  audit_("SESSION","LOGOUT","تسجيل خروج");
  return {ok:true};
}
function changePassword_(p) {
  var user=requireSession_(p.token), old=String(p.oldPassword||""), next=String(p.newPassword||"");
  if (next.length<10) throw new Error("كلمة المرور الجديدة يجب ألا تقل عن 10 أحرف");
  var sh=sheet_("Users"), rows=sh.getDataRange().getValues();
  for(var i=1;i<rows.length;i++) if(String(rows[i][0])===user.id) {
    if(hashPassword_(old,String(rows[i][5]))!==String(rows[i][4])) throw new Error("كلمة المرور الحالية غير صحيحة");
    var salt=Utilities.getUuid();
    sh.getRange(i+1,5,1,2).setValues([[hashPassword_(next,salt),salt]]);
    sh.getRange(i+1,13).setValue(false);
    audit_(user.username,"PASSWORD_CHANGED","تغيير كلمة المرور");
    return {ok:true};
  }
  throw new Error("المستخدم غير موجود");
}
function bootstrap_(p) {
  var u=requireSession_(p.token);
  return {user:publicUser_(u), roles:ERP.ROLES};
}
function listUsers_() {
  var rows=sheet_("Users").getDataRange().getValues(), out=[];
  for(var i=1;i<rows.length;i++) out.push({id:rows[i][0],name:rows[i][1],employeeNo:rows[i][2],username:rows[i][3],role:rows[i][6],status:rows[i][7],createdAt:rows[i][8],lastLogin:rows[i][9]});
  return out;
}
function saveUser_(u, actor) {
  if(!u || !u.name || !u.username || ERP.ROLES.indexOf(u.role)<0) throw new Error("بيانات المستخدم غير مكتملة");
  var sh=sheet_("Users"), rows=sh.getDataRange().getValues();
  for(var i=1;i<rows.length;i++) if(String(rows[i][3]).toLowerCase()===String(u.username).toLowerCase()) throw new Error("اسم المستخدم مستخدم مسبقاً");
  if(u.role==="SUPER_ADMIN" && rows.slice(1).some(function(r){return r[6]==="SUPER_ADMIN" && r[7]==="نشط";})) throw new Error("يوجد مدير عام نشط بالفعل");
  var pass=randomPassword_(), salt=Utilities.getUuid();
  sh.appendRow([Utilities.getUuid(),u.name,String(u.employeeNo||""),u.username,hashPassword_(pass,salt),salt,u.role,"نشط",new Date(),"",0,"",true]);
  audit_(actor.username,"USER_CREATED","إنشاء المستخدم "+u.username);
  return {temporaryPassword:pass, username:u.username};
}
function disableUser_(id, actor) {
  var sh=sheet_("Users"), rows=sh.getDataRange().getValues();
  for(var i=1;i<rows.length;i++) if(String(rows[i][0])===String(id)) {
    if(rows[i][6]==="SUPER_ADMIN") throw new Error("لا يمكن تعطيل المدير العام");
    sh.getRange(i+1,8).setValue("موقوف");
    audit_(actor.username,"USER_DISABLED","تعطيل "+rows[i][3]);
    return {ok:true};
  }
  throw new Error("المستخدم غير موجود");
}
function resetPassword_(id, actor) {
  var sh=sheet_("Users"), rows=sh.getDataRange().getValues();
  for(var i=1;i<rows.length;i++) if(String(rows[i][0])===String(id)) {
    var pass=randomPassword_(), salt=Utilities.getUuid();
    sh.getRange(i+1,5,1,2).setValues([[hashPassword_(pass,salt),salt]]);
    sh.getRange(i+1,13).setValue(true);
    audit_(actor.username,"PASSWORD_RESET","إعادة تعيين كلمة مرور "+rows[i][3]);
    return {temporaryPassword:pass};
  }
  throw new Error("المستخدم غير موجود");
}
function findUserById_(id) {
  var rows=sheet_("Users").getDataRange().getValues();
  for(var i=1;i<rows.length;i++) if(String(rows[i][0])===String(id)) return {id:rows[i][0],name:rows[i][1],employeeNo:rows[i][2],username:rows[i][3],role:rows[i][6],status:rows[i][7]};
  return null;
}
function publicUser_(u) { return {id:u.id,name:u.name,employeeNo:u.employeeNo,username:u.username,role:u.role}; }
function requireRole_(u, roles) { if(roles.indexOf(u.role)<0) throw new Error("ليس لديك صلاحية لتنفيذ هذه العملية"); }
function hashPassword_(password,salt) {
  var bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(salt)+"|"+String(password),Utilities.Charset.UTF_8);
  return bytes.map(function(b){var v=(b<0?b+256:b).toString(16);return v.length===1?"0"+v:v;}).join("");
}
function randomPassword_() {
  var chars="ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%";
  var out=""; for(var i=0;i<16;i++) out+=chars.charAt(Math.floor(Math.random()*chars.length));
  return out;
}
function sheet_(name) {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if(!sh) throw new Error("ورقة "+name+" غير موجودة؛ شغّل setup()");
  return sh;
}
function seedSettings_() {
  var sh=sheet_("Settings");
  if(sh.getLastRow()>1) return;
  var now=new Date();
  sh.getRange(2,1,5,4).setValues([
    ["COMPANY_NAME","اسم المنشأة","بيانات المنشأة",now],
    ["CURRENCY","SAR","العملة الأساسية",now],
    ["VAT_RATE","0.15","نسبة ضريبة داخلية اختيارية",now],
    ["POLL_SECONDS","5","مدة تحديث البيانات بالثواني",now],
    ["TIMEZONE","Asia/Riyadh","المنطقة الزمنية",now]
  ]);
}
function audit_(user,action,details) {
  try { sheet_("AuditLog").appendRow([new Date(),user,action,details,"غير متاح عبر Apps Script","Web App"]); } catch(e) {}
}
function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
