const APP_NAME = 'Revival Labs Classroom';
// IMPORTANT: Change this before running setupClassroom() for the first time.
// Use a long, unique password. After setup, the admin password is stored as a hash in Sheets.
const INITIAL_ADMIN_USERNAME = 'admin';
const INITIAL_ADMIN_PASSWORD = 'CHANGE_THIS_TO_A_LONG_UNIQUE_PASSWORD_123!';
const SESSION_SECONDS = 21600; // 6 hours

const TABLES = {
  Users: ['id','username','displayName','salt','passwordHash','role','classId','createdAt','active'],
  Classes: ['id','name','level','description','duration','active'],
  Lessons: ['id','classId','title','content','resourceUrl','published','createdBy','createdAt'],
  Assignments: ['id','classId','title','instructions','points','dueDate','resourceUrl','published','createdBy','createdAt'],
  Submissions: ['id','assignmentId','studentId','studentName','content','projectUrl','submittedAt','status','grade','feedback','gradedBy','gradedAt'],
  Announcements: ['id','title','body','createdBy','createdAt','published']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once from the Apps Script editor. Creates the database and initial admin account. */
function setupClassroom() {
  if (INITIAL_ADMIN_PASSWORD.indexOf('CHANGE_THIS_') === 0 || INITIAL_ADMIN_PASSWORD.length < 14) {
    throw new Error('Edit INITIAL_ADMIN_PASSWORD near the top of Code.gs to a unique password (at least 14 characters), save, then run setupClassroom() again.');
  }
  const props = PropertiesService.getScriptProperties();
  let ss;
  const existingId = props.getProperty('CLASSROOM_SPREADSHEET_ID');
  if (existingId) {
    ss = SpreadsheetApp.openById(existingId);
  } else {
    ss = SpreadsheetApp.create('Revival Labs Classroom Database');
    props.setProperty('CLASSROOM_SPREADSHEET_ID', ss.getId());
  }
  Object.keys(TABLES).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, TABLES[name].length).setValues([TABLES[name]]);
      sheet.setFrozenRows(1);
      sheet.getRange(1,1,1,TABLES[name].length).setFontWeight('bold').setBackground('#153c35').setFontColor('#ffffff');
      sheet.autoResizeColumns(1, TABLES[name].length);
    }
  });
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) ss.deleteSheet(defaultSheet);
  seedClasses_();
  const users = readTable_('Users');
  if (!users.some(u => u.username === INITIAL_ADMIN_USERNAME)) {
    const salt = Utilities.getUuid().replace(/-/g, '');
    appendRow_('Users', {
      id: id_(), username: INITIAL_ADMIN_USERNAME, displayName: 'Revival Labs Admin',
      salt: salt, passwordHash: hashPassword_(INITIAL_ADMIN_PASSWORD, salt),
      role: 'admin', classId: 'all', createdAt: now_(), active: 'TRUE'
    });
  }
  Logger.log('Classroom database created: ' + ss.getUrl());
  Logger.log('Admin username: ' + INITIAL_ADMIN_USERNAME);
  Logger.log('After confirming setup, change or remove the initial password constant and deploy the web app.');
  return 'Setup complete. Database: ' + ss.getUrl() + '. Admin username: ' + INITIAL_ADMIN_USERNAME;
}

function seedClasses_() {
  const rows = readTable_('Classes');
  const defaults = [
    {id:'class-a',name:'Class A — Foundations',level:'Beginner',description:'Learn the fundamentals of revival architecture, HTTP, JSON, and your first local backend.',duration:'6 weeks',active:'TRUE'},
    {id:'class-b',name:'Class B — Backend Development',level:'Intermediate',description:'Build APIs, persistent player data, and connect a client to a backend.',duration:'8 weeks',active:'TRUE'},
    {id:'class-c',name:'Class C — Systems Engineering',level:'Advanced',description:'Design, test, and document a complete original revival-style backend prototype.',duration:'10 weeks',active:'TRUE'}
  ];
  defaults.forEach(item => { if (!rows.some(r => r.id === item.id)) appendRow_('Classes', item); });
}

/** All browser requests enter through this dispatcher. */
function api(action, payload) {
  payload = payload || {};
  try {
    if (action === 'login') return login_(payload);
    if (action === 'register') return register_(payload);
    if (action === 'bootstrapStatus') return {ok:true, initialized:!!PropertiesService.getScriptProperties().getProperty('CLASSROOM_SPREADSHEET_ID')};
    const user = requireUser_(payload.token);
    switch (action) {
      case 'getDashboard': return dashboard_(user);
      case 'getClassroom': return classroom_(user, payload.classId);
      case 'createLesson': requireAdmin_(user); return createLesson_(user, payload);
      case 'createAssignment': requireAdmin_(user); return createAssignment_(user, payload);
      case 'submitAssignment': return submitAssignment_(user, payload);
      case 'gradeSubmission': requireAdmin_(user); return gradeSubmission_(user, payload);
      case 'createAnnouncement': requireAdmin_(user); return createAnnouncement_(user, payload);
      case 'listSubmissions': requireAdmin_(user); return listSubmissions_(payload.assignmentId);
      case 'setEnrollment': requireAdmin_(user); return setEnrollment_(payload.username, payload.classId);
      case 'logout': CacheService.getScriptCache().remove('session:' + payload.token); return {ok:true};
      default: throw new Error('Unknown action.');
    }
  } catch (err) {
    return {ok:false, error: safeError_(err)};
  }
}

function register_(p) {
  ensureSetup_();
  const username = String(p.username || '').trim().toLowerCase();
  const displayName = String(p.displayName || '').trim();
  const password = String(p.password || '');
  const classId = String(p.classId || 'class-a');
  if (!/^[a-z0-9_]{3,24}$/.test(username)) throw new Error('Username must be 3–24 characters using letters, numbers, or underscores.');
  if (displayName.length < 2 || displayName.length > 48) throw new Error('Display name must be 2–48 characters.');
  if (password.length < 10 || password.length > 100) throw new Error('Password must be 10–100 characters.');
  if (!readTable_('Classes').some(c => c.id === classId && c.active === 'TRUE')) throw new Error('Choose a valid class.');
  if (readTable_('Users').some(u => u.username.toLowerCase() === username)) throw new Error('That username is already taken.');
  const salt = Utilities.getUuid().replace(/-/g, '');
  appendRow_('Users', {id:id_(), username, displayName, salt, passwordHash:hashPassword_(password,salt), role:'student', classId, createdAt:now_(), active:'TRUE'});
  const user = readTable_('Users').find(u => u.username === username);
  return sessionFor_(user);
}

function login_(p) {
  ensureSetup_();
  const username = String(p.username || '').trim().toLowerCase();
  const password = String(p.password || '');
  const user = readTable_('Users').find(u => u.username.toLowerCase() === username && u.active === 'TRUE');
  if (!user || hashPassword_(password, user.salt) !== user.passwordHash) throw new Error('Incorrect username or password.');
  return sessionFor_(user);
}

function sessionFor_(user) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  const safeUser = {id:user.id, username:user.username, displayName:user.displayName, role:user.role, classId:user.classId};
  CacheService.getScriptCache().put('session:' + token, JSON.stringify(safeUser), SESSION_SECONDS);
  return {ok:true, token, user:safeUser, expiresIn:SESSION_SECONDS};
}

function requireUser_(token) {
  if (!token) throw new Error('Your session has expired. Please sign in again.');
  const raw = CacheService.getScriptCache().get('session:' + token);
  if (!raw) throw new Error('Your session has expired. Please sign in again.');
  const session = JSON.parse(raw);
  const user = readTable_('Users').find(u => u.id === session.id && u.active === 'TRUE');
  if (!user) throw new Error('This account is disabled.');
  return {id:user.id, username:user.username, displayName:user.displayName, role:user.role, classId:user.classId};
}
function requireAdmin_(user) { if (user.role !== 'admin' && user.role !== 'instructor') throw new Error('Instructor permission required.'); }

function dashboard_(user) {
  const classes = readTable_('Classes').filter(c => c.active === 'TRUE');
  const allowed = user.role === 'admin' || user.role === 'instructor' ? classes : classes.filter(c => c.id === user.classId);
  const assignments = readTable_('Assignments').filter(a => a.published === 'TRUE' && allowed.some(c => c.id === a.classId)).map(a => {
    const sub = readTable_('Submissions').find(s => s.assignmentId === a.id && s.studentId === user.id);
    return Object.assign({}, a, {mySubmission: sub ? publicSubmission_(sub) : null, className:(classes.find(c=>c.id===a.classId)||{}).name || ''});
  });
  return {ok:true, user, classes:allowed, assignments, lessons:readTable_('Lessons').filter(l => l.published === 'TRUE' && allowed.some(c => c.id === l.classId)).slice(-8).reverse(), announcements:readTable_('Announcements').filter(a=>a.published==='TRUE').slice(-5).reverse(), stats:{assignments:assignments.length, submitted:assignments.filter(a=>a.mySubmission).length, graded:assignments.filter(a=>a.mySubmission && a.mySubmission.status==='graded').length}};
}

function classroom_(user, classId) {
  const cls = readTable_('Classes').find(c => c.id === classId && c.active === 'TRUE');
  if (!cls) throw new Error('Class not found.');
  if (user.role === 'student' && user.classId !== classId) throw new Error('You are not enrolled in this class.');
  const lessons = readTable_('Lessons').filter(l => l.classId === classId && (l.published === 'TRUE' || user.role !== 'student')).reverse();
  const assignments = readTable_('Assignments').filter(a => a.classId === classId && (a.published === 'TRUE' || user.role !== 'student')).reverse().map(a => {
    const sub = readTable_('Submissions').find(s => s.assignmentId === a.id && s.studentId === user.id);
    return Object.assign({}, a, {mySubmission:sub ? publicSubmission_(sub) : null});
  });
  return {ok:true, classInfo:cls, lessons, assignments};
}

function createLesson_(user,p) {
  const classId = validClass_(p.classId);
  const title = requiredText_(p.title,'Lesson title',120);
  const content = requiredText_(p.content,'Lesson content',12000);
  const resourceUrl = safeUrl_(p.resourceUrl || '');
  const published = p.published === false ? 'FALSE' : 'TRUE';
  const row = {id:id_(),classId,title,content,resourceUrl,published,createdBy:user.username,createdAt:now_()};
  appendRow_('Lessons',row); return {ok:true, lesson:row};
}
function createAssignment_(user,p) {
  const classId = validClass_(p.classId);
  const title = requiredText_(p.title,'Assignment title',120);
  const instructions = requiredText_(p.instructions,'Instructions',12000);
  const points = Math.max(1, Math.min(1000, Number(p.points) || 100));
  let dueDate = String(p.dueDate || '').trim();
  if (dueDate && !/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) throw new Error('Use a valid due date.');
  const resourceUrl = safeUrl_(p.resourceUrl || '');
  const row = {id:id_(),classId,title,instructions,points,dueDate,resourceUrl,published:p.published===false?'FALSE':'TRUE',createdBy:user.username,createdAt:now_()};
  appendRow_('Assignments',row); return {ok:true, assignment:row};
}
function submitAssignment_(user,p) {
  if (user.role !== 'student') throw new Error('Only student accounts can submit assignments.');
  const assignment = readTable_('Assignments').find(a => a.id === String(p.assignmentId || '') && a.published === 'TRUE');
  if (!assignment || assignment.classId !== user.classId) throw new Error('Assignment not found in your class.');
  const content = String(p.content || '').trim().slice(0,12000);
  const projectUrl = safeUrl_(p.projectUrl || '');
  if (!content && !projectUrl) throw new Error('Add a written response or a project link.');
  const existing = readTable_('Submissions').find(s => s.assignmentId === assignment.id && s.studentId === user.id);
  const row = {id:existing ? existing.id : id_(), assignmentId:assignment.id, studentId:user.id, studentName:user.displayName, content, projectUrl, submittedAt:now_(), status:'submitted', grade:'', feedback:'', gradedBy:'', gradedAt:''};
  if (existing) updateRow_('Submissions', 'id', existing.id, row); else appendRow_('Submissions',row);
  return {ok:true, submission:publicSubmission_(row)};
}
function listSubmissions_(assignmentId) {
  const all = readTable_('Submissions').filter(s => !assignmentId || s.assignmentId === assignmentId);
  const assignments = readTable_('Assignments');
  return {ok:true, submissions:all.map(s => Object.assign({},publicSubmission_(s),{assignmentTitle:(assignments.find(a=>a.id===s.assignmentId)||{}).title||'',classId:(assignments.find(a=>a.id===s.assignmentId)||{}).classId||''})).reverse()};
}
function gradeSubmission_(user,p) {
  const id = String(p.submissionId || '');
  const submission = readTable_('Submissions').find(s=>s.id===id);
  if (!submission) throw new Error('Submission not found.');
  const assignment = readTable_('Assignments').find(a=>a.id===submission.assignmentId);
  const grade = Number(p.grade);
  if (!Number.isFinite(grade) || grade < 0 || grade > Number(assignment.points)) throw new Error('Grade must be between 0 and ' + assignment.points + '.');
  const feedback = String(p.feedback || '').trim().slice(0,4000);
  const updated = Object.assign({},submission,{grade:String(grade),feedback,status:'graded',gradedBy:user.username,gradedAt:now_()});
  updateRow_('Submissions','id',id,updated); return {ok:true};
}
function createAnnouncement_(user,p) {
  const row = {id:id_(),title:requiredText_(p.title,'Announcement title',120),body:requiredText_(p.body,'Announcement body',4000),createdBy:user.username,createdAt:now_(),published:'TRUE'};
  appendRow_('Announcements',row); return {ok:true,announcement:row};
}
function setEnrollment_(username,classId) {
  const target = String(username||'').trim().toLowerCase();
  const cls = validClass_(classId);
  const user = readTable_('Users').find(u=>u.username.toLowerCase()===target && u.role==='student');
  if (!user) throw new Error('Student username not found.');
  updateRow_('Users','id',user.id,Object.assign({},user,{classId:cls}));
  return {ok:true};
}

function validClass_(id) { const cls=readTable_('Classes').find(c=>c.id===String(id||'') && c.active==='TRUE'); if(!cls) throw new Error('Choose a valid class.'); return cls.id; }
function requiredText_(v,label,max) { const s=String(v||'').trim(); if(!s) throw new Error(label+' is required.'); if(s.length>max) throw new Error(label+' is too long.'); return s; }
function safeUrl_(v) { const s=String(v||'').trim(); if(!s) return ''; if(s.length>1000 || !/^https:\/\//i.test(s)) throw new Error('Links must start with https://'); return s; }
function publicSubmission_(s) { return {id:s.id,assignmentId:s.assignmentId,studentId:s.studentId,studentName:s.studentName,content:s.content,projectUrl:s.projectUrl,submittedAt:s.submittedAt,status:s.status,grade:s.grade,feedback:s.feedback,gradedBy:s.gradedBy,gradedAt:s.gradedAt}; }
function ensureSetup_() { if(!PropertiesService.getScriptProperties().getProperty('CLASSROOM_SPREADSHEET_ID')) throw new Error('Classroom has not been initialized. Run setupClassroom() in the Apps Script editor.'); }
function db_() { ensureSetup_(); return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('CLASSROOM_SPREADSHEET_ID')); }
function readTable_(name) {
  const sh=db_().getSheetByName(name); if(!sh) throw new Error('Missing database table: '+name);
  const values=sh.getDataRange().getValues(); if(values.length<2) return [];
  const headers=values[0].map(String);
  return values.slice(1).filter(r=>r.some(v=>v!=='' && v!==null)).map(r=>{const o={}; headers.forEach((h,i)=>o[h]=r[i] instanceof Date ? Utilities.formatDate(r[i],Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss') : String(r[i]===null?'':r[i])); return o;});
}
function appendRow_(name,obj) { const sh=db_().getSheetByName(name); const headers=TABLES[name]; sh.appendRow(headers.map(h=>obj[h]===undefined?'':obj[h])); }
function updateRow_(name,key,value,obj) { const sh=db_().getSheetByName(name); const data=sh.getDataRange().getValues(); const headers=data[0].map(String); const keyCol=headers.indexOf(key); if(keyCol<0) throw new Error('Invalid table key.'); for(let i=1;i<data.length;i++){if(String(data[i][keyCol])===String(value)){sh.getRange(i+1,1,1,headers.length).setValues([headers.map(h=>obj[h]===undefined?'':obj[h])]);return;}} throw new Error('Record not found.'); }
function hashPassword_(password,salt) { const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,salt+'|'+String(password),Utilities.Charset.UTF_8); return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join(''); }
function id_() { return Utilities.getUuid(); }
function now_() { return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'); }
function safeError_(err) { const s=String(err && err.message ? err.message : err); return s.length>240 ? s.slice(0,240) : s; }
