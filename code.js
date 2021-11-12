// customer ID は https://support.google.com/a/answer/10070793 の指示にしたがって手動で調査
const customerId = 'xxxxxxxx';

const bookId = 'xxxxxxxx';

const domainName = 'example.jp';

const calendarAdministratorGroup = '';

const headerRow = 1;

const logNotifyAddress = calendarAdministratorGroup;

let logMessages = [];

// main
function updateResourcesAndBuildings(){
  logMessages = [];
  updateBuildings();
  updateResources();
  deleteResources();
  if (logMessages.length > 0){
    if (logNotifyAddress) {
      GmailApp.sendEmail(logNotifyAddress, `Resource Calendar Management Result for ${domainName}`, logMessages.join("\n"));
    }
  }
}

function pushLogMessage(msg){
  logMessages.push(msg);
  Logger.log('%s', msg);
}

function updateBuildings(){
  const book = openBook();
  const sheet = book.getSheetByName('Buildings');
  const dict = getHeader(sheet, headerRow);
  const last = sheet.getLastRow();
  for (let row=headerRow + 1; row<=last; row++){
    let id = sheet.getRange(row, dict['Building Id']).getValue().toString();
    let name = sheet.getRange(row, dict['Building Name']).getValue().toString();
    let floors = sheet.getRange(row, dict['Floors']).getValue().toString().split(',');
    updateBuilding(row, id, name, floors);
  }
}

function openBook(){
  if (! openBook.cache) {
    openBook.cache = SpreadsheetApp.openById(bookId);
  }
  return openBook.cache;
}

function getHeader(sheet, row){
  let dict = {};
  const header = sheet.getRange(row, 1, 1, sheet.getMaxColumns()).getValues().shift();
  for (let i=0; i<=header.length; i++){
    dict[header[i]] = i+1;
  }
  return dict;
}

function updateBuilding(row, buildingId, buildingName, floorNames){
  if(! buildingId){
    pushLogMessage(`建物IDは省略できません: row=${row}`)
    return;
  }
  if(! buildingName){
    pushLogMessage(`建物名は省略できません: row=${row}, buildingId=${buildingId}`);
    return;
  }
  if(! floorNames){
    pushLogMessage(`階数の指定は省略できません: row=${row}, buildingId=${buildingId}, buildingName=${buildingName}`)
    return;
  }
  let building;
  try {
    building = AdminDirectory.Resources.Buildings.get(customerId, buildingId);
  } catch(e) {
    // Do nothing.
  }
  if(building){
    let change = 0;
    if (building.buildingName != buildingName) {
      pushLogMessage(`建物名が更新されました: buildingId=${buildingId}, old=${building.buildingName}, new=${buildingName}`);
      building.buildingName = buildingName;
      change++;
    }
    if (building.floorNames.toString() != floorNames.toString()) {
      pushLogMessage(`階数が更新されました: buildingId=${buildingId}, buildingName=${buildingName}, old=${building.floorNames}, new=${floorNames}`);
      building.floorNames = floorNames;
      change++;
    }
    if (change > 0) {
      AdminDirectory.Resources.Buildings.update(building, customerId, buildingId);
    }
  } else {
    building = {
      buildingId: buildingId,
      buildingName: buildingName,
      floorNames: floorNames
    }
    AdminDirectory.Resources.Buildings.insert(building, customerId);
  }
}

function updateResources(){
  const book = openBook();
  const sheet = book.getSheetByName('Resources');
  const dict = getHeader(sheet, headerRow);
  const last = sheet.getLastRow();
  for (let row=headerRow + 1; row<=last; row++){
    let id = normalizeResourceId(sheet.getRange(row, dict['Resource Id']).getValue().toString());
    //let owner = sheet.getRange(row, dict['Resource Owner']).getValue().toString();
    let group = getResourceUserGroup(sheet.getRange(row, dict['Resource User']).getValue().toString());
    let name = sheet.getRange(row, dict['Resource Name']).getValue().toString();
    let building = getBuildingId(sheet.getRange(row, dict['Building Name']).getValue().toString());
    let category = getResourceCategory(sheet.getRange(row, dict['Resource Category']).getValue().toString());
    let type = sheet.getRange(row, dict['Resource Type']).getValue().toString();
    let floor = sheet.getRange(row, dict['Floor Name']).getValue().toString();
    let capacity = Number(sheet.getRange(row, dict['Capacity']).getValue().toString());
    let room = sheet.getRange(row, dict['Floor Section']).getValue().toString();
    let resource = updateResource(row, id, group, name, building, category, type, floor, capacity, room);
    if (resource) {
      sheet.getRange(row, dict['Resource Id']).setValue(resource.resourceId);
      sheet.getRange(row, dict['Resource Calendar']).setValue('https://calendar.google.com/calendar/embed?src='+resource.resourceEmail);
    }
  }
}

function updateResource(row, resourceId, resourceGroup, resourceName, buildingId, resourceCategory, resourceType, floorName, capacity, floorSection){
  if(! resourceName){
    pushLogMessage(`リソース名は省略できません: row=${row}`);
    return;
  }
  if(! resourceCategory){
    pushLogMessage(`リソース種別は省略できません: row=${row}, resourceName=${resourceName}`);
    return;
  }
  if(! buildingId){
    pushLogMessage(`建物名は省略できません: row=${row}, resourceName=${resourceName}`);
    return;
  }
  if(floorName && !validateFloorName(buildingId, floorName)){
    pushLogMessage(`指定された階は存在しません: row=${row}, resourceName=${resourceName}, floorName=${floorName}`);
    return;
  }
  if(resourceCategory == 'CONFERENCE_ROOM' && !capacity){
    pushLogMessage(`会議室の定員は省略できません: row=${row}, resourceName=${resourceName}`);
    return;
  }
  let resource = false;
  try {
    resource = AdminDirectory.Resources.Calendars.get(customerId, resourceId);
  } catch(e) {
    // Do nothing.
  }
  if (resource) {
    let change = 0;
    if (resource.resourceName != resourceName) {
      pushLogMessage(`リソース名が更新されました: resourceId=${resourceId}, old=${resource.resourceName}, new=${resourceName}`);
      resource.resourceName = resourceName;
      change++;
    }
    if (resource.buildingId != buildingId) {
      pushLogMessage(`リソースの建物IDが更新されました: resourceId=${resourceId}, resourceName=${resourceName}, old=${resource.buildingId}, new=${buildingId}`);
      resource.buildingId = buildingId;
      change++;
    }
    if (resource.resourceCategory != resourceCategory) {
      pushLogMessage(`リソースの分類が更新されました: resourceId=${resourceId}, resourceName=${resourceName}, old=${resource.resourceCategory}, new=${resourceCategory}`);
      resource.resourceCategory = resourceCategory;
      change++;
    }
    if (! safeStringEqual(resource.resourceType, resourceType)) {
      pushLogMessage(`リソースの種別が更新されました: resourceId=${resourceId}, resourceName=${resourceName}, old=${resource.resourceType}, new=${resourceType}`);
      resource.resourceType = resourceType;
      change++;
    }
    if (! safeStringEqual(resource.floorName, floorName)) {
      pushLogMessage(`リソースの建物階が更新されました: resourceId=${resourceId}, resourceName=${resourceName}, old=${resource.floorName}, new=${floorName}`);
      resource.floorName = floorName;
      change++;
    }
    if (! safeStringEqual(resource.floorSection, floorSection)) {
      pushLogMessage(`リソースの建物区画が更新されました: resourceId=${resourceId}, resourceName=${resourceName}, old=${resource.floorSection}, new=${floorSection}`);
      resource.floorSection = floorSection;
      change++;
    }    
    if (! safeNumberEqual(resource.capacity, capacity)) {
      pushLogMessage(`リソースの定員が更新されました: resourceId=${resourceId}, resourceName=${resourceName}, old=${resource.capacity}, new=${capacity}`);
      resource.capacity = capacity;
      change++;
    }
    if (change > 0) {
      AdminDirectory.Resources.Calendars.update(resource, customerId, resourceId);
    }
  } else {
    const newId = createResourceId();
    resource = AdminDirectory.Resources.Calendars.insert({
      resourceId: newId,
      resourceName: resourceName,
      buildingId: buildingId,
      resourceCategory: resourceCategory,
      resourceType: resourceType,
      floorName: floorName,
      floorSection: floorSection,
      capacity: capacity
    }, customerId);
    pushLogMessage(`リソースを作成しました: resourceId=${newId}, resourceName=${resourceName}, calendarId=${resource.resourceEmail}, resourceGroup=${resourceGroup}`);
  }
  // リソースが作成されるまで待機
  let cal = false;
  do {
    try {
      cal = Calendar.Calendars.get(resource.resourceEmail);
    } catch(e){
      Utilities.sleep(100);
    }
  } while(! cal);
  cal.setTimeZone("Asia/Tokyo");
  updateResourceCalendarACL(resource, resourceGroup);
  //showResourceCalendarACL(resource);
  return resource;
}

// undefined を空文字列に変換してから厳密な等価比較
function safeStringEqual(x, y) {
  if (x === void 0) {
    x = '';
  }
  if (y === void 0) {
    y = '';
  }
  return x === y;
}

// undefined を0に変換してから厳密な等価比較
function safeNumberEqual(x, y) {
  if (x === void 0) {
    x = 0;
  }
  if (y === void 0) {
    y = 0;
  }
  return x === y;
}

function deleteResources(){
  let resources = listAllResources();
  resources.forEach(function(resource){
    let id = lookupSheet(Number(resource.resourceId), 'Resources', 'Resource Id', 'Resource Id');
    if(! id){
      AdminDirectory.Resources.Calendars.remove(customerId, resource.resourceId);
      pushLogMessage(`リソースを削除しました: resourceId=${resource.resourceId}, resourceName=${resource.resourceName}, calendarId=${resource.resourceEmail}`);
    }
  });
}

function createResourceId() {
  do {
    let candidate = getRandomInt(1000000000, 99999999999);
    try {
      let res = AdminDirectory.Resources.Calendars.get(customerId, candidate);
    } catch(e) {
      return normalizeResourceId(String(candidate));
    }
  } while(true);
}

function getRandomInt(min, max) {
  min = Math.ceil(min);
  max = Math.floor(max);
  return Math.floor(Math.random() * (max - min) + min); //The maximum is exclusive and the minimum is inclusive
}

function normalizeResourceId(str) {
  if (str.length < 9) {
    str = ('000000000' + str).slice(-9);
  }
  return str;
}

function getBuildingId(name){
  return lookupSheet(name, 'Buildings', 'Building Name', 'Building Id');
}

function validateFloorName(buildingId, floorName){
  let possibleFloors = lookupSheet(buildingId, 'Buildings', 'Building Id', 'Floors').split(',');
  if (possibleFloors.indexOf(floorName) >= 0) {
    return true;
  }
  return false;
}

function getResourceCategory(name){
  return lookupSheet(name, 'ResourceCategories', '和名', '英名');
}

function getResourceUserGroup(name){
  return lookupSheet(name, 'ResourceUserGroups', 'Group Name', 'Group Address');
}

function lookupSheet(target, sheetName, keyName, valueName) {
  const book = openBook();
  const sheet = book.getSheetByName(sheetName);
  const header = sheet.getRange(headerRow, 1, 1, sheet.getMaxRows()).getValues().shift();
  const keyColumn = header.indexOf(keyName) + 1;
  const valueColumn = header.indexOf(valueName) + 1;
  if (keyColumn <= 0 || valueColumn <= 0) {
    return undefined;
  }
  const keys = sheet.getRange(headerRow + 1, keyColumn, sheet.getLastRow()).getValues();
  const row = keys.findIndex(function(elem){return target==elem[0];}) + headerRow + 1;
  if (row > headerRow) {
    return sheet.getRange(row, valueColumn).getValue().toString();
  }
  return undefined;
}

// https://developers.google.com/admin-sdk/directory/reference/rest/v1/resources.calendars/list
function listAllResources() {
  let pageToken;
  let resources = [];
  do {
    let page = AdminDirectory.Resources.Calendars.list(customerId, {
      maxResults: 100,
      pageToken: pageToken
    });
    let items = page.items;
    if (items) {
      for (var i = 0; i < items.length; i++) {
        resources.push(items[i]);
      }
    } else {
      Logger.log('No items found.');
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  return resources;
}

function showAllResources() {
  let resources = listAllResources();
  resources.forEach(function(resource){
    Logger.log('%s (%s, %s)', resource.resourceName, resource.resourceId, resource.resourceEmail);
    showResourceCalendarACL(resource);
  })
}

function resetAllResourceCalendarACL() {
  let resources = listAllResources();
  resources.forEach(function(resource){
    updateResourceCalendarACL(resource, false);
    showResourceCalendarACL(resource);
  })
}

function updateResourceCalendarACL(resource, resourceGroup) {
  const calendarId = resource.resourceEmail;
  let validAclIds = ['user:'+calendarId];
  validAclIds.push(setResourceCalndarACL(resource, 'default', '__public_principal__@public.calendar.google.com', 'none'));
  validAclIds.push(setResourceCalndarACL(resource, 'domain', domainName, 'none'));
  if (resourceGroup) {
    validAclIds.push(setResourceCalndarACL(resource, 'group', resourceGroup, 'writer'));
  }
  if (calendarAdministratorGroup) {
    validAclIds.push(setResourceCalndarACL(resource, 'group', calendarAdministratorGroup, 'owner'));
  }
  const acls = listResourceCalendarACL(resource);
  acls.forEach(function(acl){
    if (validAclIds.indexOf(acl.id) < 0) {
      Calendar.Acl.remove(calendarId, acl.id, {sendNotifications: false});
      pushLogMessage(`ACLが削除されました: resourceId=${resource.resourceId}, resourceName=${resource.resourceName}, scopeType=${acl.scope.type}, scopeValue=${acl.scope.value}`);
    }
  });
}

function setResourceCalndarACL(resource, scopeType, scopeValue, role) {
  const calendarId = resource.resourceEmail;
  let acl;
  try {
    acl = Calendar.Acl.get(calendarId, scopeType+':'+scopeValue);
  } catch(e) {
    // Do nothing.
  }
  if (acl) {
    if (acl.role != role) {
      Calendar.Acl.update(acl, calendarId, acl.id, {sendNotifications: false});
      pushLogMessage(`ACLが更新されました: resourceId=${resource.resourceId}, resourceName=${resource.resourceName}, scopeType=${scopeType}, scopeValue=${scopeValue}, old=${acl.role}, new=${role}`);
    }
  } else {
    if (role != 'none') {
      acl = Calendar.Acl.insert({
        scope: {
          type: scopeType,
          value: scopeValue
        }, 
        role: role
      }, calendarId, {sendNotifications: false});
      pushLogMessage(`ACLが作成されました: resourceId=${resource.resourceId}, resourceName=${resource.resourceName}, scopeType=${scopeType}, scopeValue=${scopeValue}, role=${role}`);
    }
  }
  if (acl) {
    return acl.id;
  } else {
    return null;
  }
}

function showResourceCalendarACL(resource) {
  const acls = listResourceCalendarACL(resource);
  acls.forEach(function(acl){
    Logger.log('scopeType=%s, scopeValue=%s, role=%s', acl.scope.type, acl.scope.value, acl.role);
  })
}

// https://developers.google.com/calendar/v3/reference/acl
function listResourceCalendarACL(resource) {
  const calendarId = resource.resourceEmail;
  let pageToken;
  let acls = [];
  do {
    let page = Calendar.Acl.list(calendarId, {
      maxResults: 100,
      pageToken: pageToken
    });
    let items = page.items;
    if (items) {
      for (var i = 0; i < items.length; i++) {
        acls.push(items[i]);
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  return acls;
}
