// customer ID は https://support.google.com/a/answer/10070793 の指示にしたがって手動で調査
const customerId = 'xxxxxxxx';

const bookId = 'xxxxxxxx';

const domainName = 'example.jp';

const calendarAdministratorGroup = '';

const headerRow = 1;

// main
function updateResourcesAndBuildings(){
  const book = SpreadsheetApp.openById(bookId);
  updateBuildings(book);
  updateResources(book);
  //deleteResources();
}

function updateBuildings(book){
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
    Logger.log("%s行: 建物IDは省略できません", row);
    return;
  }
  if(! buildingName){
    Logger.log("%s行: 建物名は省略できません", row);
    return;
  }
  if(! floorNames){
    Logger.log("%s行: 階数の指定は省略できません", row);
    return;
  }
  let building;
  try {
    building = AdminDirectory.Resources.Buildings.get(customerId, buildingId);
  } catch(e) {
    // Do nothing.
  }
  if(building){
    building.buildingName = buildingName;
    building.floorNames = floorNames;
    AdminDirectory.Resources.Buildings.update(building, customerId, buildingId);
  } else {
    building = {
      buildingId: buildingId,
      buildingName: buildingName,
      floorNames: floorNames
    }
    AdminDirectory.Resources.Buildings.insert(building, customerId);
  }
}

function updateResources(book){
  const sheet = book.getSheetByName('Resources');
  const dict = getHeader(sheet, headerRow);
  const last = sheet.getLastRow();
  for (let row=headerRow + 1; row<=last; row++){
    let id = normalizeResourceId(sheet.getRange(row, dict['Resource Id']).getValue().toString());
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
    Logger.log("%d行: リソース名は省略できません", row);
    return;
  }
  if(! resourceCategory){
    Logger.log("%d行: リソース種別は省略できません", row);
    return;
  }
  if(! buildingId){
    Logger.log("%d行: 建物名は省略できません", row);
    return;
  }
  if(floorName && !validateFloorName(buildingId, floorName)){
    Logger.log("%d行: 指定された階は存在しません", row);
    return;
  }
  if(resourceCategory == 'CONFERENCE_ROOM' && !capacity){
    Logger.log("%d行: 会議室の収容人数は省略できません", row);
    return;
  }
  let resource = false;
  try {
    resource = AdminDirectory.Resources.Calendars.get(customerId, resourceId);
  } catch(e) {
    // Do nothing.
  }
  if (resource) {
    resource.resourceName = resourceName;
    resource.buildingId = buildingId;
    resource.resourceCategory = resourceCategory;
    resource.resourceType = resourceType;
    resource.floorName = floorName;
    resource.floorSection = floorSection;
    resource.capacity = capacity;
    AdminDirectory.Resources.Calendars.update(resource, customerId, resourceId);
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
  Logger.log('Resource updated: resourceName=%s, resourceId=%s, calendarId=%s, calendarTimezone=%s, resourceGroup=%s', resource.resourceName, resource.resourceId, resource.resourceEmail, cal.getTimeZone(), resourceGroup);
  resetResourceCalendarACL(resource.resourceEmail, resourceGroup);
  showResourceCalendarACL(resource.resourceEmail);
  return resource;
}

function deleteResources(){
  let resources = listAllResources();
  resources.forEach(function(resource){
    let id = lookupSheet(Number(resource.resourceId), 'Resources', 'Resource Id', 'Resource Id');
    if(! id){
      AdminDirectory.Resources.Calendars.remove(customerId, resource.resourceId);
      Logger.log('Resource removed: resourceName=%s, resourceId=%s, calendarId=%s', resource.resourceName, resource.resourceId, resource.resourceEmail);
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
  const book = SpreadsheetApp.openById(bookId);
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
    showResourceCalendarACL(resource.resourceEmail);
  })
}

function resetAllResourceCalendarACL() {
  let resources = listAllResources();
  resources.forEach(function(resource){
    Logger.log('%s (%s, %s)', resource.resourceName, resource.resourceId, resource.resourceEmail);
    resetResourceCalendarACL(resource.resourceEmail, false);
    showResourceCalendarACL(resource.resourceEmail);
  })
}

function resetResourceCalendarACL(calendarId, resourceGroup) {
  setResourceCalndarACL(calendarId, 'default', '__public_principal__@public.calendar.google.com', 'none');
  setResourceCalndarACL(calendarId, 'domain', domainName, 'none');
  if (resourceGroup) {
    setResourceCalndarACL(calendarId, 'group', resourceGroup, 'writer');
  }
  if (calendarAdministratorGroup) {
    setResourceCalndarACL(calendarId, 'group', calendarAdministratorGroup, 'owner');
  }
}

function setResourceCalndarACL(calendarId, scopeType, scopeValue, role) {
  let acl;
  try {
    acl = Calendar.Acl.get(calendarId, scopeType+':'+scopeValue);
  } catch(e) {
    // Do nothing.
  }
  if (acl) {
    acl.role = role;
    Calendar.Acl.update(acl, calendarId, acl.id, {sendNotifications: false});
  } else {
    if (role != 'none') {
      Calendar.Acl.insert({
        scope: {
          type: scopeType,
          value: scopeValue
        }, 
        role: role
      }, calendarId, {sendNotifications: false});
    }
  }
}

function showResourceCalendarACL(calendarId) {
  let acls = listResourceCalendarACL(calendarId);
  acls.forEach(function(acl){
    Logger.log('scopeType=%s, scopeValue=%s, role=%s', acl.scope.type, acl.scope.value, acl.role);
  })
}

// https://developers.google.com/calendar/v3/reference/acl
function listResourceCalendarACL(calendarId) {
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
