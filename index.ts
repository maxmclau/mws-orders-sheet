const MENU_TITLE = 'MWS';
const CONFIG_PROPERTY = 'configuration';

interface ReportConfiguration {
  seller_id: string,
  auth_token: string,
  frequency: number
}

var config:ReportConfiguration = {
  seller_id: undefined,
  auth_token: undefined,
  frequency: 5
}

function onOpen(e: MessageEvent) {
  SpreadsheetApp.getUi()
    .createMenu(MENU_TITLE)
    .addItem('Configure Report', 'showConfiguration')
    .addItem('Update Report', 'fetchOrders')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Utilities')
      .addItem('Reset Configuration', 'resetProperties')
      .addItem('Clear Triggers', 'clearTriggers'))
    .addToUi()
}

function showConfiguration() {
  var html = HtmlService.createHtmlOutputFromFile('modal')
      .setWidth(240)
      .setHeight(350)

  SpreadsheetApp.getUi()
      .showModalDialog(html, `${MENU_TITLE} Configuration`)
}

function getConfiguration():ReportConfiguration {
  const property = PropertiesService.getDocumentProperties().getProperty(CONFIG_PROPERTY)

  if(property) {
    config = JSON.parse(property) as ReportConfiguration
  }

  return config
}

function setConfiguration(updated: ReportConfiguration):Boolean {
  try{
    PropertiesService.getDocumentProperties().setProperty(CONFIG_PROPERTY, JSON.stringify(updated))
    SpreadsheetApp.getActiveSpreadsheet().toast(`Script configuration saved.`)

    return false
  } catch(err) {
    console.log(err)

    SpreadsheetApp.getActiveSpreadsheet().toast(`Error saving configuration.`)

    return true
  }
}


function resetProperties() {
  PropertiesService.getDocumentProperties().deleteAllProperties()

  SpreadsheetApp.getActiveSpreadsheet().toast(`Credentials for script cleared.`)
}

function clearTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Deleted all script triggers.`)
}

function promptForSecret(secret: Secret) {
  const ui = SpreadsheetApp.getUi()

  const result = ui.prompt(
    `Set Developer Credential`,
    `Please set your ${secret.name}`,
    ui.ButtonSet.OK_CANCEL)

  const button = result.getSelectedButton()
  const response = result.getResponseText()

  switch(button) {
    case ui.Button.OK:
      const i = secrets.findIndex(item => item.id == secret.id);
      secrets[i].value = response;

      const properties = JSON.stringify(secrets);

      PropertiesService.getDocumentProperties().setProperty(SECRETS_PROPERTY, properties);
      SpreadsheetApp.getActiveSpreadsheet().toast(`Credential for ${secret.name} modified.`)
      break
    case ui.Button.CANCEL:
    case ui.Button.CLOSE:
    default:
      SpreadsheetApp.getActiveSpreadsheet().toast(`Credentials for ${secret.name} unchanged.`)
      break
  }
}

function setTrigger() {
  const ui = SpreadsheetApp.getUi()

  const result = ui.prompt(
    `Trigger Interval`,
    `Set interval between fetches. (minutes)`,
    ui.ButtonSet.OK_CANCEL)

  const button = result.getSelectedButton()
  const response = result.getResponseText()

  switch(button) {
    case ui.Button.OK:
      ScriptApp.newTrigger('fetchOrders')
        .timeBased()
        .everyMinutes(Number(response))
        .create();

      SpreadsheetApp.getActiveSpreadsheet().toast(`Set new trigger interal`)
      break
    case ui.Button.CANCEL:
    case ui.Button.CLOSE:
    default:
      SpreadsheetApp.getActiveSpreadsheet().toast(`Trigger unchanged`)
      break
  }
}

function fetchOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  sheet.appendRow(['Hello World']);
}