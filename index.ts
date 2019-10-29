const MENU_TITLE = 'Configuration';
const SECRETS_PROPERTY = 'secrets';

interface Secret {
  readonly id: number,
  readonly name: string,
  readonly handler: string,
  value: string
}

enum Credentials {
  SellerID,
  AuthToken
}

const secrets:Array<Secret> = [
  {
    id: Credentials.SellerID,
    name: 'Seller ID',
    handler: 'setSellerId',
    value: null
  },
  {
    id: Credentials.AuthToken,
    name: 'Auth Token',
    handler: 'setAuthToken',
    value: null
  }
]

function onOpen(e: MessageEvent) {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu(MENU_TITLE)

  secrets.forEach((secret: Secret) => {
    menu.addItem(secret.name, secret.handler)
  })

  menu.addSeparator()
    .addSubMenu(ui
      .createMenu('Triggers')
      .addItem('Set Interval', 'setTrigger')
      .addItem('Reset Triggers', 'resetTriggers')
    )

  menu.addSeparator()
    .addItem('Reset Credentials', 'resetProperties')

  menu.addSeparator()
    .addItem('Force Update', 'fetchOrders')
    .addToUi()
}

function getSecrets():Array<Secret> {
  const properties = PropertiesService.getDocumentProperties().getProperty(SECRETS_PROPERTY)
  const secrets = JSON.parse(properties)

  return secrets
}

function setSellerId() {
  const i = secrets.findIndex(item => item.id == Credentials.SellerID);

  promptForSecret(secrets[i])
}

function setAuthToken() {
  const i = secrets.findIndex(item => item.id == Credentials.AuthToken);

  promptForSecret(secrets[i])
}

function resetProperties() {
  PropertiesService.getDocumentProperties().deleteAllProperties()

  SpreadsheetApp.getActiveSpreadsheet().toast(`Credentials for script cleared.`)
}

function resetTriggers() {
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
  /*var query = '"Apps Script" stars:">=100"';
  var url = 'https://api.github.com/search/repositories'
    + '?sort=stars'
    + '&q=' + encodeURIComponent(query);

  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  Logger.log(response);
  */

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  sheet.appendRow(['Hello World']);
}