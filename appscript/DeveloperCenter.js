//DeveloperCenter.gs

/**
 * Returns GitHub configuration stored in Script Properties.
 *
 * @returns {Object}
 */
function developerCenterGetConfiguration() {

  const props = PropertiesService.getScriptProperties();

  return {

    bridge: props.getProperty("BRIDGE_URL"),

    owner: props.getProperty("GITHUB_OWNER"),

    repository: props.getProperty("GITHUB_REPOSITORY"),

    branch: props.getProperty("GITHUB_BRANCH")

  };

}


function openDeveloperCenter() {

  SpreadsheetApp.getUi()

    .createMenu("Developer Center")

    .addItem("Push Project", "dcPushProject")

    .addSeparator()

    .addItem("Test Bridge", "dcTestBridge")

    .addToUi();

}


function dcRepositoryStatus() {

  SpreadsheetApp.getUi().alert(
    "GarageOS\n\nRepository Status\n\nComing Soon"
  );

}


function dcPullProject() {

  SpreadsheetApp.getUi().alert(
    "Pull Project\n\nComing Soon"
  );

}


function dcPushProject() {

  const config = developerCenterGetConfiguration();

  const response = UrlFetchApp.fetch(
    config.bridge + "/github/push",
    {
      method: "post",
      muteHttpExceptions: true
    }
  );

  SpreadsheetApp.getUi().alert(
    response.getContentText()
  );

}


function dcTestBridge() {

  const config = developerCenterGetConfiguration();

  const response = UrlFetchApp.fetch(config.bridge);

  Logger.log(response.getContentText());

}


function dcDiagnostics() {

  SpreadsheetApp.getUi().alert(
    "Diagnostics\n\nComing Soon"
  );

}


/**
 * Tests Developer Center configuration.
 */
function developerCenterTestConfiguration() {

  const config = developerCenterGetConfiguration();

  Logger.log(JSON.stringify(config, null, 2));

}



