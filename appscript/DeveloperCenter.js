//DeveloperCenter.gs

/**
 * ============================================================
 * GarageOS Developer Center
 * Version : 0.1.0
 * Module  : GitHub Connector
 * ============================================================
 */

/**
 * Returns GitHub configuration stored in Script Properties.
 *
 * @returns {Object}
 */
function developerCenterGetConfiguration() {

  const props = PropertiesService.getScriptProperties();

  return {
    owner: props.getProperty("GITHUB_OWNER"),
    repository: props.getProperty("GITHUB_REPOSITORY"),
    branch: props.getProperty("GITHUB_BRANCH"),
    token: props.getProperty("GITHUB_TOKEN")
  };

}


/**
 * Tests Developer Center configuration.
 */
function developerCenterTestConfiguration() {

  const config = developerCenterGetConfiguration();

  Logger.log(JSON.stringify(config, null, 2));

}


/**
 * Returns GitHub request headers.
 *
 * @returns {Object}
 */
function developerCenterGetGitHubHeaders() {

  const config = developerCenterGetConfiguration();

  return {
    "Authorization": `Bearer ${config.token}`,
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28"
  };

}


/**
 * Tests GitHub connection.
 */
function developerCenterTestGitHubConnection() {

  const response = UrlFetchApp.fetch(
    "https://api.github.com/user",
    {
      method: "get",
      headers: developerCenterGetGitHubHeaders(),
      muteHttpExceptions: true
    }
  );

  Logger.log(response.getResponseCode());
  Logger.log(response.getContentText());

}


/**
 * ============================================================
 * GitHub API
 * Version : 0.1.0
 * ============================================================
 */

/**
 * Executes a request against the GitHub REST API.
 *
 * @param {string} endpoint API endpoint beginning with "/"
 * @param {string} method HTTP method
 * @param {Object=} payload Optional payload
 * @returns {Object}
 */
function githubApiExecuteRequest(endpoint, method, payload = null) {

  const config = developerCenterGetConfiguration();

  const options = {
    method: method,
    headers: {
      Authorization: `Bearer ${config.token}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
    options.contentType = "application/json";
  }

  const response = UrlFetchApp.fetch(
    `https://api.github.com${endpoint}`,
    options
  );

  return {
    code: response.getResponseCode(),
    body: JSON.parse(response.getContentText())
  };

}



/**
 * Returns repository information.
 */
function githubRepositoryGetInformation() {

  const config = developerCenterGetConfiguration();

  return ghRequest(
    `/repos/${config.owner}/${config.repository}`,
    "get"
  );

}



function githubRepositoryTestInformation() {

  const repo = githubRepositoryGetInformation();

  Logger.log(JSON.stringify(repo, null, 2));

}



/**
 * Returns the contents of a repository path.
 *
 * @param {string} repositoryPath
 * @returns {Object}
 */
function githubRepositoryGetContents(repositoryPath) {

  const config = developerCenterGetConfiguration();

  return githubApiExecuteRequest(
    `/repos/${config.owner}/${config.repository}/contents/${repositoryPath}?ref=${config.branch}`,
    "get"
  );

}


/**
 * Tests reading the repository appscript folder.
 */
function githubRepositoryTestContents() {

  const response =
    githubRepositoryGetContents("appscript");

  Logger.log(JSON.stringify(response, null, 2));

}
