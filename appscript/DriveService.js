/**
 * GARAGE OS
 * DriveService.gs
 * Central service for Google Drive operations.
 */
const DriveService = (() => {

  function getVehicleImagesFolderId() {
    const folderId = PropertiesService
      .getScriptProperties()
      .getProperty("VEHICLE_IMAGES_FOLDER_ID");

    if (!folderId) {
      throw new Error("Missing Script Property: VEHICLE_IMAGES_FOLDER_ID");
    }

    return folderId;
  }

  function getVehicleImagesFolder() {
    return DriveApp.getFolderById(getVehicleImagesFolderId());
  }

  function buildImageFilename(imageName) {
    return `${imageName}.${DRIVE.IMAGE_EXTENSION}`;
  }

  function findImage(imageName) {
    const filename = buildImageFilename(imageName);
    const files = getVehicleImagesFolder().getFilesByName(filename);

    if (!files.hasNext()) return null;

    return files.next();
  }

  function saveImageBlob(imageName, blob) {
    const existingFile = findImage(imageName);

    if (existingFile) {
      return existingFile;
    }

    const filename = buildImageFilename(imageName);
    const folder = getVehicleImagesFolder();

    blob.setName(filename);

    const file = folder.createFile(blob);
    file.setName(filename);

    return file;
  }

  return Object.freeze({
    getVehicleImagesFolderId,
    getVehicleImagesFolder,
    findImage,
    saveImageBlob
  });

})();


function testVehicleImagesFolder() {
  const folder = DriveService.getVehicleImagesFolder();

  Logger.log(folder.getName());
  Logger.log(folder.getId());
  Logger.log(folder.getUrl());
}