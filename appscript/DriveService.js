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

    return files.hasNext() ? files.next() : null;
  }

  function getFileById(fileId) {
    try {
      if (!fileId) return null;
      return DriveApp.getFileById(fileId);
    } catch (e) {
      return null;
    }
  }

  function saveImageBlob(imageName, blob) {
    const existingFile = findImage(imageName);

    if (existingFile) return existingFile;

    const filename = buildImageFilename(imageName);
    const folder = getVehicleImagesFolder();

    blob.setName(filename);

    const file = folder.createFile(blob);
    file.setName(filename);

    return file;
  }

  function trashFile(fileId) {
    const file = getFileById(fileId);

    if (!file) return false;

    file.setTrashed(true);
    return true;
  }

  return Object.freeze({
    getVehicleImagesFolderId,
    getVehicleImagesFolder,
    buildImageFilename,
    findImage,
    getFileById,
    saveImageBlob,
    trashFile
  });

})();