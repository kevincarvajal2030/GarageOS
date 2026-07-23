/**
 * ============================================================
 * GARAGE OS
 * DriveService.gs
 *
 * Central service for Google Drive operations.
 * ============================================================
 */

const DriveService = (() => {

  /**
   * Returns the Vehicle Images folder ID.
   */
  function getVehicleImagesFolderId() {

    return PropertiesService
      .getScriptProperties()
      .getProperty("VEHICLE_IMAGES_FOLDER_ID");

  }

  /**
   * Returns the Vehicle Images folder.
   */
  function getVehicleImagesFolder() {

    return DriveApp.getFolderById(
      getVehicleImagesFolderId()
    );

  }

  /**
   * Finds an image by filename.
   *
   * Returns:
   * Drive File
   * or
   * null
   */
  function findImage(imageName) {

    const filename =
      `${imageName}.${DRIVE.IMAGE_EXTENSION}`;

    const files =
      getVehicleImagesFolder()
        .getFilesByName(filename);

    if (!files.hasNext()) {

      return null;

    }

    return files.next();

  }

  return Object.freeze({

    getVehicleImagesFolderId,
    getVehicleImagesFolder,
    findImage

  });

})();


function testFindImage() {

  const file = DriveService.findImage(
    "gmc-sierra-2020-black"
  );

  if (!file) {

    Logger.log("Image not found");
    return;

  }

  Logger.log(file.getName());
  Logger.log(file.getId());

}